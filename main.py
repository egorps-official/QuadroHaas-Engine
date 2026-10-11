import json, os, queue, sys, threading, time, tkinter as tk
from tkinter import ttk, messagebox
import numpy as np, sounddevice as sd, pystray, pygame
from PIL import Image, ImageDraw
import pythoncom
from win32com.client import Dispatch

SAMPLE_RATE = 44100
MAX_DELAY_SAMPLES = 44100
ICON_PATH = "icon.ico"
TICK_PATH = "tick.wav"
OST_PATH = "coconutmall.mp3"
ENGINE_VER = "v1.2"
NON_COMPATIBLE_VERS = ["v1.0", "v1.1", None]
ECHO_DELAY_MS = 200
ECHO_FEEDBACK = 0.4
TICK_INTERVAL = 0.5

PRESET_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "QuadroHaas")
PRESET_PATH = os.path.join(PRESET_DIR, "preset.json")
STARTUP_DIR = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
STARTUP_LNK = os.path.join(STARTUP_DIR, "QuadroHaas.lnk")

audio_input_stream = None
stream_dev1 = None
stream_dev2 = None
is_engine_active = False

reverb_amount = 0.0
echo_amount = 0.0
device1_delay_ms = 0.0
device2_delay_ms = 0.0
device1_offset_ms = 0.0
device2_offset_ms = 0.0

queue_dev1 = queue.Queue(maxsize=4)
queue_dev2 = queue.Queue(maxsize=4)

buffer_device1 = np.zeros((MAX_DELAY_SAMPLES, 2), dtype=np.float32)
buffer_device2 = np.zeros((MAX_DELAY_SAMPLES, 2), dtype=np.float32)
buffer_reverb = np.zeros((MAX_DELAY_SAMPLES, 2), dtype=np.float32)
buffer_echo = np.zeros((MAX_DELAY_SAMPLES, 2), dtype=np.float32)

pointer_device1 = 0
pointer_device2 = 0
pointer_reverb = 0
pointer_echo = 0

output_device_ids = []
output_device_names = []

tray_icon = None
tray_message_shown = False
autorun_state = False
autorun_changed = False

tick_thread = None
tick_stop = threading.Event()

def clamp(v, lo, hi):
    return max(lo, min(hi, v))

def resource_path(name):
    base = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, name)

def audio_input_callback(indata, frames, time_info, status):
    global pointer_device1, pointer_device2, pointer_reverb, pointer_echo
    if not is_engine_active:
        return
    processed_input = indata.copy()

    reverb_delay_samples = int(0.08 * SAMPLE_RATE)
    indices_rev = (
        np.arange(pointer_reverb - reverb_delay_samples, pointer_reverb - reverb_delay_samples + frames)
        % MAX_DELAY_SAMPLES
    )
    delayed_signal = buffer_reverb[indices_rev]
    processed_input += delayed_signal * reverb_amount
    write_indices_rev = np.arange(pointer_reverb, pointer_reverb + frames) % MAX_DELAY_SAMPLES
    buffer_reverb[write_indices_rev] = processed_input + delayed_signal * (reverb_amount * 0.5)
    pointer_reverb = (pointer_reverb + frames) % MAX_DELAY_SAMPLES

    echo_delay_samples = int((ECHO_DELAY_MS / 1000.0) * SAMPLE_RATE)
    indices_echo = (
        np.arange(pointer_echo - echo_delay_samples, pointer_echo - echo_delay_samples + frames)
        % MAX_DELAY_SAMPLES
    )
    echo_delayed = buffer_echo[indices_echo]
    processed_input += echo_delayed * echo_amount
    write_indices_echo = np.arange(pointer_echo, pointer_echo + frames) % MAX_DELAY_SAMPLES
    buffer_echo[write_indices_echo] = processed_input + echo_delayed * (ECHO_FEEDBACK * echo_amount)
    pointer_echo = (pointer_echo + frames) % MAX_DELAY_SAMPLES

    total_d1 = (-device2_offset_ms) + device1_delay_ms
    delay1_samples = clamp(int((total_d1 / 1000.0) * SAMPLE_RATE), 0, MAX_DELAY_SAMPLES - 1)
    write_indices_d1 = np.arange(pointer_device1, pointer_device1 + frames) % MAX_DELAY_SAMPLES
    buffer_device1[write_indices_d1] = processed_input
    read_indices_d1 = (write_indices_d1 - delay1_samples) % MAX_DELAY_SAMPLES
    output_device1 = buffer_device1[read_indices_d1]
    pointer_device1 = (pointer_device1 + frames) % MAX_DELAY_SAMPLES
    total_d2 = (-device1_offset_ms) + device2_delay_ms
    delay2_samples = clamp(int((total_d2 / 1000.0) * SAMPLE_RATE), 0, MAX_DELAY_SAMPLES - 1)
    write_indices_d2 = np.arange(pointer_device2, pointer_device2 + frames) % MAX_DELAY_SAMPLES
    buffer_device2[write_indices_d2] = processed_input
    read_indices_d2 = (write_indices_d2 - delay2_samples) % MAX_DELAY_SAMPLES
    output_device2 = buffer_device2[read_indices_d2]
    pointer_device2 = (pointer_device2 + frames) % MAX_DELAY_SAMPLES
    try:
        queue_dev1.put_nowait(output_device1)
    except queue.Full:
        try:
            queue_dev1.get_nowait()
            queue_dev1.put_nowait(output_device1)
        except queue.Empty:
            pass
    try:
        queue_dev2.put_nowait(output_device2)
    except queue.Full:
        try:
            queue_dev2.get_nowait()
            queue_dev2.put_nowait(output_device2)
        except queue.Empty:
            pass

def callback_output_dev1(outdata, frames, time_info, status):
    try:
        data = queue_dev1.get_nowait()
        if len(data) < frames:
            outdata[: len(data)] = data
            outdata[len(data):] = 0
        else:
            outdata[:] = data[:frames]
    except queue.Empty:
        outdata.fill(0)

def callback_output_dev2(outdata, frames, time_info, status):
    try:
        data = queue_dev2.get_nowait()
        if len(data) < frames:
            outdata[: len(data)] = data
            outdata[len(data):] = 0
        else:
            outdata[:] = data[:frames]
    except queue.Empty:
        outdata.fill(0)

def update_devices_list(*args):
    global output_device_ids, output_device_names
    selected_api_idx = combo_api.current()
    all_devices = sd.query_devices()
    output_device_ids = []
    output_device_names = []
    for i, d in enumerate(all_devices):
        if d['max_output_channels'] >= 2 and d['hostapi'] == selected_api_idx:
            output_device_ids.append(i)
            output_device_names.append(d['name'])
    combo_device1.config(values=output_device_names)
    combo_device2.config(values=output_device_names)
    if output_device_names:
        combo_device1.current(0)
        combo_device2.current(1 if len(output_device_names) > 1 else 0)
    else:
        combo_device1.set('')
        combo_device2.set('')

def toggle_engine():
    global audio_input_stream, stream_dev1, stream_dev2, is_engine_active
    if is_engine_active:
        is_engine_active = False
        while not queue_dev1.empty():
            queue_dev1.get()
        while not queue_dev2.empty():
            queue_dev2.get()
        for s in [audio_input_stream, stream_dev1, stream_dev2]:
            if s:
                try:
                    s.stop()
                    s.close()
                except Exception:
                    pass
        button_start.config(text="Start")
        return
    selected_api_idx = combo_api.current()
    virtual_device_id = None
    for i, d in enumerate(sd.query_devices()):
        if d['hostapi'] == selected_api_idx and "cable" in d['name'].lower() and d['max_input_channels'] >= 2:
            virtual_device_id = i
            break
    if virtual_device_id is None:
        messagebox.showerror("Engine Error", "VB-Cable input not found")
        return
    current_block_size = int(combo_block_size.get())
    try:
        dev1_id = output_device_ids[combo_device1.current()]
        dev2_id = output_device_ids[combo_device2.current()]
    except IndexError:
        messagebox.showerror("Engine Error", "Device not selected")
        return
    is_engine_active = True
    try:
        stream_dev1 = sd.OutputStream(
            samplerate=SAMPLE_RATE, blocksize=current_block_size, channels=2,
            device=dev1_id, latency='low', callback=callback_output_dev1,
        )
        stream_dev2 = sd.OutputStream(
            samplerate=SAMPLE_RATE, blocksize=current_block_size, channels=2,
            device=dev2_id, latency='low', callback=callback_output_dev2,
        )
        audio_input_stream = sd.InputStream(
            samplerate=SAMPLE_RATE, blocksize=current_block_size, channels=2,
            device=virtual_device_id, latency='low', callback=audio_input_callback,
        )
        stream_dev1.start()
        stream_dev2.start()
        audio_input_stream.start()
        button_start.config(text="Stop")
    except Exception as e:
        is_engine_active = False
        messagebox.showerror("Engine Error", str(e))

def _sync_entry(entry, value, decimals=1):
    text = f"{value:.{decimals}f}"
    if entry.get() != text:
        entry.delete(0, tk.END)
        entry.insert(0, text)

def update_parameters(*args):
    global reverb_amount, echo_amount, device1_delay_ms, device2_delay_ms
    global device1_offset_ms, device2_offset_ms
    for name in ('slider_reverb', 'slider_echo', 'slider_delay1', 'slider_delay2',
                 'slider_offset1', 'slider_offset2',
                 'entry_reverb', 'entry_echo', 'entry_delay1', 'entry_delay2',
                 'entry_offset1', 'entry_offset2'):
        if name not in globals():
            return
    reverb_amount = round(float(slider_reverb.get()), 1) / 100.0
    echo_amount = round(float(slider_echo.get()), 1) / 100.0
    device1_delay_ms = round(float(slider_delay1.get()), 1)
    device2_delay_ms = round(float(slider_delay2.get()), 1)
    device1_offset_ms = round(float(slider_offset1.get()), 1)
    device2_offset_ms = round(float(slider_offset2.get()), 1)
    _sync_entry(entry_reverb, reverb_amount * 100.0)
    _sync_entry(entry_echo, echo_amount * 100.0)
    _sync_entry(entry_delay1, device1_delay_ms)
    _sync_entry(entry_delay2, device2_delay_ms)
    _sync_entry(entry_offset1, device1_offset_ms)
    _sync_entry(entry_offset2, device2_offset_ms)

def _on_entry_commit(entry, slider, lo, hi, decimals=1):
    try:
        v = float(entry.get().replace(',', '.'))
    except ValueError:
        v = float(slider.get())
    v = clamp(v, lo, hi)
    v = round(v, decimals)
    slider.set(v)
    _sync_entry(entry, v, decimals)
    update_parameters()

def make_entry(parent, width=8):
    return ttk.Entry(parent, width=width, justify="center")

def bind_entry(entry, slider, lo, hi, decimals=1):
    entry.bind("<Return>", lambda ev: _on_entry_commit(entry, slider, lo, hi, decimals))
    entry.bind("<FocusOut>", lambda ev: _on_entry_commit(entry, slider, lo, hi, decimals))

def _install_autorun():
    try:
        if not getattr(sys, "frozen", False):
            messagebox.showerror("Autorun Error", "Autorun requires compiled .exe")
            return False, "Autorun requires compiled .exe"
        os.makedirs(STARTUP_DIR, exist_ok=True)
        pythoncom.CoInitialize()
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(STARTUP_LNK)
        shortcut.TargetPath = sys.executable
        shortcut.Arguments = "-autorun"
        shortcut.WorkingDirectory = os.path.dirname(sys.executable)
        shortcut.IconLocation = sys.executable
        shortcut.Save()
        pythoncom.CoUninitialize()
        return True, None
    except Exception as e:
        messagebox.showerror("Autorun Error", str(e))
        return False, str(e)

def _uninstall_autorun():
    try:
        if os.path.isfile(STARTUP_LNK):
            os.remove(STARTUP_LNK)
        return True, None
    except Exception as e:
        messagebox.showerror("Autorun Error", str(e))
        return False, str(e)

def _sync_autorun(enabled):
    global autorun_state
    if enabled == autorun_state:
        return True, None
    if enabled:
        ok, err = _install_autorun()
        if ok:
            autorun_state = True
        return ok, err
    ok, err = _uninstall_autorun()
    if ok:
        autorun_state = False
    return ok, err

def save_preset_to_disk():
    data = {
        "EngineVer": ENGINE_VER,
        "hostapi": combo_api.get(),
        "blocksize": combo_block_size.get(),
        "device1": combo_device1.get(),
        "device2": combo_device2.get(),
        "reverb": float(slider_reverb.get()),
        "echo": float(slider_echo.get()),
        "delay1": float(slider_delay1.get()),
        "delay2": float(slider_delay2.get()),
        "offset1": float(slider_offset1.get()),
        "offset2": float(slider_offset2.get()),
        "autorun": bool(var_autorun.get()),
    }
    try:
        os.makedirs(PRESET_DIR, exist_ok=True)
        with open(PRESET_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True, None
    except Exception as e:
        err = f"Save preset failed: {e}"
        messagebox.showerror("Save Error", err)
        return False, err

def _apply_preset(data):
    if "hostapi" in data and data["hostapi"] in api_names:
        combo_api.set(data["hostapi"])
        update_devices_list()
    if "blocksize" in data and data["blocksize"] in block_sizes:
        combo_block_size.set(data["blocksize"])
    if "device1" in data and data["device1"] in output_device_names:
        combo_device1.set(data["device1"])
    if "device2" in data and data["device2"] in output_device_names:
        combo_device2.set(data["device2"])
    def set_slider(slider, entry, value, lo, hi):
        try:
            v = float(value)
        except (TypeError, ValueError):
            return
        v = clamp(v, lo, hi)
        slider.set(v)
        _sync_entry(entry, v)
    if "reverb" in data:
        set_slider(slider_reverb, entry_reverb, data["reverb"], 0, 50)
    if "echo" in data:
        set_slider(slider_echo, entry_echo, data["echo"], 0, 50)
    if "delay1" in data:
        set_slider(slider_delay1, entry_delay1, data["delay1"], 0, 100)
    if "delay2" in data:
        set_slider(slider_delay2, entry_delay2, data["delay2"], 0, 100)
    if "offset1" in data:
        set_slider(slider_offset1, entry_offset1, data["offset1"], -200, 0)
    if "offset2" in data:
        set_slider(slider_offset2, entry_offset2, data["offset2"], -200, 0)
    if "autorun" in data:
        var_autorun.set(bool(data["autorun"]))
    update_parameters()

def load_preset():
    global autorun_state
    if not os.path.isfile(PRESET_PATH):
        ok, err = save_preset_to_disk()
        if not ok:
            return False, err
        return True, None
    try:
        with open(PRESET_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        err = f"Preset read failed: {e}"
        messagebox.showerror("Load Error", err)
        return False, err
    if data.get("EngineVer", None) in NON_COMPATIBLE_VERS:
        try:
            os.remove(PRESET_PATH)
        except Exception as e:
            err = f"Failed to remove old preset: {e}"
            messagebox.showerror("Load Error", err)
            return False, err
        ok, err = save_preset_to_disk()
        if not ok:
            return False, err
        return True, None
    try:
        _apply_preset(data)
        if "autorun" in data:
            autorun_state = bool(data["autorun"])
        return True, None
    except Exception as e:
        err = f"Preset apply failed: {e}"
        messagebox.showerror("Load Error", err)
        return False, err

def _show_window(icon=None, item=None):
    global tray_icon
    if tray_icon:
        tray_icon.stop()
        tray_icon = None
    root.after(0, lambda: (
        root.deiconify(),
        root.lift(),
        root.attributes("-topmost", True),
        root.after(100, lambda: root.attributes("-topmost", False)),
        root.focus_force(),
    ))

def _quit_from_tray(icon=None, item=None):
    global tray_icon
    if tray_icon:
        tray_icon.stop()
        tray_icon = None
    root.after(0, close_application)

def _run_tray():
    global tray_icon
    path = resource_path(ICON_PATH)
    if os.path.isfile(path):
        try:
            img = Image.open(path)
        except Exception:
            img = Image.new('RGB', (64, 64), color=(30, 30, 30))
            d = ImageDraw.Draw(img)
            d.ellipse((12, 12, 52, 52), fill=(80, 200, 120))
    else:
        img = Image.new('RGB', (64, 64), color=(30, 30, 30))
        d = ImageDraw.Draw(img)
        d.ellipse((12, 12, 52, 52), fill=(80, 200, 120))
    menu = pystray.Menu(
        pystray.MenuItem("Show", _show_window, default=True),
        pystray.MenuItem("Exit", _quit_from_tray),
    )
    tray_icon = pystray.Icon("QuadroHaas", img, "QuadroHaas Engine", menu)
    tray_icon.run()

def minimize_to_tray(event=None):
    global tray_message_shown
    if event is not None and event.widget is not root:
        return
    if root.state() == 'iconic':
        if not tray_message_shown:
            tray_message_shown = True
            messagebox.showinfo("Info", "This program will be moved into Tray")
        root.withdraw()
        threading.Thread(target=_run_tray, daemon=True).start()

def _tick_loop():
    tick_sound = pygame.mixer.Sound(resource_path(TICK_PATH))
    next_time = time.perf_counter()
    while not tick_stop.is_set():
        tick_sound.play()
        next_time += TICK_INTERVAL
        sleep_time = next_time - time.perf_counter()
        if sleep_time > 0:
            time.sleep(sleep_time)
        else:
            next_time = time.perf_counter()

def on_tick_toggle():
    global tick_thread
    if var_tick.get():
        tick_stop.clear()
        tick_thread = threading.Thread(target=_tick_loop, daemon=True)
        tick_thread.start()
    else:
        tick_stop.set()
        tick_thread = None

def on_ost_toggle():
    if var_ost.get():
        pygame.mixer.music.load(resource_path(OST_PATH))
        pygame.mixer.music.play(-1)
    else:
        pygame.mixer.music.stop()

def close_application():
    global is_engine_active, autorun_changed
    tick_stop.set()
    pygame.mixer.music.stop()
    ok, err = save_preset_to_disk()
    if not ok:
        messagebox.showerror("Close Error", err)
    if autorun_changed:
        ok, err = _sync_autorun(bool(var_autorun.get()))
        if not ok:
            messagebox.showerror("Close Error", err)
        autorun_changed = False
    is_engine_active = False
    for s in [audio_input_stream, stream_dev1, stream_dev2]:
        if s:
            try:
                s.stop()
                s.close()
            except Exception:
                pass
    root.destroy()

pygame.mixer.pre_init(frequency=SAMPLE_RATE, size=-16, channels=2, buffer=512)
pygame.mixer.init()

root = tk.Tk()
root.title("QuadroHaas Engine v1.2")
root.geometry("480x720")
root.resizable(False, False)
icon_path_full = resource_path(ICON_PATH)
if os.path.isfile(icon_path_full):
    try:
        root.iconbitmap(icon_path_full)
    except Exception:
        pass
style = ttk.Style()
style.theme_use('winnative')
host_apis = sd.query_hostapis()
api_names = [api['name'] for api in host_apis]
block_sizes = ["128", "256", "512", "1024", "2048", "4096", "8192"]

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=5, pady=5)

tab_sound = ttk.Frame(notebook)
tab_test = ttk.Frame(notebook)
notebook.add(tab_sound, text="Sound")
notebook.add(tab_test, text="Test")

frame_api = ttk.LabelFrame(tab_sound, text=" Audio Settings ")
frame_api.pack(fill="x", padx=20, pady=5)
ttk.Label(frame_api, text="Audio Driver (Host API):").pack(anchor="w", padx=10, pady=(2, 0))
combo_api = ttk.Combobox(frame_api, values=api_names, state="readonly", width=45)
combo_api.pack(pady=5, padx=10)
combo_api.bind("<<ComboboxSelected>>", update_devices_list)
ttk.Label(frame_api, text="Block Size (Samples):").pack(anchor="w", padx=10, pady=(2, 0))
combo_block_size = ttk.Combobox(frame_api, values=block_sizes, state="readonly", width=45)
combo_block_size.pack(pady=5, padx=10)
combo_block_size.current(3)

frame_devices = ttk.LabelFrame(tab_sound, text=" Audio outputs ")
frame_devices.pack(fill="x", padx=20, pady=5)
combo_device1 = ttk.Combobox(frame_devices, values=[], state="readonly", width=45)
combo_device1.pack(pady=5, padx=10)
combo_device2 = ttk.Combobox(frame_devices, values=[], state="readonly", width=45)
combo_device2.pack(pady=5, padx=10)

frame_offsets = ttk.LabelFrame(tab_sound, text=" Device Delay Offset (ms) ")
frame_offsets.pack(fill="x", padx=20, pady=5)
row_o1 = ttk.Frame(frame_offsets)
row_o1.pack(fill="x", padx=5, pady=3)
ttk.Label(row_o1, text="Device 1 offset:", width=15).pack(side="left", padx=5)
slider_offset1 = ttk.Scale(row_o1, from_=-200, to=0, orient="horizontal", command=update_parameters)
slider_offset1.pack(side="left", fill="x", expand=True)
slider_offset1.set(0)
entry_offset1 = make_entry(row_o1, width=8)
entry_offset1.pack(side="right", padx=5)
entry_offset1.insert(0, "0.0")
row_o2 = ttk.Frame(frame_offsets)
row_o2.pack(fill="x", padx=5, pady=3)
ttk.Label(row_o2, text="Device 2 offset:", width=15).pack(side="left", padx=5)
slider_offset2 = ttk.Scale(row_o2, from_=-200, to=0, orient="horizontal", command=update_parameters)
slider_offset2.pack(side="left", fill="x", expand=True)
slider_offset2.set(0)
entry_offset2 = make_entry(row_o2, width=8)
entry_offset2.pack(side="right", padx=5)
entry_offset2.insert(0, "0.0")
bind_entry(entry_offset1, slider_offset1, -200, 0)
bind_entry(entry_offset2, slider_offset2, -200, 0)

if api_names:
    combo_api.current(0)
    if "MME" in api_names:
        combo_api.set("MME")
    if "Windows DirectSound" in api_names:
        combo_api.set("Windows DirectSound")
    update_devices_list()

frame_effects = ttk.LabelFrame(tab_sound, text=" Effects ")
frame_effects.pack(fill="x", padx=20, pady=5)
frame_rev = ttk.Frame(frame_effects)
frame_rev.pack(fill="x", padx=5, pady=3)
ttk.Label(frame_rev, text="Reverb:", width=15).pack(side="left", padx=5)
slider_reverb = ttk.Scale(frame_rev, from_=0, to=50, orient="horizontal", command=update_parameters)
slider_reverb.pack(side="left", fill="x", expand=True)
slider_reverb.set(0)
entry_reverb = make_entry(frame_rev, width=8)
entry_reverb.pack(side="right", padx=5)
entry_reverb.insert(0, "0.0")
bind_entry(entry_reverb, slider_reverb, 0, 50)
frame_echo = ttk.Frame(frame_effects)
frame_echo.pack(fill="x", padx=5, pady=3)
ttk.Label(frame_echo, text="Echo:", width=15).pack(side="left", padx=5)
slider_echo = ttk.Scale(frame_echo, from_=0, to=50, orient="horizontal", command=update_parameters)
slider_echo.pack(side="left", fill="x", expand=True)
slider_echo.set(0)
entry_echo = make_entry(frame_echo, width=8)
entry_echo.pack(side="right", padx=5)
entry_echo.insert(0, "0.0")
bind_entry(entry_echo, slider_echo, 0, 50)
frame_del1 = ttk.Frame(frame_effects)
frame_del1.pack(fill="x", padx=5, pady=3)
ttk.Label(frame_del1, text="Device 1 delay:", width=15).pack(side="left", padx=5)
slider_delay1 = ttk.Scale(frame_del1, from_=0, to=100, orient="horizontal", command=update_parameters)
slider_delay1.pack(side="left", fill="x", expand=True)
slider_delay1.set(0)
entry_delay1 = make_entry(frame_del1, width=8)
entry_delay1.pack(side="right", padx=5)
entry_delay1.insert(0, "0.0")
bind_entry(entry_delay1, slider_delay1, 0, 100)
frame_del2 = ttk.Frame(frame_effects)
frame_del2.pack(fill="x", padx=5, pady=3)
ttk.Label(frame_del2, text="Device 2 delay:", width=15).pack(side="left", padx=5)
slider_delay2 = ttk.Scale(frame_del2, from_=0, to=100, orient="horizontal", command=update_parameters)
slider_delay2.pack(side="left", fill="x", expand=True)
slider_delay2.set(0)
entry_delay2 = make_entry(frame_del2, width=8)
entry_delay2.pack(side="right", padx=5)
entry_delay2.insert(0, "0.0")
bind_entry(entry_delay2, slider_delay2, 0, 100)

frame_preset = ttk.LabelFrame(tab_sound, text=" Autorun ")
frame_preset.pack(fill="x", padx=20, pady=5)
var_autorun = tk.BooleanVar(value=False)
chk_autorun = ttk.Checkbutton(frame_preset, text="Autorun with current preset", variable=var_autorun, command=lambda: globals().update(autorun_changed=True))
chk_autorun.pack(anchor="w", padx=10, pady=5)

button_start = ttk.Button(tab_sound, text="Start", command=toggle_engine)
button_start.pack(pady=15, fill="x", padx=20, ipady=5)

frame_tick = ttk.LabelFrame(tab_test, text=" Metronome ")
frame_tick.pack(fill="x", padx=20, pady=10)
var_tick = tk.BooleanVar(value=False)
chk_tick = ttk.Checkbutton(frame_tick, text="Tick 120 BPM", variable=var_tick, command=on_tick_toggle)
chk_tick.pack(anchor="w", padx=10, pady=10)

frame_ost = ttk.LabelFrame(tab_test, text=" Music ")
frame_ost.pack(fill="x", padx=20, pady=10)
var_ost = tk.BooleanVar(value=False)
chk_ost = ttk.Checkbutton(frame_ost, text="Test Ost", variable=var_ost, command=on_ost_toggle)
chk_ost.pack(anchor="w", padx=10, pady=10)

root.bind("<Unmap>", minimize_to_tray)
root.protocol("WM_DELETE_WINDOW", close_application)

load_preset()

if "-autorun" in sys.argv:
    root.after(100, lambda: (
        globals().update(tray_message_shown=True),
        toggle_engine() if not is_engine_active else None,
        root.withdraw(),
        threading.Thread(target=_run_tray, daemon=True).start(),
    ))

root.mainloop()
