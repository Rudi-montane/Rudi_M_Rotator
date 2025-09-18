# --- Imports ---
import tkinter as tk
from tkinter import ttk
import threading
import time
import random
import os
import json
import win32gui               # Window handling via pywin32
from pynput.keyboard import Key, Controller, GlobalHotKeys  # Keyboard emulation

# Keyboard controller instance from pynput
keyboard_controller = Controller()

# --- Global state variables ---
macro_active = False           # Tracks whether the macro (with global hotkey) is active
rotation_active = False        # Tracks whether the skill rotation loop is running
sequence_keys = ["6", "7", "8", "9", "0", "'", "^", "1", "2", "3", "follow", "4", "5", "6"]  
# Default sequence of keys to press, can be changed by user
target_window_title = "retailpartz"   # Default window title to target
altf_enabled = True             # Whether the Alt+F follow hotkey is enabled
config_file = "macro_config.json"  # Default configuration file

hotkeys = None                  # Will hold GlobalHotKeys instance
ontop_var = None                # Keeps GUI on top if enabled

# --- Helper functions ---

def console_output(text):
    """Writes a message into the console box inside the GUI."""
    console_text.insert(tk.END, text + "\n")
    console_text.see(tk.END)


def send_key(key):
    """Simulates pressing and releasing a single key."""
    try:
        keyboard_controller.press(key)
        keyboard_controller.release(key)
    except Exception as e:
        console_output(f"Error sending {key}: {e}")


def send_hotkey(keys):
    """Sends a key combination, e.g. ['alt', 'f']."""
    try:
        for k in keys:
            if k.lower() == 'alt':
                keyboard_controller.press(Key.alt)
            else:
                keyboard_controller.press(k)
        for k in reversed(keys):
            if k.lower() == 'alt':
                keyboard_controller.release(Key.alt)
            else:
                keyboard_controller.release(k)
    except Exception as e:
        console_output(f"Error sending hotkey {keys}: {e}")


def is_target_window_active():
    """Checks if the currently active foreground window matches the configured title."""
    active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
    return target_window_title.lower() in active_title

# --- Rotation logic ---

def combined_loop():
    """
    Main loop that cycles through the key sequence while rotation is active.
    Adds random sleep intervals to mimic natural delays.
    """
    global rotation_active
    seq_index = 0
    while rotation_active:
        if is_target_window_active():
            current = sequence_keys[seq_index].strip().lower()
            console_output(f"Sending {current}")
            if current in ["follow", "followall", "alt+f"]:
                if altf_enabled:
                    send_hotkey(['alt', 'f'])
                else:
                    console_output("Alt+F disabled")
            else:
                # Example: small delay before skill press
                send_key(Key.f7)  
                time.sleep(random.uniform(0.18, 0.22))
                send_key(sequence_keys[seq_index])
            time.sleep(random.uniform(0.18, 0.22))
            seq_index = (seq_index + 1) % len(sequence_keys)
        else:
            console_output("Target window not active, waiting...")
            time.sleep(0.5)


def toggle_rotation():
    """Turns the rotation loop on or off."""
    global rotation_active
    if rotation_active:
        rotation_active = False
        console_output("Rotation stopped")
    else:
        rotation_active = True
        console_output("Rotation started")
        threading.Thread(target=combined_loop, daemon=True).start()

# --- Macro activation (hotkey binding) ---

def toggle_macro():
    """
    Enables or disables the macro entirely.
    Registers/unregisters global hotkeys (R toggles rotation).
    """
    global macro_active, hotkeys
    if macro_active:
        macro_active = False
        console_output("Macro off (hotkey disabled)")
        if hotkeys:
            hotkeys.stop()
            hotkeys = None
    else:
        macro_active = True
        console_output("Macro on (R key is active)")
        hotkeys = GlobalHotKeys({'r': toggle_rotation})
        hotkeys.start()

# --- Config load/save ---

def load_config():
    """Loads configuration (sequence, window title, follow option) from JSON file."""
    global sequence_keys, target_window_title, altf_enabled
    file_name = config_filename_var.get().strip()
    if not file_name:
        console_output("No filename provided, cannot load")
        return
    if os.path.exists(file_name):
        try:
            with open(file_name, "r", encoding="utf-8") as f:
                data = json.load(f)
            sequence_keys = data.get("sequence_keys", sequence_keys)
            target_window_title = data.get("window_title", target_window_title)
            altf_enabled = data.get("altf_enabled", altf_enabled)
            sequence_entry.delete(0, tk.END)
            sequence_entry.insert(0, ", ".join(sequence_keys))
            altf_var.set(1 if altf_enabled else 0)
            refresh_window_list()
            window_combobox.set(target_window_title)
            console_output(f"Config loaded from {file_name}")
        except Exception as e:
            console_output(f"Error loading config: {e}")
    else:
        console_output(f"File {file_name} does not exist")


def save_config():
    """Saves current configuration to JSON file."""
    global sequence_keys, target_window_title, altf_enabled
    file_name = config_filename_var.get().strip()
    if not file_name:
        console_output("No filename provided, cannot save")
        return
    data = {
        "sequence_keys": sequence_keys,
        "window_title": target_window_title,
        "altf_enabled": altf_enabled
    }
    try:
        with open(file_name, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        console_output(f"Config saved to {file_name}")
    except Exception as e:
        console_output(f"Error saving config: {e}")

# --- GUI events ---

def update_config():
    """Reads the sequence from the entry field and updates the in-memory list."""
    global sequence_keys
    seq_str = sequence_entry.get()
    sequence_keys = [s.strip() for s in seq_str.split(",") if s.strip()]
    console_output(f"Updated sequence: {sequence_keys}")


def toggle_altf():
    """Enables or disables the Alt+F follow hotkey."""
    global altf_enabled
    altf_enabled = bool(altf_var.get())
    console_output(f"Alt+F {'enabled' if altf_enabled else 'disabled'}")


def on_close():
    """Stops rotation, hotkeys, and closes the application window cleanly."""
    global rotation_active, hotkeys, macro_active
    rotation_active = False
    macro_active = False
    if hotkeys:
        hotkeys.stop()
    root.destroy()


def toggle_always_on_top():
    """Keeps the window always on top when enabled."""
    root.attributes("-topmost", bool(ontop_var.get()))

# --- Window handling ---

def enum_windows():
    """Enumerates all visible windows and returns their titles."""
    results = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            wtext = win32gui.GetWindowText(hwnd)
            if wtext:
                results.append(wtext)
    win32gui.EnumWindows(enum_handler, None)
    return results


def refresh_window_list():
    """Refreshes the combobox with the list of active window titles."""
    window_combobox['values'] = enum_windows()


def select_window(event):
    """Sets the selected window title as the active target."""
    global target_window_title
    chosen = window_combobox.get()
    target_window_title = chosen
    console_output(f"Target window set to: {target_window_title}")

# --- GUI setup ---

root = tk.Tk()
root.title("rudi m. rotator")
root.geometry("600x600")
root.configure(bg="#222")

style = ttk.Style()
style.configure("TButton", foreground="red", background="black")
style.configure("TLabel", foreground="pink", background="#222")
style.configure("TEntry", foreground="black", fieldbackground="pink")
style.configure("TCombobox", foreground="black", fieldbackground="pink")
style.configure("TCheckbutton", foreground="pink", background="#222")

info_label = ttk.Label(root, text="Macro toggle button (R key controls rotation)")
info_label.pack(pady=5)

window_label = ttk.Label(root, text="Target window title")
window_label.pack()

window_combobox = ttk.Combobox(root, width=50)
window_combobox.pack(pady=2)
window_combobox.bind("<<ComboboxSelected>>", select_window)

refresh_button = ttk.Button(root, text="Refresh window list", command=refresh_window_list)
refresh_button.pack(pady=5)

sequence_label = ttk.Label(root, text="Sequence (comma-separated)")
sequence_label.pack()
sequence_entry = ttk.Entry(root, width=50)
sequence_entry.insert(0, ", ".join(sequence_keys))
sequence_entry.pack()

update_button = ttk.Button(root, text="Update config", command=update_config)
update_button.pack(pady=5)

altf_var = tk.IntVar(value=1 if altf_enabled else 0)
altf_check = ttk.Checkbutton(root, text="Enable Alt+F follow", variable=altf_var, command=toggle_altf)
altf_check.pack(pady=5)

ontop_var = tk.IntVar(value=0)
ontop_check = ttk.Checkbutton(root, text="Always on top", variable=ontop_var, command=toggle_always_on_top)
ontop_check.pack(pady=5)

config_filename_frame = ttk.Frame(root)
config_filename_frame.pack(pady=5)

config_filename_label = ttk.Label(config_filename_frame, text="Config file (JSON)")
config_filename_label.grid(row=0, column=0, padx=5)

config_filename_var = tk.StringVar(value=config_file)
config_filename_entry = ttk.Entry(config_filename_frame, width=35, textvariable=config_filename_var)
config_filename_entry.grid(row=0, column=1, padx=5)

frame_buttons = ttk.Frame(root)
frame_buttons.pack(pady=5)

save_button = ttk.Button(frame_buttons, text="Save", command=save_config)
save_button.grid(row=0, column=0, padx=5)

load_button = ttk.Button(frame_buttons, text="Load", command=load_config)
load_button.grid(row=0, column=1, padx=5)

macro_button = ttk.Button(root, text="Macro toggle", command=toggle_macro)
macro_button.pack(pady=5)

console_label = ttk.Label(root, text="Console output")
console_label.pack()

console_text = tk.Text(root, height=10, width=70, bg="black", fg="red")
console_text.pack()

root.after(100, load_config)

refresh_window_list()
window_combobox.set(target_window_title)

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()