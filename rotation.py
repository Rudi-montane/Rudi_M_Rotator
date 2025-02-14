######################################
# erstellt von big dick rudi montane #
######################################
# schwör lerne macht spass           #
######################################
 #imports 
import tkinter as tk # Gui 
from tkinter import ttk # Gui
import threading # threading für makro 
import time # sleep für makro unterbruch
import random # random für sleeps
import os # safe und laden von datei
import json # speicher für json datei
import win32gui  # pywin32: pip install pywin32 # window graber
from pynput.keyboard import Key, Controller, GlobalHotKeys  # pip install pynput # tasten emulation

# tastaturcontroller init
keyboard_controller = Controller()

##################################
# globale variablen
##################################
macro_active = False  # ob das makro (inkl hotkey) aktiv ist
rotation_active = False  # ob rotation aktiv ist
sequence_keys = ["6", "7", "8", "9", "0", "'", "^", "1", "2", "3", "follow", "4", "5", "6"] # standart rotation # change wenn willst
target_window_title = "retailpartz"  # standardfenster
altf_enabled = True  # ob altf aktiv ist
config_file = "macro_config.json"  # standard datei

hotkeys = None  # globalhotkeys instanz
ontop_var = None  # always on top

##################################
# hilfsfunktionen
##################################

def console_output(text):
    """fügt nachricht in gui konsole hinzu"""
    console_text.insert(tk.END, text + "\n")
    console_text.see(tk.END)


def send_key(key):
    """sendet eine taste"""
    try:
        keyboard_controller.press(key)
        keyboard_controller.release(key)
    except Exception as e:
        console_output(f"fehler senden {key} {e}")


def send_hotkey(keys):
    """sendet eine hotkey-kombination"""
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
        console_output(f"fehler hotkey {keys} {e}")


def is_target_window_active():
    """prüft fenstertitel"""
    active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
    return target_window_title.lower() in active_title

##################################
# rotation
##################################

def combined_loop():
    """schleife für tasten sequenz"""
    global rotation_active
    seq_index = 0
    while rotation_active:
        if is_target_window_active():
            current = sequence_keys[seq_index].strip().lower()
            console_output(f"sende {current}")
            if current in ["follow", "followall", "alt+f"]:
                if altf_enabled:
                    send_hotkey(['alt', 'f'])
                else:
                    console_output("alt+f aus")
            else:
                send_key(Key.f7)
                time.sleep(random.uniform(0.18, 0.22))
                send_key(sequence_keys[seq_index])
            time.sleep(random.uniform(0.18, 0.22))
            seq_index = (seq_index + 1) % len(sequence_keys)
        else:
            console_output("ziel fenster nicht aktiv warte")
            time.sleep(0.5)

def toggle_rotation():
    """schaltet rotation an oder aus"""
    global rotation_active
    if rotation_active:
        rotation_active = False
        console_output("rotation aus")
    else:
        rotation_active = True
        console_output("rotation an")
        threading.Thread(target=combined_loop, daemon=True).start()

##################################
# makro an/aus (hotkey aktivieren)
##################################

def toggle_macro():
    """aktiviert oder deaktiviert das makro (globalhotkey r)"""
    global macro_active, hotkeys
    if macro_active:
        macro_active = False
        console_output("makro aus (kein hotkey)")
        if hotkeys:
            hotkeys.stop()
            hotkeys = None
    else:
        macro_active = True
        console_output("makro an (hotkey r aktiv)")
        hotkeys = GlobalHotKeys({'r': toggle_rotation})
        hotkeys.start()

##################################
# config laden/speichern
##################################

def load_config():
    """lädt config"""
    global sequence_keys, target_window_title, altf_enabled
    file_name = config_filename_var.get().strip()
    if not file_name:
        console_output("kein dateiname kein laden")
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
            console_output(f"config aus {file_name} geladen")
        except Exception as e:
            console_output(f"fehler laden config {e}")
    else:
        console_output(f"datei {file_name} nicht da")


def save_config():
    """speichert config"""
    global sequence_keys, target_window_title, altf_enabled
    file_name = config_filename_var.get().strip()
    if not file_name:
        console_output("kein dateiname kein speichern")
        return
    data = {
        "sequence_keys": sequence_keys,
        "window_title": target_window_title,
        "altf_enabled": altf_enabled
    }
    try:
        with open(file_name, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        console_output(f"config in {file_name} gespeichert")
    except Exception as e:
        console_output(f"fehler speichern config {e}")

##################################
# gui events
##################################

def update_config():
    """holt seq aus eingabe"""
    global sequence_keys
    seq_str = sequence_entry.get()
    sequence_keys = [s.strip() for s in seq_str.split(",") if s.strip()]
    console_output(f"seq {sequence_keys}")

def toggle_altf():
    """schaltet alt+f um"""
    global altf_enabled
    altf_enabled = bool(altf_var.get())
    console_output(f"alt+f {'an' if altf_enabled else 'aus'}")

def on_close():
    """beendet programm"""
    global rotation_active, hotkeys, macro_active
    rotation_active = False
    macro_active = False
    if hotkeys:
        hotkeys.stop()
    root.destroy()

def toggle_always_on_top():
    """schaltet always on top um"""
    root.attributes("-topmost", bool(ontop_var.get()))

##################################
# fenster handling
##################################

def enum_windows():
    """listet sichtbare fenster auf"""
    results = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            wtext = win32gui.GetWindowText(hwnd)
            if wtext:
                results.append(wtext)
    win32gui.EnumWindows(enum_handler, None)
    return results

def refresh_window_list():
    """aktualisiert combobox liste"""
    window_combobox['values'] = enum_windows()

def select_window(event):
    """übernimmt fensterauswahl"""
    global target_window_title
    chosen = window_combobox.get()
    target_window_title = chosen
    console_output(f"aktives fenster {target_window_title}")

##################################
# gui leech von GPT. Kein Bock GUI langweilig 
##################################
root = tk.Tk()
root.title("rudi m. rotator") #titel
root.geometry("600x600") #grösse Fenster
root.configure(bg="#222") #BG gray

style = ttk.Style()
style.configure("TButton", foreground="red", background="black") #ok coloring ist fun
style.configure("TLabel", foreground="pink", background="#222")
style.configure("TEntry", foreground="black", fieldbackground="pink")
style.configure("TCombobox", foreground="black", fieldbackground="pink")
style.configure("TCheckbutton", foreground="pink", background="#222")

info_label = ttk.Label(root, text="button makro an aus  taste r rotation an aus")
info_label.pack(pady=5) #label text info

window_label = ttk.Label(root, text="fenstertitel")
window_label.pack()

window_combobox = ttk.Combobox(root, width=50)
window_combobox.pack(pady=2)
window_combobox.bind("<<ComboboxSelected>>", select_window)

refresh_button = ttk.Button(root, text="fensterliste aktualisieren", command=refresh_window_list)
refresh_button.pack(pady=5)

sequence_label = ttk.Label(root, text="sequenz (komma)")
sequence_label.pack()
sequence_entry = ttk.Entry(root, width=50)
sequence_entry.insert(0, ", ".join(sequence_keys))
sequence_entry.pack()

update_button = ttk.Button(root, text="konfig update", command=update_config)
update_button.pack(pady=5) #update config

altf_var = tk.IntVar(value=1 if altf_enabled else 0)
altf_check = ttk.Checkbutton(root, text="alt+f follow an", variable=altf_var, command=toggle_altf)
altf_check.pack(pady=5) #follow checker

ontop_var = tk.IntVar(value=0)
ontop_check = ttk.Checkbutton(root, text="always on top", variable=ontop_var, command=toggle_always_on_top)
ontop_check.pack(pady=5) # always on top enabler

config_filename_frame = ttk.Frame(root)
config_filename_frame.pack(pady=5)

config_filename_label = ttk.Label(config_filename_frame, text="json datei")
config_filename_label.grid(row=0, column=0, padx=5) #config datei .json

config_filename_var = tk.StringVar(value=config_file)
config_filename_entry = ttk.Entry(config_filename_frame, width=35, textvariable=config_filename_var)
config_filename_entry.grid(row=0, column=1, padx=5)

frame_buttons = ttk.Frame(root)
frame_buttons.pack(pady=5)

save_button = ttk.Button(frame_buttons, text="speichern", command=save_config)
save_button.grid(row=0, column=0, padx=5) #safer

load_button = ttk.Button(frame_buttons, text="laden", command=load_config)
load_button.grid(row=0, column=1, padx=5) #loader

macro_button = ttk.Button(root, text="makro an aus", command=toggle_macro)
macro_button.pack(pady=5) #global ein aus 

console_label = ttk.Label(root, text="konsole")
console_label.pack() #console 

console_text = tk.Text(root, height=10, width=70, bg="black", fg="red")
console_text.pack() #console

root.after(100, load_config) #loader

refresh_window_list()
window_combobox.set(target_window_title)

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
