# produced by big dick rudi montane
# import tkinter libary as tk
import tkinter as tk
# import ttk from tkinter for widget styling
from tkinter import ttk
# import threading for background tasks
import threading
# import time for sleep and timing
import time
# import random for random intervals
import random
# import os for file handling
import os
# import json for config file handling
import json
# import win32gui for window handling pip install pywin32
import win32gui  # pywin32 pip install pywin32
# import keyboard control and global hotkeys pip install pynput
from pynput.keyboard import Key, Controller, GlobalHotKeys  # pip install pynput

# init tastatur controller
keyboard_controller = Controller()

##################################
# globale variablen
##################################
# makro aktiv false heisst makro hotkey is aus
macro_active = False  # ob das makro inkl hotkey aktiv ist
# rotation aktiv false heisst rotation is aus
rotation_active = False  # ob rotation aktiv ist
# liste von tasten fuer sequence rotation
sequence_keys = ["6", "7", "8", "9", "0", "'", "^", "1", "2", "3", "follow", "4", "5", "6"]
# ziel fenster titel standard
target_window_title = "retailpartz"  # standardfenster
# alt f funktion enabled true
altf_enabled = True  # ob altf an
# dateiname fuer config
config_file = "macro_config.json"  # standard datei
# variable fuer globalhotkeys instanz
hotkeys = None  # globalhotkeys instanz
# variable fuer fenster immer im vordergrund
ontop_var = None  # always on top

##################################
# hilfsfunktionen
##################################
def console_output(text):
    # zeigt text in der gui konsohle an
    console_text.insert(tk.END, text + "\n")
    console_text.see(tk.END)

def send_key(key):
    # sendet einzelne taste ueber tastatur controller
    try:
        keyboard_controller.press(key)
        keyboard_controller.release(key)
    except Exception as e:
        console_output(f"fehler senden {key} {e}")

def send_hotkey(keys):
    # sendet hotkey kombination ueber tastatur controller
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
    # prueft ob das ziel fenster aktiv is
    active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow()).lower()
    return target_window_title.lower() in active_title

##################################
# rotation
##################################
def combined_loop():
    # hauptschleife fuer sequence rotation wenn fenster aktiv is
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
    # schaltet rotation an oder aus
    global rotation_active
    if rotation_active:
        rotation_active = False
        console_output("rotation aus")
    else:
        rotation_active = True
        console_output("rotation an")
        threading.Thread(target=combined_loop, daemon=True).start()

##################################
# macro an aus hotkey aktivieren
##################################
def toggle_macro():
    # schaltet makro hotkey r an oder aus
    global macro_active, hotkeys
    if macro_active:
        macro_active = False
        console_output("makro aus kein hotkey")
        if hotkeys:
            hotkeys.stop()
            hotkeys = None
    else:
        macro_active = True
        console_output("makro an hotkey r aktiv")
        hotkeys = GlobalHotKeys({'r': toggle_rotation})
        hotkeys.start()

##################################
# config laden speichern
##################################
def load_config():
    # laedt konfiguration aus json datei
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
    # speichert konfiguration in json datei
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
    # aktualisiert sequence liste aus eingabefeld
    global sequence_keys
    seq_str = sequence_entry.get()
    sequence_keys = [s.strip() for s in seq_str.split(",") if s.strip()]
    console_output(f"seq {sequence_keys}")

def toggle_altf():
    # schaltet alt f funktion an oder aus
    global altf_enabled
    altf_enabled = bool(altf_var.get())
    console_output(f"alt f {'an' if altf_enabled else 'aus'}")

def on_close():
    # beendet programm sauber und stoppt alle loops
    global rotation_active, hotkeys, macro_active
    rotation_active = False
    macro_active = False
    if hotkeys:
        hotkeys.stop()
    root.destroy()

def toggle_always_on_top():
    # schaltet fenster immer im vordergrund an oder aus
    root.attributes("-topmost", bool(ontop_var.get()))

##################################
# fenster handling
##################################
def enum_windows():
    # listet alle sichtbaren fenster auf
    results = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            wtext = win32gui.GetWindowText(hwnd)
            if wtext:
                results.append(wtext)
    win32gui.EnumWindows(enum_handler, None)
    return results

def refresh_window_list():
    # aktualisiert fenster liste in combobox
    window_combobox['values'] = enum_windows()

def select_window(event):
    # uebernimmt fensterauswahl aus combobox
    global target_window_title
    chosen = window_combobox.get()
    target_window_title = chosen
    console_output(f"aktives fenster {target_window_title}")

##################################
# gui
##################################
# erstellt hauptfenster
root = tk.Tk()
# setzt fenstertitel
root.title("rudi m. rotator")
# setzt fenster groesse
root.geometry("600x600")
# setzt hintergrund farbe
root.configure(bg="#222")

# konfiguriert style fuer widgets
style = ttk.Style()
style.configure("TButton", foreground="red", background="black")
style.configure("TLabel", foreground="white", background="#222")
style.configure("TEntry", foreground="black", fieldbackground="white")
style.configure("TCombobox", foreground="black", fieldbackground="white")
style.configure("TCheckbutton", foreground="white", background="#222")

# erstellt info label mit hinweisen
info_label = ttk.Label(root, text="button makro an aus  taste r rotation an aus")
info_label.pack(pady=5)

# erstellt label fuer fenstertitel
window_label = ttk.Label(root, text="fenstertitel")
window_label.pack()

# erstellt combobox zur fensterauswahl
window_combobox = ttk.Combobox(root, width=50)
window_combobox.pack(pady=2)
window_combobox.bind("<<ComboboxSelected>>", select_window)

# erstellt button um fensterliste zu aktualisieren
refresh_button = ttk.Button(root, text="fensterliste aktualisieren", command=refresh_window_list)
refresh_button.pack(pady=5)

# erstellt label fuer sequence eingabe
sequence_label = ttk.Label(root, text="sequenz (komma)")
sequence_label.pack()
# erstellt eingabefeld fuer sequence
sequence_entry = ttk.Entry(root, width=50)
sequence_entry.insert(0, ", ".join(sequence_keys))
sequence_entry.pack()

# erstellt button um konfig update auszufuhren
update_button = ttk.Button(root, text="konfig update", command=update_config)
update_button.pack(pady=5)

# erstellt checkbutton fuer alt f follow funktion
altf_var = tk.IntVar(value=1 if altf_enabled else 0)
altf_check = ttk.Checkbutton(root, text="alt+f follow an", variable=altf_var, command=toggle_altf)
altf_check.pack(pady=5)

# erstellt checkbutton fuer always on top funktion
ontop_var = tk.IntVar(value=0)
ontop_check = ttk.Checkbutton(root, text="always on top", variable=ontop_var, command=toggle_always_on_top)
ontop_check.pack(pady=5)

# erstellt frame fuer json datei eingabe
config_filename_frame = ttk.Frame(root)
config_filename_frame.pack(pady=5)

# erstellt label fuer json datei
config_filename_label = ttk.Label(config_filename_frame, text="json datei")
config_filename_label.grid(row=0, column=0, padx=5)

# erstellt eingabefeld fuer dateiname
config_filename_var = tk.StringVar(value=config_file)
config_filename_entry = ttk.Entry(config_filename_frame, width=35, textvariable=config_filename_var)
config_filename_entry.grid(row=0, column=1, padx=5)

# erstellt frame fuer speichern und laden buttons
frame_buttons = ttk.Frame(root)
frame_buttons.pack(pady=5)

# erstellt button um config zu speichern
save_button = ttk.Button(frame_buttons, text="speichern", command=save_config)
save_button.grid(row=0, column=0, padx=5)

# erstellt button um config zu laden
load_button = ttk.Button(frame_buttons, text="laden", command=load_config)
load_button.grid(row=0, column=1, padx=5)

# erstellt button um makro hotkey an oder aus zu schalten
macro_button = ttk.Button(root, text="makro an aus", command=toggle_macro)
macro_button.pack(pady=5)

# erstellt label fuer konsohle ausgabe
console_label = ttk.Label(root, text="konsole")
console_label.pack()

# erstellt text widget fuer konsohle ausgabe
console_text = tk.Text(root, height=10, width=70, bg="black", fg="red")
console_text.pack()

# laedt config nach kurzer zeit
root.after(100, load_config)

# aktualisiert fenster liste und setzt standard fenster
refresh_window_list()
window_combobox.set(target_window_title)

# definiert schliess event und beendet programm
root.protocol("WM_DELETE_WINDOW", on_close)
# startet gui hauptloop
root.mainloop()
