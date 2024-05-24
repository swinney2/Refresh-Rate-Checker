import ctypes
import json
import os
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox
import winsound
import pystray
from PIL import Image, ImageDraw
from pystray import MenuItem as item
import win32api
import win32com.client
import threading
from screeninfo import get_monitors

# Constants
DISPLAY_DEVICE_ACTIVE = 0x00000001
ENUM_CURRENT_SETTINGS = -1

class DEVMODE(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", ctypes.c_wchar * 32),
        ("dmSpecVersion", ctypes.c_ushort),
        ("dmDriverVersion", ctypes.c_ushort),
        ("dmSize", ctypes.c_ushort),
        ("dmDriverExtra", ctypes.c_ushort),
        ("dmFields", ctypes.c_ulong),
        ("dmOrientation", ctypes.c_short),
        ("dmPaperSize", ctypes.c_short),
        ("dmPaperLength", ctypes.c_short),
        ("dmPaperWidth", ctypes.c_short),
        ("dmScale", ctypes.c_short),
        ("dmCopies", ctypes.c_short),
        ("dmDefaultSource", ctypes.c_short),
        ("dmPrintQuality", ctypes.c_short),
        ("dmColor", ctypes.c_short),
        ("dmDuplex", ctypes.c_short),
        ("dmYResolution", ctypes.c_short),
        ("dmTTOption", ctypes.c_short),
        ("dmCollate", ctypes.c_short),
        ("dmFormName", ctypes.c_wchar * 32),
        ("dmLogPixels", ctypes.c_short),
        ("dmBitsPerPel", ctypes.c_ulong),
        ("dmPelsWidth", ctypes.c_ulong),
        ("dmPelsHeight", ctypes.c_ulong),
        ("dmDisplayFlags", ctypes.c_ulong),
        ("dmDisplayFrequency", ctypes.c_ulong),
        ("dmICMMethod", ctypes.c_ulong),
        ("dmICMIntent", ctypes.c_ulong),
        ("dmMediaType", ctypes.c_ulong),
        ("dmDitherType", ctypes.c_ulong),
        ("dmReserved1", ctypes.c_ulong),
        ("dmReserved2", ctypes.c_ulong),
        ("dmPanningWidth", ctypes.c_ulong),
        ("dmPanningHeight", ctypes.c_ulong)
    ]

class DISPLAY_DEVICE(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.c_ulong),
        ("DeviceName", ctypes.c_wchar * 32),
        ("DeviceString", ctypes.c_wchar * 128),
        ("StateFlags", ctypes.c_ulong),
        ("DeviceID", ctypes.c_wchar * 128),
        ("DeviceKey", ctypes.c_wchar * 128)
    ]

class RefreshRateMonitor:
    def __init__(self):
        print("Initializing RefreshRateMonitor...")
        self.refresh_rates = {}
        self.preferences = self.load_preferences()
        self.icon = None

        # Start the periodic check in a separate thread
        self.start_periodic_check()
        print("Initialized RefreshRateMonitor.")

    def load_preferences(self):
        print("Loading preferences...")
        default_preferences = {"alert_threshold": 60, "alert_sound": True}
        if not os.path.exists("preferences.json"):
            self.save_preferences(default_preferences)
            print("Created default preferences file.")
            return default_preferences
        try:
            with open("preferences.json", "r") as file:
                print("Preferences loaded.")
                return json.load(file)
        except json.JSONDecodeError:
            self.save_preferences(default_preferences)
            print("Loaded default preferences due to JSON decode error.")
            return default_preferences

    def save_preferences(self, preferences=None):
        print("Saving preferences...")
        if preferences is None:
            preferences = self.preferences
        with open("preferences.json", "w") as file:
            json.dump(preferences, file)
        print("Preferences saved.")

    def get_display_devices(self):
        print("Getting display devices...")
        devices = []
        i = 0
        while True:
            device = DISPLAY_DEVICE()
            device.cb = ctypes.sizeof(device)
            if not ctypes.windll.user32.EnumDisplayDevicesW(None, i, ctypes.byref(device), 0):
                break
            if device.StateFlags & DISPLAY_DEVICE_ACTIVE:
                devices.append(device)
            i += 1
        print(f"Found {len(devices)} active display devices.")
        return devices

    def get_refresh_rate(self, device_name):
        devmode = DEVMODE()
        devmode.dmSize = ctypes.sizeof(devmode)
        if ctypes.windll.user32.EnumDisplaySettingsW(device_name, ENUM_CURRENT_SETTINGS, ctypes.byref(devmode)):
            return devmode.dmDisplayFrequency
        return None

    def get_available_refresh_rates(self, device_name):
        print(f"Getting available refresh rates for {device_name}...")
        rates = set()
        devmode = DEVMODE()
        devmode.dmSize = ctypes.sizeof(devmode)
        i = 0
        while ctypes.windll.user32.EnumDisplaySettingsW(device_name, i, ctypes.byref(devmode)):
            rates.add(devmode.dmDisplayFrequency)
            i += 1
        return sorted(rates)

    def check_refresh_rates(self):
        print("Checking refresh rates...")
        devices = self.get_display_devices()
        for device in devices:
            rate = self.get_refresh_rate(device.DeviceName)
            if rate:
                self.refresh_rates[device.DeviceName] = rate
        print(f"Current refresh rates: {self.refresh_rates}")
        self.check_alerts()

    def check_alerts(self):
        print("Checking alerts...")
        for device_name, rate in self.refresh_rates.items():
            if rate != self.preferences.get(device_name, 60):  # Check against individual preferences
                self.show_alert(device_name, rate)

    def show_alert(self, device_name, rate):
        print(f"Showing alert for {device_name} with rate {rate}Hz")
        if self.preferences["alert_sound"]:
            winsound.Beep(1000, 500)  # Frequency 1000 Hz, duration 500 ms
        messagebox.showwarning("Low Refresh Rate",
                               f"The refresh rate of {device_name} is {rate}Hz, which is below the threshold.")

    def show_settings(self):
        def save_settings():
            for device_name in self.refresh_rates:
                self.preferences[device_name] = int(preferred_rates[device_name].get())
            self.save_preferences()
            settings_window.destroy()

        print("Showing settings window...")
        settings_window = tk.Tk()
        settings_window.title("Settings")

        row = 0
        preferred_rates = {}
        monitor_names = {m.name: m.name for m in get_monitors()}
        for device_name, rate in self.refresh_rates.items():
            monitor_name = monitor_names.get(device_name, device_name)
            ttk.Label(settings_window, text=f"{monitor_name}: Current rate {rate}Hz").grid(row=row, column=0, padx=10, pady=10)
            available_rates = self.get_available_refresh_rates(device_name)
            preferred_rates[device_name] = ttk.Combobox(settings_window, values=available_rates)
            preferred_rates[device_name].grid(row=row, column=1, padx=10, pady=10)
            preferred_rates[device_name].set(self.preferences.get(device_name, available_rates[0]))
            row += 1

        ttk.Button(settings_window, text="Save", command=save_settings).grid(row=row, column=0, columnspan=2, pady=10)

        settings_window.mainloop()
        print("Settings window closed.")

    def exit_app(self, icon, item):
        print("Exiting application...")
        self.icon.stop()
        self.save_preferences()
        sys.exit()

    def create_tray_icon(self):
        print("Creating tray icon...")
        icon_path = "rfm.ico"
        if not os.path.exists(icon_path):
            icon_image = Image.new('RGBA', (64, 64), (255, 255, 255, 0))
            draw = ImageDraw.Draw(icon_image)
            draw.rectangle((0, 0, 63, 63), outline=(0, 0, 0, 255))
            draw.text((10, 20), "RFM", fill=(0, 0, 0, 255))
            icon_image.save(icon_path)

        menu = (
            item('Settings', self.show_settings),
            item('Manual Check', self.manual_check),
            item('Exit', self.exit_app)
        )
        icon_image = Image.open(icon_path)
        self.icon = pystray.Icon("RefreshRateMonitor", icon_image, "Refresh Rate Monitor", menu)
        print("Running tray icon...")
        self.icon.run()
        print("Tray icon created.")

    def set_startup(self):
        print("Setting up startup...")
        startup_path = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        shortcut_path = os.path.join(startup_path, 'RefreshRateMonitor.lnk')
        target = os.path.abspath(sys.argv[0])
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = os.path.dirname(target)
        shortcut.IconLocation = target
        shortcut.save()
        print("Startup setup complete.")

    def restart_program(self):
        print("Restarting program...")
        self.icon.stop()
        os.execv(sys.executable, ['python'] + sys.argv)

    def manual_check(self, icon, item):
        print("Performing manual check...")
        self.check_refresh_rates()

    def start_periodic_check(self):
        def periodic_check():
            while True:
                print("Periodic check: Checking refresh rates...")
                self.check_refresh_rates()
                time.sleep(60)
        threading.Thread(target=periodic_check, daemon=True).start()
        print("Started periodic check thread.")

def main():
    print("Starting Refresh Rate Monitor...")
    monitor = RefreshRateMonitor()
    monitor.set_startup()
    monitor.create_tray_icon()
    print("Refresh Rate Monitor started.")

if __name__ == "__main__":
    main()
