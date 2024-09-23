import win32gui
import win32process
import psutil
import csv
import schedule
import time
import logging
from datetime import datetime
from pystray import Icon, Menu, MenuItem
from PIL import Image
import threading
import os

BASE_PATH = "captures"

current_csv_file = ""

# Set up logging
script_dir = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(script_dir, 'logs')
captures_dir = os.path.join(script_dir, 'captures')
os.makedirs(log_dir, exist_ok=True)
os.makedirs(captures_dir, exist_ok=True)
log_file = os.path.join(log_dir, 'window_capture.log')

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def get_window_info(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return {
            'hwnd': hwnd,
            'title': win32gui.GetWindowText(hwnd),
            'process': process.name(),
            'timestamp': datetime.now().isoformat()
        }
    except Exception as e:
        logging.error(f"Error getting window info for hwnd {hwnd}: {e}")
        return None

def capture_windows():
    windows = []
    def enum_windows(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            window_info = get_window_info(hwnd)
            if window_info:
                windows.append(window_info)
        return True
    win32gui.EnumWindows(enum_windows, None)
    
    # Sort windows by z-order (topmost first)
    windows.sort(key=lambda w: win32gui.GetWindowRect(w['hwnd'])[2], reverse=True)
    
    # Mark the foreground window
    foreground_hwnd = win32gui.GetForegroundWindow()
    for window in windows:
        window['is_active'] = (window['hwnd'] == foreground_hwnd)
    
    return windows

def save_to_csv(windows):
    global current_csv_file
    try:
        with open(current_csv_file, 'a', newline='', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=['id', 'timestamp', 'title', 'process', 'is_active', 'order'])
            if file.tell() == 0:
                writer.writeheader()
            current_timestamp = int(time.time())
            for order, window in enumerate(windows, 1):
                row = {
                    'id': current_timestamp,
                    'timestamp': window['timestamp'],
                    'title': window['title'],
                    'process': window['process'],
                    'is_active': window['is_active'],
                    'order': order
                }
                writer.writerow(row)
        logging.info(f"Saved {len(windows)} window entries to CSV")
    except Exception as e:
        logging.error(f"Error saving to CSV: {e}")

def capture_and_save():
    try:
        windows = capture_windows()
        save_to_csv(windows)
        logging.info("Capture and save completed successfully")
    except Exception as e:
        logging.error(f"Error in capture_and_save: {e}")

def run_scheduler():
    schedule.every(1).minutes.do(capture_and_save)
    while True:
        schedule.run_pending()
        time.sleep(1)

def on_quit(icon):
    logging.info("Application shutting down")
    icon.stop()

def main():
    global current_csv_file
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        current_csv_file = os.path.join(captures_dir, f"capture_{timestamp}.csv")
        
        logging.info("Application starting")
        logging.info(f"Captures will be saved to: {current_csv_file}")
        # Create system tray icon
        image = Image.new('RGB', (64, 64), color = (73, 109, 137))
        menu = Menu(MenuItem('Quit', on_quit))
        icon = Icon("Window Capture", image, "Window Capture", menu)

        # Start the scheduler in a separate thread
        scheduler_thread = threading.Thread(target=run_scheduler)
        scheduler_thread.daemon = True
        scheduler_thread.start()

        # Run the system tray icon
        icon.run()
    except Exception as e:
        logging.critical(f"Critical error in main function: {e}")

if __name__ == "__main__":
    main()