import keyboard
import pyautogui
import time
import json
import cv2  # For capturing webcam images
from PIL import ImageGrab  # For capturing screenshots
from win32gui import GetWindowText, GetForegroundWindow  # For active window detection
from win32com.client import Dispatch  # For minimizing active windows
import os
import pandas as pd  # For logging data to Excel

LOG_FILE = "flagged_content.log"  # File to store flagged content
KEYSTROKE_FILE = "keystrokes.log"  # File to store all keystrokes
KEYWORDS_FILE = "keywords.json"  # JSON file containing flagged keywords
SCREENSHOT_DIR = "screenshots"  # Directory to store screenshots
WEBCAM_DIR = "webcam_images"  # Directory to store webcam images
EXCEL_LOG_FILE = "flagged_content.xlsx"  # Excel file to store detailed logs

# Load Keywords from JSON File
def load_keywords():
    try:
        with open(KEYWORDS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
            # Retrieve the combined flagged words list from the JSON file
            combined_flagged_words = data.get("keywords", [])
            if combined_flagged_words:
                return combined_flagged_words
            else:
                print(f"Warning: No 'keywords' found in {KEYWORDS_FILE}.")
                return []
    except FileNotFoundError:
        print(f"Error: {KEYWORDS_FILE} not found. Please create the file.")
        return []
    except json.JSONDecodeError:
        print(f"Error: {KEYWORDS_FILE} contains invalid JSON.")
        return []

# Example of using the function
keywords = load_keywords()
print(keywords)  # Prints the loaded flagged keywords

# Show Popup Alert
def show_popup_alert(content):
    pyautogui.alert(
        text=f"Flagged content detected: '{content}'. Activity has been blocked.",
        title="Alert",
        button="OK"
    )

# Log Flagged Content
def log_flagged_content(content, timestamp, platform):
    # Log to text file
    with open(LOG_FILE, "a") as log_file:
        log_file.write(f"[{timestamp}] {platform} - {content}\n")
    print(f"Logged: {content}")

    # Log to Excel
    if not os.path.exists(EXCEL_LOG_FILE):
        df = pd.DataFrame(columns=["Timestamp", "Platform", "Flagged Content"])
        df.to_excel(EXCEL_LOG_FILE, index=False)

    df = pd.read_excel(EXCEL_LOG_FILE)
    new_entry = {"Timestamp": timestamp, "Platform": platform, "Flagged Content": content}
    df = df.append(new_entry, ignore_index=True)
    df.to_excel(EXCEL_LOG_FILE, index=False)
    print("Logged to Excel.")

# Check if the active window is a browser
def is_browser_active():
    browser_names = ["chrome", "firefox", "msedge", "opera", "brave", "safari"]
    active_window = GetWindowText(GetForegroundWindow()).lower()
    for browser in browser_names:
        if browser in active_window:
            return True, active_window
    return False, "Unknown"

# Log Keystrokes Locally
def log_keystroke(key, timestamp):
    with open(KEYSTROKE_FILE, "a") as keystroke_file:
        keystroke_file.write(f"[{timestamp}] {key}\n")

# Capture Screenshot
def capture_screenshot(timestamp):
    screenshot = ImageGrab.grab()
    screenshot.save(f"{SCREENSHOT_DIR}/screenshot_{timestamp}.png")
    print("Screenshot captured.")

# Capture Webcam Image
def capture_webcam_image(timestamp):
    webcam = cv2.VideoCapture(0)  # Access the default webcam
    ret, frame = webcam.read()
    if ret:
        cv2.imwrite(f"{WEBCAM_DIR}/webcam_{timestamp}.png", frame)
        print("Webcam image captured.")
    webcam.release()

# Minimize All Windows
def minimize_all_windows():
    try:
        print("Minimizing all windows...")
        # Press Windows + D to show desktop (minimizes all windows)
        pyautogui.hotkey("win", "d")
    except Exception as e:
        print(f"Error while minimizing windows: {e}")


# Block Activity
def block_activity():
    print("Blocking activity...")
    minimize_all_windows()  # Minimize all windows

# Monitor Keyboard Input
def monitor_keyboard():
    print("Monitoring keyboard for flagged content and logging keystrokes... (Press ESC to quit)")
    flagged_keywords = load_keywords()
    if not flagged_keywords:
        print("No keywords to monitor. Exiting.")
        return

    typed_text = ""
    while True:
        try:
            event = keyboard.read_event(suppress=False)
            if event.event_type == "down":  # Key pressed
                # Check if the active window is a browser
                browser_active, platform = is_browser_active()
                if not browser_active:
                    continue

                timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")

                # Log every keystroke
                log_keystroke(event.name, timestamp)

                # Track typed text
                if event.name == "space":
                    typed_text += " "
                elif event.name == "enter":
                    typed_text += "\n"
                elif event.name == "backspace":
                    typed_text = typed_text[:-1]
                else:
                    typed_text += event.name

                # Check for keywords
                for keyword in flagged_keywords:
                    if keyword.lower() in typed_text.lower():
                        print(f"Flagged content detected: {typed_text.strip()}")

                        # Trigger actions
                        show_popup_alert(typed_text.strip())
                        log_flagged_content(typed_text.strip(), timestamp, platform)
                        capture_screenshot(timestamp)
                        capture_webcam_image(timestamp)
                        block_activity()

                        # Clear text buffer after alert
                        typed_text = ""

                # Exit on ESC key
                if event.name == "esc":
                    print("Exiting monitoring...")
                    return
        except KeyboardInterrupt:
            print("Keyboard monitoring interrupted.")
            break

# Main Function
if __name__ == "__main__":
    # Create directories for storing screenshots and webcam images
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    os.makedirs(WEBCAM_DIR, exist_ok=True)

    try:
        monitor_keyboard()
    except Exception as e:
        print(f"An error occurred: {e}")