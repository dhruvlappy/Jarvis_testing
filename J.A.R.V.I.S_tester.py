# ==========================================
# J.A.R.V.I.S. V3.2 ULTIMATE MONOLITH
# ==========================================
# Optimized, Portable, and Wireless-First Assistant.
# ==========================================

import os
import time
import datetime
import threading
import logging
import subprocess
import re
import ctypes
import smtplib
import shutil
import json
import webbrowser
import winreg
import concurrent.futures
import xml.etree.ElementTree as ET
import csv
from dotenv import load_dotenv

import cv2
import face_recognition
import numpy as np
import pyttsx3
import winsound
import speech_recognition as sr
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from cvzone.HandTrackingModule import HandDetector
import pywhatkit
import pyautogui
import pygetwindow as gw
import keyboard
import screen_brightness_control as sbc
import psutil
import requests
import wikipedia
from git import Repo

import urllib.parse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

# Advanced Health dependencies
try:
    import wmi
    import gputil
except ImportError:
    wmi = None
    gputil = None

# OCR support
try:
    import pytesseract

    TESS_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(TESS_PATH):
        pytesseract.pytesseract.tesseract_cmd = TESS_PATH
except ImportError:
    pytesseract = None

# ==========================================
# 1. CORE CONFIGURATION & LOGGING
# ==========================================

jarvis_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'jarvis.env')
load_dotenv(jarvis_env_path)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - JARVIS: %(message)s',
    handlers=[
        logging.FileHandler("jarvis.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("JARVIS")

# Global Settings
OPERA_PATH = os.getenv("OPERA_PATH")
SPOTIFY_EXE_PATH = os.getenv("SPOTIFY_EXE_PATH")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD")
SENTRY_CONTACT_NUMBER = os.getenv("SENTRY_CONTACT_NUMBER")

# Thread-safe flags
is_processing = threading.Event()
is_call_active = threading.Event()
is_muted = threading.Event()

# Legacy Handlers & Assets
contacts = {
    'father': {'email': 'ujha78@gmail.com', 'number': None},
    'mother': {'email': 'jhaaarti83@gmail.com', 'number': None},
    'didi': {'email': 'shailtinu2000@gmail.com', 'number': None},
    'myself': {'email': 'dhruvajha06@gmail.com', 'number': None},
    'siddhi': {'email': 'ao2025.siddhi.parmar@ves.ac.in', 'number': None},
    'riddhi': {'email': '', 'number': None},
    'radhu': {'email': '', 'number': None}
}
labeled_contacts = ['father', 'mother', 'didi', 'myself', 'siddhi', 'riddhi', 'radhu']


def search_memory_core(target_name, filename="contacts.csv"):
    target_name = target_name.lower()
    filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    try:
        with open(filepath, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            for row in reader:
                row_data = " ".join(row).lower()
                if target_name in row_data:
                    numbers = re.findall(r'\+?\d[\d\-\s]{8,14}\d', row_data)
                    if numbers:
                        clean_num = re.sub(r"\D", "", numbers[0])
                        if len(clean_num) >= 10: return clean_num[-10:]
        return None
    except:
        return False


def spotify_search_and_play(target_song=""):
    time.sleep(4)
    spotify_wins = gw.getWindowsWithTitle('Spotify')
    if spotify_wins:
        try:
            spotify_wins[0].activate();
            spotify_wins[0].maximize()
        except:
            pass
    if not target_song:
        keyboard.send("play/pause media");
        return
    try:
        pyautogui.hotkey('ctrl', 'l');
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a');
        pyautogui.press('backspace')
        pyautogui.write(target_song, interval=0.05);
        time.sleep(1.5)
        pyautogui.press('enter');
        time.sleep(1);
        pyautogui.press('enter')
    except:
        pass


def wish_user():
    current_time = time.strftime("%I:%M %p")
    try:
        os.startfile(SPOTIFY_EXE_PATH)
    except:
        pass
    threading.Thread(target=spotify_search_and_play, args=("",), daemon=True).start()
    speak(f"Welcome back sir. It's {current_time}")


def send_email(to_email, content):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, f"Subject: JARVIS Message\n\n{content}")
        server.close()
        speak("Email sent successfully.")
    except Exception as e:
        speak("Sir, I was unable to send the email.")


# ==========================================
# 2. SPEECH ENGINE & AUDIO UTILS
# ==========================================

speech_lock = threading.Lock()


def speak(audio):
    """Speaks the audio string via TTS, using a lock for thread safety."""
    with speech_lock:
        is_processing.set()
        logger.info(f"Speaking: {audio}")
        try:
            engine = pyttsx3.init('sapi5')
            voices = engine.getProperty('voices')
            if voices:
                engine.setProperty('voice', voices[0].id)
            engine.say(str(audio))
            engine.runAndWait()
            engine.stop()
        except Exception as e:
            logger.error(f"TTS Error: {e}")
        finally:
            is_processing.clear()


def set_system_volumes(level: float):
    """Uses pycaw to change the master volume for all applications EXCEPT this script."""
    try:
        sessions = AudioUtilities.GetAllSessions()
        current_pid = os.getpid()
        for session in sessions:
            if session.Process:
                volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                if session.Process.pid == current_pid or "python" in session.Process.name().lower():
                    volume.SetMasterVolume(1.0, None)
                else:
                    volume.SetMasterVolume(level, None)
    except Exception as e:
        logger.error("Failed to modify system volumes", exc_info=True)


def play_beep(frequency, duration=200):
    try:
        winsound.Beep(frequency, duration)
    except:
        pass


def take_command():
    """Listens to microphone and blocks until command is interpreted."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        logger.info("Listening for response...")
        play_beep(1000)
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=0.5)
        try:
            audio = r.listen(source, timeout=5)
        except sr.WaitTimeoutError:
            return None

    try:
        play_beep(400)
        query = r.recognize_google(audio, language='en-US')
        logger.info(f"User said: {query}")
        return query.lower()
    except:
        return None


# ==========================================
# 3. WIRELESS ADB & MOBILE ENGINE
# ==========================================

DEVICE_ID = "DISCONNECTED"


def scan_network_for_adb_device(silent=False):
    """Scans local mDNS broadcast data and checks existing device links for a zero-input wireless handshake."""
    global DEVICE_ID
    try:
        res = subprocess.run("adb devices", capture_output=True, text=True, shell=True)
        for line in res.stdout.split("\n"):
            if "\tdevice" in line:
                match = re.search(r'(\d{1,3}(?:\.\d{1,3}){3}:\d{3,5})', line)
                if match:
                    DEVICE_ID = match.group(1)
                    if not silent: logger.info(f"Stealth-link active: {DEVICE_ID}")
                    return DEVICE_ID

        subprocess.run("adb mdns services", capture_output=True, text=True, shell=True, timeout=5)
        time.sleep(1)

        try:
            result = subprocess.run("adb mdns services", capture_output=True, text=True, shell=True, timeout=5)
            for line in result.stdout.split("\n"):
                if "_adb-tls-connect._tcp" in line or "_adb._tcp" in line:
                    match = re.search(r'(\d{1,3}(?:\.\d{1,3}){3}:\d{3,5})', line)
                    if match:
                        active_device = match.group(1)
                        if not silent: logger.info(f"Handshake Initialized via mDNS: {active_device}")
                        subprocess.run(f"adb connect {active_device}", shell=True, capture_output=True, timeout=5)
                        DEVICE_ID = active_device
                        return DEVICE_ID
        except subprocess.TimeoutExpired:
            if not silent: logger.warning("mDNS Discovery timed out. Phone likely offline.")
    except Exception as e:
        if not silent: logger.error(f"Airspace scan failed: {e}")
    return None


scan_network_for_adb_device()


def run_adb(command):
    """Executes an ADB command with built-in auto-discovery routing and self-healing."""
    global DEVICE_ID
    full_cmd = f"adb -s {DEVICE_ID} {command}"
    try:
        result = subprocess.run(full_cmd, shell=True, capture_output=True, text=True, encoding='utf-8', errors='ignore',
                                timeout=5)
        if result.returncode != 0 and (
                "not found" in result.stderr or "offline" in result.stderr or "DISCONNECTED" in full_cmd):
            if DEVICE_ID != "DISCONNECTED":
                logger.warning(f"Link severed at {DEVICE_ID}. Initiating stealth re-discovery...")
            new_id = scan_network_for_adb_device(silent=True)
            if new_id:
                full_cmd = f"adb -s {new_id} {command}"
                result = subprocess.run(full_cmd, shell=True, capture_output=True, text=True, encoding='utf-8',
                                        errors='ignore', timeout=5)
                logger.info(f"Stealth-link restored to {new_id} flawlessly.")
            else:
                DEVICE_ID = "DISCONNECTED"
        return result
    except:
        return subprocess.CompletedProcess(args=full_cmd, returncode=1, stdout="", stderr="Link timeout.")


# ---------- New Helper Functions ----------

def safe_set_brightness(level: int):
    """Set screen brightness safely, clamping between 0‑100% and logging any errors."""
    try:
        level = max(0, min(100, level))
        sbc.set_brightness(level)
        logger.info(f"Brightness set to {level}%")
    except Exception as e:
        logger.error(f"Failed to set brightness: {e}")


def adb_search_and_call(contact_name: str):
    """Dial a phone number via ADB using the stored contacts dictionary.
    Looks up `contacts` for the given `contact_name` and issues an ADB CALL intent.
    """
    try:
        # contacts dict is defined elsewhere in the script
        if contact_name in contacts:
            number = contacts[contact_name].get('phone')
            if number:
                run_adb(f"shell am start -a android.intent.action.CALL -d tel:{number}")
                speak(f"Calling {contact_name}.")
                return True
        speak(f"Contact {contact_name} not found.")
        return False
    except Exception as e:
        logger.error(f"adb_search_and_call error: {e}")
        speak("Failed to place the call.")
        return False


def force_speaker():
    import xml.etree.ElementTree as ET
    run_adb("shell uiautomator dump /sdcard/view.xml")
    run_adb("pull /sdcard/view.xml .")
    try:
        if not os.path.exists('view.xml'): return False
        tree = ET.parse('view.xml')
        root = tree.getroot()
        for node in root.iter('node'):
            text = node.get('text', '')
            desc = node.get('content-desc', '')
            if "Speaker" in text or "Speaker" in desc:
                bounds = node.get('bounds')
                if bounds:
                    coords = re.findall(r'\d+', bounds)
                    x = (int(coords[0]) + int(coords[2])) // 2
                    y = (int(coords[1]) + int(coords[3])) // 2
                    run_adb(f"shell input tap {x} {y}")
                    return True
    except:
        pass
    return False


def monitor_call_state():
    global DEVICE_ID
    subprocess.run(f"adb -s {DEVICE_ID} logcat -c", shell=True)
    log_cmd = f"adb -s {DEVICE_ID} logcat -v brief -b radio -T 1"
    process = subprocess.Popen(log_cmd, shell=True, stdout=subprocess.PIPE, text=True, encoding='utf-8',
                               errors='ignore')
    try:
        for line in iter(process.stdout.readline, ''):
            if any(sig in line for sig in ["PRECISE_CALL_STATE_IDLE", "mCallState=0", "DISCONNECT"]):
                if is_call_active.is_set():
                    logger.info("Call Link Terminated.")
                    if is_muted.is_set():
                        run_adb("shell input keyevent 91")
                        is_muted.clear()
                    is_call_active.clear()
                    set_system_volumes(1.0)
                    keyboard.send("play/pause media")
                break
    finally:
        process.terminate()


def incoming_call_daemon():
    logger.info("Call Intercept Active.")
    while True:
        if is_call_active.is_set():
            time.sleep(2)
            continue
        try:
            registry = run_adb("shell dumpsys telephony.registry")
            if "mCallState=1" in registry.stdout:
                logger.info("LIVE INCOMING CALL DETECTED!")
                incoming_num = re.search(r'mCallIncomingNumber=(\d+)', registry.stdout)
                caller_name = "an unknown number"
                if incoming_num:
                    caller_name = reverse_search_memory_core(incoming_num.group(1))

                set_system_volumes(0.0)
                keyboard.send("play/pause media")
                speak(f"Sir, you have an incoming call from {caller_name}. Answer on speaker, or decline?")

                command = take_command()
                if command and any(word in command for word in ["answer", "pick up", "accept"]):
                    speak("Connecting link.")
                    run_adb("shell input keyevent 5")
                    is_call_active.set()
                    if "speaker" in command:
                        time.sleep(1)
                        force_speaker()
                    threading.Thread(target=monitor_call_state, daemon=True).start()
                elif command and any(word in command for word in ["decline", "reject"]):
                    speak("Call rejected.")
                    run_adb("shell input keyevent 6")
                    set_system_volumes(1.0)
                    keyboard.send("play/pause media")

                while "mCallState=1" in run_adb("shell dumpsys telephony.registry").stdout:
                    time.sleep(1)
        except:
            pass
        time.sleep(1.5)


def battery_guardian_daemon():
    logger.info("Battery Guardian Online.")
    warned_at_15 = warned_at_5 = False
    while True:
        try:
            battery_data = run_adb("shell dumpsys battery")
            if battery_data.stdout:
                level_match = re.search(r'level:\s+(\d+)', battery_data.stdout)
                status_match = re.search(r'status:\s+(\d+)', battery_data.stdout)
                if level_match and status_match:
                    level = int(level_match.group(1))
                    is_charging = (int(status_match.group(1)) in [2, 5])
                    if is_charging:
                        warned_at_15 = warned_at_5 = False
                    elif level <= 15 and level > 5 and not warned_at_15:
                        while is_processing.is_set(): time.sleep(2)
                        set_system_volumes(0.1)
                        keyboard.send("play/pause media")
                        speak(f"Sir, mobile power at {level} percent.")
                        keyboard.send("play/pause media")
                        set_system_volumes(1.0)
                        warned_at_15 = True
                    elif level <= 5 and not warned_at_5:
                        while is_processing.is_set(): time.sleep(2)
                        set_system_volumes(0.1)
                        keyboard.send("play/pause media")
                        speak(f"Emergency. Mobile power at {level} percent.")
                        keyboard.send("play/pause media")
                        set_system_volumes(1.0)
                        warned_at_5 = True
        except:
            pass
        time.sleep(60)


# ==========================================
# 4. SECURITY & VISION SENSORS (SENTRY)
# ==========================================

ADMIN_ENCODINGS = []
admin_images = ['dhruv_front.jpg', 'dhruv_left.jpg', 'dhruv_right.jpg']

for img_name in admin_images:
    p = os.path.join(os.getcwd(), img_name)
    if os.path.exists(p):
        try:
            img = face_recognition.load_image_file(p)
            enc = face_recognition.face_encodings(img)
            if enc:
                ADMIN_ENCODINGS.append(enc[0])
                logger.info(f"Loaded Sentry encoding for {img_name}")
        except:
            pass


def normalize_lighting(img):
    lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    cl = clahe.apply(l)
    return cv2.cvtColor(cv2.merge((cl, a, b)), cv2.COLOR_LAB2BGR)


def analyze_frame_for_sentry(img):
    master_present = intruder_present = False
    norm_img = normalize_lighting(img)
    small_frame = cv2.resize(norm_img, (0, 0), fx=0.25, fy=0.25)
    rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
    face_locations = face_recognition.face_locations(rgb_small_frame)
    if not face_locations: return False, False
    intruder_present = True
    if ADMIN_ENCODINGS:
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
        for face_encoding in face_encodings:
            results = face_recognition.compare_faces(ADMIN_ENCODINGS, face_encoding, tolerance=0.50)
            if True in results:
                master_present = True;
                intruder_present = False;
                break
    return master_present, intruder_present


def sentry_ghost_dispatch(message):
    os.system("taskkill /f /im chrome.exe /t >nul 2>&1")
    os.system("taskkill /f /im chromedriver.exe /t >nul 2>&1")
    time.sleep(1)

    opt = Options()
    opt.add_argument(r"--user-data-dir=C:\Jarvis_Final_Session")
    opt.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
    opt.add_argument("--headless=new")
    opt.add_argument("--window-size=1920,1080")
    opt.add_argument("--no-sandbox")

    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opt)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        driver.get(f"https://web.whatsapp.com/send?phone=919511069126&text={urllib.parse.quote(message)}")
        wait = WebDriverWait(driver, 50)

        xpath = '//div[@contenteditable="true"][@data-tab="10"] | //div[@title="Type a message"] | //div[@title="Type a message"]//p'
        chat_box = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))

        chat_box.click()
        time.sleep(2)
        chat_box.send_keys(Keys.ENTER)

        logger.info(f"Intruder Message injected. Waiting for server confirmation...")
        wait.until(EC.presence_of_element_located(
            (By.XPATH, '//span[@data-icon="msg-check"] | //span[@data-icon="msg-dblcheck"]')))
        logger.info(f"SECURITY DISPATCH CONFIRMED: {time.strftime('%H:%M:%S')}")

        time.sleep(2)
        driver.quit()

    except Exception as e:
        logger.error(f"Sentry Dispatch failed: {e}")
        try:
            driver.quit()
        except:
            pass


# --- Section 5: INTELLIGENCE & PRODUCTIVITY ---

def send_invisible_whatsapp(contact_name, message, is_first_run=False):
    target_number = search_memory_core(contact_name)
    if not target_number: return
    chrome_options = Options()
    chrome_options.add_argument(r"--user-data-dir=C:\Jarvis_Chrome_Data")
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    chrome_options.add_argument(f"user-agent={user_agent}")
    if not is_first_run:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--log-level=3")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    try:
        url = f"https://web.whatsapp.com/send?phone=91{target_number}&text={urllib.parse.quote(message)}"
        driver.get(url)
        wait = WebDriverWait(driver, 65)
        wait.until(EC.invisibility_of_element_located((By.XPATH, '//div[@id="startup"]')))
        text_box = wait.until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
        text_box.send_keys(Keys.SPACE + Keys.BACKSPACE + Keys.ENTER)
        time.sleep(5)
        logger.info(f"Ghost Transmission Confirmed.")
    except Exception as e:
        driver.save_screenshot("headless_error_vision.png")
    finally:
        driver.quit()


def reverse_search_memory_core(number):
    """Scrapes Android contacts via ADB to identify the caller."""
    logger.info(f"Reverse searching number: {number}")
    cmd = f"shell content query --uri content://com.android.contacts/data --projection display_name --where \"mimetype='vnd.android.cursor.item/phone_v2' AND data1 LIKE '%{number}%'\""
    result = run_adb(cmd)
    if result.stdout:
        match = re.search(r'display_name=(.*?)$', result.stdout, re.MULTILINE)
        if match: return match.group(1).strip()
    return "Unknown"


def get_advanced_health_score():
    speak("Analyzing system vitals.")
    score = 100;
    report = []
    cpu_usage = psutil.cpu_percent(interval=1)
    if cpu_usage > 85: score -= 20
    if wmi:
        try:
            w = wmi.WMI(namespace="root\\wmi")
            temp = w.MSAcpi_ThermalZoneTemperature()[0].CurrentTemperature
            score -= 15 if (temp / 10.0 - 273.15) > 80 else 0
        except:
            pass
    mem = psutil.virtual_memory()
    if mem.percent > 90: score -= 15
    speak(f"System Health Score is {score} out of 100.")
    logger.info(f"Health Check: {score}%")


def ollama_chat(prompt):
    try:
        response = requests.post("http://localhost:11434/api/generate",
                                 json={"model": "llama3", "prompt": prompt, "stream": False}, timeout=15)
        if response.status_code == 200:
            speak(response.json().get('response', ''))
        else:
            speak("Local AI offline.")
    except:
        speak("AI Bridge failure.")


def ai_code_reviewer():
    speak("Locating recent code for review.")
    try:
        files = [f for f in os.listdir('.') if f.endswith('.py')]
        if files:
            latest = max(files, key=os.path.getmtime)
            with open(latest, 'r') as f: code = f.read()
            ollama_chat(f"Review this Python code briefly: \n\n{code[:1500]}")
    except:
        pass


def voice_calculator(query):
    if any(word in query for word in ["convert", "miles", "km"]):
        ollama_chat(query)
    else:
        try:
            expr = re.sub(r'[a-zA-Z]', '', query).strip()
            speak(f"The result is {eval(expr)}")
        except:
            ollama_chat(query)


def organize_downloads():
    path = os.path.expanduser("~/Downloads")
    if not os.path.exists(path): return
    speak("Organizing downloads.")
    ext_map = {"Images": [".jpg", ".png"], "Docs": [".pdf", ".docx"], "Apps": [".exe", ".msi"]}
    for f in os.listdir(path):
        _, ext = os.path.splitext(f)
        for folder, exts in ext_map.items():
            if ext.lower() in exts:
                dest = os.path.join(path, folder);
                os.makedirs(dest, exist_ok=True)
                try:
                    shutil.move(os.path.join(path, f), os.path.join(dest, f))
                except:
                    pass


def toggle_windows_dark_mode():
    try:
        path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
        val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, 1 - val);
        winreg.CloseKey(key)
        speak("Theme toggled.")
    except:
        pass


def lockdown_mode():
    speak("Lockdown engaged.");
    ctypes.windll.user32.LockWorkStation()


# --- Section 6: CORE ROUTER ---

def conversation_flow(query):
    query = query.lower()

    if any(word in query for word in ["wake up", "briefing"]):
        wish_user()
        get_advanced_health_score()
    elif "play" in query:
        song_name = query.replace("play", "").strip()
        speak(f"Playing {song_name} on Spotify")
        threading.Thread(target=spotify_search_and_play, args=(song_name,), daemon=True).start()
    elif any(word in query for word in ["pc", "system", "laptop"]):
        if "shutdown" in query:
            os.system("shutdown /s /t 1")
        elif "lock" in query:
            lockdown_mode()
    elif 'wikipedia' in query:
        speak('Searching Wikipedia...')
        query = query.replace("wikipedia", "")
        results = wikipedia.summary(query, sentences=2)
        speak("According to Wikipedia");
        speak(results)
    elif "brightness" in query:
        if 'max' in query:
            sbc.set_brightness(100)
        elif 'half' in query:
            sbc.set_brightness(50)
    elif any(word in query for word in ["resume", "pause", "stop"]):
        keyboard.send("play/pause media");
        speak("Done.")
    elif any(word in query for word in ["next", "forward"]):
        keyboard.send("ctrl+right")
    elif any(word in query for word in ["last", "previous"]):
        keyboard.send("ctrl+left")
    elif "whatsapp" in query or "message" in query:
        speak("To whom, sir?")
        target = take_command()
        if target:
            speak("What is the message?")
            msg = take_command()
            if msg:
                speak("Initiating ghost protocol in background.")
                threading.Thread(target=send_invisible_whatsapp, args=(target, msg, False), daemon=True).start()
    elif any(word in query for word in ["call", "dial", "phone"]):
        target = None
        for label in labeled_contacts:
            if label in query: target = label; break
        if not target:
            speak("Who should I call, sir?")
            target = take_command()
        if target:
            adb_search_and_call(target)
            if "speaker" in query: force_speaker()
    elif 'open youtube' in query:
        speak("Opening YouTube.")
        subprocess.Popen([OPERA_PATH, "https://www.youtube.com"])
        speak("Do you have a specific search in mind, sir?")
        search_query = take_command()
        if search_query:
            for word in ["search", "for", "about", "yeah"]: search_query = search_query.replace(word, "").strip()
            opera_win = gw.getWindowsWithTitle("Opera")
            if opera_win:
                opera_win[0].activate();
                opera_win[0].maximize();
                time.sleep(1)
            pyautogui.click(x=1003, y=100);
            keyboard.press_and_release('ctrl+a')
            keyboard.press_and_release('delete');
            pyautogui.write(search_query, interval=0.005)
            keyboard.press_and_release('enter');
            time.sleep(5)
            # OCR Vision Extraction
            if pytesseract:
                speak("Analyzing results for the most popular video.")
                search_region = (400, 150, 1200, 850)
                screenshot = pyautogui.screenshot(region=search_region)
                img = np.array(screenshot);
                gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
                _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
                data = pytesseract.image_to_data(thresh, output_type=pytesseract.Output.DICT)
                max_views = -1;
                target_x, target_y = 700, 400

                def parse_views(text):
                    clean = "".join(c for c in text.lower() if c.isdigit() or c in ['k', 'm', '.'])
                    try:
                        if 'm' in clean: return float(clean.replace('m', '')) * 1_000_000
                        if 'k' in clean: return float(clean.replace('k', '')) * 1_000
                        return float(clean)
                    except:
                        return 0

                for i, text in enumerate(data['text']):
                    if 'view' in text.lower():
                        val = parse_views(data['text'][i])
                        if val == 0 and i > 0: val = parse_views(data['text'][i - 1])
                        if val > max_views:
                            max_views = val
                            target_x = search_region[0] + data['left'][i]
                            target_y = search_region[1] + data['top'][i] - 180
                if max_views > 0:
                    speak(f"Opening video with {int(max_views):,} views.");
                    pyautogui.click(x=target_x, y=target_y)
                else:
                    speak("Vision search inconclusive. Opening top result.");
                    pyautogui.click(x=700, y=450)
    elif 'open opera' in query:
        os.startfile(OPERA_PATH);
        speak("Opening Opera GX")
    elif 'open spotify' in query:
        os.startfile(SPOTIFY_EXE_PATH);
        speak("Opening Spotify")
    elif 'close' in query:
        if 'youtube' in query:
            keyboard.press_and_release('ctrl+w')
        elif 'spotify' in query:
            os.system("taskkill /F /IM Spotify.exe")
        elif 'opera' in query or 'browser' in query:
            os.system("taskkill /F /IM opera.exe")
        elif 'window' in query:
            keyboard.press_and_release('alt+f4')
        speak("Closed.")
    elif 'switch windows' in query:
        speak("Switching windows");
        keyboard.press_and_release('alt+tab')
    elif 'the time' in query:
        speak(f"Sir, the time is {datetime.datetime.now().strftime('%I:%M %p')}")
    elif 'send a mail' in query or 'send email' in query:
        recipient_email, recipient_name = None, None
        for name in labeled_contacts:
            if name in query:
                recipient_email, recipient_name = contacts[name]['email'], name;
                break
        if not recipient_email:
            speak("To whom shall I send the mail, sir?")
            target = take_command()
            if target and target in contacts:
                recipient_email, recipient_name = contacts[target]['email'], target
        if recipient_email:
            speak(f"What is the message for {recipient_name}?")
            content = take_command()
            if content: send_email(recipient_email, content)
    elif any(word in query for word in ["ask", "ai", "ollama", "is", "what", "who"]):
        ollama_chat(query)
    elif "health" in query:
        get_advanced_health_score()
    elif "lockdown" in query:
        lockdown_mode()
    elif "dark mode" in query:
        toggle_windows_dark_mode()


def continuous_listener():
    r = sr.Recognizer();
    r.dynamic_energy_threshold = True
    wake_words = ["jarvis", "jervis", "travis", "garbage", "chavez", "java", "service", "nervous", "charvis"]
    with sr.Microphone() as source:
        logger.info("--- JARVIS: Awareness Online ---")
        speak("System online.")
        r.adjust_for_ambient_noise(source, duration=1)
        while True:
            if is_processing.is_set(): time.sleep(0.1); continue
            try:
                audio = r.listen(source, timeout=None, phrase_time_limit=4)
                query = r.recognize_google(audio, language='en-US').lower()
                logger.info(f"[Mic Captured]: {query}")
                detected = next((w for w in wake_words if w in query), None)
                if detected:
                    is_processing.set()
                    try:
                        cmd = query.split(detected)[-1].strip()
                        if cmd: play_beep(1000); conversation_flow(cmd)
                    except:
                        pass
                    finally:
                        is_processing.clear()
            except:
                pass


def run_sentry_loop():
    cap = cv2.VideoCapture(0)
    if not cap.isOpened(): return
    logger.info("--- Sentry Vision Online ---")
    detector = None
    try:
        detector = HandDetector(detectionCon=0.8, maxHands=1)
    except:
        pass

    alert_fired = False;
    frame_skip = 0;
    absence_timer = 0
    sentry_result = [False, False]
    sentry_busy = threading.Event()

    def _run_sentry_analysis(frame):
        try:
            sentry_result[0], sentry_result[1] = analyze_frame_for_sentry(frame)
        except:
            pass
        finally:
            sentry_busy.clear()

    while True:
        try:
            ret, img = cap.read()
            if not ret: break
            img = cv2.flip(img, 1)

            # Gesture Controls
            if detector:
                hands, img = detector.findHands(img, draw=False)
                if hands:
                    fingers = detector.fingersUp(hands[0])
                    if fingers == [0, 1, 1, 0, 0]:
                        keyboard.send("play/pause media")
                    elif fingers == [1, 1, 1, 1, 1]:
                        set_system_volumes(0.0)
                    elif fingers == [0, 0, 0, 0, 0]:
                        set_system_volumes(1.0)

            frame_skip += 1
            if frame_skip % 10 == 0 and not sentry_busy.is_set():
                sentry_busy.set()
                threading.Thread(target=_run_sentry_analysis, args=(img.copy(),), daemon=True).start()

            master_present, intruder_present = sentry_result
            if master_present:
                if absence_timer > 300: safe_set_brightness(100)
                absence_timer = 0
            else:
                absence_timer += 1
                if absence_timer > 300: safe_set_brightness(10)

            if intruder_present and not master_present and not alert_fired:
                logger.warning("INTRUDER ALERT!")
                threading.Thread(target=sentry_ghost_dispatch, args=("Intruder detected!",), daemon=True).start()
                alert_fired = True
            elif master_present:
                alert_fired = False

            if cv2.waitKey(1) & 0xFF == ord('q'): break
        except:
            pass


def main():
    threading.Thread(target=continuous_listener, daemon=True).start()
    threading.Thread(target=incoming_call_daemon, daemon=True).start()
    threading.Thread(target=battery_guardian_daemon, daemon=True).start()
    run_sentry_loop()


if __name__ == "__main__":
    main()
