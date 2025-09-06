import pyttsx3 
import speech_recognition as sp
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import cv2 
import requests
import time
import pythoncom
import win32com.client as win32 
import pyautogui
import re
from win32com.client import Dispatch
import shutil
import zipfile
import ctypes
import subprocess
import random
import feedparser
import speech_recognition as sr
from gtts import gTTS
from tempfile import NamedTemporaryFile
from twilio.rest import Client
import threading
import psutil
import screen_brightness_control as sbc
import pyperclip
import PyPDF2
import docx
import threading
import sounddevice as sd
from scipy.io.wavfile import write
import numpy as np
from scipy.io import wavfile
import warnings
warnings.filterwarnings("ignore")
import pygame
from dotenv import load_dotenv
import urllib.parse
load_dotenv()
import webbrowser
import urllib.parse
from translate import Translator
import time

# Background-only passive listening loop
import pythoncom
pythoncom.CoInitialize()

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def ask_and_speak_multilingual():
    speak("Which language should I speak?")
    lang_input = takeCommand().lower()

    lang_map = {
        "english": "en",
        "hindi": "hi",
        "marathi": "mr",
        "french": "fr",
        "spanish": "es",
        "german": "de",
        "japanese": "ja",
        "chinese": "zh",
        "arabic": "ar",
        "korean": "ko"
    }

    lang_code = lang_map.get(lang_input)
    if not lang_code:
        speak("Sorry, that language is not supported yet.")
        return

    speak(f"What should I say in {lang_input}?")
    phrase = takeCommand()

    try:
        translator = Translator(to_lang=lang_code)
        translated = translator.translate(phrase)
        print(f"[Translated] {translated}")
        speak_lang(translated, lang=lang_code)
    except Exception as e:
        print(f"[Translation Error] {e}")
        speak("Sorry, I couldn't translate that.")

def speak_lang(text, lang='en'):
    """Speak the given text in the specified language using gTTS."""
    try:
        tts = gTTS(text=text, lang=lang)
        with NamedTemporaryFile(delete=True) as fp:
            tts.save(fp.name + ".mp3")
            pygame.mixer.init()
            pygame.mixer.music.load(fp.name + ".mp3")
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                continue
            pygame.mixer.music.unload()
    except Exception as e:
        speak("Sorry, I couldn't speak in that language.")
        print(f"Error in speak_lang: {e}")

def passive_listen_for_wake_word(wake_words=["hey vois", "hey voice", "hey voice assistant","hey boys"]):
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Passive listening for wake word...")
        r.adjust_for_ambient_noise(source, duration=1)
        try:
            audio = r.listen(source, timeout=None)
            query = r.recognize_google(audio).lower()
            print(f"Heard: {query}")
            for word in wake_words:
                if word in query:
                    return True
            return False
        except:
            return False

def record_voice(filename, duration=5):  
    fs = 44100
    print(f"Recording for {duration} seconds...")
    try:
        recording = sd.rec(int(duration * fs), samplerate=fs, channels=1, dtype='float32')
        sd.wait()
        write(filename, fs, (recording * 32767).astype(np.int16))
        return True
    except Exception as e:
        print(f"Recording error: {e}")
        return False
    
def match_voice(sample1, sample2, threshold=0.45):  
    """Voice matching with normalization and noise tolerance"""
    try:
        sr1, data1 = wavfile.read(sample1)
        sr2, data2 = wavfile.read(sample2)

        if data1.ndim > 1:
            data1 = data1.mean(axis=1)
        if data2.ndim > 1:
            data2 = data2.mean(axis=1)

        if sr1 != sr2:
            from scipy.signal import resample
            target_rate = min(sr1, sr2)
            if sr1 != target_rate:
                data1 = resample(data1, int(len(data1) * target_rate / sr1))
            if sr2 != target_rate:
                data2 = resample(data2, int(len(data2) * target_rate / sr2))

        min_len = min(len(data1), len(data2))
        data1 = data1[:min_len]
        data2 = data2[:min_len]

        def normalize(sig):
            sig = sig - np.mean(sig)
            sig = sig / (np.std(sig) + 1e-9)
            return sig / np.max(np.abs(sig) + 1e-9)  

        data1 = normalize(data1)
        data2 = normalize(data2)

        similarity = np.corrcoef(data1, data2)[0, 1]
        print(f"Similarity score: {similarity:.2f}")

        return similarity >= 0

    except Exception as e:
        print(f"Matching error: {e}")
        return False


def authenticate_user():
    print("\nVoice Authentication.Say some phrase to authenticate.")
    print("------------------------")

    if not os.path.exists("myvoice.wav"):
        speak("No voice found. Please register your voice.")
        record_voice("myvoice.wav", 5)  
        return True  

    speak("Please repeat your phrase to authenticate.")
    
    if record_voice("input_voice.wav", 5):
        if match_voice("myvoice.wav", "input_voice.wav"):
            print("Access granted.")
            return True 
        
        else:
            speak("Access denied. Please register your voice again.")
            record_voice("myvoice.wav", 5)  
            return authenticate_user()  

    return False

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour < 12:
        speak("Good Morning!")
    elif hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am VOIS AI. How may I assist you?")

def takeCommand():
    r = sp.Recognizer()
    with sp.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=1)  
        print("Listening...")
        r.pause_threshold = 0.8
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=7)
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
            return query.lower()
        except sp.UnknownValueError:
            print("Didn't catch that. Please repeat.")
            return takeCommand()
        except sp.RequestError:
            print("Recognition service unavailable.")
            return "None"
        except Exception as e:
            print(f"Error: {e}")
            return "None"

def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        command = recognizer.recognize_google(audio)
        print("You said:", command)
        return command.lower()
    except sr.UnknownValueError:
        print("Sorry, I couldn't understand.")
        return ""
    except sr.RequestError:
        print("Speech recognition service error.")
        return ""
    
def search_duckduckgo(query):
    url = f"https://api.duckduckgo.com/?q={query}&format=json"
    response = requests.get(url)
    data = response.json()
    
    answer = data.get("AbstractText")
    
    if answer:
        speak(answer)
        print("Answer:", answer)
    else:
        speak("Sorry, I couldn't find an answer.")
        print("No direct answer found.")

def get_system_status():
    cpu_usage = psutil.cpu_percent(interval=1)
    ram_usage = psutil.virtual_memory().percent
    battery = psutil.sensors_battery().percent if psutil.sensors_battery() else "N/A"
    return f"CPU: {cpu_usage}%, RAM: {ram_usage}%, Battery: {battery}%"

def control_brightness(level):
    try:
        sbc.set_brightness(level)
        speak(f"Brightness set to {level} percent.")
    except Exception as e:
        speak("Failed to adjust brightness.")

def validate_email(email):
    """Validate email format using regex."""
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w{2,}$'
    return re.match(pattern, email) is not None

def send_email():
    try:
        speak("To whom should I send the email?")
        recipient = takeCommand().replace(" at ", "@").replace(" dot ", ".").replace(" ", "")
        if not validate_email(recipient):
            speak("That doesn't seem like a valid email.")
            return

        speak("What is the subject?")
        subject = takeCommand()

        speak("What should I say in the email?")
        body = takeCommand()

        # Gmail direct compose link
        gmail_url = (
            f"https://mail.google.com/mail/?view=cm&fs=1"
            f"&to={recipient}"
            f"&su={urllib.parse.quote(subject)}"
            f"&body={urllib.parse.quote(body)}"
        )

        webbrowser.open(gmail_url)
        speak("Opening Gmail compose window with your message. Please press send.")

    except Exception as e:
        print(f"[Email Error] {e}")
        speak("Something went wrong while preparing the email.")

def openWordAndWrite():
    pythoncom.CoInitialize()
    speak("Opening Word")
    os.system("start winword")
    time.sleep(3)
    word = win32.Dispatch("Word.Application")
    doc = word.Documents.Add()
    speak("Start speaking. Say 'stop writing' to finish.")
    while True:
        text = takeCommand()
        if 'stop writing' in text:
            speak("Stopped writing.")
            break
        doc.Content.Text += text + " "
    doc.SaveAs(os.path.expanduser("~/Desktop/speech_doc.docx"))
    doc.Close()
    word.Quit()
    speak("Saved on Desktop.")

def copy_text(text):
    pyperclip.copy(text)
    print("Text copied to clipboard!")

def paste_text():
    pasted = pyperclip.paste()
    print("Pasted text:", pasted)
    return pasted

def takePicture():
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        speak("Camera not accessible")
        return
    speak("Say Cheese!")
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        cv2.imshow('Camera', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            cv2.imwrite("captured_image.jpg", frame)
            speak("Picture saved")
            break
    cap.release()
    cv2.destroyAllWindows()

def get_news():
    speak("Getting the latest news headlines...")
    feed_url = "https://timesofindia.indiatimes.com/rssfeedstopstories.cms"
    news_feed = feedparser.parse(feed_url)

    r = sr.Recognizer()

    for i, entry in enumerate(news_feed.entries[:10]):
        speak(f"News {i + 1}: {entry.title}")

        # Wait a moment before listening (to avoid catching own voice)
        time.sleep(0.5)

        print("Listening briefly for 'stop' command...")
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=0.5)
            try:
                audio = r.listen(source, timeout=2, phrase_time_limit=2)
                command = r.recognize_google(audio).lower()
                print(f"Heard: {command}")
                if 'stop' in command:
                    speak("News reading stopped.")
                    return
            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                continue
            except Exception as e:
                print(f"[News listen error] {e}")
                continue

    speak("That's all for the news.")

def takeScreenshot():
    try:
        path = os.path.join(os.path.expanduser("~"), "Desktop", "Sonic_Screenshots")
        os.makedirs(path, exist_ok=True)
        screenshot_path = os.path.join(path, f"screenshot_{int(time.time())}.png")
        screenshot = pyautogui.screenshot()
        screenshot.save(screenshot_path)
        print(f"Saved: {screenshot_path}")
    except Exception as e:
        speak("Failed to take screenshot.")
        print(f"Error: {e}")

def get_weather_wttr(city):
    try:
        url = f"https://wttr.in/{city}?format=3"
        res = requests.get(url)
        return res.text
    except:
        return "Couldn't fetch weather."
    
def systemControl(command):
    if command == 'shutdown':
        speak("Shutting down")
        os.system("shutdown /s /t 1")
    elif command == 'restart':
        speak("Restarting")
        os.system("shutdown /r /t 1")
    elif command == 'sleep':
        speak("Sleeping")
        os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def tellDateTime():
    now = datetime.datetime.now()
    date = now.strftime("%A, %d %B %Y")
    time_str = now.strftime("%I:%M %p")
    speak(f"Today is {date}, and time is {time_str}.")

def openCalculator():
    speak("Opening Calculator")
    os.system("calc")

def openSettings():
    speak("Opening Settings")
    pyautogui.hotkey('win', 'i')

def enableHotspot():
    speak("Opening Network Settings")
    try:
        subprocess.run(["start", "ms-settings:network"], shell=True)
    except Exception as e:
        speak("Sorry, I couldn't open the network settings.")

def closeWindow():
    speak("Closing the window.")
    pyautogui.hotkey('alt', 'f4') 

def openExcel():
    speak("Opening Excel")
    os.system("start excel")

def openPowerPoint():
    speak("Opening PowerPoint")
    os.system("start powerpnt")

def enable_wifi():
    speak("Opening Wi-Fi settings.")
    os.system("start ms-settings:network-wifi")

def disable_wifi():
    speak("Opening Wi-Fi settings to turn off Wi-Fi.")
    os.system("start ms-settings:network-wifi")

def enable_bluetooth():
    speak("Opening Bluetooth settings.")
    os.system("start ms-settings:bluetooth")

def disable_bluetooth():
    speak("Opening Bluetooth settings to turn off Bluetooth.")
    os.system("start ms-settings:bluetooth")

def increase_volume():
    for _ in range(5):
        pyautogui.press("volumeup")

def decrease_volume():
    for _ in range(5):
        pyautogui.press("volumedown")

def get_astrology_prediction(rashi):
    predictions = {
        "aries": "New beginnings await you.",
        "taurus": "Be patient today.",
        "gemini": "Opportunities are coming.",
        "cancer": "Focus on emotions.",
        "leo": "Lead with confidence.",
        "virgo": "Hard work pays off.",
        "libra": "Maintain balance.",
        "scorpio": "Trust instincts.",
        "sagittarius": "Adventure ahead.",
        "capricorn": "Stay disciplined.",
        "aquarius": "Be creative.",
        "pisces": "Listen to your heart."
    }
    return predictions.get(rashi.lower(), "Unknown Rashi.")

def set_alarm_time(alarm_time_str):
    try:
        now = datetime.datetime.now()
        alarm_time = datetime.datetime.strptime(alarm_time_str, "%H:%M")
        alarm_time = now.replace(hour=alarm_time.hour, minute=alarm_time.minute, second=0, microsecond=0)

        if alarm_time < now:
            alarm_time += datetime.timedelta(days=1)

        seconds_left = (alarm_time - now).total_seconds()

        def alarm_action():
            time.sleep(seconds_left)
            speak("Alarm ringing! Wake up!")

        threading.Thread(target=alarm_action).start()

        speak(f"Alarm set for {alarm_time.strftime('%I:%M %p')}")

    except ValueError:
        speak("Sorry, I didn't understand the time format. Please say something like '7:30'.")

def make_call():
    speak("Say the phone number.")
    phone_number = takeCommand().replace(" ", "")
    if phone_number.isdigit() and len(phone_number) >= 10:
        os.system(f"start tel:{phone_number}")
        speak(f"Calling {phone_number}")
    else:
        speak("Invalid number.")

def send_sms_alert():
    def register_emergency_number():
        speak("Please say the emergency phone number including country code.")
        number = takeCommand().replace(" ", "").replace("plus", "+")

        if number.startswith('+') and number[1:].isdigit():
            with open("emergency_contact.txt", "w") as f:
                f.write(number)
            speak(f"Emergency contact number {number} registered successfully.")
            return number
        else:
            speak("That doesn't seem like a valid phone number. Please try again later.")
            return None

    # Step 1: Check if number is registered
    if not os.path.exists("emergency_contact.txt"):
        speak("No emergency number found. Please register now.")
        to_number = register_emergency_number()
        if not to_number:
            return
    else:
        with open("emergency_contact.txt", "r") as f:
            to_number = f.read().strip()

    # Step 2: Load Twilio credentials
    account_sid = os.getenv("TWILIO_SID")
    auth_token = os.getenv("TWILIO_TOKEN")
    twilio_number = os.getenv("TWILIO_PHONE")

    if not account_sid or not auth_token or not twilio_number:
        speak("Twilio credentials are missing. Please check your setup.")
        return

    client = Client(account_sid, auth_token)

    # Step 3: Get user location
    try:
        response = requests.get('http://ip-api.com/json/').json()
        city = response.get('city')
        region = response.get('regionName')
        country = response.get('country')
        lat = response.get('lat')
        lon = response.get('lon')
        location_link = f"https://maps.google.com/?q={lat},{lon}"
        message_body = f"Emergency! Please check on me.\nLocation: {city}, {region}, {country}\nMap: {location_link}"
    except:
        message_body = "Emergency! Please check on me. Location unavailable."

    # Step 4: Send SMS
    try:
        message = client.messages.create(
            body=message_body,
            from_=twilio_number,
            to=to_number
        )
        speak("Emergency message sent successfully.")
        print("Message SID:", message.sid)
    except Exception as e:
        print("Error:", e)
        speak("Failed to send the emergency SMS. Let's try registering a different number.")
        to_number = register_emergency_number()
        if to_number:
            try:
                message = client.messages.create(
                    body=message_body,
                    from_=twilio_number,
                    to=to_number
                )
                speak("Emergency message sent successfully.")
                print("Message SID:", message.sid)
            except Exception as e:
                speak("Sorry, SMS still failed after retrying. Please try again later.")
                print("Final Error:", e)

reminders = []

def setReminder():
    speak("What should I remind you about?")
    task = takeCommand()
    speak("At what time?")
    time_str = takeCommand()
    
    try:
        reminder_time = datetime.datetime.strptime(time_str, "%H:%M").time()
        now = datetime.datetime.now()
        reminder_datetime = datetime.datetime.combine(now.date(), reminder_time)
        
        if reminder_datetime < now:
            reminder_datetime += datetime.timedelta(days=1)

        delay = (reminder_datetime - now).total_seconds()

        def remind():
            time.sleep(delay)
            speak(f"Reminder: {task}")

        threading.Thread(target=remind).start()
        speak(f"Reminder set for {task} at {reminder_time.strftime('%I:%M %p')}")
        reminders.append((task, reminder_time.strftime("%I:%M %p")))

    except ValueError:
        speak("Sorry, I couldn't understand the time format.")

def find_file(filename, search_dir):
    for root, dirs, files in os.walk(search_dir):
        if filename in files:
            return os.path.join(root, filename)
    return None

def open_file_voice():
    speak("Please say the name of the file you want to open.")
    query = takeCommand().lower()

    file_name = query.replace(" dot ", ".").replace(" ", "")
    print(f"[DEBUG] Looking for: {file_name}")
    
    home_dir = os.path.expanduser("~")
    file_path = find_file(file_name, home_dir)

    if file_path:
        os.startfile(file_path)
        speak(f"Opening {file_name}")
    else:
        speak("Sorry, I could not find the file.")

def calculate_expression(expression):
    try:
        expression = expression.replace("plus", "+").replace("minus", "-")
        expression = expression.replace("into", "").replace("multiplied by", "")
        expression = expression.replace("divided by", "/").replace("by", "/")
        expression = expression.replace("modulus", "%").replace("mod", "%")
        expression = expression.replace("power", "**").replace("raise to", "%")

        result = eval(expression)
        speak(f"The result is {result}")
    except:
        speak("Sorry, I couldn't calculate that. Please try again.")

def create_folder():
    speak("Please tell me the folder path where you want to create the new folder.")
    parent_path = takeCommand().lower()

    parent_path = parent_path.replace("sea/see/se", "c")
    parent_path = parent_path.replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    speak("What should be the name of the new folder?")
    folder_name = takeCommand().lower().replace(" ", "_")

    if not folder_name:
        speak("No folder name provided.")
        return

    full_path = os.path.join(parent_path, folder_name)

    try:
        os.makedirs(full_path, exist_ok=True)
        speak(f"Folder '{folder_name}' has been created successfully at {parent_path}.")
    except Exception as e:
        speak("I couldn't create the folder. Please check the path and try again.")


def create_file():
    speak("Please tell me the folder path where I should create the file.")
    folder_path = takeCommand().lower().replace(" ", "").replace("colon", ":").replace("slash", "\\")
    folder_path = folder_path.replace("sea", "c").replace("see", "c").replace("se", "c")
    folder_path = folder_path.replace("colon", ":").replace("slash", "\\").replace("backslash", "\\")
    speak("What should be the name of the file?")
    name = takeCommand().lower()

    if not name:
        speak("No name provided.")
        return

    speak("What should I write in the file?")
    content = takeCommand()

    try:
        full_path = os.path.join(folder_path, f"{name}.txt")
        with open(full_path, "w") as file:
            file.write(content)
        speak(f"File {name}.txt created successfully at {folder_path}.")
    except Exception as e:
        speak("I couldn't create the file. Please check the folder path and try again.")

def rename_folder():
    speak("Please tell me the path where the folder is located.")
    folder_path = takeCommand().lower()
    folder_path = folder_path.replace("sea", "c").replace("see", "c").replace("se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    speak("Tell me the name of the folder you want to rename.")
    old_name = takeCommand().lower().replace(" ", "")

    speak("What should be the new name of the folder?")
    new_name = takeCommand().lower().replace(" ", "")

    try:
        old_folder_path = os.path.join(folder_path, old_name)
        new_folder_path = os.path.join(folder_path, new_name)
        os.rename(old_folder_path, new_folder_path)
        speak(f"Folder {old_name} has been renamed to {new_name}.")
    except Exception as e:
        speak("I couldn't rename the folder. Please check the path and folder name.")

def delete_folder():
    speak("Please tell me the path where the folder is located.")
    base_path = takeCommand().lower()
    base_path = base_path.replace("sea", "c").replace("see", "c").replace("se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    speak("Tell me the name of the folder you want to delete.")
    folder_name = takeCommand().lower().replace(" ", "")

    folder_path = os.path.join(base_path, folder_name)

    try:
        shutil.rmtree(folder_path)
        speak(f"Folder {folder_name} has been deleted successfully.")
    except Exception as e:
        speak("I couldn't delete the folder. Please check the path and folder name.")

def rename_file():
    speak("Please tell me the folder path where the file is located.")
    folder_path = takeCommand().lower().replace(" ", "").replace("colon", ":").replace("slash", "\\")
    folder_path = folder_path.replace("sea", "c").replace("see", "c").replace("se", "c")
    folder_path = folder_path.replace("colon", ":").replace("slash", "\\").replace("backslash", "\\")
    speak("Tell me the current name of the file to rename.")
    old_name = takeCommand().lower()

    speak("What should be the new name?")
    new_name = takeCommand().lower()

    try:
        old_path = os.path.join(folder_path, f"{old_name}.txt")
        new_path = os.path.join(folder_path, f"{new_name}.txt")
        os.rename(old_path, new_path)
        speak(f"File renamed from {old_name} to {new_name}.")
    except Exception as e:
        speak("I couldn't rename the file. Please check the folder and file name.")

def delete_file():
    speak("Please tell me the folder path where the file is located.")
    folder_path = takeCommand().lower().replace(" ", "").replace("colon", ":").replace("slash", "\\")
    folder_path = folder_path.replace("sea", "c").replace("see", "c").replace("se", "c")
    folder_path = folder_path.replace("colon", ":").replace("slash", "\\").replace("backslash", "\\")
    speak("Tell me the name of the file you want to delete.")
    file_name = takeCommand().lower()

    try:
        full_path = os.path.join(folder_path, f"{file_name}.txt")
        os.remove(full_path)
        speak(f"File {file_name}.txt has been deleted from {folder_path}.")
    except:
        speak("I couldn't delete the file. Please check the file name and folder path.")

def read_file():
    speak("Please tell me the folder path of the file.")
    folder_path = takeCommand().lower()
    folder_path = folder_path.replace("sea", "c").replace("see", "c").replace("se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    speak("Now, tell me the file name with extension.")
    name = takeCommand().lower()

    full_path = os.path.join(folder_path, name)

    if name.endswith(".pdf"):
        try:
            with open(full_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                text = ""
                for page in reader.pages[:2]:  # Read first 2 pages
                    text += page.extract_text()
                speak(text if text else "The PDF seems to be empty or not readable.")
        except:
            speak("I couldn't read the PDF file.")
    elif name.endswith(".docx"):
        try:
            doc = docx.Document(full_path)
            text = "\n".join([para.text for para in doc.paragraphs])
            speak(text[:400] if text else "The document is empty.")
        except:
            speak("I couldn't read the Word file.")
    elif name.endswith(".txt"):
        try:
            with open(full_path, 'r') as f:
                content = f.read()
                speak(content[:400] if content else "The file is empty.")
        except:
            speak("I couldn't read the text file.")
    else:
        speak("Unsupported file format.")

def zip_folder():
    speak("Tell me the full path of the folder you want to zip.")
    folder_path = takeCommand().lower().replace("sea/see/se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")
    speak("Tell me the name of the zip file to create.")
    zip_name = takeCommand().lower().replace(" ", "") + ".zip"

    try:
        shutil.make_archive(zip_name.replace(".zip", ""), 'zip', folder_path)
        speak(f"Folder zipped successfully as {zip_name}.")
    except Exception as e:
        speak("I couldn't zip the folder. Please check the path and try again.")

def unzip_file():
    speak("Tell me the path of the zip file.")
    zip_path = takeCommand().lower().replace("sea/see/se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    speak("Tell me the folder path where to extract.")
    extract_path = takeCommand().lower().replace("sea/see/se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    try:
        if not os.path.exists(zip_path):
            speak("The zip file path does not exist.")
            return
        if not os.path.exists(extract_path):
            os.makedirs(extract_path)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
        speak("Zip file extracted successfully.")
    except Exception as e:
        speak("I couldn't unzip the file. Please check the paths or the file format.")
    
def change_wallpaper():
    speak("Tell me the folder path where your wallpapers are stored.")
    folder_path = takeCommand().lower().replace("sea", "c").replace("see", "c").replace("sequence", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    speak("Tell me the image name with extension. For example, sunset dot jpg.")
    image_name = takeCommand().lower().replace(" ", "").replace("dot", ".")

    try:
        image_path = os.path.join(folder_path, image_name)
        
        if not os.path.isfile(image_path):
            speak("The image was not found in the given folder.")
            return

        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
        speak("Wallpaper changed successfully.")

    except Exception as e:
        speak("I couldn't change the wallpaper. Please check the folder path and image name.")

def create_folder_shortcut():
    speak("Tell me the base path where your folder is located.")
    base_path = takeCommand().lower().replace("sea/see/se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    speak("Tell me the name of the folder.")
    folder_name = takeCommand().strip().replace(" ", "")

    folder_path = os.path.join(base_path, folder_name)

    speak("What should be the name of the shortcut?")
    shortcut_name = takeCommand().strip().replace(" ", "_") + ".lnk"

    speak("Tell me the path where you want to create the shortcut.")
    shortcut_location = takeCommand().lower().replace("sea/see/se", "c").replace("colon", ":").replace("slash", "\\").replace("backslash", "\\").replace(" ", "")

    shortcut_path = os.path.join(shortcut_location, shortcut_name)

    try:
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = folder_path
        shortcut.WorkingDirectory = folder_path
        shortcut.IconLocation = folder_path + '\\folder.ico'
        shortcut.save()
        speak("Shortcut created successfully.")
    except Exception as e:
        speak("I couldn't create the shortcut. Please check the paths provided.")

def list_remaining_files(extension=None, search_dir="C:/Users"):
    matched_files = []
    for root, dirs, files in os.walk(search_dir):
        for file in files:
            if extension is None or file.lower().endswith(extension.lower()):
                matched_files.append(os.path.join(root, file))
    return matched_files

if __name__ == "__main__":
    while True:
        print("Waiting for 'Hey Vois'...")
        if passive_listen_for_wake_word():
            speak("Welcome. Please authenticate.")

            if not os.path.exists("myvoice.wav"):
                speak("Please say your phrase to register.")
                record_voice("myvoice.wav", duration=5)

            authenticated = False
            for attempt in range(3):
                if authenticate_user():
                    authenticated = True
                    break
                else:
                    remaining = 2 - attempt
                    if remaining > 0:
                        speak(f"Authentication failed. {remaining} attempts remaining.")
                    else:
                        speak("Authentication failed. Returning to listening mode.")

            if authenticated:
                wishMe()
                listening = True

                while listening:
                    query = takeCommand().lower()

                    if 'stop' in query:
                        speak("I will stop listening.")
                        listening = False

                    elif 'start' in query:
                        speak("Listening resumed.")
                        listening = True

                    elif "your name" in query:
                        speak("I am your personal assistant.")
                    
                    elif "How is going on" in query:
                        speak("I'm doing great. How about you?")

                    elif "i am fine" in query or "i am good" in query:
                        speak("Glad to hear that!")

                    elif "thank you" in query:
                        speak("You're welcome!")

                    elif listening:
                        if 'open camera' in query:
                            takePicture()

                        elif 'take screenshot' in query:
                            takeScreenshot()

                        elif 'date and time' in query:
                            tellDateTime()

                        elif 'open word' in query:
                            openWordAndWrite()

                        elif "send email" in query:
                            send_email()

                        elif 'wikipedia' in query:
                            try:
                                import speech_recognition as sr
                                speak('Searching Wikipedia...')
                                search_term = query.replace("wikipedia", "").strip()

                                results = wikipedia.summary(search_term, sentences=2)
                                speak("According to Wikipedia")

                                r = sr.Recognizer()
                                sentences = results.split('. ')
                                for sentence in sentences:
                                    if sentence.strip():
                                        speak(sentence.strip())
                                        time.sleep(0.2)

                                        with sr.Microphone() as source:
                                            r.adjust_for_ambient_noise(source, duration=0.3)
                                            try:
                                                audio = r.listen(source, timeout=1.8, phrase_time_limit=2)
                                                command = r.recognize_google(audio).lower()
                                                if 'stop' in command:
                                                    speak("Okay, stopping the answer.")
                                                    break
                                            except sr.WaitTimeoutError:
                                                continue
                                            except sr.UnknownValueError:
                                                continue
                                            except Exception as e:
                                                print(f"[Wikipedia Stop Listener Error] {e}")
                                                continue
                            except Exception as e:
                                print(f"[Wikipedia Error] {e}")
                                speak("Sorry, I couldn't find anything on Wikipedia.")

                        elif 'shutdown' in query:
                            systemControl('shutdown')
                        
                        elif "open command prompt" in query or "open cmd" in query:
                            os.system("start cmd")
                            speak("Opening Command Prompt.")

                        elif "open browser" in query or "open chrome" in query:
                            os.system("start chrome")
                            speak("Opening Google Chrome.")

                        elif "open notepad" in query:
                            os.system("notepad.exe")
                            speak("Opening Notepad.")

                        elif "open task manager" in query:
                            os.system("taskmgr")
                            speak("Opening Task Manager.")

                        elif "open file explorer" in query:
                            os.system("explorer")
                            speak("Opening File Explorer.")

                        elif "open spotify" in query:
                            os.system("start spotify")
                            speak("Opening Spotify.")

                        elif 'restart' in query:
                            systemControl('restart')
                            
                        elif 'increase volume' in query:
                            increase_volume()

                        elif 'decrease volume' in query:
                            decrease_volume()

                        elif "mute" in query:
                            pyautogui.press("volumemute")
                            speak("Muted.")
                        
                        elif "unmute" in query:
                            pyautogui.press("volumemute")
                            speak("Unmuted.")

                        elif 'open calculator' in query:
                            openCalculator()

                        elif 'open google' in query:
                            webbrowser.open("google.com")

                        elif 'open youtube' in query:
                            webbrowser.open("youtube.com")

                        elif 'open code' in query:
                            os.startfile("D:\\Microsoft VS Code\\Code.exe")

                        elif 'open setting' in query:
                            openSettings()

                        elif 'enable hotspot' in query:
                            enableHotspot()

                        elif 'close window' in query:
                            closeWindow()
                        
                        elif "open facebook" in query:
                            webbrowser.open("https://www.facebook.com")
                            speak("Opening Facebook.")

                        elif "search for" in query:
                            query = query.replace("search for", "").strip()
                            webbrowser.open(f"https://www.google.com/search?q={query}")
                            speak(f"Searching for {query}.")

                        elif "open whatsapp" in query:
                            os.system("start whatsapp")
                            speak("Opening WhatsApp.")

                        elif 'open excel' in query:
                            openExcel()

                        elif 'open powerpoint' in query:
                            openPowerPoint()

                        elif 'my day' in query:
                            speak("Tell me your Rashi.")
                            rashi = takeCommand()
                            prediction = get_astrology_prediction(rashi)
                            speak(prediction)

                        elif 'call' in query:
                            make_call()

                        elif 'move mouse' in query:
                            if 'up' in query:
                                pyautogui.moveRel(0, -100)
                            elif 'down' in query:
                                pyautogui.moveRel(0, 100)
                            elif 'left' in query:
                                pyautogui.moveRel(-100, 0)
                            elif 'right' in query:
                                pyautogui.moveRel(100, 0)

                        elif 'click left' in query:
                            pyautogui.click(button='left')

                        elif 'click right' in query:
                            pyautogui.click(button='right')

                        elif 'double click' in query:
                            pyautogui.doubleClick()

                        elif "speak" in query or "say" in query:
                            ask_and_speak_multilingual()

                        elif 'set reminder' in query:
                            setReminder()

                        elif 'list pdf files' in query:
                            pdfs = list_remaining_files('.pdf')
                            if pdfs:
                                speak(f"Found {len(pdfs)} PDF files.")
                                for i, f in enumerate(pdfs[:5]):
                                    print(f"{i+1}. {f}")
                            else:
                                speak("No PDF files found.")
                        
                        elif 'open file' in query:
                            open_file_voice()
                        elif "brightness" in query:
                            level = int(query.split()[-1])
                            control_brightness(level)
                        elif "system status" in query:
                            speak(get_system_status())
                        elif "copy to clipboard" in query:
                            speak("What text should I copy?")
                            text = listen()
                            copy_text(text)
                            speak("Text copied to clipboard.")
                        elif "paste from clipboard" in query:
                            text = paste_text()
                            speak("Pasted text is: " + text)
                        elif "set alarm" in query:
                            speak("At what time should I set the alarm?Please say the time in 24-hour format, like 7:30 or 19:00.")
                            alarm_time = listen()  
                            set_alarm_time(alarm_time)
                            
                        elif any(word in query for word in ["question", "who", "what", "tell", "how", "whom", "why", "when", "where"]):
                            try:
                                import speech_recognition as sr
                                search_query = query.strip()  

                                print(f"User asked: {search_query}")
                                url = f"https://api.duckduckgo.com/?q={search_query}&format=json"
                                print(f"Fetching: {url}")
                                response = requests.get(url)
                                data = response.json()

                                answer = data.get("AbstractText")
                                print(f"AbstractText: {answer}")

                                if not answer:
                                    print("Trying RelatedTopics fallback...")
                                    related = data.get("RelatedTopics", [])
                                    for topic in related:
                                        if "Topics" in topic:
                                            for subtopic in topic["Topics"]:
                                                if isinstance(subtopic, dict) and "Text" in subtopic:
                                                    answer = subtopic["Text"]
                                                    break
                                        elif isinstance(topic, dict) and "Text" in topic:
                                            answer = topic["Text"]
                                            break
                                        if answer:
                                            break

                                if answer:
                                    print("Answer: ", answer)
                                    r = sr.Recognizer()
                                    sentences = answer.split('. ')
                                    for sentence in sentences:
                                        if sentence.strip():
                                            speak(sentence.strip())
                                            time.sleep(0.2)

                                            with sr.Microphone() as source:
                                                r.adjust_for_ambient_noise(source, duration=0.3)
                                                try:
                                                    audio = r.listen(source, timeout=1.8, phrase_time_limit=2)
                                                    command = r.recognize_google(audio).lower()
                                                    if 'stop' in command:
                                                        speak("Okay, stopping the answer.")
                                                        break
                                                except sr.WaitTimeoutError:
                                                    continue
                                                except sr.UnknownValueError:
                                                    continue
                                                except Exception as e:
                                                    print(f"[Listening Error] {e}")
                                                    continue
                                else:
                                    speak("Sorry, I couldn't find an answer.")

                            except Exception as e:
                                print(f"[DuckDuckGo API Error] {e}")
                                speak("There was a problem accessing the answer. Please try again.")

                        elif 'news' in query:
                            get_news()     

                        elif 'weather' in query:
                            speak("For which city?")
                            city = takeCommand()
                            if city:
                                weather_report = get_weather_wttr(city)
                                speak(weather_report)
                            else:
                                speak("I couldn't understand the city name.")
                        elif "create file" in query or "make a new file" in query or "create document" in query:
                            create_file()
                        elif "rename file" in query or "change file name" in query:
                            rename_file()
                        elif 'delete file' in query:
                            delete_file()
                        elif 'create folder' in query or 'make folder' in query:
                            create_folder()
                        elif 'rename folder' in query:
                            rename_folder()

                        elif 'delete folder' in query:
                            delete_folder()
                        elif 'read file' in query:
                            read_file()
                        elif "calculate" in query or "what is" in query:
                            speak("Tell me the expression to calculate.")
                            expr = takeCommand().lower()
                            calculate_expression(expr)
                        elif 'zip folder' in query:
                            zip_folder()
                        elif 'unzip file' in query:
                            unzip_file()
                        elif "change wallpaper" in query:
                            change_wallpaper()
                        elif 'create shortcut' in query or 'make shortcut' in query:
                            create_folder_shortcut()
                        
                        elif 'enable wifi' in query:
                            enable_wifi()

                        elif 'disable wifi' in query:
                            disable_wifi()

                        elif 'turn on bluetooth' in query or 'enable bluetooth' in query:
                            enable_bluetooth()

                        elif 'turn off bluetooth' in query or 'disable bluetooth' in query:
                            disable_bluetooth()

                        elif 'stop' in query:
                            speak("I will stop listening.")
                            listening = False

                        elif 'exit' in query or 'quit' in query or 'exit assistant' in query:
                            speak("VOIS is shutting down. Goodbye!")
                            os._exit(0)

                    else:
                        speak("Access denied. Maximum attempts reached.")