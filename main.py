import time
import instaloader as instaloader
import pyautogui
import pyttsx3
import speech_recognition as sr
import datetime
import os
import cv2
import wikipedia
import webbrowser
import pywhatkit as kit
import sys
import random
import pyjokes
import requests
from PyQt5 import QtWidgets,QtCore,QtGui
from PyQt5.QtCore import QTimer,QTime,QDate,Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from gui


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[1].id)


# text to speech
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()


# to convert voice to text
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("listening...")
        r.pause_threshold = 1
        audio = r.listen(source, timeout=9, phrase_time_limit=9)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        # print(f"user said: {query}")
        # speak(query)
    except Exception as e:
        speak("Say that again please...")
        return "none"
    return query


# opening apps


def NewsFromBBC():  # BBC news api
    # following query parameters are used
    # source, sortBy and apiKey
    query_params = {
        "source": "bbc-news",
        "sortBy": "top",
        "apiKey": "55a8f017716041c89662b67114c99d5c"
    }
    main_url = " https://newsapi.org/v1/articles"

    # fetching data in json format
    res = requests.get(main_url, params=query_params)
    open_bbc_page = res.json()

    # getting all articles in a string article
    article = open_bbc_page["articles"]

    # empty list which will
    # contain all trending news
    results = []

    for ar in article:
        results.append(ar["title"])

    for i in range(len(results)):
        # printing all trending news
        print(i + 1, results[i])

    # to read the news out loud for us
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(results)


# to wish
def wish():
    hour = int(datetime.datetime.now().hour)

    if hour >= 0 and hour <= 12:
        speak("Good Morning")
    elif hour > 12 and hour < 18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak("I am Teddy. Please tell me how can I help you")


# .......................................... instructions...............................................
if __name__ == "__main__":
    wish()
    while True:
        # if 1:

        query = takecommand().lower()

        # notepad
        if "open notepad" in query:
            npath = "C:\\WINDOWS\\system32\\notepad.exe"
            os.startfile(npath)
        elif "close notepad" in query:
            speak("okay sir, closing notepad")
            os.system("taskkill /f /im notepad.exe")
        # vs code
        elif "open code editor" in query:
            bpath = "C:\\Users\\VICKY\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(bpath)
        elif "close code editor" in query:
            speak("closing code editor")
            os.system("taskkill /f /im  code.exe")
        # opening excel
        elif "open excel" in query:
            dpath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(dpath)
        elif "close excel" in query:
            speak(" closing excel")
            os.system("taskkill /f /im  EXCEL.EXE")
        # opening brave
        elif "open brave" in query:
            epath = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
            os.startfile(epath)
        elif "close brave" in query:
            speak(" closing brave")
            os.system("taskkill /f /im  brave.exe")
        # command prompt
        elif "open command" in query:
            os.system("start cmd")
        # opening word document
        elif "open document" in query:
            speak("Opening document")
            cpath = "C:\\Program Files\\Microsoft Office\\root\Office16\\WINWORD.EXE"
            os.startfile(cpath)
        elif "close document" in query:
            speak(" closing document")
            os.system("taskkill /f /im  WINWORD.EXE")
        # camera
        elif "open camera" in query:
            cap = cv2.VideoCapture(0)
            while True:
                ret, img = cap.read()
                cv2.imshow("webcam", img)
                k = cv2.waitKey(50)
                if k == 27:
                    break
            cap.release()
            cv2.destroyAllWindows()
        # play music
        elif "play music" in query:
            music_dir = "E:\\music"
            songs = os.listdir(music_dir)
            # rd = random.choice(songs)
            for song in songs:
                if song.endswith('.mp3'):
                    os.startfile(os.path.join(music_dir, song))
        # open youtube
        elif "open youtube" in query:
            webbrowser.open("www.youtube.com")
        elif "play songs on youtube" in query:
            speak("Please tell me which song to play")
            sn = takecommand().lower()
            kit.playonyt(f"{sn}")
        # wikipedia
        elif "wikipedia" in query:
            print("What should i search about")
            search = takecommand().lower()
            speak("searching wikipedia")
            # query = query.replace("Wikipedia", "")
            result = wikipedia.summary(search, sentences=10)
            speak("According to wikipedia")
            speak(result)
            print(result)
        # open gmail
        elif "open email" in query:
            webbrowser.open("www.gmail.com")

        # open google

        elif "open google" in query:
            speak("what should i search on google ?")
            cb = takecommand().lower()
            webbrowser.open(f"{cb}")
        # send msg
        elif "send message" in query:
            number = input("Enter the number  you pig")
            speak("Speak the msg please")
            ms = takecommand().lower()
            current_time = datetime.datetime.now()
            current_hour = current_time.hour
            current_minute = current_time.minute
            kit.sendwhatmsg(f" +91 {number}", f"{ms}", current_hour, current_minute + 1)
        # to find a joke
        elif "tell me a joke" in query:
            My_joke = pyjokes.get_joke(language="en", category="neutral")
            speak(My_joke)
        # shutdown
        elif "shut down the system" in query:
            os.system("shutdown /s /t 5")
        # restart
        elif "restart the system" in query:
            os.system("shutdown /r /t 5")
        # sleep
        elif "system sleep" in query:
            os.system("rundll32.exe powrprof.dll, SetSuspendState 0,1,0")

        elif 'switch the window' in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")
        # news
        elif "tell me news" in query:
            speak("please wait sir, fetching the latest news")
            NewsFromBBC()
        # insta
        elif "instagram profile" in query or "profile on instagram" in query:
            speak("Please enter the user name ")
            name = input("Enter username here:")
            webbrowser.open(f"www.instagram.com/{name}")
            speak(f"Here is the profile of the user {name}")
            time.sleep(3)
            speak("sir would you like to download profile picture of this account.")
            condition = takecommand().lower()
            if "yes" in condition:
                mod = instaloader.Instaloader()
                mod.download_profile(name, profile_pic_only=True)
                speak("i am done sir, profile picture is saved in our main folder. now i am ready")
            else:
                pass
        elif "take screenshot" in query or "take a screenshot" in query:
            speak("sir, please tell me the name for this screenshot file")
            name = takecommand().lower()
            speak("please sir hold the screen for few seconds, i am taking screenshot")
            time.sleep(3)
            img = pyautogui.screenshot()
            img.save(f'E:\\Image\\{name}.png')
            speak("Screen shot saved")
        # open lms
        elif "open college site" in query:
            webbrowser.open("https://lmsone.iiitkottayam.ac.in/login/index.php")
        # open online compiler
        elif "open online compiler" in query:
            webbrowser.open("https://www.onlinegdb.com/online_bash_shell#")

        elif "close" in query:
            speak("Thank You for using me")
            sys.exit()












