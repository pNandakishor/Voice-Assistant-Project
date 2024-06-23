import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import operator
import random
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import tkinter as tk
import time
import requests
import shutil
import cv2
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
#______________________

# Initialize the recognizer
recognizer = sr.Recognizer()

# Initialize the pyttsx3 engine
engine = pyttsx3.init()

# Function to convert text to speech
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to recognize speech
def recognize_speech():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio)
        print(f"User said: {query}")
        return query
    except Exception as e:
        print(e)
        print("Sorry, I couldn't understand what you said.")
        speak("Sorry, I couldn't understand what you said.")


# Function to execute commands
def execute_command(command):

    if "hello" in command:
        print("Hello! How can I help you?")
        speak("Hello! How can I help you?")

    elif "how are you" in command:
        print("I'm fine, thank you!")
        speak("I'm fine, thank you!")
    elif "what is your name" in command:
        print("My name is jarvis.")
        speak("My name is jarvis.")
    elif 'open youtube' in command:
        speak("Here you go to Youtube\n")
        webbrowser.open("https://www.youtube.com/")
    elif 'open google' in command:
        speak("Heer you go to Google\n")
        webbrowser.open("google.com")
    elif 'open stackoverflow' in command:
        speak("Here you go to Stack Over flow. Happy coding")
        webbrowser.open("stackoverflow.com")
    elif 'play music' in command or "play song" in command:
        speak("Here you go with music")
        # music_dir = "G:\\Song"
        music_dir = r"C:\Users\Manu\Downloads\testingsound"
        songs = os.listdir(music_dir)
        print(songs)
        random = os.startfile(os.path.join(music_dir, songs[0]))
    elif 'what is the time' in command:
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"Sir, the time is {strTime}")
    elif 'search' in command or 'play' in command:
        query = command.replace("search", "")
        query = query.replace("play", "")
        webbrowser.open(query)
    elif 'tell me joke' in command:
        speak(pyjokes.get_joke())

    elif 'what day is it' in command:
        current_day = datetime.datetime.now().strftime("%A")
        speak(f"Sir, today is {current_day}")
    elif 'what is the date' in command:
        current_date = datetime.datetime.now().strftime("%Y-%m-%d")
        speak(f"Sir, today's date is {current_date}")
    elif 'how are you' in command:
        speak("I am fine, Thank you")
        speak("How are you, Sir")
    elif 'open chrome' in command:
        codePath = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        os.startfile(codePath)
    elif "who made you" in command or "who created you" in command:
        print("I have been developed by the great Nandu garu .")
        speak("I have been created by the great Nandu garu .")
    elif "what's your name" in command or "What is your name" in command:
        print("My friends call me jarvis")
        speak("My friends call me jarvis")

    elif "camera" in command or "take a photo" in command:
        speak("Opening camera")
        #ubprocess.Popen("start microsoft.windows.camera:")
        os.startfile("microsoft.windows.camera:")
    elif "where is" in command:
        query = command.replace("where is", "")
        location = query
        speak("User asked to locate")
        speak(location)
        webbrowser.open("https://www.google.com/maps/place/" + location)
    elif "lock windows" in command:
        speak('windows locking now')
        os.system("rundll32.exe user32.dll,LockWorkStation")

    elif "write a note" in command:
        speak("What should I write, sir?")
        note = recognize_speech()
        with open(r'jarvis.txt', 'a') as file:
            file.write(note + '\n')
        speak("Note written successfully")
    elif 'what is news' in command:
        webbrowser.open("https://epaper.eenadu.net/")


    elif 'empty recycle bin' in command:
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
        speak("Recycle Bin Recycled")
    elif 'shutdown system' in command:
        speak("Hold On a Sec! Your system is on its way to shut down")
        subprocess.call('shutdown / p /f')


# Simulate Win+L key press to lock Windows
#pyautogui.hotkey('win', 'l')
    elif "ok bye" in command or "exit" in command:
        speak("Goodbye!")
        exit()

def on_button_click():
    command = recognize_speech().lower()
    execute_command(command)

window = tk.Tk()
window.title("Voice Assistant")
# Main function to run the assistant
window.geometry("400x300")

output_text = tk.StringVar()
output_label = tk.Label(window, textvariable=output_text)
output_label.pack()

button = tk.Button(window, text="Listen", command=on_button_click)
button.pack()



def main():
    speak("Hello! How can I help you?")
    while True:
        command = recognize_speech().lower()
        execute_command(command)

        
if __name__ == "__main__":
    main()
