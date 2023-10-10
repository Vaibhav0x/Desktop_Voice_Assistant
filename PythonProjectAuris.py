from sys import platform
import tkinter as tk
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import subprocess
import psutil
import pyaudio
import requests
import json
import numpy as np
import matplotlib.pyplot as plt
import time
import pygame
import pyqtgraph as pg
from pyqtgraph.Qt import QtGui, QtCore
from PyQt5 import QtGui, QtCore, QtWidgets
import os
import random
import wolframalpha
import operator
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import shutil
import threading
import sys
import re
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from tkinter import scrolledtext

camera_process = None
camera_open = False
google_open = False
sys.stdout.reconfigure(encoding='utf-8')

def speak(audio):
    conversation_area.insert(tk.END, 'Auris: ' + audio + '\n')
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = datetime.datetime.now().hour
    if 4 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 16:
        speak("Good Afternoon!")
    elif 16 <= hour < 20:
        speak("Good Evening!")
    else:
        speak("Good Night!")
     
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
            conversation_area.insert(tk.END, 'User: ' + query + '\n')
            speak("I heard you say: " + query)
        except sr.UnknownValueError:
            print("Auris could not understand audio")
            return "None"
        except Exception as e:
            print("Say that again, please...")
            return "None"

        return query

def open_camera():
    global camera_process, camera_open
    if not camera_open:
        camera_process = subprocess.Popen('start microsoft.windows.camera:', shell=True)
        camera_open = True
    else:
        print("The camera is already open")

def open_google():
    global google_open
    webbrowser.open("https://www.google.com")
    google_open = True

def get_random_advice():
    res = requests.get("https://api.adviceslip.com/advice").json()
    return res['slip']['advice']

def get_random_joke():
    headers = {
        'Accept': 'application/json'
    }
    res = requests.get("https://icanhazdadjoke.com/", headers=headers).json()
    return res["joke"]

class MusicVisualizer(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Music Visualizer")
        self.resize(1800, 1200)

        # Create a PlotWidget for the music visualization
        self.plot_widget = pg.PlotWidget()
        self.plot_widget.setBackground("w")
        self.plot_widget.showGrid(x=True, y=True)
        self.plot_widget.setLabel("bottom", "Time (s)")
        self.plot_widget.setLabel("left", "Amplitude")
        self.plot_curve = self.plot_widget.plot()

        # Layout the widgets
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.plot_widget)
        self.setLayout(layout)

        self.is_running = False
        self.stream = None

    def start_visualizer(self):
        if not self.is_running:
            self.is_running = True
            self.thread = threading.Thread(target=self.update_visualizer)
            self.thread.start()

    def stop_visualizer(self):
        if self.is_running:
            self.is_running = False
            self.thread.join()
            self.stream.stop_stream()
            self.stream.close()

    def update_visualizer(self):
        CHUNK = 1024
        RATE = 44100
        FORMAT = pyaudio.paInt16

        p = pyaudio.PyAudio()
        self.stream = p.open(format=pyaudio.paInt16, channels=1, rate=RATE, input=True, frames_per_buffer=CHUNK)

        while self.is_running:
            data = np.frombuffer(self.stream.read(CHUNK), dtype=np.int16)

            try:
                if self.plot_widget.isUnderMouse():
                    self.plot_curve.setData(data)
            except RuntimeError as e:
                print(f"RuntimeError: {e}")

        QtCore.QCoreApplication.processEvents()
        time.sleep(0.01)
        p.terminate()

def stop_music():
    pygame.mixer.music.stop()

def send_command():
    command = Auris.get()
    response = takeCommand()
    conversation_area.insert(tk.END, 'Jarvis: ' + response + '\n')

def get_weather(api_key, city):
    base_url = "http://api.openweathermap.org/data/2.5/weather"
    params = {
        "q": city,
        "appid": api_key,
        "units": "metric"
    }

    response = requests.get(base_url, params=params)
    weather_data = json.loads(response.text)

    if "weather" in weather_data:
        main_weather = weather_data["weather"][0]["main"]
        description = weather_data["weather"][0]["description"]
        temperature = weather_data["main"]["temp"]
        humidity = weather_data["main"]["humidity"]
        wind_speed = weather_data["wind"]["speed"]

        speak(f"Weather in {city}: {main_weather}, {description}.")
        speak(f"Temperature: {temperature}°C.")
        speak(f"Humidity: {humidity}%.")
        speak(f"Wind speed: {wind_speed} m/s.")

        result_text = f"Weather in {city}: {main_weather}, {description}. " \
                      f"Temperature: {temperature}°C. " \
                      f"Humidity: {humidity}%. " \
                      f"Wind speed: {wind_speed} m/s."

        print(result_text.encode('utf-8'))  # Encode the result_text using UTF-8 encoding before printing
    else:
        speak("Sorry, I couldn't fetch the weather information.")

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('rajjvaibhavv121@gmail.com', 'password')
    server.sendmail('rajjvaibhavv121@gmail.com', to, content)
    server.close()

def exit_program():
    global exit_program
    exit_program = True
    window.destroy()

def calculate(expression):
    # Remove non-numeric characters and convert spoken words to their numeric representation
    expression = re.sub(r'[^\d+\-*/(). ]+', '', expression, flags=re.IGNORECASE)
    expression = expression.lower()
    expression = expression.replace('plus', '+').replace('minus', '-').replace('multiply', '*').replace('times','*').replace('divide', '/')
    expression = expression.replace('multiply by', '*').replace('x', '*')  # Corrected replacement

    try:
        result = eval(expression)
        speak(f"The result is {result}")
    except ZeroDivisionError:
        speak("Error: Cannot divide by zero")
    except Exception as e:
        speak(f"Error: {str(e)}")

def listen_and_respond():
    wishMe()
    speak("How can I Help you, Sir")
    api_key = "fb6f75b3d1a87d2606371ebef16bd928"  # Replace with your OpenWeatherMap API key
    cities = ["New Delhi"]  # Replace with the desired city name
    while(1):
        query = takeCommand().lower()
        try:    
            if 'wikipedia' in query:
                try:
                    speak("According to Wikipedia....")
                    query = query.replace("wikipedia", "")
                    results = wikipedia.summary(query, sentences=3)
                    speak("According to Wikipedia")
                    print(results)
                    speak(results)
                except wikipedia.exceptions.PageError:
                    speak("Sorry, I couldn't find any Wikipedia page matching your query.")
                except wikipedia.exceptions.DisambiguationError as e:
                    options = e.options[:3]  # Get the first three options from the disambiguation error
                    speak("Multiple options found. Here are the top three:")
                    for i, option in enumerate(options):
                        print(f"{i+1}. {option}")
                        speak(f"{i+1}. {option}")
                except (wikipedia.exceptions.WikipediaException, ConnectionError):
                    speak("Sorry, I couldn't fetch information from Wikipedia. Please check your internet connection.") 
        except UnicodeEncodeError:
            print("UnicodeEncodeError: Cannot display some characters in the output.")
            print("Please ensure your terminal or console supports Unicode encoding.")
            speak("Sorry, I encountered an error while processing the output.")


        if 'open camera' in query:
            speak("Opening Camera")
            open_camera()

        elif 'open google' in query:
                speak("Opening the google")
                open_google()

        elif 'play music' in query:
            music_dir = "C:\Songs\Songs"
            songs = os.listdir(music_dir)
            random.shuffle(songs)
            pygame.mixer.init()
            pygame.mixer.music.load(os.path.join(music_dir, songs[0]))
            pygame.mixer.music.play()
            app = QtWidgets.QApplication([])
            window = MusicVisualizer()
            window.show()
            window.start_visualizer()
            app.exec_()

        elif 'stop music' in query:
            stop_music()
            
        elif 'email to vaibhav' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "email_id"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry, we can't send your email!")

        elif 'provide advice' in query or 'advice' in query or 'provide a advice' in query:
            speak(f"Here's an advice for you, sir")
            advice = get_random_advice()
            print(advice)
            speak(advice)
            
        elif 'provide joke' in query or 'joke' in query or 'provide a joke' in query:
            speak(f"Hope you like this one sir")
            joke = get_random_joke()
            print(joke)
            speak(joke)
        
        elif 'open google maps' in query:
            webbrowser.open("https://www.google.com/maps/@28.7063024,77.2485325,10.43z?entry=ttu")

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
        
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
        elif 'open gmail' in query:
            webbrowser.open("https://mail.google.com/")
        
        elif 'calculate' in query:
            calculation = query.replace("calculate", "").strip()
            if calculation:
                calculate(calculation)
            else:
                speak("Sorry, I didn't catch the calculation. Could you please repeat?")

        elif 'weather update' in query:
            for city in cities:
                get_weather(api_key,city)

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")
 
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
 
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")
 
        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            name="Auris"
            speak(name)
            print("My friends call me", name)
        
        elif "who i am" in query:
            speak("If you talk then definitely your human.")
 
        elif "why you came to world" in query:
            speak("Thanks to Vaibhav Raj. further It's a secret")
 
        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")
 
        elif "who are you" in query:
            speak("I am your virtual assistant created by Vaibhav Raj")
 
        elif 'reason for you' in query:
            speak("I was created as a Minor project by Vaibhav Raj")
 
        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "Location of wallpaper",
                                                       0)
            speak("Background changed successfully")

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('auris.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
         
        elif "show note" in query:
            speak("Showing Notes")
            file = open("auris.txt", "r")
            print(file.read())
            speak(file.read(6))       
        
        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Vaibhav Raj.")
             
        elif 'exit' in query:
            speak('Goodbye Sir!')
            window.after(100, window.destroy)  # Close the Tkinter window
            break

# Create the main window
window = tk.Tk()
window.title("Auris Voice Assistant")

# Create the GUI components
conversation_area = tk.Text(window, bg="olive", fg="white", font=("Arial", 12))
Auris = tk.Entry(window, font=("Arial", 12))
send_button = tk.Button(window, text="Send", font=("Arial", 12), command=send_command, bg="blue", fg="white")

# Add the components to the window
conversation_area.pack()
Auris.pack(side=tk.LEFT)
send_button.pack(side=tk.RIGHT)

# Initialize the speech synthesis engine
engine = pyttsx3.init()
# Set Rate
engine.setProperty('rate', 190)
# Set Volume
engine.setProperty('volume', 1.0)
# Set Voice
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Create a thread for listening and responding
thread = threading.Thread(target=listen_and_respond)
thread.daemon = True
thread.start()

# Start the GUI event loop
window.mainloop()
