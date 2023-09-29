import win32com.client
import os
import speech_recognition as sr
import webbrowser
import openai
import datetime
import random

speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Speak("hello I'm Audi AI")

def takeCommand():

    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing !!")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said:{query}")
            return query
        except Exception as e:
            return "Sorry I didn't recognize you"



while 1:
    print("Listening......")
    query = takeCommand()
    speaker.Speak(query)
    if "Thank you Audi".lower() in query.lower():
        speaker.Speak(f"thank you madam hope you have a great experience with me")
        break

    #todo: Add more sites
    sites = [["google","https://www.google.com"],["youtube","https://youtube.com"],["wikipedia","https://www.wikipedia.com"],["linkedin","https://www.linkedin.com"],["gmail","https://www.gmail.com"]]
    for site in sites:
        if f"Open {site[0]}".lower() in query.lower():
            speaker.Speak(f"Opening {site[0]}")
            webbrowser.open(site[1])

    if "the time" in query:
        hour = datetime.datetime.now().strftime("%H")
        min = datetime.datetime.now().strftime(("%M"))
        speaker.Speak(f"Madam the time is {hour} hour and {min} minutes")

