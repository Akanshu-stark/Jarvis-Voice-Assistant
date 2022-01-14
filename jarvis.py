import datetime
import os
import smtplib
import webbrowser
import json
import pyttsx3
import requests
import speech_recognition as sr
import wikipedia
import pyjokes
from PyDictionary import PyDictionary

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am Jarvis . Please tell me how may I help you")

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source,duration=1)
        r.dynamic_energy_threshold = True
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query

def NewsFromBBC():
    main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=enter your key"
    open_bbc_page = requests.get(main_url).json()
    article = open_bbc_page["articles"]
    results = []
    for ar in article:
        results.append(ar["title"])
    for i in range(len(results)):
        print(i + 1, results[i])
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(results)

def Weatherget():
    api_key = "enter your key"
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    speak("Please Enter city name")
    city_name = input("Enter city name : ")
    complete_url = base_url + "appid=" + api_key + "&q=" + city_name
    response = requests.get(complete_url)
    x = response.json()
    if x["cod"] != "404":
        y = x["main"]
        current_temperature = y["temp"]
        current_pressure = y["pressure"]
        current_humidiy = y["humidity"]
        z = x["weather"]
        weather_description = z[0]["description"]
        print(" Temperature (in kelvin unit) = " +
              str(current_temperature) +
              "\n atmospheric pressure (in hPa unit) = " +
              str(current_pressure) +
              "\n humidity (in percentage) = " +
              str(current_humidiy) +
              "\n description = " +
              str(weather_description))
        speak(f" Temperature {str(current_temperature)} , atmospheric pressure {str(current_pressure)} , humidity {str(current_humidiy)} , description {str(weather_description)}")

    else:
        print(" City Not Found ")

def sendEmail(to, content):
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login('enter_your_email', 'your_password')
    server.sendmail('enter_your_email', to, content)
    server.quit()

def log_list(list_items):
    f = open("listofitems.txt","w")
    f.write(list_items)
    f.close()

def tell_list():
    f = open("listofitems.txt","r")
    speak(f.read())
    f.close()
if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'hello jarvis' in query:
            speak("Hello Sir! ")

        elif 'how are you' in query:
            speak("I am fine Sir! How about you?")

        elif 'send email' in query:
            try:
                speak("what should I say?")
                content = takeCommand()
                to = "enter_email"
                sendEmail(to,content)
                speak("email has been successfully sent")
            except exception as e:
                print(e)
                speak("Sorry I am not able to send the email")

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'open facebook' in query:
            webbrowser.open("facebook.com")

        elif 'open twitter' in query:
            webbrowser.open("twitter.com")

        elif 'open whatsapp' in query:
            webbrowser.open("web.whatsapp.com")

        elif 'open google meet' in query:
            webbrowser.open("meet.google.com")

        elif 'open classroom' in query:
            webbrowser.open("classroom.google.com")

        elif 'open instagram' in query:
            webbrowser.open("instagram.com")

        elif 'open online compiler' in query:
            webbrowser.open("twitter.com")

        elif 'play music' in query:
            music_dir = 'F:\music'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif ' time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open code' in query:
            codePath = "C:\Program Files\CodeBlocks\codeblocks.exe"
            os.startfile(codePath)

        elif 'joke' in query:
            jokes1= pyjokes.get_joke('en','neutral')
            speak(jokes1)

        elif 'created you' in query:
            speak("My beloved Sir. Akanshu Rana created me")

        elif 'i love you' in query:
            speak("Oh! Thank you so much I also love you")

        elif 'can we have cup of coffee' in query:
            speak("Thank you but I don't take coffee. But we can have some volts of electricity if you want")

        elif 'on which pc you are running' in query:
            speak("I am running on Dell Inspiron 15 3000 ")

        elif 'marry me' in query:
            speak("No thanks I am married with my work!")

        elif 'tell me top 10 news' in query:
            NewsFromBBC()

        elif 'tell me weather' in query:
            Weatherget()
        elif 'open dictionary' in query:
            speak("Tell me the word ")
            word = PyDictionary(takeCommand())
            print(word.getMeanings())
            print(word.translateTo("hi"))
            from win32com.client import Dispatch

            speak = Dispatch("SAPI.Spvoice")
            speak.Speak(word.getMeanings())

        elif 'make shopping list' in query:
            speak("please tell me the list")
            list_items = takeCommand()
            log_list(list_items)
            speak("thank you I learned your list")
            print(list_items)

        elif 'tell me the list' in query:
            tell_list()
        
        elif 'jarvis exit' in query:
            speak("ok sir")
            exit(0)






