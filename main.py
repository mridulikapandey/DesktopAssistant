import webbrowser
import openai
import speech_recognition as sr
import os
import win32com.client
from AppOpener import open
from AppOpener import close


speaker = win32com.client.Dispatch("SAPI.SpVoice")
def say(text):
    speaker.Speak(text)


def takeCommand():
   r = sr.Recognizer()
   with sr.Microphone() as source:
         #r.pause_threshold =1
         audio = r.listen(source)
         try:
           query = r.recognize_google(audio,language="en-in")
           print(f"USER SAID : {query}")
           return query
         except Exception as e:
             return "SORRY SOME ERROR OCCURED"


if __name__ == '__main__':


     #say="Hello MRIDULIKA CUTIE Welcome To JARVIS A.I. YOU ARE THE BEST AND RHYTHM IS WORSTTTTTTTTT "
     #speaker.Speak(say)
     while True:
        print("listening")
        #say("hello")

        query = takeCommand()
        #speaker.Speak(text)
        #break
        sites = [["youtube","https://youtube.com"],["google","https://google.com"],["wikipedia","https://wikipedia.com"],["kiit sap","https://kiit.ac.in/sap/"]]
        for site in sites:
         if f"Open {site[0]}".lower() in query.lower():
            say(f"opening {site[0]}")
            webbrowser.open_new_tab(site[1])


         if "Open music" in query:
             say("opening music")
             musicPath = "/Users/KIIT/Downloads/Hamare-sath-shree-raghunath (2).mp3"
             say("opening music")
             speaker.Speak(f"open{musicPath}")


         if "time" in query:
             say("telling time")
             #musicPath = "/Users/KIIT/Downloads/Hamare-sath-shree-raghunath (2).mp3"
             strfTime = datetime.datetime.now().strftime("%Hours:%Minutes:%Seconds ")
             say(f"time is{strfTime}")
             break

        apps =[["WhatsApp"], ["Google Chrome"],["Google Drive"]]
        for app in apps:
            if f"Open {app[0]}".lower() in query.lower():
             say(f"opening{app[0]}")
             open(app[0])
            if f"close {app[0]}".lower() in query.lower():
             say(f"closing{app[0]}")
             close(app[0])


