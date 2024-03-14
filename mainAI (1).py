import webbrowser
import openai
import datetime
import speech_recognition as sr
import subprocess,os,time
import win32com.client
from testconfig import name,apikey

speaker = win32com.client.Dispatch("SAPI.SpVoice")

# while 1:
#     print("enter the word you want to hear")
#     s = input()
#     speaker.Speak(s)

def ai(query):
    openai.api_key = apikey
    query += f" in short"
    content = f"OpenAI response for the promt: {query}  \n ****************** \n\n"
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "system",
                "content": query
            },
            {
                "role": "user",
                "content": ""
            }
        ],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: wrap this in try catch block
    answer = response["choices"][0]["message"]["content"]
    print(answer)
    say(answer)

    # text += response["choices"][0]["text"]
    # print(text)

def say(text):
    # os.system(f"say {text}")
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        #r.pause_threshold = 0.8
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said {query}")
            return query
        except Exception as e:
            return "I'm sorry I didn't quite catch that"

if __name__ == "__main__":
    print('Pycharm')
    say(f"Hello I am {name}, Your Personal A.I. Voice Based assistant")
    time.sleep(0.5)
    while True:
        say("Im listening")
        print("Listening...")
        query = takeCommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"]]

        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} ......")
                webbrowser.open(site[1])

        if "the time" in query.lower():
            hour = datetime.datetime.now().strftime("%H")
            hour = str(int(hour) % 12)
            min = datetime.datetime.now().strftime("%M")
            say(f"The time is {hour} hours and {min} minutes")

        if "open camera" in query.lower():
            say("opening camera")
            subprocess.run('start microsoft.windows.camera:', shell=True)

        if "close camera" in query.lower():
            say("closing camera")
            subprocess.run('Taskkill /IM WindowsCamera.exe /F', shell=True)

        if "shut down" in query.lower() or "shutdown" in query.lower():
            say("Shutting down")
            break

        if "using artificial intelligence" in query.lower():
            ai(query)

        say(query)
