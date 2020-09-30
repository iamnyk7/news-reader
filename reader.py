
from win32com.client import Dispatch
import requests
import json

def speak(str):
    speak=Dispatch("SAPI.SpVoice")
    speak.speak(str)

if __name__ == '__main__':
    speak(" welcome to big bulletin I am your news anchor Ranganna OP")
    url="http://newsapi.org/v2/top-headlines?country=in&apiKey=9f6878be6ff343cf8dee9b022e186d9c"
    news=requests.get(url).text
    news_py=json.loads(news)
    article=news_py['articles']
    last=news_py['totalResults']
    speak("Our first news is")
    v=0
    for art in article:
        v+=1
        if v== last:
            speak("Moving on last news")
            speak(art['title'])
            break
#         speak(art['title'])
#         speak("Alright mundakke hogona")
#     speak("Thank you have a goood day")
