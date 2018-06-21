import pyttsx
#import speech_recognition as sr
from time import ctime
import time
import os
import webbrowser
from gtts import gTTS
import win32gui
import win32con
from os import getpid, system
from threading import Timer

speech_engine = pyttsx.init()
def speak(text):
#import pyttsx
    #engine = pyttsx.init()
    speech_engine.say(text)
    speech_engine.runAndWait()
 
def bot_input():
    text = raw_input('ask your question:')
    speak(text)
    #speech_engine.runAndWait()
    return text

def jarvis(data):
    if "how are you" in data:
        speak("I am fine")
 
    if "what time is it" in data:
        speak(ctime())
 
    if "where is" in data:
        data = data.split(" ")
        location = data[2]
        speak("Hold on kanishk, I will show you where " + location + " is.")
        #os.system("mozilla https://www.google.nl/maps/place/" + location + "/&amp;")
        webbrowser.open_new_tab("https://www.google.co.in/maps/place/" + location)

    elif "open Media player" in data:
        speak("Hold on, I will open media player")
        os.startfile('C:\Program Files (x86)\Windows Media Player\wmplayer.exe', 'open')

    elif "search for" in data:
        data = data.split(" ")
        searchfor = data[2]
        speak("Hold on, I will search for " + searchfor)
        webbrowser.open_new_tab('https://www.google.co.in/?gfe_rd=cr&ei=c5cTWLeuOqLnugSekrfQAw#q=' + searchfor) 

    elif "open power point" in data:
        speak("Hold on, I will open power point")
        os.startfile('C:\Program Files (x86)\Microsoft Office\Office14\POWERPNT.exe', 'open')

    elif "open word" in data:
        speak("Hold on, I will open word")
        os.startfile('C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.exe', 'open')

    elif "open excel" in data:
        speak("Hold on, I will open excel")
        os.startfile('C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.exe', 'open')

    elif "display off" in data:#sys.platform.startswith('win'):
        def force_exit():
            pid = getpid()
            system('taskkill /pid %s /f' % pid)

        t = Timer(1, force_exit)
        t.start()
        SC_MONITORPOWER = 0xF170
        win32gui.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, SC_MONITORPOWER, 2)
        t.cancel()

"""
    elif "display off" or "display of" in data:
        def runScreensaver():
          strComputer = "."
          objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
          objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
          colItems = objSWbemServices.ExecQuery("Select * from Win32_Desktop")
          for objItem in colItems:
             if objItem.ScreenSaverExecutable:
                os.system(objItem.ScreenSaverExecutable + " /start")
                break
"""


        

time.sleep(2)
speak("Hey Kanishk, what can I do for you?")
while 1:
    data = bot_input()
    jarvis(data)
"""
import pyttsx
engine = pyttsx.init()
#print "ask your question"
text = raw_input('ask your question')
engine.say(text)
engine.runAndWait()
"""
