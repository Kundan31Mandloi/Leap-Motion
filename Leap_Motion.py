#Import Libraries
import serial
import speech_recognition as sr
import pyautogui as pt
import time
import pyperclip
import pyttsx3
import win32com.client as wincl
import winsound

#Set Frequency and Duration for Voice
frequency = 2500  
duration = 500  
sp=wincl.Dispatch('SAPI.SpVoice')

#Set Message to be Generated
sp.Speak('hello User, welcome to the program. Please listen the instructions Carefully')
sp.Speak('From Now ownwards, This system will be controlled with the help of Ultrasonic Sensors')
sp.Speak('For Left click, put your hand on left Sensor between 10 to 25 centimeter')
sp.Speak('For Right click, put your hand on right Sensor between 10 to 25 centimeter')
sp.Speak('For Left Move, put your hand on left Sensor between 25 to 60 centimeter')
sp.Speak('For right Move, put your hand on right Sensor between 25 to 60 centimeter')
sp.Speak('For down Move, put your hands on both Sensors together between 10 to 25 centimeter')
sp.Speak('For up Move, put your hands on both Sensors together between 25 to 60 centimeter')
sp.Speak('For typing, keep your one hand on left sensor between 10 to 25 centimeter and other hand on right sensor together between 25 to 60 centimeter')
sp.Speak('If you forget the rule than keep your hands on both the sensors, Rules will be repeated again')

#Connect With Arduino 
au=serial.Serial('COM6',9600)
sp=wincl.Dispatch('SAPI.SpVoice')

while True:
    #print('Distance ')
    x=au.readline()
    y=au.readline()
    x=x.decode().strip()    
    y=y.decode().strip()
    print(x,y)
    x=int(x)
    y=int(y)
    
    #print(x,y)
    #print(x,y)
    #//x=au.readline()
    #print(x.decode().strip(),y.decode().strip())
    #x=int(x.decode().strip())
    #y=int(y.decode().strip())
    #time.sleep(1)
    #print(y.decode().strip(),x.decode().strip())
    #print(x,y)
    if(20<x<35 and y>75):
        pt.click(button='left')
        print('left click')
    elif(60>x>35 and y>75):
        while(True):
            x=au.readline()
            y=au.readline()
            x=int(x.decode().strip())
            y=int(y.decode().strip())
            if(x>80 and y>80):
                pt.moveRel(-10,0)
            if(abs(x-y)<25 and x<75 and y<75):
                #time.sleep(0.4)
                break
        print('left move')
    
    elif(20<y<35 and x>75):
        pt.click(button='right')
        print('right click')

    elif(75>x>60and y>75):
        while(True):
            x=au.readline()
            y=au.readline()
            x=int(x.decode().strip())
            y=int(y.decode().strip())
            if(x>75 and y>75):
                pt.moveRel(0,-10)   
            if(abs(x-y)<25 and x<75 and y<75):
                #time.sleep(0.4)
                break

    elif(75<x and 60<y<75):
        while(True):
            x=au.readline()
            y=au.readline()
            x=int(x.decode().strip())
            y=int(y.decode().strip())
            if(x>75 and y>75):
                pt.moveRel(0,10)
            if(abs(x-y)<25 and x<75 and y<75):
                #time.sleep(0.4)
                break
    
    elif(35<y<60 and x>75):
        pt.moveRel(10,0)
        print('right move')
        while(True):
            x=au.readline()
            y=au.readline()
            x=int(x.decode().strip())
            y=int(y.decode().strip())
            if(x>60 and y>60):
                pt.moveRel(10,0)
            if(abs(x-y)<25 and x<75 and y<75):
                #time.sleep(0.4)
                break
    
    elif(20<x<35 and 35<y<60):
        r=sr.Recognizer()
        while(True):
            x=au.readline()
            y=au.readline()
            x=int(x.decode().strip())
            y=int(y.decode().strip())
            text=str()
            with sr.Microphone() as source:
                #sp.Speak('Give Instruction')
                winsound.Beep(frequency, duration)
                print('Give Instruction')
                #engi.runAndWait()
                r.adjust_for_ambient_noise(source)
                audio=r.listen(source)
                try:
                    text=r.recognize_google(audio)
                    print(text)
                    if(text.upper()=='VOLUME UP'):
                        pt.press('volumeup')
                    elif(text.upper()=='SEARCH'):
                        pt.press('enter')
                    elif(text.upper()=='SPACE'):
                        pt.press('space')
                    
                        #sp.Speak('Volume Increased')
                        #engi.runAndWait()
                    elif(text.upper()=='VOLUME DOWN'):
                        pt.press('volumedown')
                        #sp.Speak('Volume Decreased')
                        #engi.runAndWait()
                    elif(text.upper()=='SILENCE'):
                        pt.press('volumemute')
                        #sp.Speak('Mute')
                        #engi.runAndWait()
                    elif(text.upper()=='SELECT ALL'):
                        pt.hotkey('ctrl','a')
                        #sp.Speak('Selected All')
                        #engi.runAndWait()
                    elif(text.upper()=='COPY'):
                        pt.hotkey('ctrl','c')
                        #sp.Speak('Copied')
                        #engi.runAndWait()
                    elif(text.upper()=='PASTE'):
                        pt.hotkey('ctrl','v')
                        #sp.Speak('Pasted')
                        #engi.runAndWait()
                    elif(text.upper()=='MINIMISE'):
                        pt.hotkey('alt','space','n')

                    elif(text.upper()=='LEFT'):
                         pt.hotkey('left')

                    elif(text.upper()=='RIGHT'):
                         pt.hotkey('right')

                    elif(text.upper()=='UP'):
                         pt.hotkey('up')
             
                    elif(text.upper()=='DOWN'):
                         pt.hotkey('down')

                    elif(text.upper()=='EDIT'):
                         pt.hotkey('backspace')

                    elif(text.lower()=='ulta'):
                        pt.moveRel(-10,0)
                        while(True):
                            pt.moveRel(-10,0)
                            y=au.readline()
                            x=au.readline()
                            x=int(x.decode().strip())
                            y=int(y.decode().strip())
                            if(x<50 or y<50):
                                break

                    elif(text.lower()=='sidha'):
                        pt.moveRel(10,0)
                        while(True):
                            pt.moveRel(10,0)
                            y=au.readline()
                            x=au.readline()
                            x=int(x.decode().strip())
                            y=int(y.decode().strip())
                            if(x<50 or y<50):
                                break

                    elif(text.lower()=='upar'):
                        pt.moveRel(0,-10)
                        while(True):
                            pt.moveRel(0,-10)
                            y=au.readline()
                            x=au.readline()
                            x=int(x.decode().strip())
                            y=int(y.decode().strip())
                            if(x<50 or y<50):
                                break

                    elif(text.lower()=='niche'):
                        pt.moveRel(0,10)
                        while(True):
                            pt.moveRel(0,10)
                            y=au.readline()
                            x=au.readline()
                            x=int(x.decode().strip())
                            y=int(y.decode().strip())
                            if(x<50 or y<50):
                                break
                 
                    elif(text.upper()=='LEFT CLICK'):
                        pt.click(button='left')

                    elif(text.upper()=='RIGHT CLICK'):
                        pt.click(button='right')

                    elif(text.upper()=='NUMBER LOCK'):
                        pt.hotkey('numlock')

                    elif(text.upper()=='CAPS LOCK'):
                        pt.hotkey('capslock')

                    elif(text.upper()=='DELETE'):
                        pt.hotkey('del')


                    elif(text.upper()=='CLOSE'):
                        pt.hotkey('alt','f4')
                        #sp.Speak('Window Closed')
                        #engi.runAndWait()
                    elif(text.upper()=='SCREENSHOT'):
                        pt.press('PRTSC')
                    elif(text.upper()=='TYPING'):
                        #sp.Speak('Now you can type')
                        winsound.Beep(frequency, duration)
                        #time.sleep(0.6)
                        #engi.runAndWait()
                        r=sr.Recognizer()
                        with sr.Microphone() as source:
                            r.adjust_for_ambient_noise(source)
                            audio=r.listen(source)
                            text=r.recognize_google(audio)
                        pyperclip.copy(text)
                        #time.sleep(0.5)
                        pt.hotkey('ctrl','v')
                        print(text)
                except:
                    pass
            if(text.upper()=='GET OUT'):
                time.sleep(1)
                break
    elif(x<5 and y<5):
        sp.Speak('For Left click, put your hand on left Sensor between 10 to 25 centimeter')
        sp.Speak('For Right click, put your hand on right Sensor between 10 to 25 centimeter')
        sp.Speak('For Left Move, put your hand on left Sensor between 25 to 60 centimeter')
        sp.Speak('For right Move, put your hand on right Sensor between 25 to 60 centimeter')
        sp.Speak('For down Move, put your hands on both Sensors together between 10 to 25 centimeter')
        sp.Speak('For up Move, put your hands on both Sensors together between 25 to 60 centimeter')
        sp.Speak('For typing, keep your one hand on left sensor between 10 to 25 centimeter and other hand on right sensor together between 25 to 60 centimeter')
        sp.Speak('If you forget the rule than keep your hands on both the sensors, Rules will be repeated again')
