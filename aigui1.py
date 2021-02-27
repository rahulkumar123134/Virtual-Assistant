import src.SpeechRecognition.text
import src.wiki.wikipedia
import src.SearchWeb.search
import src.youtube.auto
import src.mail.mailgui
import src.app.open
import os
#import src.face.recognizer
#import src.wolframalpha.wolf
from tkinter import *
import cv2
import smtplib
import numpy as np
import datetime
import pyttsx3
from openpyxl import Workbook
from PIL import ImageTk,Image

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


    
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<16:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("Hi Rahul, Please tell me how may I help you")       


def recognize():

    recognizer = cv2.face.LBPHFaceRecognizer_create();
    recognizer.read('trainner/trained data.yml')
    faceCascade =cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

    book=Workbook()
    sheet=book.active

    sheet['A1'].value="UID"
    sheet['B1'].value="NAME"
    sheet['C1'].value="P/A"

    sheet['C2'].value="P"

    cam = cv2.VideoCapture(0)
    font                   = cv2.FONT_HERSHEY_SIMPLEX
    #bottomLeftCornerOfText = (10,500)
    fontScale              = 1
    fontColor              = (208,225,30)
    lineType               = 2

    # Load present date and time
    now= datetime.datetime.now()
    today=now.day
    month=now.month

    while True:
        ret, im =cam.read()
        if ret is True:
            gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
        else:
            continue
        faces=faceCascade.detectMultiScale(gray, 1.2,5)
        for(x,y,w,h) in faces:
            cv2.rectangle(im,(x,y),(x+w,y+h),(229,2,54),2)
            Id,conf = recognizer.predict(gray[y:y+h,x:x+w])
            if(Id==1):
                Id="Rahul"
                sheet['A2'].value="1"
                sheet['B2'].value="Rahul"
                sheet['C2'].value="P"
            
            elif(Id==2):
                Id="Anurag"
                sheet['A3'].value="2"
                sheet['B3'].value="Anurag"
                sheet['C3'].value="P"
            elif(Id==3):
                Id="Kabir"
                sheet['A4'].value="3"
                sheet['B4'].value="Kabir"
                sheet['C4'].value="P"
            elif(Id==4):
                Id="Priyam"
                sheet['A5'].value="4"
                sheet['B5'].value="Priyam"
                sheet['C5'].value="P"
            elif(Id==5):
                Id="Amal"
                sheet['A5'].value="5"
                sheet['B5'].value="Amal"
                sheet['C5'].value="P"
            else:
                Id="Unknown"
            cv2.putText(im,str(Id),(x,y+h),font,fontScale,fontColor,lineType)
        cv2.imshow('im',im) 
        book.save(str(today)+'_'+str(month)+'.xlsx')
        if cv2.waitKey(10) & 0xFF==ord('q'):
            break
    cam.release()
    cv2.destroyAllWindows()



def onClick():
    text1=e.get()
    Label(innerframe, text=text1,font=("Arial",13,"bold")).pack(side="top",anchor="nw")

    try:
    
        if("search" in text1):
            y = text1.replace('search','')
            print(y)
            speak("Searching " +y)
            answer=src.wiki.wikipedia.find(y)
            Label(innerframe, text=answer,font=("Arial",11,"bold")).pack(side="top")
        
        elif("songs" in text1):
            #y = text1.replace('play ','')
            #src.youtube.auto.find(y)
            music_dir = 'F:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[229]))
        
        elif("mail" in text1):
            src.mail.mailgui.send()
        
        elif("open" in text1 ):
            y = text1.replace('open ','')
            speak("openeing" + y)
            src.SearchWeb.search.find(y)
        
        elif("start" in text1):
            src.app.open.app(text1)
        
        elif("recognise" in text1):
            recognize()
        
        elif("what" and "can" and "you" and "do" in text1):
            speak("""I can do searches for you,        I can open camera for you,
                         i can schedule a task for you,         I can play songs for you,
                       i can open web for you,             so what do you want me to do.
                """)

        elif ('youtube' in text1.lower()): 
            speaks("Opening in youtube") 
            indx = text1.lower().split().index('youtube') 
            query = text1.split()[indx + 1:] 
            driver.get("http://www.youtube.com/results?search_query =" + '+'.join(query)) 

            
        elif("quit" in text1 or "close" in text1):
            speak("Quitting your personal assistance, Have a nice day Sir")
            exit()

    except:
        pass
        


    
def onClick1():
    text1=src.SpeechRecognition.text.ask()
    Label(innerframe, text=text1).pack(side="top",anchor="nw")

    try:
        
        if("search" in text1):
            y = text1.replace('search','')
            print(y)
            speak("Searching "+y)
            answer=src.wiki.wikipedia.find(y)
            Label(innerframe, text=answer,font=("Arial",11,"bold")).pack(side="top")
        
        elif("songs" in text1):
            #y = text1.replace('play ','')
            #src.youtube.auto.find(y)
            music_dir = 'F:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[229]))
        
        elif("mail" in text1):
            src.mail.mailgui.send()
        
        elif("open" in text1 ):
            y = text1.replace('open ','')
            speak("openeing" + y)
            src.SearchWeb.search.find(y)
                
        elif("start" in text1):
            src.app.open.app(text1)
        
        elif("recognise" in text1):
            recognize()

        elif("what" and "can" and "you" and "do" in text1):
            speak("""I can do searches for you,          I can open camera for you,
                    i can schedule a task for you,       I can play songs for you,
                       i can open web for you,         so what do you want me to do.
                """)

            
        elif ('youtube' in text1): 
            speaks("Opening in youtube") 
            indx = text1.lower().split().index('youtube') 
            query = text1.split()[indx + 1:] 
            driver.get("http://www.youtube.com/results?search_query =" + '+'.join(query)) 
            

        elif("quit" in text1 or "close" in text1):
            speak("Quitting your personal assistance, Have a nice day Sir")
            exit()


            
    except:
        pass


if __name__ == "__main__":
    wishMe()   
    
window=Tk()

canvas=Canvas(window,width=500,height=600)
image=ImageTk.PhotoImage(Image.open("F:\\College\\Semester 6\\Project\\virtual\\virtual\\bg.jpg"))
canvas.create_image(0,0,anchor=NW,image=image)
canvas.pack()

window.title("Virtual Assistant")
window.configure(background='black')
window.geometry("500x600")
text1=StringVar()
mainframe=Frame(window,bd=1,relief=GROOVE)
mainframe.place(x=10,y=10,height=450)


canvas=Canvas(mainframe)
innerframe=Frame(canvas)
hscroller= Scrollbar(mainframe,command=canvas.xview,orient="horizontal")
scroller= Scrollbar(mainframe,command=canvas.yview,orient="vertical")
scroller.pack(side="right",fill="y")
hscroller.pack(side="bottom",fill="x")


def EventFunc(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
    
canvas.configure(yscrollcommand=scroller.set)
canvas.pack(side="left")
canvas.create_window((0,0),window=innerframe,anchor="nw")
innerframe.bind("<Configure>",EventFunc)

canvas.configure(xscrollcommand=hscroller.set)


Label(window, text="Enter your command ").place(x=150,y=500)

Button(window,text="Search",width=6,command=onClick).place(x=300,y=525)
e=Entry(window,width=45, bg="white")
e.place(x=10,y=525)


photo=ImageTk.PhotoImage(Image.open("F:\\College\\Semester 6\\Project\\virtual\\virtual\\mic.png"))
Button(window,text="Voice",width=25,height=25,image=photo,command=onClick1).place(x=200,y=470)


window.mainloop()
