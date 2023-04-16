# import required modules
import tkinter.filedialog
from tkinter import *
import customtkinter
import pyttsx3
import docx2txt
from  tkinter import  ttk
from tkinter.filedialog import askdirectory
import tkinter as tk

# initialize text-to-speech engine
txt = pyttsx3.init()
win = Tk()
win.geometry("800x600")
win.title("SpeechEase")
win.configure(bg="#252A4B")
sidebar = Label(win,bg="#202442",padx=80,pady=600)
sidebar.place(x=0,y=0)
win.resizable(False,False)
win.iconbitmap('C:\\Users\\hp\\Documents\\tkinter\\Txtv2\\Txt.ico')

# create a label widget
label = customtkinter.CTkLabel(master=win,text="Voice setting",text_font="Courier",bg="#202442",text_color="white")
label.place(x=10,y=200)

# define a function to clear the textbox
def cle_r():
    box.textbox.delete("0.1",END)

# define a function to save the audio file
def saver():
    inp = box.textbox.get("0.1",END)
    path =  tkinter.filedialog.asksaveasfilename(title="Save as",filetypes=(("MP3 files", "*.mp3"), ("WAV files", "*.wav"), ("All files", "*.*"))
)
    txt.save_to_file(inp,path)
    txt.runAndWait()

# define a function to read the text aloud
def say():
    inp = box.textbox.get("0.1", END)
    txt.say(inp)
    txt.runAndWait()

# define a function to select the female voice
def female():
    voices = txt.getProperty('voices')
    txt.setProperty('voice', voices[1].id)

    txt.runAndWait()

# define a function to select the male voice
def male():
    voices = txt.getProperty('voices')
    txt.setProperty('voice', voices[0].id)

    txt.runAndWait()

# define a function to open a file
def open_file():
    file_path = tk.filedialog.askopenfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    if file_path:
        with open(file_path, 'rb') as doc:
            text = docx2txt.process(doc)
            box.textbox.delete("1.0", "end")
            box.textbox.insert("end", text)

# create a textbox widget
box = customtkinter.CTkTextbox(master=win,fg_color="#2D325A",width=520,height=242,border_color="#7033FF",border_width=1,corner_radius=10,text_color="white",
  text_font="Roboto"                                 )
box.place(x=202,y=113)


# create button widgets

b1 = customtkinter.CTkButton(master=win,text_color="white",text_font="Courier",fg_color="#2D325A",bg_color="#252A4B",text="Clear",corner_radius=10,
         border_color="#43B5B5",border_width=1,command=cle_r)
b1.place(x=435,y=391)
b2 = customtkinter.CTkButton(master=win,text_color="white",text_font="Courier",fg_color="#2D325A",bg_color="#252A4B",text="Read aloud",corner_radius=10,
         border_color="#43B5B5",border_width=1,command=say)
b2.place(x=636,y=390)
b3 = customtkinter.CTkButton(master=win,text_color="white",text_font="Courier",fg_color="#2D325A",bg_color="#252A4B",text="Save audio",corner_radius=10,
         border_color="#FFEA2D",border_width=1,command=saver,hover_color="#DED581")
b3.place(x=10,y=120)
b4 = customtkinter.CTkButton(master=win,text_color="white",text_font="Courier",fg_color="#2D325A",bg_color="#252A4B",text="Open text",corner_radius=10,
         border_color="#FF8412",border_width=1,command=open_file,hover_color="#DE6F53")
b4.place(x=10,y=45)

b5 = customtkinter.CTkButton(master=win,text_color="white",text_font="Courier",fg_color="#2D325A",bg_color="#252A4B",text="Female",corner_radius=10,
         border_color="#83f52c",border_width=1,command=female,hover_color="#CBF5AB")
b5.place(x=10,y=250)

b6 = customtkinter.CTkButton(master=win,text_color="white",text_font="Courier",fg_color="#2D325A",bg_color="#252A4B",text="Male",corner_radius=10,
         border_color="#006369",border_width=1,command=male)
b6.place(x=10,y=300)

win.mainloop()