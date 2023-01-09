import pyttsx3, docx2txt, os
from pdfminer.high_level import extract_text
from settings import VOLUME, VOICE, RATE
from tkinter import *
from tkinter import filedialog

choice = 0
#Function that allows the user to select his PDF file
def get_file_path_and_type():
    file = filedialog.askopenfile(title="Please select a file to convert.")
    if file:
        filePath = os.path.abspath(file.name)
        if ".pdf" in filePath:
            choice = 0
            return filePath, choice
        elif ".docx" in filePath:
            choice = 1
            return filePath, choice
        else:
            print("Invalid file type")

#Function that modifies the settings of the TTS voice using variables from the settings.py file
def voice_settings():
    voices = engine.getProperty('voices')
    engine.setProperty("voice", voices[VOICE].id)
    engine.setProperty("volume", VOLUME)
    engine.setProperty("rate", RATE)

#Extracts text from the PDF gotten from get_file_path_and_type

def convert_text():
    if (choice == 0):
        text = extract_text(filePath)
        return text
    elif (choice == 1):
        text = docx2txt.process(filePath)
        return text

#Creates a Tkinter instance, hides it and changes the icon
root = Tk()
root.withdraw()
root.iconbitmap("icon.ico")

#Creates a PyTTSX3 instance
engine = pyttsx3.init()

#Calls voice_settings function
voice_settings()

#Calls the get_file_path_and_type function
filePath, choice = get_file_path_and_type()

#Calls the convert_text function
text = convert_text()

#Saves the TTS to an MP3 file
engine.save_to_file(text, "audio.mp3")

#Runs the PyTTSX3 module and waits until the file creation finishes
engine.runAndWait()