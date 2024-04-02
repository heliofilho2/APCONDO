import os
import glob
import zipfile
import cv2
import pytesseract
import openpyxl
import time
import tkinter as tk

from datetime import datetime
from tkinter import font
from PIL import Image, ImageTk
from winotify import Notification, audio
from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service


import tkinter as tk
from tkinter import font

from functools import partial
from processamento import iniciar_processamento

def criar_interface(iniciar_programa):
    root = tk.Tk()
    root.title("Tela Inicial")
    root.geometry("900x300")
    root.configure(bg="#F0F0F0")
    root.resizable(False, False)
    
    fonte_titulo = font.Font(family="Helvetica", size=16, weight="bold")
    fonte_texto = font.Font(family="Helvetica", size=12)

    icone_caixa = Image.open("caixa.png")
    icone_caixa = icone_caixa.resize((100, 100))
    icone_caixa = ImageTk.PhotoImage(icone_caixa)

    label_icone = tk.Label(root, image=icone_caixa, bg="#F0F0F0")
    label_icone.pack(pady=10)

    label = tk.Label(root, text="Bem-vindo!", font=fonte_titulo, bg="#F0F0F0", fg="#333333")
    label.pack()

    instrucoes = tk.Label(root, text="Clique em 'Iniciar' para começar.", font=fonte_texto, bg="#F0F0F0", fg="#666666")
    instrucoes.pack()

    button = tk.Button(root, text="Iniciar", command=iniciar_programa, bg="#4CAF50", fg="white", relief="flat")
    button.config(width=10, height=2, font=fonte_texto)
    button.pack(pady=20)

    root.mainloop()


    