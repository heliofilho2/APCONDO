import os
import glob
import zipfile
import cv2
import pytesseract
import openpyxl
import time
import tkinter as tk
import difflib


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

#from utils import criar_planilha_entrega, registrar_entrega
from tkinter import messagebox



def iniciar_processamento(nomes_moradores, abrir_forms_e_preencher):

    camera = cv2.VideoCapture(0)

    while True:
        validacao, frame = camera.read()

        if not validacao:
            print("Erro ao capturar o próximo frame.")
            break

        cv2.imshow("Video", frame)

        preprocessed_frame = preprocessar_imagem(frame)
        texto_extraido = pytesseract.image_to_string(preprocessed_frame)

        for nome, info in nomes_moradores.items():
            if nome in texto_extraido.lower():
                abrir_forms_e_preencher(nome, info)

        key = cv2.waitKey(1)
        if key == 27:
            break

    camera.release()
    cv2.destroyAllWindows()


def preprocessar_imagem(imagem):
    gray = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)

    # Aplicar suavização para reduzir o ruído
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    return binary

    