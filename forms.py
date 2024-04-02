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

def abrir_forms_e_preencher(nome, info):
    url_do_forms = "https://forms.gle/av9HDTPS1FBEmPJZ7"
    service = Service(executable_path='chromedriver.exe')
    driver = webdriver.Chrome(service=service)
    driver.get(url_do_forms)
    time.sleep(2)

    elemento_texto_nome = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
    elemento_apartamento = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
    elemento_bloco = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input')

    elemento_texto_nome.send_keys(nome)
    elemento_apartamento.send_keys(info["apartamento"])
    elemento_bloco.send_keys(info["bloco"])

    #time.sleep(2)
   # driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span').click()

    driver.quit()  