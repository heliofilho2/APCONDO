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
from forms import abrir_forms_e_preencher

import tkinter as tk
from tkinter import font

from functools import partial
from processamento import iniciar_processamento

def criar_janela_registro_manual(nomes_moradores):
    # Função para criar a segunda janela de registro manual
    def registrar_manual(nomes_moradores):
        nome = entry_nome.get().strip().lower()  # Remover espaços em branco extras e converter para minúsculas
        print("Nome inserido manualmente:", nome)
        # Consultar a planilha de moradores para obter as informações do morador
        if nome in nomes_moradores:
            info_morador = nomes_moradores[nome]
            apartamento = info_morador["apartamento"]
            bloco = info_morador["bloco"]
            print(f"Nome: {nome}, Apartamento: {apartamento}, Bloco: {bloco}")
            abrir_forms_e_preencher(nome, info_morador)  # Abrir o formulário com as informações do morador
        else:
            print("Morador não encontrado na planilha de moradores.")
            print(nomes_moradores)

    # Configuração da segunda janela
    janela_registro_manual = tk.Toplevel()
    janela_registro_manual.title("Registro Manual")
    janela_registro_manual.geometry("600x300")
    janela_registro_manual.configure(bg="#F0F0F0")
    janela_registro_manual.resizable(False, False)

    fonte_titulo = font.Font(family="Helvetica", size=16, weight="bold")
    fonte_texto = font.Font(family="Helvetica", size=12)

    label_titulo = tk.Label(janela_registro_manual, text="Registro Manual", font=fonte_titulo, bg="#F0F0F0", fg="#333333")
    label_titulo.pack(pady=10)

    label_nome = tk.Label(janela_registro_manual, text="Nome:", font=fonte_texto, bg="#F0F0F0", fg="#333333")
    label_nome.pack()

    entry_nome = tk.Entry(janela_registro_manual, font=fonte_texto)
    entry_nome.pack()

    label_apartamento = tk.Label(janela_registro_manual, text="Apartamento:", font=fonte_texto, bg="#F0F0F0", fg="#333333")
    label_apartamento.pack()

    entry_apartamento = tk.Entry(janela_registro_manual, font=fonte_texto)
    entry_apartamento.pack()

    button_registrar = tk.Button(janela_registro_manual, text="Registrar", command=lambda: registrar_manual(nomes_moradores), bg="#4CAF50", fg="white", relief="flat")
    button_registrar.config(width=10, height=2, font=fonte_texto)
    button_registrar.pack(pady=20)




def criar_interface(iniciar_programa, nomes_moradores):
    def iniciar_programa_com_registro_manual():
        # Função para iniciar o programa com o registro manual
        criar_janela_registro_manual(nomes_moradores)

    root = tk.Tk()
    root.title("EntregAp")
    root.geometry("600x300")
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

    button_iniciar = tk.Button(root, text="Iniciar", command=lambda: iniciar_programa(com_camera=True), bg="#4CAF50", fg="white", relief="flat")
    button_iniciar.config(width=10, height=2, font=fonte_texto)
    button_iniciar.pack(pady=10)

    button_registro_manual = tk.Button(root, text="Registro Manual", command=iniciar_programa_com_registro_manual, bg="#FF5733", fg="white", relief="flat")
    button_registro_manual.config(width=15, height=2, font=fonte_texto)
    button_registro_manual.pack(pady=10)

    root.mainloop()


# Testar a interface
#criar_interface(lambda: print("Iniciar programa"))



    