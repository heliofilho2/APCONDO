import pandas as pd
import pytesseract
from PIL import Image
import openpyxl
import time 
import cv2

from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
# Carregar a planilha do Excel com os nomes dos moradores
planilha = openpyxl.load_workbook('PLANILHA.xlsx')
folha = planilha.active

# Criar um dicionário para armazenar os nomes dos moradores
nomes_moradores = {}
for linha in folha.iter_rows(values_only=True):
    nome = linha[0].lower()  # Supondo que o nome esteja na primeira coluna e transformando para minúsculas
    nomes_moradores[nome] = {"apartamento": linha[1], "bloco": linha[2]}

# Inicializar a câmera
camera = cv2.VideoCapture(0)

while True:
    # Capturar o próximo frame da câmera
    validacao, frame = camera.read()

    # Verificar se o frame foi capturado corretamente
    if not validacao:
        print("Erro ao capturar o próximo frame.")
        break

    # Exibir o frame na janela
    cv2.imshow("Video", frame)

    # Converter o frame para escala de cinza para melhorar o desempenho do OCR
    frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # Realizar OCR para extrair o texto do frame
    texto_extraido = pytesseract.image_to_string(frame_gray)

    # Comparar os nomes extraídos com os nomes dos moradores na planilha
    for nome, info in nomes_moradores.items():
        if nome in texto_extraido.lower():  # Converter para minúsculas para garantir a comparação
            print("Nome do destinatário encontrado:", nome)
            print("Apartamento:", info["apartamento"])
            print("Bloco:", info["bloco"])

            # Abrir o formulário no navegador
            url_do_forms = "https://forms.gle/av9HDTPS1FBEmPJZ7"
            service = Service(executable_path='chromedriver.exe')
            driver = webdriver.Chrome(service=service)
            driver.get(url_do_forms)
            time.sleep(10)

            # Preencher as informações do formulário com base na correspondência encontrada
            elemento_texto_nome = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
            elemento_apartamento = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
            elemento_bloco = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input')

            elemento_texto_nome.send_keys(nome)
            elemento_apartamento.send_keys(info["apartamento"])
            elemento_bloco.send_keys(info["bloco"])

            time.sleep(3)
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span').click()
            driver.quit()  # Fechar o navegador após o preenchimento

    # Verificar se a tecla ESC foi pressionada para sair do loop
    key = cv2.waitKey(1)
    if key == 27:  # ESC
        break

# Liberar a câmera e fechar todas as janelas
camera.release()
cv2.destroyAllWindows()
