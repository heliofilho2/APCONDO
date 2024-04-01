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

# Carregar a imagem da embalagem
#imagem_embalagem = Image.open('emb1.jpg')

# Configurações para o Tesseract OCR
myconfig = r"--psm 11 --oem 3"

# Inicializa a câmera
camera = cv2.VideoCapture(0)

# Verifica se a câmera foi aberta corretamente
if not camera.isOpened():
    print("Erro ao abrir a câmera.")
    exit()

# Loop para captura de vídeo
while True:
    # Captura o próximo frame
    validacao, frame = camera.read()
    
    # Verifica se o frame foi capturado corretamente
    if not validacao:
        print("Erro ao capturar o próximo frame.")
        break
    
    # Mostra o vídeo em uma janela
    cv2.imshow("Video", frame)
    
    # Verifica se a tecla ESC foi pressionada
    key = cv2.waitKey(1)
    if key == 27:  # ESC
        # Salva a imagem capturada
        cv2.imwrite("testexx.png", frame)
        break

# Libera a câmera e fecha todas as janelas
camera.release()
cv2.destroyAllWindows()


# Utilizar o Tesseract para extrair o texto da imagem
texto_extraido = pytesseract.image_to_string("testexx.png")

# Abrir a planilha do Excel com os nomes dos moradores
planilha = openpyxl.load_workbook('PLANILHA.xlsx')
folha = planilha.active

# Criar uma lista para armazenar os nomes dos moradores
nomes_moradores = {}

# Extrair os nomes dos moradores da planilha
for linha in folha.iter_rows(values_only=True):
    nome = linha[0]  # Supondo que o nome esteja na primeira coluna
    #nomes_moradores.append(nome)
    nomes_moradores[nome.lower()] = {"apartamento": linha[1], "bloco": linha[2]}

# Comparar os nomes extraídos da imagem com os nomes dos moradores na planilha
for nome, info in nomes_moradores.items():
    if nome in texto_extraido.lower():
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
        break  # Parar de procurar assim que uma correspondência for encontrada
