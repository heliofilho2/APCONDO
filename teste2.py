import pytesseract
from PIL import Image
import openpyxl

# Carregar a imagem da embalagem
imagem_embalagem = Image.open('testexx.png')

# Utilizar o Tesseract para extrair o texto da imagem
texto_extraido = pytesseract.image_to_string(imagem_embalagem)

# Abrir a planilha do Excel com os nomes dos moradores
planilha = openpyxl.load_workbook('PLANILHA.xlsx')
folha = planilha.active

# Criar uma lista para armazenar os nomes dos moradores
nomes_moradores = []

# Extrair os nomes dos moradores da planilha
for linha in folha.iter_rows(values_only=True):
    nome = linha[0]  # Supondo que o nome esteja na primeira coluna
    nomes_moradores.append(nome)

# Comparar os nomes extraídos da imagem com os nomes dos moradores na planilha
for nome in nomes_moradores:
    if nome.lower() in texto_extraido.lower():  # Comparação de strings, ignorando maiúsculas/minúsculas
        print("Nome do destinatário encontrado:", nome)
        # Aqui você pode continuar com o processamento, como preencher o sistema do condomínio, etc.

