import cv2
import pytesseract
import re
import nltk


from nltk import ne_chunk, pos_tag, word_tokenize
from nltk.tree import Tree

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('maxent_ne_chunker')
nltk.download('words')

# Configurações para o Tesseract OCR
myconfig = r"--psm 11 --oem 3"


imagem = cv2.imread("emb1.jpg")
texto = pytesseract.image_to_string(imagem, config=myconfig)
print("Texto extraído:")
print(texto)


padrao_nome = r"Nome (.+?)"
padrao_apartamento = r"Ap. (\d+)"
padrao_rua = r"Rua (\S+)"


match_nome = re.search(padrao_nome, texto)
match_apartamento = re.search(padrao_apartamento, texto)
match_rua= re.search(padrao_rua, texto)


# Extrai as informações encontradas
nome = match_nome.group(1) if match_nome else None
apartamento = match_apartamento.group(1) if match_apartamento else None
rua = match_rua.group(1) if match_rua else None

# Imprime as informações extraídas
print("Nome:", nome)
print("Apartamento:", apartamento)
print("Rua:", rua)



#nltk_results = ne_chunk(pos_tag(word_tokenize(texto)))
#for nltk_result in nltk_results:
 #   if type(nltk_result) == Tree:
  #      name = ''
   #     for nltk_result_leaf in nltk_result.leaves():
    #        name += nltk_result_leaf[0] + ' '
     #   print ('Type: ', nltk_result.label(), 'Name: ', name)