import cv2
import pytesseract

# Configurações para o Tesseract OCR
myconfig = r"--psm 6 --oem 3"

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
        cv2.imwrite("foto.png", frame)
        break

# Libera a câmera e fecha todas as janelas
camera.release()
cv2.destroyAllWindows()

# Lê o texto da imagem usando o Tesseract OCR
imagem = cv2.imread("foto.png")
#imagem_cinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
#imagem_threshold = cv2.threshold(imagem_cinza, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
texto = pytesseract.image_to_string(imagem, lang = 'por' , config=myconfig)
print("Texto extraído:")
print(texto)
