from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import urllib
import pandas as pd
from datetime import datetime

# Função para obter a saudação com base no horário atual
def obter_saudacao():
    agora = datetime.now()
    hora = agora.hour
    if hora < 12:
        return "Bom dia"
    elif hora < 18:
        return "Boa tarde"
    else:
        return "Boa noite"

# Inicializa o navegador Chrome
navegador = webdriver.Chrome()

# Carrega os dados do arquivo Excel
contatos_df = pd.read_excel("lista_contatos.xlsx")

# Abre o WhatsApp Web
navegador.get("https://web.whatsapp.com/")

# Espera até que o elemento 'side' esteja visível na página
wait = WebDriverWait(navegador, 30)
wait.until(EC.visibility_of_element_located((By.ID, "side")))

# Loop para enviar mensagens para cada contato
for i, pessoa in contatos_df.iterrows():
    nome_completo = pessoa['Nome']
    primeiro_nome = nome_completo.split()[0]  # Obtém apenas o primeiro nome
    numero = pessoa['Telefone']
    
    # Verifica se o número de telefone está presente e não vazio
    if numero and str(numero).strip():
        try:
            # Constrói a mensagem personalizada
            mensagem = f"{obter_saudacao()} {primeiro_nome}." #Personaliza a mensagem.
            texto = urllib.parse.quote(mensagem)
            
            # Abre o link do WhatsApp Web com a mensagem
            link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
            navegador.get(link)
            
            # Aguarda até que o campo de entrada de texto esteja visível
            campo_mensagem = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p')))
            
            # Envia a mensagem pressionando Enter
            campo_mensagem.send_keys(Keys.ENTER)
            
            # Aguarda um tempo antes de enviar a próxima mensagem
            time.sleep(30)
        except Exception as e:
            print(f"Erro ao enviar mensagem para {nome_completo}: {e}")
            continue
    else:
        print(f"Número de telefone ausente ou inválido para {nome_completo}. Pulando para o próximo.")

# Fecha o navegador ao final do processo
navegador.quit()
