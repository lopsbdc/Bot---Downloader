#Ativando o módulo para fazer os passos de forma sincronizada com o navegador (faz e espera)
from playwright.sync_api import sync_playwright

#módulo de geração de log
import logging

#módulo para arredondar numero
import math

#módulo de tempo de espera
import time

import cred

#**************************** Inicio ****************************

#iniciando o navegador
with sync_playwright() as p:
    navegador = p.chromium.launch(headless=False)  #Headless = não fazer no modo invisvel.
    pagina = navegador.new_page()

    #configs finais do navegador
    pagina.goto("https://empresa.holmesdoc.com/#home")

    pagina.locator('//*[@id="tiUser"]').fill(cred.email)
    pagina.locator('//*[@id="tiPass"]').fill(cred.senha)
    pagina.locator('//*[@id="login"]/div[4]/button').click()

    #log config básico. W é de Write, A de append. W ele apaga o ultimo log, A, ele vai somando
    logging.basicConfig(filename='Downloader.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s')

    #entrar na página que precisa baixar (pagina de erros)
    pagina.goto("https://empresa.holmesdoc.com/#search/+_nature%253A%2522Incompleto%2522%2520+_uploaddate%253A%255B20161128%2520TO%2520*%255D/p1/sortedby/_uploaddate/DESC/GRID")

    #Calcular quantidade de páginas
    qtd = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
    #tratando os dados
    qtd0 = qtd.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
    qtd1 = int(qtd0)
    qtd2 = math.ceil(qtd1/54)  #arredondar para cima
    #aviso no log
    logging.warning('************ - Total de páginas a serem baixadas: ' + str(qtd2) + " - ************")

    i = 1

    #iniciar a repetição, de acordo com o tamanho da tabela (usando enumerate)
    while i <= int(qtd2):

        try:
            caminhox = ('//*[@id="' + str(i) + '"]')
            pagina.locator(caminhox).click()
            time.sleep(4)

            #renomeador
            nome = ('Pagina ' + str(i))

            #identificando onde ta a informação no excel, e o i, indica a linha. Conversão em String para evitar erro
            pagina.locator('//*[@id="trackbar"]/div[7]/div/div[7]/span[1]').click()
            time.sleep(1)
            pagina.locator('//*[@id="trackbar"]/div[7]/div/div[4]/span[1]').click()

            with pagina.expect_download() as download_info:
                pagina.locator('//*[@id="holmes"]/div[7]/div[2]/span/button').click()
                download = download_info.value
                # selecionando o caminho
                download.save_as(r'C:\Downloads\ ' + nome + ".zip")

            #salvando no log em texto, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('Página ' + str(i) + " foi baixada com sucesso!")

            time.sleep(1)

        except:
            #salvando erro no log, de acordo com o numero da filial na planilha (usado como identificador)
            logging.warning('Erro ao baixar a Página ' + str(i) + " !")

        i = i+1


    #salvando no log,informando que finalizou o bot
    logging.warning('Bot finalizado')
