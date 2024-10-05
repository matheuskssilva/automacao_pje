import tkinter as tk
from tkinter import ttk
from tkinter import messagebox  # Importa a biblioteca de mensagens
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.select import Select
import openpyxl
import os
import subprocess

# Função que inicia a automação após capturar os valores de entrada
def iniciar_automacao(oab, uf):
    numero_oab = oab
    sigla_uf = uf

    # Carrega a planilha de consulta
    planilha_dados_consulta = openpyxl.load_workbook('dados_da_consulta.xlsx')
    pagina_processos = planilha_dados_consulta['processos']

    # Inicia o navegador
    driver = webdriver.Chrome()
    driver.get('https://pje-consulta-publica.tjmg.jus.br/')
    sleep(5)

    # Localiza o campo OAB e insere o valor
    campo_numero_oab = driver.find_element(By.XPATH, "//input[@id='fPP:Decoration:numeroOAB']")
    sleep(2)
    campo_numero_oab.click()
    sleep(1)
    campo_numero_oab.send_keys(numero_oab)
    sleep(1)

    # Seleciona a UF
    selecao_uf = driver.find_element(By.XPATH, "//select[@id='fPP:Decoration:estadoComboOAB']")
    sleep(1)
    opcoes_uf = Select(selecao_uf)
    sleep(1)
    opcoes_uf.select_by_visible_text(sigla_uf)
    sleep(1)

    # Clica no botão pesquisar
    campo_pesquisar = driver.find_element(By.XPATH, "//input[@id='fPP:searchProcessos']")
    sleep(1)
    campo_pesquisar.click()
    sleep(5)

    # Captura os links dos processos
    links_abrir_processos = driver.find_elements(By.XPATH, "//a[@title='Ver Detalhes']")

    for link in links_abrir_processos:
        janela_principal = driver.current_window_handle
        link.click()
        sleep(6)

        # Captura todas as janelas abertas
        janelas_abertas = driver.window_handles
        for janela in janelas_abertas:
            if janela != janela_principal:
                driver.switch_to.window(janela)
                sleep(6)

                # Captura o número do processo
                numero_processo_element = driver.find_elements(By.XPATH, "//div[@class='propertyView']//div[@class='col-sm-12']")
                if numero_processo_element:
                    numero_processo = numero_processo_element[0].text

                    # Captura os nomes dos participantes
                    nome_participantes = driver.find_elements(By.XPATH, "//tbody[contains(@id,'processoPartesPoloAtivoResumidoList:tb')]//span[@class='text-bold']") 

                    lista_participantes = []
                    for participante in nome_participantes:
                        lista_participantes.append(participante.text)

                    # Verifica se há dados a serem salvos
                    if len(lista_participantes) == 1:
                        pagina_processos.append([numero_oab, numero_processo, lista_participantes[0]])
                    else:
                        pagina_processos.append([numero_oab, numero_processo, ','.join(lista_participantes)])                 

                    # Salva os dados na planilha
                    planilha_dados_consulta.save('dados_da_consulta.xlsx')

                driver.close()
                sleep(3)
        driver.switch_to.window(janela_principal)

# Função que pega os valores dos inputs e inicia a automação
def infos_inputs():
    value_oab = entrada_oab.get()
    value_uf = entrada_uf.get().upper()

    # Verifica se os campos estão preenchidos
    if value_oab == '' or value_uf == '':
        messagebox.showerror("Erro", "Preencha todos os campos!")  # Mensagem de erro
        return
    
    # Atualiza o status na interface
    label_status.config(text="Iniciando automação...")
    progress['value'] = 0
    janela.update_idletasks()

    # Inicia a automação passando os valores
    iniciar_automacao(value_oab, value_uf)
    
    # Atualiza o status após conclusão
    label_status.config(text="Automação concluída com sucesso!")
    progress['value'] = 100
    janela.update_idletasks()
    
    ver_planilha_button.pack(pady=10)  # Exibe o botão após a conclusão da automação

def ver_planilha():
    caminho_planilha = r'C:\Users\mathe\Desktop\automação de consulta\dados da consulta.xlsx'

    # Verifica se o arquivo existe
    if os.path.exists(caminho_planilha):
        # Tenta abrir o arquivo com o aplicativo padrão
        try:
            subprocess.Popen(['start', caminho_planilha], shell=True)
        except Exception as e:
            print(f"Erro ao abrir o arquivo: {e}")
    else:
        print("Arquivo não encontrado!")

# Interface gráfica com Tkinter
janela = tk.Tk()
janela.title('Automação de Consulta PJE')
janela.geometry('400x400')

janela.configure(bg="#f0f0f0") 

label_oab = tk.Label(janela, text="Digite o número OAB", font=("Arial", 12, "bold"), fg="#333")
label_oab.pack(pady=10)  

entrada_oab = tk.Entry(janela, width=30, font=("Arial", 12), bg="#e0e0e0")
entrada_oab.pack(pady=10) 

label_uf = tk.Label(janela, text="Digite o UF", font=("Arial", 12, "bold"), fg="#333")
label_uf.pack(pady=10)

entrada_uf = tk.Entry(janela, width=30, font=("Arial", 12), bg="#e0e0e0")
entrada_uf.pack(pady=10) 

# Botão para capturar os dados e iniciar a automação
button = tk.Button(janela, text="Enviar", font=("Arial", 12, "bold"), bg="#0093E9", fg="white", relief="raised", command=infos_inputs)
button.pack(pady=10)

# Label para status
label_status = tk.Label(janela, text="", font=("Arial", 10), fg="#333", bg="#f0f0f0")
label_status.pack(pady=10)

# Barra de progresso
progress = ttk.Progressbar(janela, orient='horizontal', length=300, mode='determinate')
progress.pack(pady=10)

# Botão para ver a planilha
ver_planilha_button = tk.Button(janela, text='Ver as Alterações na Planilha', font=("Arial", 12, "bold"), bg="#0093E9", fg="white", relief="raised", command=ver_planilha)
ver_planilha_button.pack_forget()  # Inicialmente oculta o botão

# Inicia o loop principal da janela
janela.mainloop()
