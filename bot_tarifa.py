import tkinter as tk                        #Criar interfaces   
import win32com.client as win32             #Manipular o windows
import sys                                  #Manipular informações do sistemas
import pandas as pd                         #Analise de dados (tabelas)
import fitz                                 #Interface Mupdf, que permite trabalhar com pdf
import PyPDF2                               #Trabalha com pdf
import re                                   #Padrão de busca, expressões regulares
import os                                   #Interagir com o sistema operacional
from tkinter import filedialog, messagebox  #Criar dialogo na interface do tkinter

# codigo de executavel (python -m PyInstaller --onefile bot_tarifa.py)

def merge_pdf(pdf_files, output_path):
    pdf_merger = PyPDF2.PdfMerger()
    for pdf_file in pdf_files:
        pdf_merger.append(pdf_file)
    with open(output_path, 'wb') as output_file:
        pdf_merger.write(output_file)

def process_pdfs(email, password, folder_path):
    # Lista de arquivos na pasta
    pdf_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.lower().endswith('.pdf')]

    if not pdf_files:
        messagebox.showinfo("Info", "Não foram encontrados arquivos PDF na pasta especificada.")
        return

    for pdf_file in pdf_files:
        print("Executando código para:", pdf_file)
        # leitor de PDF
    # Abra o arquivo PDF em modo de leitu
        valorTarifa_total = 0
        data_count = 0
        with open(pdf_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)

            for page in pdf_reader.pages:
                text = page.extract_text()
                lines = text.split('\n')
                
                # print(lines) # Imprime todas as linhas do PDF para depuração
                
                for i in range(len(lines) - 1):
                    match = re.search(r'\d{2}/\d{2}/\d{4}', lines[i])
                    if match:
                        data_count += 1
                        if data_count == 4:
                            break
                    if ("TARIFA" in lines[i] or "DOC/TED INTERNET" in lines[i]) and ("TARIFA OPERACAO" not in lines[i] and "DOC/TED INTERNET" not in lines[i+1]):
                        numero_negativo = re.findall(r'-\d+,\d+', lines[i + 1])
                        if numero_negativo:
                            for num in numero_negativo:
                                num = num.replace(',', '.')
                                valorTarifa_total += float(num) * -1
                                print("-" * 60)
                                print(lines[i + 1])
                                print("o valor da tarifa é: " + str(num))
                    if "ENCARGOS DESCOBERTO" in lines[i] or "TAR " in lines[i] or "TARIFA OPERACAO" in lines[i]:
                        match = re.search(r'-\d+,\d+', lines[i])
                        if match:
                            num = match.group().replace(',', '.')  # Substituindo ',' por '.'
                            valorTarifa_total += float(num) * -1  # Somando ao total das tarifas
                            print("-" * 60)
                            print(lines[i])  # Imprimindo a linha atual
                            print("O valor da tarifa é: " + str(num))  
                                
                #Encontrar numero da Conta e FIlial
            
                    if "CC:" in lines[i]:
                        texto = lines[i]
                        if "CC" in lines[i]:
                            partes = texto.split("|")  # Dividir o texto em partes usando o caractere "|"
                            for parte in partes:
                                if "CC:" in parte:
                                    numero_cc = parte.split(":")[1].split()[-1]  #
                                    numero_cc_limpo = numero_cc.replace("-", "")  # Remove o caractere "-"
                                    numeros = []  # Lista para armazenar os números
                                    encontrado = False
                                    for digito in numero_cc_limpo:
                                        if digito != "0":
                                            encontrado = True
                                        if encontrado:
                                            numeros.append(digito)
                                    if numeros:
                                        numero_cc_formatado = "".join(numeros)  # Converte a lista de números em uma string
                                        print("Número após 'CC': ", numero_cc)
                                    break

                        # Encontrar o índice do primeiro dígito diferente de zero
                        indice_primeiro_numero = next((i for i, x in enumerate(numero_cc_formatado) if x != '0'), None)

                        if indice_primeiro_numero is not None:
                            # Se houver dígitos diferentes de zero, remover os zeros à esquerda até esse índice
                            numero_cc_formatado = numero_cc_formatado[indice_primeiro_numero:]
                        numero_cc_formatado = numero_cc_formatado[:-1]
                        print("O código da Filial é:", numero_cc_formatado)
                
                            #bot filial
                        
                        tabelaFIliais = [
                            ('35655', 'BANCO BRADESCO S.A', 'RECIFE'),
                            ('12260', 'BANCO BRADESCO S.A', 'CORDEIROPOLI'),
                            ('10162', 'BANCO BRADESCO S.A', 'GUARULHOS'),
                            ('11700', 'BANCO BRADESCO S.A', 'BENEVIDES'),
                            ('71644', 'BANCO BRADESCO S.A', 'BENEVIDES'),
                            ("90718", 'BANCO BRADESCO S.A', 'FORTALEZA'),
                            ('7326', 'BANCO BRADESCO S.A', 'PARAUAPEBAS'),
                            ('423', 'BANCO BRADESCO S.A', 'SIMOES FILHO'),
                            ('461', 'BANCO BRADESCO S.A', 'MINAS GERAIS'),
                            ('18885', 'BANCO BRADESCO S.A', 'TRANSULPARTI'),
                            ('27830', 'BANCO DO BRASIL S.A', 'MATRIZ'),
                            ('610', 'CAIXA ECONOMICA FEDERAL', 'MATRIZ'),
                            ('554', 'BANCO  BRADESCO S.A', 'MATRIZ'),
                            ('543','BANCO BRADESCO S.A', 'MATRIZ'),
                            ('55000', 'BANCO BRADESCO S.A', 'MATRIZ'),
                            ('25779', 'BANCO BRADESCO S.A', 'MATRIZ'),
                            ('18200', 'BANCO BRADESCO S.A', 'MATRIZ'),
                            ('12473', 'BANCO ITAUCARD S.A.', 'MATRIZ'),
                            ('10293', 'BANCO BRADESCO S.A', 'MATRIZ'),
                            ('20540', 'BANCO ITAUCARD S.A.', 'MATRIZ'),
                            ('28169', 'BANCO DO NORDESTE DO BRASIL S.A', 'MATRIZ'),
                            ('130001633', 'BANCO SANTANDER (BRASIL) S.A.', 'MATRIZ'),
                            ('18180', 'BANCO BRADESCO S.A', 'TRANSULOG MA'),
                            ('25123', 'BANCO BRADESCO S.A', 'RODONORTE MT'),
                            ('25122', 'BANCO BRADESCO S.A', 'RODONORTE MT'),
                            ('505', 'BANCO BRADESCO S.A', 'MATRIZ'),
                            ('1120', 'BANCO DO BRASIL S.A', 'MATRIZ')
                        ]
                        filial_df = pd.DataFrame(tabelaFIliais, columns=['Número', 'Banco', 'Filial'])
                        
                        linha = filial_df[filial_df['Número'] == numero_cc_formatado]

                        if not linha.empty:
                            # Obtendo o terceiro elemento da linha encontrada
                            filial = linha.iloc[0, 2]  # O terceiro elemento está na terceira coluna (índice 2)
                            fornecedor = linha.iloc[0,1]
                            print(f"o fornecedor é: " + fornecedor)

                            print("a filial é {}: {}".format(numero_cc_formatado, filial))
                        else:
                            print("Valor não encontrado na tabela.")       #bot Emissão de data
                        def encontrar_linha_lancamentos(texto_pdf):
                            # Dividir o texto em linhas
                            linhas = texto_pdf.split('\n')

                            # Iterar sobre as linhas para encontrar aquela que contém os lançamentos
                            for i, linha in enumerate(linhas):
                                if 'SALDO' in linha:
                                    return i + 1  # Retorna o índice da próxima linha após "SALDO"

                        def encontrar_primeira_data(texto):
                            # Encontrar todas as ocorrências de data no texto
                            datas_encontradas = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', texto)
                            # Retornar a primeira data encontrada, se houver alguma
                            return datas_encontradas[0] if datas_encontradas else None
                        
                        #Caminho do arquivo PDF
                        caminho_arquivo_pdf = pdf_file

                        #Inicialize uma variável para armazenar o texto a partir da linha de lançamentos
                        texto_lancamentos = ""

                        #Abra o arquivo PDF
                        with fitz.open(caminho_arquivo_pdf) as documento_pdf:
                            # Itere sobre cada página do PDF
                            for pagina_atual in documento_pdf:
                                # Extrair o texto da página
                                texto_pagina = pagina_atual.get_text()

                                # Tenta encontrar a linha que contém os lançamentos
                                linha_inicio_lancamentos = encontrar_linha_lancamentos(texto_pagina)
                                if linha_inicio_lancamentos:
                                    # Adiciona todas as linhas abaixo da linha de lançamentos ao texto_lancamentos
                                    texto_lancamentos += "\n".join(texto_pagina.split('\n')[linha_inicio_lancamentos:])
                                    texto_lancamentos += "\n"  # Adiciona uma quebra de linha entre as páginas


                        #Encontrar e imprimir a primeira data nos lançamentos no formato XX/XX/XXXX
                        primeira_data = encontrar_primeira_data(texto_lancamentos)
                        if primeira_data:
                            print("'Data da emissão:", primeira_data)
                        else:
                            print("Nenhuma data encontrada nos lançamentos.")
                
                            

                                        #bot de soma de tarifa
                                
            total_formatado = round(valorTarifa_total, 3)
            print("-" * 60)
            print("a soma de todas as tarifas: "+ str(total_formatado))
            
            from selenium import webdriver
            import pyautogui
            from pynput.mouse import Controller, Button
            from webdriver_manager.chrome import ChromeDriverManager
            from selenium.webdriver.chrome.service import Service
            import time
            from selenium.webdriver.common.by import By
            from selenium.webdriver.common.action_chains import ActionChains

            servico =  Service(ChromeDriverManager().install())
            navegador = webdriver.Chrome(service=servico)
            
            navegador.maximize_window()


            navegador.get("https://webtrans.saas.gwsistemas.com.br/")
            navegador.find_element(By.XPATH, '//*[@id="login"]').send_keys(email)
            navegador.find_element(By.XPATH, '//*[@id="senha"]').send_keys(password)
            navegador.find_element(By.XPATH, '//*[@id="form_login"]/div[3]/div/button').click()
            time.sleep(3)

            navegador.get("https://webtrans.saas.gwsistemas.com.br/consultadespesa?acao=iniciar")
            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="novo"]').click()
            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="especie"]').send_keys("FIN")

            navegador.find_element(By.XPATH, '//*[@id="localiza_filial"]').click()
            time.sleep(3)
            pyautogui.write(filial)
            time.sleep(3)
            pyautogui.press("enter")
            time.sleep(4)
            pyautogui.click(289,439)
            time.sleep(8)


            navegador.find_element(By.XPATH, '//*[@id="localiza_fornecedor"]').click()
            time.sleep(4)
            pyautogui.write(fornecedor)
            time.sleep(3)
            time.sleep(2)
            pyautogui.press("enter")
            time.sleep(10)
            
            pyautogui.click(289,439)
            time.sleep(5)

            navegador.find_element(By.XPATH, '//*[@id="descricao_historico"]').send_keys("TARIFA BANCARIA")
            time.sleep(3)

            elemento_data = navegador.find_element(By.XPATH, '//*[@id="dtemissao"]')
            actions = ActionChains(navegador)
            actions.click(elemento_data)
            actions.click(elemento_data)
            actions.click(elemento_data)
            actions.perform()
            pyautogui.write(primeira_data)
            time.sleep(4)
            pyautogui.press("tab")
            pyautogui.write(primeira_data)

            time.sleep(3)

            elemento_Nf = navegador.find_element(By.XPATH, '//*[@id="valor"]')
            actions = ActionChains(navegador)
            actions.click(elemento_Nf)
            actions.click(elemento_Nf)
            actions.click(elemento_Nf)
            actions.perform()
            total_tarifas_str = str(total_formatado)
            pyautogui.write(total_tarifas_str)
            time.sleep(5)

            navegador.find_element(By.XPATH, '//*[@id="criarDups"]').click()
            elemento_data_duplicata = navegador.find_element(By.XPATH, '//*[@id="dupVenc1"]')
            actions = ActionChains(navegador)
            actions.click(elemento_data_duplicata)
            actions.click(elemento_data_duplicata)
            actions.click(elemento_data_duplicata)
            actions.perform()
            pyautogui.write(primeira_data)
            time.sleep(4)

            navegador.find_element(By.XPATH, '//*[@id="btAddPl"]').click()
            time.sleep(2)
            pyautogui.write("tarifas")
            pyautogui.press("enter")
            time.sleep(10)
            
            pyautogui.click(289,439)
            time.sleep(4)

            navegador.find_element(By.XPATH, '//*[@id="botaoAddUnidadeCusto_1"]').click()
            time.sleep(4)
            pyautogui.write("ADM")
            pyautogui.press("enter")
            time.sleep(10)
            
            pyautogui.click(289,439)

            navegador.find_element(By.XPATH, '//*[@id="salvar"]').click()
            time.sleep(10)

            navegador.get("https://webtrans.saas.gwsistemas.com.br/bxcontaspagar?acao=iniciar")
            time.sleep(2)

            data_Baixa = navegador.find_element(By.XPATH, '//*[@id="dtinicial"]')
            actions = ActionChains(navegador)
            actions.click(data_Baixa)
            actions.click(data_Baixa)
            actions.click(data_Baixa)
            actions.perform()
            pyautogui.write(primeira_data)
            time.sleep(3)

            pyautogui.press("tab")
            pyautogui.write(primeira_data)

            time.sleep(10)
            navegador.find_element(By.XPATH, '//*[@id="localiza_clifor"]').click()
            time.sleep(7)
            pyautogui.write(fornecedor)
            pyautogui.press("enter")

            time.sleep(10)
            
            pyautogui.click(289,439)

            navegador.find_element(By.XPATH, '//*[@id="idfilial"]').click()
            pyautogui.write(filial)
            pyautogui.press('enter')
            pyautogui.press('tab')

            pyautogui.write(total_tarifas_str)
            time.sleep(2)
            pyautogui.press("tab")
            pyautogui.write(total_tarifas_str)
            time.sleep(3)
            pyautogui.press("tab")

            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="visualizar"]').click()

            time.sleep(7)
            navegador.find_element(By.XPATH, '//*[@id="chk_0"]').click()

            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="conta"]').click()
            pyautogui.write(numero_cc_formatado)
            time.sleep(3)
            pyautogui.press("enter")

            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="fpag"]').click()
            time.sleep(3)
            pyautogui.press('up')
            time.sleep(2)
            pyautogui.press('enter')

            time.sleep(3)
            data_Baixa = navegador.find_element(By.XPATH, '//*[@id="dtchq"]')
            actions = ActionChains(navegador)
            actions.click(data_Baixa)
            actions.click(data_Baixa)
            actions.click(data_Baixa)
            actions.perform()
            pyautogui.write(primeira_data)
            time.sleep(3)
            pyautogui.press("tab")

            time.sleep(3)
            data_Baixa = navegador.find_element(By.XPATH, '//*[@id="dtpago"]')
            actions = ActionChains(navegador)
            actions.click(data_Baixa)
            actions.click(data_Baixa)
            actions.click(data_Baixa)
            actions.perform()
            pyautogui.write(primeira_data)
            time.sleep(3)
            pyautogui.press("tab")

            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="chkconciliado"]').click()
            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="baixar"]').click()
            time.sleep(10)
            pyautogui.press('up')
            pyautogui.press('up')
            navegador.quit()
            time.sleep(4)

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        process_pdfs(email_entry.get(), password_entry.get(), folder_path)

# Configuração da interface Tkinter
root = tk.Tk()
root.title("Processamento de PDFs")

# Criando um frame para organizar os elementos
frame = tk.Frame(root)
frame.pack(expand=True)

# Criando os elementos
email_label = tk.Label(frame, text="Email:")
email_entry = tk.Entry(frame)

password_label = tk.Label(frame, text="Senha:")
password_entry = tk.Entry(frame, show="*")

browse_button = tk.Button(frame, text="Escolher Pasta", command=browse_folder)

# Centralizando os elementos
email_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
email_entry.grid(row=0, column=1, padx=5, pady=5)

password_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
password_entry.grid(row=1, column=1, padx=5, pady=5)

browse_button.grid(row=2, column=0, columnspan=2, padx=5, pady=20)

# Obtendo a largura e a altura da janela
window_width = 400
window_height = 200

# Obtendo a largura e a altura da tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calculando a posição x e y para centralizar a janela
x = int((screen_width / 2) - (window_width / 2))
y = int((screen_height / 2) - (window_height / 2))

# Definindo a geometria da janela para centralizá-la
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

root.mainloop()