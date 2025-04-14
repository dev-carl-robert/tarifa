import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
import sys
import pandas as pd
import fitz
import PyPDF2
import re

def receive_pdf(email, password, pdf_path):
    # Aqui você pode implementar a lógica para receber o PDF
    print("PDF recebido para o e-mail:", email)
    print("Senha:", password)
    print("Caminho do PDF:", pdf_path)
    message_label.config(text="Pronto! caso precise, envie outro arquivo")


                                # leitor de PDF
    # Abra o arquivo PDF em modo de leitu
    import PyPDF2
    with open(pdf_path, 'rb') as file:
    # Crie um objeto PdfReader
        pdf_reader = PyPDF2.PdfReader(file)

        # Obtenha a primeira página
        first_page = pdf_reader.pages[0]

        # Obtenha o texto da primeira página
        text = first_page.extract_text()

        # Divida o texto em linhas
        lines = text.split('\n')
            
        valorTarifa1_total = 0
        valorTarifa2_total = 0
        valorTarifa3_total = 0
        valorTarifa4_total = 0
        valorTarifa5_total = 0
        valorTarifa_1 = 0
        valorTarifa_2 = 0
        valorTarifa_3 = 0
        valorTarifa_4 = 0
        valorTarifa_5 = 0
        total_tarifas = 0

        for line_num, line in enumerate(lines):

                if "TRANSF PGTO" in line or "TIT.BX.DECURSO PRAZO" in line or "PLANO EXCLUSIVO" in line:
                    linha_tarifa = line.split(" ")
                    listaTarifa = pd.DataFrame([linha_tarifa])
                    valorTarifa_1 = listaTarifa[3]
                    valorTarifa1 = float(valorTarifa_1.iloc[0].replace(",", ".").replace("-", "-"))
                    print(line)
                    print("o valor da tarifa é : " + str(valorTarifa1))  # Converta o valor para string antes de concatenar
                    valorTarifa1_total += valorTarifa1 * -1 # Adiciona o valor da tarifa à soma total

                if "TAR.MANUT.C" in line:
                    linha_tarifa = line.split(" ")
                    listaTarifa = pd.DataFrame([linha_tarifa])
                    valorTarifa_2 = listaTarifa[1]
                    valorTarifa2 = float(valorTarifa_2.iloc[0].replace(",", ".").replace("-", "-"))
                    print(line)
                    print("o valor da tarifa é : " + str(valorTarifa2))  # Converta o valor para string antes de concatenar
                    valorTarifa2_total += valorTarifa2  * -1 # Adiciona o valor da tarifa à soma total

                if "QUANDO DO REGISTRO" in line or "PAGAMENTO FUNCs NET EMPRESA" in line:
                    linha_tarifa = line.split(" ")
                    listaTarifa = pd.DataFrame([linha_tarifa])
                    valorTarifa_3 = listaTarifa[4]
                    valorTarifa3 = float(valorTarifa_3.iloc[0].replace(",", ".").replace("-", "-"))
                    print(line)
                    print("o valor da tarifa é : " + str(valorTarifa3))  # Converta o valor para string antes de concatenar
                    valorTarifa3_total += valorTarifa3  * -1
                if "TAR CC REAL TIME PAGFOR" in line or "TAR SERV TED STR PAGFOR" in line:
                    linha_tarifa = line.split(" ")
                    listaTarifa = pd.DataFrame([linha_tarifa])
                    valorTarifa_4 = listaTarifa[6]
                    valorTarifa4 = float(valorTarifa_4.iloc[0].replace(",", ".").replace("-", "-"))
                    print(line)
                    print("o valor da tarifa é : " + str(valorTarifa4))  # Converta o valor para string antes de concatenar
                    valorTarifa4_total += valorTarifa4  * -1
                if "TRANSFER VIA NET" in line:
                    linha_tarifa = line.split(" ")
                    listaTarifa = pd.DataFrame([linha_tarifa])
                    valorTarifa_4 = listaTarifa[3]
                    valorTarifa4 = float(valorTarifa_4.iloc[0].replace(".", "").replace(",", ".").replace("-", "-"))
                    print(line)
                    print("o valor da tarifa é : " + str(valorTarifa4))  # Converta o valor para string antes de concatenar
                    valorTarifa4_total += valorTarifa4 * -1
                      
               
        # Encontrar numero da Conta e FIlial
                if "CC:" in line:
                    texto = line
                    if "CC" in line:
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
                    import pandas as pd
                    tabelaFIliais = [
                        ('35655', 'BANCO BRADESCO S.A', 'RECIFE'),
                        ('12260', 'BANCO BRADESCO S.A', 'CORDEIROPOLI'),
                        ('10162', 'BANCO BRADESCO S.A', 'GUARULHOS'),
                        ('11700', 'BANCO BRADESCO S.A', 'BENEVIDES'),
                        ('71644', 'BANCO BRADESCO S.A', 'BENEVIDES'),
                        ("90718", 'BANCO BRADESCO S.A', 'FORTALEZA'),
                        ('7326', 'BANCO BRADESCO S.A', 'PARAUPEBAS'),
                        ('423', 'BANCO BRADESCO S.A', 'SIMOES FILHO'),
                        ('461', 'BANCO BRADESCO S.A', 'MINAS GERAIS'),
                        ('18885', 'BANCO BRADESCO S.A', 'TRANSULPARTI'),
                        ('27830', 'BANCO DO BRASIL S.A', 'MATRIZ'),
                        ('610', 'CAIXA ECONOMICA FEDERAL', 'MATRIZ'),
                        ('554', 'BANCO  SAFRA S.A', 'MATRIZ'),
                        ('55000', 'BANCO BRADESCO S.A', 'MATRIZ'),
                        ('25779', 'BANCO BRADESCO S.A', 'MATRIZ'),
                        ('18200', 'BANCO BRADESCO S.A', 'MATRIZ'),
                        ('12473', 'BANCO ITAUCARD S.A.', 'MATRIZ'),
                        ('10293', 'BANCO BRADESCO S.A', 'MATRIZ'),
                        ('20540-0', 'BANCO ITAUCARD S.A.', 'MATRIZ'),
                        ('28169', 'BANCO DO NORDESTE DO BRASIL S.A', 'MATRIZ'),
                        ('130001633', 'BANCO SANTANDER (BRASIL) S.A.', 'MATRIZ'),
                        ('18180-3', 'BANCO BRADESCO S.A', 'TRANSULOG MA'),
                        ('25123', 'BANCO BRADESCO S.A', 'RODONORTE MT'),
                        ('25122', 'BANCO BRADESCO S.A', 'RODONORTE MT'),
                        ('505-3', 'BANCO BRADESCO S.A', 'MATRIZ'),
                        ('1120-7', 'BANCO DO BRASIL S.A', 'MATRIZ')
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

                    import re
                    def encontrar_primeira_data(texto):
                        # Encontrar todas as ocorrências de data no texto
                        datas_encontradas = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', texto)
                        # Retornar a primeira data encontrada, se houver alguma
                        return datas_encontradas[0] if datas_encontradas else None

                    #Caminho do arquivo PDF
                    caminho_arquivo_pdf = pdf_path

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
                        print("Data da emissão:", primeira_data)
                    else:
                        print("Nenhuma data encontrada nos lançamentos.")
                                #bot de soma de tarifa
                        
        total_de_todas_tarifas = valorTarifa1_total + valorTarifa2_total + valorTarifa3_total + valorTarifa4_total + valorTarifa5_total
        total_formatado = round(total_de_todas_tarifas, 3)
        print(f"A soma de todas as tarifas é: {total_formatado}")


                        
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

        mouse = Controller()

# Coordenadas do ponto para onde queremos mover o mouse
        x_destino = 324
        y_destino = 394
        x_atual, y_atual = mouse.position
        delta_x = x_destino - x_atual
        delta_y = y_destino - y_atual



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
        time.sleep(10)
        # pyautogui.click(x=324,y=394, duration=5)
        mouse.move(delta_x, delta_y)
        time.sleep(10)
        mouse.click(Button.left)
        time.sleep(8)


        navegador.find_element(By.XPATH, '//*[@id="localiza_fornecedor"]').click()
        time.sleep(4)
        pyautogui.write(fornecedor)
        time.sleep(3)
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(10)
        
        mouse.click(Button.left)
        time.sleep(10)

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
        
        mouse.click(Button.left)
        time.sleep(4)

        navegador.find_element(By.XPATH, '//*[@id="botaoAddUnidadeCusto_1"]').click()
        time.sleep(4)
        pyautogui.write("ADM")
        pyautogui.press("enter")
        time.sleep(10)
        
        mouse.click(Button.left)

        navegador.find_element(By.XPATH, '//*[@id="salvar"]').click()
        time.sleep(10)
        navegador.get("https://webtrans.saas.gwsistemas.com.br/")

        # Baixa de tarifa
        time.sleep(4)
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
        
        mouse.click(Button.left)

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
        navegador.find_element(By.XPATH, '//*[@id="baixar"]').click()
            
            
        
# Criar a janela principal
def main():
    # Criar a janela principal
    root = tk.Tk()
    root.title("Receber PDF")

    # Ajustar tamanho da janela
    root.geometry("400x200")

    # Centralizar a janela na tela
    root.eval('tk::PlaceWindow . center')

    # Função para receber o e-mail, senha e PDF
    def receive_pdf_and_execute():
        email = email_entry.get()
        password = password_entry.get()
        pdf_path = filedialog.askopenfilename(title="Selecione o PDF")
        if pdf_path:
             receive_pdf(email, password, pdf_path)

    # Criar um frame para centralizar os elementos
    frame = tk.Frame(root)
    frame.pack(expand=True)

    # Entrada para o e-mail
    email_label = tk.Label(frame, text="E-mail:")
    email_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")

    email_entry = tk.Entry(frame, width=30)  # Aumentando o tamanho visual
    email_entry.grid(row=0, column=1, padx=10, pady=10)

    # Entrada para a senha
    password_label = tk.Label(frame, text="Senha:")
    password_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")

    password_entry = tk.Entry(frame, show="*", width=30)  # Aumentando o tamanho visual
    password_entry.grid(row=1, column=1, padx=10, pady=10)

    # Botão para receber PDF
    receive_button = tk.Button(frame, text="Receber PDF", command=receive_pdf_and_execute)
    receive_button.grid(row=2, columnspan=2, padx=10, pady=10)

    # Criar a label para exibir mensagem
    global message_label
    message_label = tk.Label(frame, text="")
    message_label.grid(row=3, columnspan=2, padx=10, pady=10)

    # Executar o loop principal da interface
    root.mainloop()

if __name__ == "__main__":
    main()