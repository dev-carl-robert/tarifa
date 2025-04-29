import os
import re
import fitz
import PyPDF2
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
import pyautogui
from pynput.mouse import Controller, Button
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

# codigo de executavel (python -m PyInstaller --onefile bot_tarifa.py)


def merge_pdf(pdf_files, output_path):
    pdf_merger = PyPDF2.PdfMerger()
    for pdf_file in pdf_files:
        pdf_merger.append(pdf_file)
    with open(output_path, 'wb') as output_file:
        pdf_merger.write(output_file)


def extrair_infos_pdf(caminho_pdf):
    with open(caminho_pdf, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        texto = ""
        for page in pdf_reader.pages:
            texto += page.extract_text()

    linhas = texto.split('\n')
    valor_tarifa_total = 0.0
    tarifas_detalhadas = []
    linhas_processadas = set()
    data_count = 0
    codigo_filial = "Desconhecido"
    conta_completa = "Conta não encontrada"

    # → Busca por código da conta (linha com "CC:")
    for linha in linhas:
        if "CC:" in linha:
            partes = linha.split("|")
            for parte in partes:
                if "CC:" in parte:
                    try:
                        conta_completa = parte.split(":")[1].strip().split()[-1]  # Ex: 0000543-6
                        numero_sem_dv = conta_completa.split('-')[0].lstrip("0")
                        codigo_filial = numero_sem_dv
                    except:
                        pass
                    break

    # → Busca por tarifas
    for i in range(len(linhas) - 1):
        # Ignorar linhas que contêm a palavra "total"
        if "total" in linhas[i].lower():
            continue

        if re.search(r'\d{2}/\d{2}/\d{4}', linhas[i]):
            data_count += 1
            if data_count == 4:
                break

        linha_atual = linhas[i]
        linha_proxima = linhas[i + 1] if i + 1 < len(linhas) else ""

        # Armazenar apenas o primeiro número negativo encontrado na linha atual
        if ("TARIFA" in linha_atual or "DOC/TED INTERNET" in linha_atual) and \
        ("TARIFA OPERACAO" not in linha_atual and "DOC/TED INTERNET" not in linha_proxima):
            if linha_proxima not in linhas_processadas:
                # Encontrar o primeiro número negativo
                valores = re.findall(r'-\d+,\d+', linha_proxima)
                if valores:  # Verifica se há valores encontrados
                    primeiro_valor_negativo = float(valores[0].replace(',', '.'))  # Armazena apenas o primeiro
                    valor_tarifa_total += primeiro_valor_negativo * -1
                    tarifas_detalhadas.append((linha_proxima, primeiro_valor_negativo * -1))
                    linhas_processadas.add(linha_proxima)

        if any(x in linha_atual for x in ["ENCARGOS DESCOBERTO", "TAR ", "TARIFA OPERACAO"]):
            if linha_atual not in linhas_processadas:
                match = re.search(r'-\d+,\d+', linha_atual)
                if match:
                    primeiro_valor_negativo = float(match.group().replace(',', '.'))  # Armazena apenas o primeiro
                    valor_tarifa_total += primeiro_valor_negativo * -1
                    tarifas_detalhadas.append((linha_atual, primeiro_valor_negativo * -1))
                    linhas_processadas.add(linha_atual)
                    
    # → Encontrar data de lançamento
    datas = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', texto)
    data_lancamento = datas[2] if len(datas) >= 2 else "Data não encontrada"

    # → Tabela de filiais
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
    df_filiais = pd.DataFrame(tabelaFIliais, columns=['Código', 'Banco', 'Filial'])
    resultado = df_filiais[df_filiais['Código'] == codigo_filial]

    banco = resultado.iloc[0]['Banco'] if not resultado.empty else "Desconhecido"
    filial = resultado.iloc[0]['Filial'] if not resultado.empty else "Desconhecida"

    return {
        "conta": conta_completa,
        "codigo_filial": codigo_filial,
        "data_lancamento": data_lancamento,
        "valor_tarifa": valor_tarifa_total,
        "banco": banco,
        "filial": filial,
        "tarifas_detalhadas": tarifas_detalhadas
    }

def process_pdfs(email, password, folder_path):
    pdf_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.lower().endswith('.pdf')]

    if not pdf_files:
        messagebox.showinfo("Info", "Não foram encontrados arquivos PDF na pasta especificada.")
        return

    for pdf_file in pdf_files:
        dados = extrair_infos_pdf(pdf_file)  # Corrigi também o nome da função que você usou

        if dados["conta"] in ["Conta não encontrada", None, ""] or \
           dados["codigo_filial"] in ["Desconhecido", None, ""] or \
           dados["valor_tarifa"] == 0.0 or \
           dados["data_lancamento"] == "Data não encontrada":
            print(f'erro no: "{os.path.basename(pdf_file)}": PDF EXTRAÍDO INCORRETAMENTE.')
            print("-" * 60)
            continue  # Alterado de return False para continue para processar o próximo PDF

        print("*" * 60)
        print(f"Executando {os.path.basename(pdf_file)}")
        print("Conta:", dados["conta"])
        print("Código da Filial:", dados["codigo_filial"])
        print("Banco:", dados["banco"])
        print("Filial:", dados["filial"])
        print("Data do Lançamento:", dados["data_lancamento"])
        print("Valor total das Tarifas:", dados["valor_tarifa"])
        print("--- Tarifas Encontradas ---")
        for linha, valor in dados["tarifas_detalhadas"]:
            print(f"Linha: {linha.strip()}")
            print(f"Valor: R$ {valor:.2f}")
        
        print(f"Total tarifas: {dados['valor_tarifa']}")

        # Mensagem ao final do processamento
        # Pode continuar seu tratamento aqui
        filial = dados["filial"]
        fornecedor = dados["banco"]
        data_lancamento = dados["data_lancamento"]
        total_tarifas = dados["valor_tarifa"]
        total_formatado = "{:.2f}".format(total_tarifas).replace('.', ',')
        conta = dados["codigo_filial"]
       
        print("*" * 60)
        servico =  Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)
        
        navegador.maximize_window()


        navegador.get("https://webtrans.saas.gwsistemas.com.br/")
        navegador.find_element(By.XPATH, '//*[@id="login"]').send_keys(email)
        time.sleep(3)
        navegador.find_element(By.XPATH, '//*[@id="senha"]').send_keys(password)
        time.sleep(3)
        navegador.find_element(By.XPATH, '//*[@id="form_login"]/div[3]/div/button').click()
        time.sleep(3)

        navegador.get("https://webtrans.saas.gwsistemas.com.br/consultadespesa?acao=iniciar")
        time.sleep(3)
        navegador.find_element(By.XPATH, '//*[@id="novo"]').click()
        time.sleep()
        navegador.find_element(By.XPATH, '//*[@id="especie"]').send_keys("FIN")

        navegador.find_element(By.XPATH, '//*[@id="localiza_filial"]').click()
        time.sleep(3)
        pyautogui.write(filial)
        time.sleep(3)
        pyautogui.press("enter")
        time.sleep(4)
        pyautogui.click(226, 347)
        time.sleep(8)


        navegador.find_element(By.XPATH, '//*[@id="localiza_fornecedor"]').click()
        time.sleep(4)
        pyautogui.write(fornecedor)
        time.sleep(3)
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(10)
        
        pyautogui.click(226, 347)
        time.sleep(5)

        navegador.find_element(By.XPATH, '//*[@id="descricao_historico"]').send_keys("TARIFA BANCARIA")
        time.sleep(3)

        elemento_data = navegador.find_element(By.XPATH, '//*[@id="dtemissao"]')
        actions = ActionChains(navegador)
        actions.click(elemento_data)
        actions.click(elemento_data)
        actions.click(elemento_data)
        actions.perform()
        pyautogui.write(data_lancamento)
        time.sleep(4)
        pyautogui.press("tab")
        pyautogui.write(data_lancamento)

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
        pyautogui.write(data_lancamento)
        time.sleep(4)

        navegador.find_element(By.XPATH, '//*[@id="btAddPl"]').click()
        time.sleep(2)
        pyautogui.write("tarifas")
        pyautogui.press("enter")
        time.sleep(10)
        
        pyautogui.click(226, 347)
        time.sleep(4)

        navegador.find_element(By.XPATH, '//*[@id="botaoAddUnidadeCusto_1"]').click()
        time.sleep(4)
        pyautogui.write("ADM")
        pyautogui.press("enter")
        time.sleep(10)
        
        pyautogui.click(226, 347)

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
        pyautogui.write(data_lancamento)
        time.sleep(3)

        pyautogui.press("tab")
        pyautogui.write(data_lancamento)

        time.sleep(10)
        navegador.find_element(By.XPATH, '//*[@id="localiza_clifor"]').click()
        time.sleep(7)
        pyautogui.write(fornecedor)
        pyautogui.press("enter")

        time.sleep(10)
        
        pyautogui.click(226, 347)

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
        pyautogui.write(conta)
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
        pyautogui.write(data_lancamento)
        time.sleep(3)
        pyautogui.press("tab")

        time.sleep(3)
        data_Baixa = navegador.find_element(By.XPATH, '//*[@id="dtpago"]')
        actions = ActionChains(navegador)
        actions.click(data_Baixa)
        actions.click(data_Baixa)
        actions.click(data_Baixa)
        actions.perform()
        pyautogui.write(data_lancamento)
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
        navegador.quit()
        time.sleep(4)
        print(f'"{os.path.basename(pdf_file)}" lançado com sucesso.')

    print("Todos os PDFs executados.")
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
