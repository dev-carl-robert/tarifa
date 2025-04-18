import fitz
import re
import pandas as pd

def extrair_infos_pdf(caminho_pdf):
    with fitz.open(caminho_pdf) as doc:
        texto = ""
        for pagina in doc:
            texto += pagina.get_text()

    # Extrair conta completa (ex: 0000543-6)
    conta_match = re.search(r'\b\d{7}-\d\b', texto)
    conta_completa = conta_match.group() if conta_match else "Conta não encontrada"

    # Derivar código da filial
    if conta_completa != "Conta não encontrada":
        numero_base = conta_completa.split('-')[0]  # pega só os números antes do hífen
        numero_sem_zeros = numero_base.lstrip('0')  # remove zeros à esquerda
        codigo_filial = numero_sem_zeros
    else:
        codigo_filial = "Desconhecido"

    # Data: sempre a 2ª encontrada no texto
    datas = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', texto)
    data_lancamento = datas[1] if len(datas) >= 2 else "Data não encontrada"

    # Valor da tarifa (até 5 linhas depois da palavra TARIFA)
    linhas = texto.split('\n')
    valor_tarifa = "Valor não encontrado"

    for i, linha in enumerate(linhas):
        if "TARIFA" in linha.upper():
            for j in range(i + 1, min(i + 6, len(linhas))):
                match = re.search(r'-\d+,\d+', linhas[j])
                if match:
                    valor_tarifa = match.group().replace(',', '.')
                    break
            break

    # Mapeamento de código da filial
    tabela_filiais = [
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
    df_filiais = pd.DataFrame(tabela_filiais, columns=['Código', 'Banco', 'Filial'])

    resultado = df_filiais[df_filiais['Código'] == codigo_filial]
    if not resultado.empty:
        banco = resultado.iloc[0]['Banco']
        filial = resultado.iloc[0]['Filial']
    else:
        banco = "Desconhecido"
        filial = "Desconhecida"

    return {
        "conta": conta_completa,
        "codigo_filial": codigo_filial,
        "data_lancamento": data_lancamento,
        "valor_tarifa": float(valor_tarifa) if valor_tarifa != "Valor não encontrado" else valor_tarifa,
        "banco": banco,
        "filial": filial
    }

# Teste
dados = extrair_infos_pdf("543 01.pdf")
print("Conta:", dados["conta"])
print("Código da Filial:", dados["codigo_filial"])
print("Banco:", dados["banco"])
print("Filial:", dados["filial"])
print("Data do Lançamento:", dados["data_lancamento"])
print("Valor da Tarifa:", dados["valor_tarifa"])
