import os
import pandas as pd
from datetime import datetime

def ler_numeros_excel(arquivo_excel):
    # Ler o arquivo Excel
    xls = pd.ExcelFile(arquivo_excel)
    
    # Ler a primeira coluna da aba Sales Orders
    sales_orders = pd.read_excel(xls, 'Sales Orders', usecols=[0], engine='openpyxl')
    sales_orders_numeros = sales_orders.iloc[:, 0].drop_duplicates().astype(str).tolist()
    
    # Ler a primeira coluna da aba Quotations
    quotations = pd.read_excel(xls, 'Quotations', usecols=[0], engine='openpyxl')
    quotations_numeros = quotations.iloc[:, 0].drop_duplicates().astype(str).tolist()
    
    return sales_orders_numeros, quotations_numeros

def buscar_arquivos(diretorio, numeros):
    arquivos_nao_correspondentes = []
    
    for arquivo in os.listdir(diretorio):
        caminho_completo = os.path.join(diretorio, arquivo)
        if os.path.isfile(caminho_completo):
            encontrado = False
            for numero in numeros:
                if numero in arquivo:
                    encontrado = True
                    break
            if not encontrado:
                arquivos_nao_correspondentes.append(arquivo)
    
    return arquivos_nao_correspondentes

def gerar_log_entry(entry_type, message):
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    return f"{timestamp} - {entry_type}: {message}\n"

def listar_arquivos(diretorios, arquivo_excel, log_file):
    sales_orders_numeros, quotations_numeros = ler_numeros_excel(arquivo_excel)
    
    with open(log_file, 'a') as log:  # Abrir o arquivo de log no modo de adição
        log.write(gerar_log_entry("INFO", "Iniciando verificação de arquivos"))
        
        for diretorio in diretorios:
            log.write(gerar_log_entry("INFO", f"Verificando arquivos no diretório {diretorio}"))
            
            # Buscar arquivos no diretório para Sales Orders e Quotations
            arquivos_nao_correspondentes_sales_orders = buscar_arquivos(diretorio, sales_orders_numeros)
            arquivos_nao_correspondentes_quotations = buscar_arquivos(diretorio, quotations_numeros)
            
            # Combinar listas de arquivos não correspondentes e remover duplicatas
            arquivos_nao_correspondentes = list(set(arquivos_nao_correspondentes_sales_orders + arquivos_nao_correspondentes_quotations))
            
            if arquivos_nao_correspondentes:
                log.write(gerar_log_entry("ERROR", f"Arquivos não correspondentes encontrados: {arquivos_nao_correspondentes}"))
            else:
                log.write(gerar_log_entry("INFO", "Todos os arquivos correspondem aos números das planilhas"))
            
            log.write("\n")

# Caminho do arquivo Excel e diretório para busca
arquivo_excel = r'C:\Users\v.santos\OneDrive - Interroll Management SA\Desktop\projeto python\atualizacao\base_de_dados.xlsx'
diretorios = [r'C:\Users\v.santos\OneDrive - Interroll Management SA\Desktop\projeto python\folder teste', r'C:\Users\v.santos\OneDrive - Interroll Management SA\Desktop\projeto python\outro folder']
# Nome do arquivo de log
log_file = 'update_excel.log'

listar_arquivos(diretorios, arquivo_excel, log_file)