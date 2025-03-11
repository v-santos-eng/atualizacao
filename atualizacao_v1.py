import os
import pandas as pd

# Definindo o caminho do arquivo Excel e do diretório
excel_path = r'C:\Users\v.santos\OneDrive - Interroll Management SA\Desktop\projeto python\atualizacao\base_de_dados.xlsx'
directory_path = r'C:\Users\v.santos\OneDrive - Interroll Management SA\Desktop\projeto python'

# Lendo as abas do arquivo Excel
sales_orders_df = pd.read_excel(excel_path, sheet_name='Sales Orders', usecols=['Sales Order'])
quotations_df = pd.read_excel(excel_path, sheet_name='Quotations', usecols=['Quotation'])

# Extraindo os números das colunas
sales_orders = sales_orders_df['Sales Order'].astype(str).tolist()
quotations = quotations_df['Quotation'].astype(str).tolist()

# Função para verificar se algum arquivo no diretório contém os números lidos
def check_files(directory, numbers):
    for root, dirs, files in os.walk(directory):
        for file in files:
            for number in numbers:
                if number in file:
                    return True
    return False

# Verificando se há arquivos que contenham os números lidos
sales_orders_found = check_files(directory_path, sales_orders)
quotations_found = check_files(directory_path, quotations)

# Exibindo a mensagem apropriada
if not sales_orders_found and not quotations_found:
    print("Nenhum novo arquivo")
else:
    print("Arquivos encontrados que correspondem aos números lidos")