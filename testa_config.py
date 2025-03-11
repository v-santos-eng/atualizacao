import configparser
import os

print("Caminho atual:", os.getcwd())  # Mostra o caminho atual

config = configparser.ConfigParser()
config.read('config.ini')  # Ou coloca o caminho completo aqui

print("Seções encontradas:", config.sections())  # Verifica se leu o arquivo
caminho = r'C:\temp\config.ini'
print("Arquivo existe:", os.path.isfile(caminho))
