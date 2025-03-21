import platform
import psutil
import gspread
import json
import wmi
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Configurar acesso ao Google Sheets
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "SUA-CREDENCIAL-GOOGLE.json"  # Substitua pelo caminho do seu arquivo JSON
SPREADSHEET_ID = "SEU-ID-DA-PLANILHA"  # Substitua pelo ID da sua planilha


# Autenticação no Google Sheets
def conectar_sheets():
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPES)
    cliente = gspread.authorize(creds)
    return cliente.open_by_key(SPREADSHEET_ID).sheet1


# Coletar informações do computador com WMI
def coletar_info():
    sistema = platform.uname()
    memoria = psutil.virtual_memory()
    disco = psutil.disk_usage('/')
    w = wmi.WMI()
    computador = w.Win32_ComputerSystem()[0]
    bios = w.Win32_BIOS()[0]
    processador = w.Win32_Processor()[0]
    so = w.Win32_OperatingSystem()[0]

    # Determinar o tipo de disco
    discos = w.Win32_DiskDrive()
    tipos_de_disco = {"SSD": "SSD", "HDD": "HDD", "NVMe": "NVMe", "Apple SSD": "Apple SSD"}
    disco_tipo = "Desconhecido"
    for d in discos:
        for chave, valor in tipos_de_disco.items():
            if chave in d.Model:
                disco_tipo = valor
                break

    info = {
        "Hostname": sistema.node,
        "Usuário": psutil.users()[0].name if psutil.users() else "Desconhecido",
        "Modelo": computador.Model,
        "Fabricante": computador.Manufacturer,
        "Serial Number": bios.SerialNumber,
        "Sistema Operacional": f"{so.Caption} {so.OSArchitecture}",
        "CPU": processador.Name,
        "RAM": round(memoria.total / (1024 ** 3)),  # Arredondado
        "ROM": round(disco.total / (1024 ** 3)),  # Arredondado
        "Disco": disco_tipo
    }
    return info


# Atualizar ou adicionar dados no Google Sheets
def atualizar_sheets():
    sheet = conectar_sheets()
    dados = coletar_info()
    serial_number = dados["Serial Number"]

    # Buscar se o serial number já existe na planilha
    registros = sheet.get_all_records()
    for i, registro in enumerate(registros, start=2):  # Começa da linha 2 (1 é cabeçalho)
        if registro["Serial Number"] == serial_number:
            sheet.update(range_name=f"A{i}:J{i}", values=[list(dados.values())])  # Atualiza a linha existente
            print("Dados atualizados no Google Sheets!")
            return

    # Se não encontrar, adiciona uma nova linha
    sheet.append_row(list(dados.values()))
    print("Novo registro adicionado ao Google Sheets!")


if __name__ == "__main__":
    atualizar_sheets()
