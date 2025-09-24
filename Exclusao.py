import pandas as pd
from tqdm import tqdm
import time
import sys
import os
from tabulate import tabulate
import re
from datetime import datetime
import glob
import shutil

# -*- coding: utf-8 -*-
# Configurar codificação para evitar problemas no Windows
if sys.platform.startswith('win'):
    import locale
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        pass

# Limpar o terminal para uma visualização limpa
os.system('cls' if os.name == 'nt' else 'clear')

def get_next_exclusao_folder():
    """
    Encontra o próximo número sequencial para a pasta Exclusão e retorna
    o nome da pasta com data no formato: Exclusão XX - DD-MM-AAAA
    """
    base_dir = 'Excluir Aniel'
    
    # Verificar se o diretório base existe
    if not os.path.exists(base_dir):
        return f'{base_dir}/Exclusão 01 - {datetime.now().strftime("%d-%m-%Y")}'
    
    # Buscar todas as pastas que começam com "Exclusão"
    pattern = os.path.join(base_dir, 'Exclusão*')
    existing_folders = glob.glob(pattern)
    
    # Extrair números das pastas existentes
    numbers = []
    for folder in existing_folders:
        folder_name = os.path.basename(folder)
        # Buscar padrão "Exclusão XX" onde XX é um número
        match = re.search(r'Exclusão\s+(\d+)', folder_name)
        if match:
            numbers.append(int(match.group(1)))
    
    # Determinar o próximo número
    if numbers:
        next_number = max(numbers) + 1
    else:
        next_number = 1
    
    # Formatear o nome da pasta com data atual
    current_date = datetime.now().strftime("%d-%m-%Y")
    folder_name = f'Exclusão {next_number:02d} - {current_date}'
    
    return os.path.join(base_dir, folder_name)

# Etapas do processo para a barra de progresso
with tqdm(total=6, ncols=92, desc="Processando", leave=True, dynamic_ncols=False) as pbar:
    # 1. Lendo CSV
    pbar.set_description("Lendo CSV QTD Solicitações")
    try:
        df_csv = pd.read_csv('Excluir Aniel/QTD Solicitações _ Recolhimento.csv', sep=',', encoding='utf-8', on_bad_lines='skip')
    except FileNotFoundError:
        print("\n \033[1;31mArquivo 'Excluir Aniel/QTD Solicitações _ Recolhimento.csv' não encontrado.\033[0m \n \033[1;34mCertifique-se de que o arquivo está no diretório 'Excluir Aniel'\033[0m")
        sys.exit("Encerrando o programa")
    except Exception as e:
        try:
            df_csv = pd.read_csv('Excluir Aniel/QTD Solicitações _ Recolhimento.csv', sep=',', encoding='latin1', on_bad_lines='skip')
        except FileNotFoundError:
            print("\n \033[1;31mArquivo 'Excluir Aniel/QTD Solicitações _ Recolhimento.csv' não encontrado.\033[0m \n \033[1;34mCertifique-se de que o arquivo está no diretório 'Excluir Aniel'\033[0m")
            sys.exit("Encerrando o programa")
        except Exception as e2:
            print(f"\n \033[1;31mErro ao ler arquivo CSV: {e2}\033[0m")
            sys.exit("Encerrando o programa")
    pbar.update(1)
    time.sleep(0.5)

    # 2. Lendo Excel
    pbar.set_description("Lendo arquivos Excel RECOLHIMENTO")
    try:
        # Buscar arquivos Excel de recolhimento
        recolhimento_files = []
        
        # Verificar se existem os arquivos específicos
        if os.path.exists('Excluir Aniel/RECOLHIMENTO.xlsx'):
            recolhimento_files.append('Excluir Aniel/RECOLHIMENTO.xlsx')
        
        if os.path.exists('Excluir Aniel/RECOLHIMENTO AGENDADO.xlsx'):
            recolhimento_files.append('Excluir Aniel/RECOLHIMENTO AGENDADO.xlsx')
        
        # Se não encontrar os novos arquivos, buscar o antigo formato
        if not recolhimento_files:
            excel_files = [f for f in os.listdir('Excluir Aniel') if f.endswith('.xlsx') and 'Painel' in f]
            if excel_files:
                recolhimento_files = [f'Excluir Aniel/{excel_files[0]}']
        
        if not recolhimento_files:
            raise FileNotFoundError("Nenhum arquivo Excel encontrado (RECOLHIMENTO.xlsx, RECOLHIMENTO AGENDADO.xlsx ou Painel de Serviços.xlsx)")
        
        # Ler e combinar todos os arquivos Excel encontrados
        dataframes_excel = []
        excel_counts = {}
        for file_path in recolhimento_files:
            df_temp = pd.read_excel(file_path, engine='openpyxl', header=1)
            dataframes_excel.append(df_temp)
            # Armazena a contagem de protocolos de cada arquivo Excel
            excel_counts[os.path.basename(file_path)] = df_temp['Nº. Ordem Serviço'].count() if 'Nº. Ordem Serviço' in df_temp.columns else len(df_temp)
        # Combinar todos os DataFrames
        df_xlsx = pd.concat(dataframes_excel, ignore_index=True)
        
    except FileNotFoundError as e:
        print(f"\n \033[1;31m{e}\033[0m \n \033[1;34mCertifique-se de que os arquivos estão no diretório 'Excluir Aniel'\033[0m")
        sys.exit("Encerrando o programa")
    except Exception as e:
        print(f"\n \033[1;31mErro ao ler arquivos Excel: {e}\033[0m")
        sys.exit("Encerrando o programa")
    pbar.update(1)
    time.sleep(0.5)

    # 3. Mesclando DataFrames
    pbar.set_description("Mesclando Planilhas")
    df_merged = pd.merge(
        df_csv[["ID Protocolo | Proxxima", "Status Protocolo", "Tipo Solicitação"]],
        df_xlsx[[
            "Nº. Ordem Serviço", "Contrato", "Projeto", "Identificação do Cliente", "Data/Hora Criação", "Status", "Tipo de Serviço"
        ]],
        left_on="ID Protocolo | Proxxima",
        right_on="Nº. Ordem Serviço",
        how="inner"
    )
    pbar.update(1)
    time.sleep(0.5)

    # 4. Montando DataFrame final
    pbar.set_description("Montando DataFrame final")
    df_final = df_merged[[
        "Contrato",
        "Projeto",
        "Identificação do Cliente",
        "Nº. Ordem Serviço",
        "Data/Hora Criação",
        "Status Protocolo",
        "Status",
        "Tipo Solicitação",
        "Tipo de Serviço"
    ]]
    df_final = df_final.rename(columns={
        "Nº. Ordem Serviço": "Ordem de Serviço",
        "Data/Hora Criação": "Data de Criação",
        "Status Protocolo": "Status Voalle",
        "Status": "Status Aniel"
    })
    pbar.update(1)
    time.sleep(0.5)

    # 5. Filtrando Status
    pbar.set_description("Filtrando Status")
    
    # Primeiro, criar arquivo específico antes de aplicar filtros
    # Formatando a coluna Data de Criação para o formato brasileiro DD/MM/AAAA HH:MM:SS
    df_final['Data de Criação'] = pd.to_datetime(df_final['Data de Criação'], dayfirst=True).dt.strftime('%d/%m/%Y %H:%M:%S')
    
    # Criando DataFrame para registros com Status Voalle = "CANCELADO" E Status Aniel = "Fechada Improdutiva" ou "Fechada Produtiva"
    df_cancelado_fechada = df_final[(df_final["Status Voalle"] == "CANCELADO") & (df_final["Status Aniel"].isin(["Fechada Improdutiva", "Fechada Produtiva"]))]
    
    # Manter na planilha Excluir_Aniel apenas Status Voalle desejados
    # Aceitar variações de caixa (Cancelado/cancelado/CANCELADO, etc.)
    allowed_voalle = {"CANCELADO", "FECHAMENTO", "ENCERRAMENTO"}
    df_final = df_final[df_final["Status Voalle"].astype(str).str.strip().str.upper().isin(allowed_voalle)]
    
    # Aplicando filtro principal - manter apenas registros que precisam de ação
    # Remover status que indicam conclusão ou não precisam de intervenção
    df_final = df_final[~df_final["Status Aniel"].isin(["Fechada Improdutiva", "Fechada Produtiva", "Cancelado"])]
    # Nota: Como não há esses status no Excel atual, o filtro não remove nada, 
    # mas mantém a compatibilidade com versões futuras
    pbar.update(1)
    time.sleep(0.5)

    # 6. Salvando arquivos
    pbar.set_description("Salvando arquivos")
    
    # Criando diretório sequencial com data
    output_dir = get_next_exclusao_folder()
    os.makedirs(output_dir, exist_ok=True)
    
    df_final_sem_status = df_final.drop(columns=["Status Voalle", "Status Aniel", "Tipo Solicitação", "Tipo de Serviço"])
    df_final_sem_status.to_excel(f'{output_dir}/Excluir_Aniel.xlsx', index=False)
    df_final.to_excel(f'{output_dir}/Excluir_Aniel_com_Status.xlsx', index=False)
    df_cancelado_fechada.to_excel(f'{output_dir}/Cancelado-Voalle_Fechada_Produtiva-Aniel.xlsx', index=False)
    
    # Criar pasta Source e mover arquivos originais
    source_dir = os.path.join(output_dir, 'Source')
    os.makedirs(source_dir, exist_ok=True)
    
    # Mover arquivo CSV para a pasta Source
    csv_source = 'Excluir Aniel/QTD Solicitações _ Recolhimento.csv'
    csv_destination = os.path.join(source_dir, 'QTD Solicitações _ Recolhimento.csv')
    
    if os.path.exists(csv_source):
        shutil.move(csv_source, csv_destination)
    
    # Mover arquivos Excel para a pasta Source
    excel_files_to_move = [
        'Excluir Aniel/RECOLHIMENTO.xlsx',
        'Excluir Aniel/RECOLHIMENTO AGENDADO.xlsx'
    ]
    
    for excel_file in excel_files_to_move:
        if os.path.exists(excel_file):
            filename = os.path.basename(excel_file)
            excel_destination = os.path.join(source_dir, filename)
            shutil.move(excel_file, excel_destination)
    
    pbar.update(1)
    time.sleep(0.5)

print(tabulate(df_final_sem_status.head(), headers='keys', tablefmt='psql', showindex=False))

# Mostrar qual pasta foi criada após o processamento
folder_name = os.path.basename(output_dir)
print(f"\n📁 Criada pasta: {folder_name}")

# RESUMO DETALHADO FINAL
print("\n=== RESUMO DETALHADO DOS DADOS ===")
print(f"- Protocolos no CSV (ID Protocolo | Proxxima): {df_csv['ID Protocolo | Proxxima'].count() if 'ID Protocolo | Proxxima' in df_csv.columns else len(df_csv)}")
for fname, count in excel_counts.items():
    print(f"- Protocolos no arquivo {fname} (Nº. Ordem Serviço): {count}")
total_excel = sum(excel_counts.values())
print(f"- Total combinado (RECOLHIMENTO.xlsx + RECOLHIMENTO AGENDADO.xlsx): {total_excel}")
print(f"- Protocolos enviados para a planilha Excluir Aniel: {len(df_final_sem_status)}")

print(f"\n📁 Arquivos salvos em '{output_dir}':")
print("- 'Excluir_Aniel.xlsx' (sem status)")
print("- 'Excluir_Aniel_com_Status.xlsx' (com status filtrado)")
print("- 'Cancelado-Voalle_Fechada_Produtiva-Aniel.xlsx' (Status Voalle = CANCELADO E Status Aniel = Fechada Improdutiva/Produtiva)")
print("\n📂 Arquivos originais movidos para 'Source/':")
print("- 'QTD Solicitações _ Recolhimento.csv' (arquivo CSV original)")
print("- 'RECOLHIMENTO.xlsx' (arquivo Excel original)")
print("- 'RECOLHIMENTO AGENDADO.xlsx' (arquivo Excel original)")
