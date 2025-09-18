import pandas as pd
from tqdm import tqdm
import time
import sys
import os
from tabulate import tabulate

# -*- coding: utf-8 -*-
# Etapas do processo para a barra de progresso
with tqdm(total=6, ncols=92) as pbar:
    # 1. Lendo CSV
    pbar.set_description("\nLendo CSV Canceladas Voalle.csv")
    try:
        df_csv = pd.read_csv('Excluir Aniel/Canceladas Voalle.csv', sep=';', encoding='utf-8', on_bad_lines='skip')
    except FileNotFoundError:
        print("\n \033[1;31mArquivo 'Excluir Aniel/Canceladas Voalle.csv' não encontrado.\033[0m \n \033[0m \n \033[1;34mCertifique-se de que o nome do arquivo está correto e no mesmo diretório do arquivo CriarExclusão.py\033[0m")
        sys.exit("Encerrando o programa")
    except Exception:
        try:
            df_csv = pd.read_csv('Excluir Aniel/Canceladas Voalle.csv', sep=';', encoding='latin1', on_bad_lines='skip')
        except FileNotFoundError:
            print("\n \033[1;31mArquivo 'Excluir Aniel/Canceladas Voalle.csv' não encontrado.\033[0m \n \033[0m \n \033[1;34mCertifique-se de que o nome do arquivo está correto e no mesmo diretório do arquivo CriarExclusão.py\033[0m")
            sys.exit("Encerrando o programa")
    pbar.update(1)
    time.sleep(1)

    # 2. Lendo Excel
    pbar.set_description("\nLendo Excluir Aniel/Excel Painel de Serviços.xlsx")
    try:
        df_xlsx = pd.read_excel('Excluir Aniel/Painel de Serviços.xlsx', engine='openpyxl', header=1)
    except FileNotFoundError:
        print("\n \033[1;31mArquivo 'Excluir Aniel/Painel de Serviços.xlsx' não encontrado.\033[0m \n \033[1;34mCertifique-se de que o nome do arquivo está correto e no mesmo diretório do arquivo CriarExclusão.py\033[0m")
        sys.exit("Encerrando o programa")
    pbar.update(1)
    time.sleep(1)

    # 3. Mesclando DataFrames
    pbar.set_description("Mesclando Planilhas")
    df_merged = pd.merge(
        df_csv[["Protocolo", "Status"]],
        df_xlsx[[
            "Nº. Ordem Serviço", "Contrato", "Projeto", "Identificação do Cliente", "Data/Hora Criação", "Status"
        ]],
        left_on="Protocolo",
        right_on="Nº. Ordem Serviço",
        how="inner"
    )
    pbar.update(1)
    time.sleep(1)

    # 4. Montando DataFrame final
    pbar.set_description("Montando DataFrame final")
    df_final = df_merged[[
        "Contrato",
        "Projeto",
        "Identificação do Cliente",
        "Nº. Ordem Serviço",
        "Data/Hora Criação",
        "Status_x",
        "Status_y"
    ]]
    df_final = df_final.rename(columns={
        "Nº. Ordem Serviço": "Ordem de Serviço",
        "Data/Hora Criação": "Data de Criação",
        "Status_x": "Status Voalle",
        "Status_y": "Status Aniel"
    })
    pbar.update(1)
    time.sleep(1)

    # 5. Filtrando Status
    pbar.set_description("Filtrando Status")
    
    # Primeiro, criar arquivo específico antes de aplicar filtros
    # Formatando a coluna Data de Criação para o formato brasileiro DD/MM/AAAA HH:MM:SS
    df_final['Data de Criação'] = pd.to_datetime(df_final['Data de Criação']).dt.strftime('%d/%m/%Y %H:%M:%S')
    
    # Criando DataFrame para registros com Status Voalle = "Cancelado" E Status Aniel = "Fechada Produtiva"

    df_cancelado_fechada = df_final[(df_final["Status Voalle"] == " Cancelado") & (df_final["Status Aniel"] == "Fechada Produtiva")]
    
    # Aplicando filtro principal para remover status indesejados
    df_final = df_final[~df_final["Status Aniel"].isin(["Fechada Improdutiva", "Fechada Produtiva","Cancelado"])]
    pbar.update(1)
    time.sleep(1)

    # 6. Salvando arquivos
    pbar.set_description("Salvando arquivos")
    
    # Criando diretório se não existir
    output_dir = 'Excluir Aniel/Exclusão'
    os.makedirs(output_dir, exist_ok=True)
    
    df_final_sem_status = df_final.drop(columns=["Status Voalle", "Status Aniel"])
    df_final_sem_status.to_excel('Excluir Aniel/Exclusão/Excluir_Aniel.xlsx', index=False)
    df_final.to_excel('Excluir Aniel/Exclusão/Excluir_Aniel_com_Status.xlsx', index=False)
    df_cancelado_fechada.to_excel('Excluir Aniel/Exclusão/Cancelado-Voalle_Fechada_Produtiva-Aniel.xlsx', index=False)
    pbar.update(1)
    time.sleep(1)

print(tabulate(df_final_sem_status.head(), headers='keys', tablefmt='psql', showindex=False))
print("Mesclagem filtrada e salva em:")
print("- 'Excluir_Aniel.xlsx' (sem status)")
print("- 'Excluir_Aniel_com_Status.xlsx' (com status filtrado)")
print("- 'Cancelado_Fechada_Produtiva.xlsx' (Status Voalle=Cancelado E Status Aniel=Fechada Produtiva)")
print("✅✅✅✅")

