import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import date

# Abrir a janela de diálogo para selecionar o arquivo Excel
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

# Verificar se o usuário selecionou um arquivo
if file_path:
    # Ler o arquivo
    df = pd.read_excel(file_path)

    # Obter a data de hoje
    today = date.today()

    # Adicionar a data e a palavra "(Formatado)" ao final do nome do arquivo
    file_name = file_path.split("/")[-1]
    new_file_name = (
        file_name.split(".xlsx")[0]
        + "_"
        + today.strftime("%Y%m%d")
        + "(Formatado).xlsx"
    )

    # Lista de colunas para excluir
    colunas_para_excluir = [
        "Cliente",
        "Descrição",
        "Grupo",
        "Email",
        "Faturável",
        "Data final",
        "Duração (decimal)",
        "Taxa Faturável (USD)",
        "Valor Faturável (USD)",
    ]

    # Verificar se as colunas existem no DataFrame
    if "Do utilizador" in df.columns:
        df.rename(columns={"Do utilizador": "Nome"}, inplace=True)
    if "Duração (h)" in df.columns:
        df.rename(columns={"Duração (h)": "Tempo Alocado"}, inplace=True)
        # Formatar a coluna "Tempo Alocado" para o formato HH:MM:SS
        df["Tempo Alocado"] = pd.to_datetime(
            df["Tempo Alocado"], format="%H:%M:%S"
        ).dt.time
    if "Tarefa" in df.columns:
        df.rename(columns={"Tarefa": "Atividade"}, inplace=True)
    if "Hora de início" in df.columns:
        df.rename(columns={"Hora de início": "Hora Inicial"}, inplace=True)
    if "Fim do tempo" in df.columns:
        df.rename(columns={"Fim do tempo": "Hora Final"}, inplace=True)
    colunas_para_excluir = [col for col in colunas_para_excluir if col in df.columns]

    # Definir a coluna "Nome" como a primeira coluna
    df = df[["Nome"] + [col for col in df.columns if col != "Nome"]]

    # Excluir as colunas
    df = df.drop(colunas_para_excluir, axis=1)

    # Somar a coluna "Tempo Alocado" e inserir em uma nova coluna chamada "Total de Tempo Alocado"
    # total_tempo_alocado = df["Tempo Alocado"].sum()
    # df["Total de Tempo Alocado"] = total_tempo_alocado

    # Salvar o DataFrame formatado em um novo arquivo
    new_file_name = (
        file_name.split(".xlsx")[0]
        + " "
        + today.strftime("%d-%m-%Y")
        + " (Formatado).xlsx"
    )
    df.to_excel(new_file_name, index=False)
