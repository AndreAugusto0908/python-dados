import random
from datetime import datetime, timedelta
import pandas as pd
import numpy as np


def tempo_medio_paciente_na_sala(df):
    df['duracao_procedimento'] = (df['retirada_paciente'] - df['entrada_paciente_sala']).dt.total_seconds() / 60
    df['mes'] = df['entrada_paciente_sala'].dt.month

    tempo_medio_por_mes = df.groupby('mes')['duracao_procedimento'].mean()

    resultado = [tempo_medio_por_mes.get(mes, 0) for mes in range(1, 7)]
    return resultado

def contar_cirurgias_por_mes(df):
    df['mes'] = df['entrada_paciente_sala'].dt.month

    contagem_por_mes = df['mes'].value_counts().sort_index()

    resultado = [contagem_por_mes.get(mes, 0) for mes in range(1, 7)]
    return resultado

def tempo_medio_montagem_da_sala(df):
    df['mes'] = df['inicio_montagem'].dt.month
    df['duracao_procedimento'] = (df['fim_montagem'] - df['inicio_montagem']).dt.total_seconds() / 60

    tempo_medio_por_mes = df.groupby('mes')['duracao_procedimento'].mean()

    resultado = [tempo_medio_por_mes.get(mes, 0) for mes in range(1, 7)]
    return resultado

def contar_cirurgias_por_sala_e_mes(df):
    df['mes'] = df['entrada_paciente_sala'].dt.month

    cirurgias_por_sala_e_mes = df.groupby(['sala', 'mes']).size().unstack(fill_value=0)

    return cirurgias_por_sala_e_mes

def media_tempo_cirurgia_por_sala(df):
    df['duracao_cirurgia'] = (df['retirada_paciente'] - df['entrada_paciente_sala']).dt.total_seconds() / 60

    media_por_sala = df.groupby('sala')['duracao_cirurgia'].mean()

    return media_por_sala



df = pd.read_csv('horarios_cirurgia.csv', parse_dates=[
    "chegada_anestesista", "chegada_cirurgiao", "entrada_paciente_sala",
    "inicio_montagem", "fim_montagem", "inicio_desmontagem", "fim_desmontagem",
    "inicio_higienizacao", "fim_higienizacao", "retirada_paciente",
    "horario_entrada_bloco", "horario_liberacao"
])

tempo_medio_paciente = tempo_medio_paciente_na_sala(df)
tempo_medio_montagem = tempo_medio_montagem_da_sala(df)
cirurgias_por_mes = contar_cirurgias_por_mes(df)
cirurgias_por_sala_mes = contar_cirurgias_por_sala_e_mes(df)
media_tempo_sala = media_tempo_cirurgia_por_sala(df)

meses = range(1, 7)
dados_resumo = pd.DataFrame({
    "Mês": meses,
    "Tempo Médio Paciente na Sala (min)": np.round(tempo_medio_paciente, 0),
    "Tempo Médio Montagem da Sala (min)": np.round(tempo_medio_montagem, 0),
    "Número de Cirurgias": cirurgias_por_mes
})

media_tempo_sala = media_tempo_sala.round(0)
cirurgias_por_sala_mes = cirurgias_por_sala_mes.round(0)

with pd.ExcelWriter('resumo_cirurgias.xlsx') as writer:
    dados_resumo.to_excel(writer, sheet_name='Resumo por Mês', index=False)
    cirurgias_por_sala_mes.to_excel(writer, sheet_name='Cirurgias por Sala', index=True)
    media_tempo_sala.to_excel(writer, sheet_name='Média Tempo por Sala', index=True)