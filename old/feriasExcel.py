# -*- coding: utf-8 -*-

import pandas as pd
import json
import os
from datetime import datetime, timedelta

def ler_quadro_ferias(ferias_dir):
    # Encontra o diretório do ano atual
    ano_atual = str(datetime.now().year)
    ano_dir = os.path.join(ferias_dir, ano_atual)

    # Encontra o arquivo do quadro de férias
    quadro_ferias = None
    for file in os.listdir(ano_dir):
        if file.startswith('Quadro de Férias Audin') and file.endswith('.xlsx'):
            quadro_ferias = os.path.join(ano_dir, file)
            break

    if not quadro_ferias:
        print(f'Arquivo do quadro de férias não encontrado em {ano_dir}')
        return None

    # Lê o arquivo Excel e armazena em um DataFrame
    try:
        df = pd.read_excel(quadro_ferias)
    except Exception as e:
        print(f'Erro ao abrir o arquivo {quadro_ferias}: {e}')
        return None

    # Encontra a linha com o valor 'SERVIDOR/COLABORADOR'
    servidor_colaborador_row = df[df.eq('SERVIDOR/COLABORADOR').any(1)].index[0]

    # Lê os dados abaixo da linha 'SERVIDOR/COLABORADOR' até a coluna 'TOTAL'
    data = {}
    for row in range(servidor_colaborador_row + 1, len(df)):
        name = df.iloc[row, 2]
        if name == 'TOTAL':
            break
        if pd.isna(name):
            continue
        data[name] = []
        for col in range(3, len(df.columns), 2):
            start_date = df.iloc[row, col]
            duration = df.iloc[row, col + 1]
            if pd.isna(start_date) or pd.isna(duration):
                continue
            end_date = pd.to_datetime(start_date) + timedelta(days=int(duration))
            data[name].append({'inicio': start_date.strftime('%Y-%m-%d'), 'fim': end_date.strftime('%Y-%m-%d')})

    # Escreve os dados em um arquivo JSON
    with open('ferias.json', 'w') as f:
        json.dump(data, f)

ler_quadro_ferias('X:\\01 - ADMINISTRATIVO\\04 - Pessoal AUDIN\\FÉRIAS\\')