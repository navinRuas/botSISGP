#Autor: Navin Ruas
import os
import json
import mysql.connector
from datetime import datetime
from extraUtils import gap

def update_ano_json():
    current_year = str(datetime.now().year)
    repo_path = None

    print("Inicializando Atualização ano.json...")

    # Localiza o repositório Gerador-Desc no disco do usuário
    for root, dirs, files in os.walk(os.path.expanduser("~")):
        if '.git' in dirs and os.path.basename(root) == 'geradorDesc':
            repo_path = root
            break

    if not repo_path:
        print("Repositório Gerador-Desc não encontrado no disco do usuário.")
        return

    # Carrega o arquivo ano.json
    with open(os.path.join(repo_path, 'ano.json'), 'r') as f:
        data = json.load(f)

    # Verifica se o ano atual está nos dados
    years = [item['value'] for item in data]
    if current_year not in years:
        # Adiciona o ano atual e remove o valor do ano mais antigo
        data.append({'value': current_year, 'text': current_year})
        min_year = min([int(year) for year in years if year.isdigit()])
        data = [item for item in data if not item['value'].isdigit() or int(item['value']) != min_year]

        # Atualiza o arquivo ano.json
        with open(os.path.join(repo_path, 'ano.json'), 'w') as f:
            json.dump(data, f)
        print("Ano do Gerador foi atualizado!")
    else:
        print("Ano do Gerador está atualizado")

def update_depara_json():
    repo_path = None

    print("Inicializando Atualização De-Para.json...")

    # Localiza o repositório Gerador-Desc no disco do usuário
    for root, dirs, files in os.walk(os.path.expanduser("~")):
        if '.git' in dirs and os.path.basename(root) == 'geradorDesc':
            repo_path = root
            break

    if not repo_path:
        print("Repositório Gerador-Desc não encontrado no disco do usuário.")
        return

    # Carrega o arquivo config.json
    with open(gap('sec\\config.json'), 'r', encoding='utf-8') as f:
        config = json.load(f)

    # Conecta ao servidor MySQL
    conn = mysql.connector.connect(
        host=config["dbHost"],
        port=config["dbPort"],
        user=config["dbUsername"],
        password=config["dbPassword"],
        database=config["dbName"]
    )
    cursor = conn.cursor()

    # Consulta o servidor MySQL para obter os dados
    cursor.execute("SELECT * FROM SISGP.`De-Para`")
    rows = cursor.fetchall()

    # Cria um novo objeto de dados com base nos dados do SQL
    data = []
    if rows is not None:
        for row in rows:
            item = {
                'CodDemanda': row[0],
                'Tipo de Demanda': row[1],
                'CodAtividade': row[2],
                'Atividade': row[3],
                'CodProduto': row[4],
                'Produto': row[5],
                'Atividade2': row[6],
                'nº da atividade': row[7],
                'Atividade PGD': row[8]
            }
            data.append(item)

    # Atualiza o arquivo depara.json
    with open(os.path.join(repo_path, 'depara.json'), 'w', encoding='utf-8') as f:
        json.dump(data, f)

    print("De-Para do Gerador foi atualizado!")

def webUP():
    repo_path = None

    print("Inicializando Atualização De-Para.json...")

    # Localiza o repositório Gerador-Desc no disco do usuário
    for root, dirs, files in os.walk(os.path.expanduser("~")):
        if '.git' in dirs and os.path.basename(root) == 'geradorDesc':
            repo_path = root
            break

    if not repo_path:
        print("Repositório Gerador-Desc não encontrado no disco do usuário.")
        return

    print("Inicializando Atualização Gerador Descrição Web...")
    update_ano_json()
    update_depara_json()
    print("Atualização Gerador Descrição Web Concluída!")

if __name__ == '__main__':
    webUP()