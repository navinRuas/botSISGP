# Autor: Navin Ruas
import os
import re

# Função para obter o caminho absoluto de um arquivo relativo
def gap(relative_path):
    # Obtém o diretório do script atual
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Concatena o diretório do script com o caminho relativo do arquivo
    absolute_path = os.path.join(script_dir, relative_path)
    return absolute_path

# Função para personalizar um arquivo HTML substituindo valores de placeholders
def personalizar_html(arquivo_html, valores):
    with open(arquivo_html, 'r') as f:
        html = f.read()
    for chave, valor in valores.items():
        # Substitui as ocorrências de {chave} no HTML pelo valor correspondente
        html = html.replace('{' + chave + '}', str(valor))
    return html

# Função para escapar caracteres especiais em HTML
def html_escape(text):
    html_escape_table = {
        '<': '&lt;',
        '>': '&gt;'
    }
    return "".join(html_escape_table.get(c, c) for c in text)

# Função para corrigir a codificação de um texto
def corrigir_codificacao(texto):
    correcoes_codificacao = {
        'Ã¡': 'á',
        'Ã ': 'à',
        'Ã¢': 'â',
        'Ã£': 'ã',
        'Ã©': 'é',
        'Ãª': 'ê',
        'Ã­': 'í',
        'Ã³': 'ó',
        'Ã²': 'ò',
        'Ã´': 'ô',
        'Ãµ': 'õ',
        'Ãº': 'ú',
        'Ã¼': 'ü',
        'Ã§': 'ç'
    }
    for codigo_errado, correcao in correcoes_codificacao.items():
        # Substitui os caracteres com codificação errada pelos caracteres corretos
        texto = texto.replace(codigo_errado, correcao)
    return texto

# Função para extrair o conteúdo de uma tag específica de um texto HTML
def stripFunc(striptext, tagname):
    regex = r'{0}>(.*?)<\/{0}'.format(tagname)
    match = re.search(regex, striptext)

    if match:
        result = match.group(1)
        if result == "":
            result = ''
        return result

# Função para remover caracteres não numéricos de um texto
def stripTrash(striptext):
    numbers = ''.join(filter(str.isdigit, striptext))
    return numbers

# Função para normalizar um texto, convertendo para minúsculas e removendo zeros à esquerda
def normalize(s):
    s = s.lower()
    s = re.sub(r'\b0+(\d)', r'\1', s)
    return s
