import requests
from openpyxl import Workbook

# Parâmetros de autenticação para a API do GLPI
USER_TOKEN = 'JTAv8pMJOYoq2sgvx1w9LG2bTlIyU5HvLMrXxiQ1'
APP_TOKEN = 'KPl03aXgI8nXmUIKT7pwYbSDnSwvwHSNWbKmjAEx'
BASE_URL = 'https://nvirtual.with18.glpi-network.cloud/apirest.php'

# Função para iniciar sessão na API do GLPI
def iniciar_sessao():
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'user_token {USER_TOKEN}',
        'App-Token': APP_TOKEN
    }
    response = requests.get(f'{BASE_URL}/initSession', headers=headers)
    response.raise_for_status()
    return response.json()['session_token']

# Função para encerrar sessão na API do GLPI
def encerrar_sessao(session_token):
    headers = {
        'Content-Type': 'application/json',
        'Session-Token': session_token,
        'App-Token': APP_TOKEN
    }
    requests.get(f'{BASE_URL}/killSession', headers=headers)

# Função para obter todos os computadores do GLPI (com paginação)
def obter_todos_computadores(session_token):
    headers = {
        'Content-Type': 'application/json',
        'Session-Token': session_token,
        'App-Token': APP_TOKEN
    }
    computadores = []
    range_inicio = 0
    batch_size = 20  # Define o tamanho do lote (ajustável)

    while True:
        response = requests.get(
            f'{BASE_URL}/Computer?expand_dropdowns=true&get_hateoas=false&range={range_inicio}-{range_inicio + batch_size - 1}',
            headers=headers
        )
        response.raise_for_status()
        dados = response.json()

        # Filtrar apenas computadores que não foram excluídos (is_deleted == 0)
        computadores.extend([comp for comp in dados if comp.get('is_deleted') == 0])

        # Verifica o cabeçalho "Content-Range" para ver se há mais dados
        content_range = response.headers.get('Content-Range')
        total_items = int(content_range.split('/')[-1]) if content_range else len(dados)
        
        # Incrementa o range para o próximo lote
        range_inicio += batch_size

        # Se já obteve todos os itens, para o loop
        if range_inicio >= total_items:
            break

    return computadores

# Função para criar e salvar dados em um arquivo Excel
def criar_arquivo_excel(dados, caminho_arquivo):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Computadores GLPI"

    # Cabeçalhos
    colunas = ["Nome", "Etiqueta", "Número de série", "Cod. Ref.", "Categoria", "Fornecedores", "Localização", 
               "Status", "Fabricante", "Marca", "Modelo", "Tags", "Data de compra", "Data de garantia", "Custo", "Descrição", "Usado por"]
    sheet.append(colunas)

    # Preenchendo dados
    for computador in dados:
        linha = [
            computador.get('name'), #Nome
            computador.get('otherserial'), #Etiqueta
            computador.get('serial'), #Serial
            computador.get('otherserial'), #ID do GLPI Ref code
            "Estações de Trabalho", #Tipo de estação
            "",
            computador.get('locations_id'), #Localização
            computador.get('states_id'), #status
            computador.get('manufacturers_id'), #Fabricante
            computador.get('manufacturers_id'), #Marca
            computador.get('computermodels_id'), #Modelo do Computador
            computador.get('groups_id'), #tags
            "",
            "",
            "",
            computador.get('contact'), #Descrição
            computador.get('contact_num') #Utilizado por
        ]
        sheet.append(linha)

    # Salvando o arquivo
    workbook.save(caminho_arquivo)
    print(f"Arquivo Excel salvo em: {caminho_arquivo}")

# Execução do script
if __name__ == "__main__":
    session_token = iniciar_sessao()
    try:
        computadores = obter_todos_computadores(session_token)
        print("Computadores recuperados do GLPI:")
        for computador in computadores:
            print(computador)
        
        # Salva os dados em um arquivo Excel
        criar_arquivo_excel(computadores, 'Computadores_GLPI.xlsx')
    finally:
        encerrar_sessao(session_token)
