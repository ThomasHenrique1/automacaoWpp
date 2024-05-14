import pandas as pd
from datetime import datetime

# Função para calcular a idade com base na data de nascimento
def calcular_idade(data_nascimento):
    data_atual = datetime.now()
    ano_atual = data_atual.year
    try:
        # Extrai o ano da data de nascimento fornecida
        ano_nascimento = data_nascimento.year
        idade = ano_atual - ano_nascimento
        return idade
    except (ValueError, TypeError):
        return None

def formatar_telefone(numero):
    if pd.isnull(numero) or numero == "":
        return None
    numero_str = str(numero)
    if "-" in numero_str:  # Verifica se contém hífen, considera formato "55(99) 99999-9999"
        numero_str = numero_str.replace("-", "")
        return f"55({numero_str[:2]}) {numero_str[2:7]}-{numero_str[7:]}"
    elif len(numero_str) == 11:  # Verifica se tem 11 dígitos, considera formato apenas números
        ddd = numero_str[:2]
        telefone = numero_str[2:]
        return f"55({ddd}) {telefone[:5]}-{telefone[5:]}"
    else:
        return "Formato de telefone inválido"

def processar_lista(nome_arquivo, colunas_necessarias):
    # Carregar o arquivo Excel para um DataFrame
    df = pd.read_excel(nome_arquivo)

    # Verificar se todas as colunas necessárias estão presentes
    colunas_presentes = df.columns.tolist()
    for coluna in colunas_necessarias:
        if coluna not in colunas_presentes:
            print(f"Coluna '{coluna}' não encontrada no arquivo '{nome_arquivo}'.")
            return  # Sair da função se alguma coluna necessária não estiver presente

    # Verificar e formatar as colunas de telefone, se necessário
    for coluna in ['Telefone', 'Telefone1', 'Telefone Residencial']:
        if coluna in colunas_presentes:
            formatado = df[coluna].apply(formatar_telefone).equals(df[coluna])
            if formatado:
                print(f"Coluna '{coluna}' do arquivo '{nome_arquivo}' já está formatada.")
            else:
                df[coluna] = df[coluna].apply(formatar_telefone)

    # Calcular e preencher a coluna 'Idade' se 'Data de Nascimento' estiver presente
    if 'Data de Nascimento' in colunas_presentes:
        df['Idade'] = pd.to_datetime(df['Data de Nascimento'], format='%d/%m/%Y', dayfirst=True, errors='coerce').apply(calcular_idade)

    # Organizar as colunas na ordem desejada
    df = df[colunas_necessarias]

    # Salvar o DataFrame de volta como arquivos CSV e Excel
    arquivo_saida_csv = nome_arquivo.replace('.xlsx', '_formatado.csv')
    arquivo_saida_excel = nome_arquivo.replace('.xlsx', '_formatado.xlsx')
    df.to_csv(arquivo_saida_csv, index=False)
    df.to_excel(arquivo_saida_excel, index=False)

    print(f"Arquivo '{arquivo_saida_csv}' (CSV) e '{arquivo_saida_excel}' (Excel) processados e salvos com sucesso.")

# Listas de arquivos e colunas necessárias
arquivos = ['lista_contatos.xlsx']
colunas_necessarias = ['Nome', 'Telefone', 'Telefone1', 'Telefone Residencial', 'Idade', 'Data de Nascimento', 'Email']

# Processar cada lista
for arquivo in arquivos:
    processar_lista(arquivo, colunas_necessarias)
