import pandas as pd

# Função para verificar duplicatas de nomes e números de telefone
def verificar_duplicatas(arquivo_excel):
    # Carrega o arquivo Excel em um DataFrame do pandas
    df = pd.read_excel(arquivo_excel)
    
    # Verifica duplicatas de nomes
    duplicatas_nomes = df[df.duplicated('Nome', keep=False)]
    
    # Verifica duplicatas de números de telefone
    duplicatas_telefones = df[df.duplicated('Telefone', keep=False)]
    
    return duplicatas_nomes, duplicatas_telefones

# Nome do arquivo Excel a ser verificado
nome_arquivo = 'lista_contatos.xlsx'

# Chama a função para verificar duplicatas
duplicatas_nomes, duplicatas_telefones = verificar_duplicatas(nome_arquivo)

# Exibe as duplicatas encontradas
print("Duplicatas de Nomes:")
print(duplicatas_nomes)
print("\nDuplicatas de Números de Telefone:")
print(duplicatas_telefones)
