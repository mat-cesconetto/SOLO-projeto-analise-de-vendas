import pandas as pd
from datetime import datetime
import re
from dateutil import parser
import matplotlib.pyplot as plt
import numpy as np


# Carregar a planilha Excel
tabela = pd.read_excel("Relacao_Produtos_e_Clientes_2024.xlsx")

# Selecionar a coluna desejada
coluna_produtos = 'Produto'  
coluna_pagamento = 'Método de Pagamento'
coluna_regiao = 'Região'
coluna_eqvenda = 'Equipe de Vendas'
coluna_valor = 'Valor da Venda'
coluna_desconto = 'Desconto (%)'
coluna_cliente = 'Cliente'
coluna_datas = 'Data da Venda'

array_produto = []
array_pagamento = []
array_regiao = []
array_valor = []
array_eqvenda = []
array_desconto = []
array_cliente = []

# Contar as ocorrências de cada valor na coluna selecionada
def venda_verificar(valor):
    try:
        float(valor[1:])
        return True
    except (ValueError, TypeError):
        return False

def trad_pagamento():
    substituicoes = {
        'cartao_credito': ['Cartão de Crédito', 'Cred.'],
        'cartao_debito': ['Cartão de Débito'],
        'transf_banc': ['Transferência Bancária', 'Tran. Bancária' ],
        'dinheiro': ['Dinheiro'],
        'cheque': ['cheque']
    }

    for padrao, variacoes in substituicoes.items():
        tabela[coluna_pagamento] = tabela[coluna_pagamento].replace(variacoes, padrao)

trad_pagamento()
for index, row in tabela.iterrows():
    # Acessar o valor da coluna 'coluna_valor' (substitua com o nome real da coluna)
    produto = row[coluna_produtos]
    pagamento = row[coluna_pagamento]
    regiao = row[coluna_regiao]
    valor = row[coluna_valor]
    eqvenda = row[coluna_eqvenda]
    cliente = row[coluna_cliente]

    array_produto.append(produto)

    array_pagamento.append(pagamento)

    array_regiao.append(regiao)

    if venda_verificar(valor):
        valor_new = re.sub(r'\$', '', valor)
        array_valor.append(float(valor_new))
    else:
        array_valor.append(0)

    array_eqvenda.append(eqvenda)

    array_cliente.append(cliente)

for desconto in tabela[coluna_desconto]:
    if desconto >= 0:
        array_desconto.append(desconto)



# Imprimir os primeiros 10 elementos de cada array para verificar
print("Array de Produtos:", array_produto)
# print("Array de Pagamentos:", array_pagamento)
# print("Array de Regiões:", array_regiao)
print("Array de Valores de Venda:", array_valor)
# print("Array de Equipes de Vendas:", array_eqvenda)
# print("Array de Desconto:", array_desconto)
# print("Array de Cliente:", array_cliente)

# def padronizar_data(data):
#     try:
#         # Verifica se é um número inteiro (não é uma data válida)
#         if isinstance(data, int):
#             return None
#         # Se a data já for um objeto Timestamp, retorna apenas a data sem hora
#         elif isinstance(data, pd.Timestamp):
#             return data.strftime('%Y-%m-%d')
#         # Se a data for uma string, converte usando parse e retorna apenas a data sem hora
#         return parser.parse(str(data)).strftime('%Y-%m-%d')
#     except Exception as e:
#         print(f"Erro ao padronizar data: {e}")
#         return None

# # Aplica a padronização na coluna de datas
# tabela['Data Padronizada'] = tabela[coluna_datas].apply(padronizar_data)

# print(tabela[[coluna_datas, 'Data Padronizada']])
# Agrupar valores por região
plt.figure(figsize=(10, 6))

# Ajuste de posição para barras laterais
posicoes = np.arange(len(array_produto))

# Plotagem das barras por região
for regiao in set(array_regiao):
    indices = [i for i, value in enumerate(array_regiao) if value == regiao]
    produtos_regiao = [array_produto[i] for i in indices]
    valores_regiao = [array_valor[i] for i in indices]
    plt.barh(produtos_regiao, valores_regiao, height=0.2, label=f'Vendas - {regiao}')

# Ajustes de legenda e labels
plt.xlabel('Valores de Venda ($)')
plt.ylabel('Produtos')
plt.title('Valores de Venda por Produto e Região')
plt.legend()  # Adiciona a legenda com base nos rótulos

plt.tight_layout() 
plt.show()