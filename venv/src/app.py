import pandas as pd
from datetime import datetime
import re
from dateutil import parser
import matplotlib.pyplot as plt
import numpy as np
import smtplib

# Carregar a planilha Excel
tabela = pd.read_excel("venv/src/Relacao_Produtos_e_Clientes_2024.xlsx")

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
    else:
        array_desconto.append(0)



# Imprimir os primeiros 10 elementos de cada array para verificar
print("Array de Produtos:", array_produto)
# print("Array de Pagamentos:", array_pagamento)
print("Array de Regiões:", array_regiao)
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


df_produtos = pd.DataFrame({
    coluna_produtos: array_produto,
})
df_valor = pd.DataFrame({
    coluna_valor: array_valor
    })
df_regiao = pd.DataFrame({
    coluna_regiao: array_regiao
})

df = pd.DataFrame({
    coluna_produtos: array_produto,
    coluna_pagamento: array_pagamento,
    coluna_regiao: array_regiao,
    coluna_valor: array_valor,
    coluna_eqvenda: array_eqvenda,
    coluna_cliente: array_cliente,
    coluna_desconto: array_desconto
})

df2 = pd.DataFrame({
    coluna_produtos: array_produto,
    coluna_regiao: array_regiao,
    coluna_valor: array_valor,
})

df2_filtrado = df2[df2[coluna_valor] != 0]

# Calculando a média dos valores de venda por produto e região
pivot_df = df2_filtrado.pivot_table(index=coluna_produtos, columns='Região', values=coluna_valor, aggfunc=lambda x: x.mean() if (x != 0).any() else None)

# Plotando o gráfico de barras agrupadas
pivot_df.plot(kind='bar', figsize=(12, 6))
plt.xlabel(coluna_produtos)
plt.ylabel(coluna_valor)
plt.title(f'Gráfico de Barras Agrupadas (Média dos Valores de {coluna_valor})')
plt.legend(title='Região')

# Exibindo o gráfico
plt.tight_layout()
plt.show()

email_sender = 'mat.cesco@gmail.com'