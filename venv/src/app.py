# Importando as bibliotecas necessárias
import pandas as pd
import matplotlib.pyplot as plt
import smtplib
import dotenv 
from email.message import EmailMessage
import os
import ssl
from datetime import datetime
from dateutil import parser
import re

# Carregando as variáveis de ambiente do arquivo .env
dotenv.load_dotenv(dotenv.find_dotenv())

# Configurações para enviar email
email_sender = os.getenv('email_sender')
email_password = os.getenv('email_password')
email_receiver = os.getenv('email_receiver')

# Configurações do email
subject = 'Gráficos de Vendas'
body = 'Segue em anexo os gráficos gerados por meio dos dados provenientes da tabela.'

# Criar mensagem de email
em = EmailMessage()
em['From'] = email_sender
em['To'] = email_receiver
em['Subject'] = subject
em.set_content(body)

# Carregar a planilha Excel
tabela = pd.read_excel("venv/src/Relacao_Produtos_e_Clientes_2024.xlsx")

# Definindo as colunas de acordo com a tabela
coluna_produtos = 'Produto'  
coluna_pagamento = 'Método de Pagamento'
coluna_regiao = 'Região'
coluna_eqvenda = 'Equipe de Vendas'
coluna_valor = 'Valor da Venda'
coluna_desconto = 'Desconto (%)'
coluna_cliente = 'Cliente'
coluna_datas = 'Data da Venda'

# Função para verificar se o valor da venda é válido
def venda_verificar(valor):
    try:
        float(valor[1:])
        return True
    except (ValueError, TypeError):
        return False

# Função para padronizar os métodos de pagamento na tabela
def trad_pagamento():
    substituicoes = {
        'Cartão de Crédito': ['Cartão de Crédito', 'Cred.'],
        'Cartão de Débito': ['Cartão de Débito'],
        'Transferência Bancária': ['Transferência Bancária', 'Tran. Bancária' ],
        'Dinheiro': ['Dinheiro'],
        'Cheque': ['cheque']
    }

    for padrao, variacoes in substituicoes.items():
        tabela[coluna_pagamento] = tabela[coluna_pagamento].replace(variacoes, padrao)

# Função para converter a data para um formato padrão
def data_convertida(date):
    try:
        data_certa = re.search(r'\b\d{1,4}[-/]\d{1,2}[-/]\d{1,4}\b', str(date))
        if data_certa:
            parsed_date = parser.parse(data_certa.group(), dayfirst=True)
            return parsed_date.strftime('%d/%m/%Y')
        else:
            return None
    except ValueError:
        return None

# Função para gerar gráficos mais facilmente
def grafico(colunax, colunalegenda, coluna_valor, tipo, titulo, legenda, nome_arquivo, rotacao):
    grafico = df.pivot_table(index=colunax, columns=colunalegenda, values=coluna_valor, aggfunc='mean')
    grafico.plot(kind=tipo, figsize=(12, 6))
    plt.xlabel(colunax)
    plt.ylabel(coluna_valor)
    plt.title(titulo)
    plt.legend(title=legenda)
    plt.xticks(rotation=rotacao)
    plt.savefig(nome_arquivo, dpi=400)
    plt.close()

# Aplicar a padronização dos métodos de pagamento
trad_pagamento()

# Inicializar listas para armazenar os dados das colunas
array_produto = []
array_pagamento = []
array_regiao = []
array_valor = []
array_eqvenda = []
array_desconto = []
array_cliente = []
array_data = []

# Preencher os arrays com os valores das colunas
for index, row in tabela.iterrows():
    produto = row[coluna_produtos]
    pagamento = row[coluna_pagamento]
    regiao = row[coluna_regiao]
    valor = row[coluna_valor]
    eqvenda = row[coluna_eqvenda]
    cliente = row[coluna_cliente]
    data = row[coluna_datas]

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
    array_data.append(data_convertida(data))

# Preencher o array de descontos
for desconto in tabela[coluna_desconto]:
    if desconto >= 0:
        array_desconto.append(desconto)
    else:
        array_desconto.append(0)

# Transformar arrays em DataFrame para facilitar a manipulação de dados
df = pd.DataFrame({
    coluna_produtos: array_produto,
    coluna_pagamento: array_pagamento,
    coluna_regiao: array_regiao,
    coluna_valor: array_valor,
    coluna_eqvenda: array_eqvenda,
    coluna_cliente: array_cliente,
    coluna_desconto: array_desconto,
    coluna_datas: array_data
})

graficos = [
    {'colunax':coluna_produtos, 'colunalegenda':coluna_regiao, 'coluna_valor':coluna_valor, 'legenda':coluna_regiao,'nome_arquivo':'grafico_produtos_por_regiao.pdf','tipo':'bar','titulo':'Gráfico de Produtos por Região', 'rotacao':'horizontal'},
    {'colunax':coluna_eqvenda, 'colunalegenda':coluna_regiao, 'coluna_valor':coluna_valor, 'legenda':coluna_regiao,'nome_arquivo':'grafico_equipe_por_regiao.pdf', 'tipo':'bar', 'titulo':'Gráfico de Equipe Por Região', 'rotacao':'horizontal'},
    {'colunax':coluna_datas, 'colunalegenda':None, 'coluna_valor':coluna_valor, 'legenda':None,'nome_arquivo':'grafico_datas_por_equipe_venda.pdf','tipo':'line','titulo':'Gráfico de Vendas por Equipe de Venda', 'rotacao':'horizontal'},
    {'colunax':coluna_regiao, 'colunalegenda':coluna_pagamento, 'coluna_valor':coluna_valor, 'legenda':coluna_pagamento,'nome_arquivo':'grafico_regiao_por_metodo_pagamento.pdf','tipo':'bar','titulo':'Gráfico de Regiao por Método de Pagamento', 'rotacao':'horizontal'}
]

for config in graficos:
    grafico(config['colunax'], config['colunalegenda'], config['coluna_valor'], config['tipo'], config['titulo'], config['legenda'], config['nome_arquivo'], config['rotacao'])

# Anexar os arquivos de gráfico ao email
for config in graficos:
    nome_arquivo = config['nome_arquivo']
    with open(nome_arquivo, 'rb') as f:
        file_data = f.read()
        em.add_attachment(file_data, maintype='application', subtype='pdf', filename=nome_arquivo)

# Enviar o email
context = ssl.create_default_context()
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_password)
    smtp.send_message(em)

print('Email enviado com sucesso!')