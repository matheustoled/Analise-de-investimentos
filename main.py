import requests
from bs4 import BeautifulSoup
import pandas as pd
from tkinter import *
import tkinter as tk
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule

#Lista com ações que o usuário escolheu
lista_acoes = []

#Definindo cores
preto = "#000000" #black
branco = "#f1ebeb" #white
azul = "#24c0eb" #blue
cinza = "#cacaca" #gray

# Cria a janela principal
janela = tk.Tk()
janela.title("Análise de Ações")
janela.geometry("375x185")

# Cria um frame para a entrada e o botão
frame = tk.Frame(janela)
frame.pack(pady=20, padx=20, fill='both', expand=True)

# Cria um campo de entrada dentro do frame
entrada = tk.Entry(frame, width=30)
entrada.pack(pady=10)

#Criando função de entrada da informação
def adicionar_acao():
    acao = entrada.get()
    if acao:  # Verifica se o campo de entrada não está vazio
        lista_acoes.append(acao)
        entrada.delete(0, tk.END)  # Limpa o campo de entrada após adicionar o nome
        print(f"Ação '{acao}' adicionado(a)!")
    else:
        print("O campo de entrada está vazio. Por favor, digite uma ação.")

# Cria um botão dentro do frame para enviar o nome
botao = tk.Button(frame, text="Adicionar", command=adicionar_acao, bg=azul, fg=branco, font=("Uvy 13 bold"), relief=RAISED, overrelief=RIDGE)
botao.pack(pady=10)

#Funçao para pegar as informações de cada ativo
def obter_dados_ativos(ativo):

    #Formatar a URL corretamente com o ativo que o usuário escolheu
    url = "https://statusinvest.com.br/acoes/{}".format(ativo)

    #Definir os cabeçalhos HTTP para imitar um navegador real
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    #Fazer request do site para puxar as infos
    requisicao = requests.get(url, headers=headers)

    #ler todo o conteúdo HTML da página
    site = BeautifulSoup(requisicao.content, "html.parser")

    #Encontrar as informações desejadas
    preco_atual = site.find('strong', {'class': 'value'}).text.strip()
    infos = site.find_all('strong', {'class': 'value d-block lh-4 fs-4 fw-700'})
    dy = infos[0].text.strip()
    pl = infos[1].text.strip()
    roe = infos[24].text.strip()

    print(infos)

    # Organizar os dados em um dicionário
    return {
        'Ativo': ativo,
        'Preço Atual': preco_atual,
        'Dividend Yield': dy,
        'P/L': pl,
        'ROE': roe
    }

#Função para salvar os dados em Excel
def salvar_dados_excel():
    
    # Lista para armazenar os dados de cada ativo separadamente
    dados_ativos = []

    # Obter os dados para cada ativo
    for i in lista_acoes:
        i = i.strip()  # Remover espaços em branco extras
        try:
            dados = obter_dados_ativos(i)
            dados_ativos.append(dados)
        except Exception as e:
            print(f"Erro ao obter dados para o ativo {i}: {e}")

    # Criar um DataFrame a partir do dicionário
    df = pd.DataFrame(dados_ativos)

    # Exportar o DataFrame para um arquivo Excel
    df.to_excel('Investimentos.xlsx', index=False)
    print("Dados salvos!")

# Cria um botão para salvar os dados em Excel
botao_salvar = tk.Button(frame, text="Salvar Dados", command=salvar_dados_excel, bg=azul, fg=branco, font=("Uvy 13 bold"), relief=RAISED, overrelief=RIDGE)
botao_salvar.pack(pady=10)

# Executa o loop principal da interface
janela.mainloop()