#Requests e web scraping
import requests
from bs4 import BeautifulSoup
#Manipulação de planilhas
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl import load_workbook
#Criação de interface
from tkinter import *
import tkinter as tk

#Carregar planilha
planilha = load_workbook('Investimentos.xlsx')

#Lista com ações que o usuário escolheu
lista_acoes = []

#Definindo cores
preto = "#000000" #black
branco = "#f1ebeb" #white
azul = "#24c0eb" #blue
cinza = "#cacaca" #gray

#Define cores para a formatação condicional
green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

#Cria a janela principal
janela = tk.Tk()
janela.title("Análise de Ações")
janela.geometry("375x185")

#Cria um frame para a entrada e o botão
frame = tk.Frame(janela)
frame.pack(pady=20, padx=20, fill='both', expand=True)

#Cria um campo de entrada dentro do frame
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

#Cria um botão dentro do frame para enviar o nome
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
    infos2 = site.find_all('strong', {'class': 'value'})
    margem_liquida = infos[23].text.strip()
    div_liquida_patrimonio = infos[14].text.strip()
    roic = infos[26].text.strip()
    div_liquida_ebitda = infos[15].text.strip()
    tag_along = infos2[6].text.strip()
    pl = infos[1].text.strip()
    p_vp = infos[3].text.strip()
    dy = infos[0].text.strip()
    liq_corrente = infos[19].text.strip()
    roe = infos[24].text.strip()
    ev_ebitda = infos[4].text.strip()
    lpa = infos[10].text.strip()

    print(infos)

    # Organizar os dados em um dicionário
    return {
        'Ativo': ativo,
        'Preço Atual': preco_atual,
        'Margem Líquida': margem_liquida,
        'Dívida Líquida/Patrimônio': div_liquida_patrimonio,
        'ROIC': roic,
        'Dívida Líquida/Ebitida': div_liquida_ebitda,
        'Tag Along': tag_along,
        'P/L': pl,
        'P/VP': p_vp,
        'D.Y': dy,
        'LIQ. CORRENTE': liq_corrente,
        'ROE': roe,
        'EV/EBITDA': ev_ebitda,
        'LPA': lpa
    }

#Função para salvar os dados em Excel
def salvar_dados_excel():
    
    #Lista para armazenar os dados de cada ativo separadamente
    dados_ativos = []

    #Obter os dados para cada ativo
    for i in lista_acoes:
        i = i.strip()  #Remover espaços em branco extras
        try:
            dados = obter_dados_ativos(i)
            dados_ativos.append(dados)
        except Exception as e:
            print(f"Erro ao obter dados para o ativo {i}: {e}")

    #Criar um DataFrame a partir do dicionário
    df = pd.DataFrame(dados_ativos)

    #Exportar o DataFrame para um arquivo Excel
    df.to_excel('Investimentos.xlsx', index=False)
    print("Dados salvos!")

#Cria um botão para salvar os dados em Excel
botao_salvar = tk.Button(frame, text="Salvar Dados", command=salvar_dados_excel, bg=azul, fg=branco, font=("Uvy 13 bold"), relief=RAISED, overrelief=RIDGE)
botao_salvar.pack(pady=10)

#Executa o loop principal da interface
janela.mainloop()