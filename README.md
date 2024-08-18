# Análise de Ações da Bolsa de Valores
Este projeto é uma aplicação em Python que permite ao usuário analisar ações da bolsa de valores, obtendo informações importantes como preço atual, Dividend Yield, P/L e ROE. A interface gráfica é construída com Tkinter, e os dados são extraídos do site Status Invest e salvos em um arquivo Excel.

## Funcionalidades
- Adicionar ações para análise através de uma interface gráfica.
- Obter dados atualizados sobre as ações selecionadas.
- Salvar os dados coletados em um arquivo Excel.
## Tecnologias Utilizadas
- Python 3
- Tkinter para a interface gráfica
- Requests para realizar requisições HTTP
- BeautifulSoup para parsing do HTML
- Pandas para manipulação de dados
- Openpyxl para manipulação de arquivos Excel

# Financial Data Automation Project

Este projeto foi desenvolvido em Python para automatizar a coleta de dados financeiros de ações a partir do site [Status Invest](https://statusinvest.com.br/). Ele realiza cálculos e análises diretamente em uma planilha Excel, facilitando a análise fundamentalista para investidores, economizando tempo e minimizando erros de digitação.

## Funcionalidades

- **Coleta Automática de Dados:** Usa web scraping para extrair informações financeiras importantes das ações no Status Invest.
- **Manipulação de Planilhas:** Organiza os dados coletados em uma planilha Excel, com diferentes abas para facilitar o acesso e a análise.
- **Análise Fundamentalista:** Realiza cálculos automáticos para métricas financeiras como valuation e margem de segurança, diretamente na planilha.
- **Interface Gráfica (em desenvolvimento):** Um aplicativo com interface gráfica, utilizando Tkinter, para tornar a interação mais intuitiva.

## Estrutura do Projeto

- **main.py:** Script principal que faz o web scraping, manipula os dados e realiza os cálculos na planilha Excel.
- **gui.py:** (Em desenvolvimento) Interface gráfica para facilitar a interação com o usuário.
- **requirements.txt:** Lista de dependências do projeto.

## Como Usar

1. Clone este repositório:
   ```bash
   git clone https://github.com/seu-usuario/financial-data-automation.git
