from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from openpyxl import load_workbook
import pandas as pd

# definição das variaveis da string do exporta dados
with open('empresa_mae.txt') as f_1:
    empresa_mae = f_1.read()

with open('chave.txt') as f_2:
    chave_ed = f_2.read()

data_inicio = '01/01/2022'
data_fim = '07/07/2022'

# definição dos caminhos para leitura e criação dos arquivos excel
leitura = r'C:\Users\desenvolvimento02\Desktop\esocial_teste.xlsx'
criacao = rf'C:\Users\desenvolvimento02\Desktop\Recibos eSocial\Recibos_eSocial_'

# leitura do arquivo excel com nome e cod das empresas
empresas_df = pd.read_excel(leitura)

# lista de eventos para virar DF
evento = []

# condição para não inserir na lista caso só tenha o cabeçalho
condicao = 'EVENTO;DATAGERACAO;UNIDADE;NOMEUNIDADE;CODIGOGED;CODIGOARQUIVOGED;' \
            'NOMEARQUIVO;FUNCIONARIO;NOMEFUNCIONARIO;STATUSEVENTO;CODIGOLOTE;NRRECIBO;' \
            'IDARQUIVO;ERRO;SEQUENCIAL;AMBIENTEPRODUCAO;CODIGOERROESOCIAL;DATAINICIOCONDICAO;CARGAINICIAL'

# estrutura de repetição para os codigos e nomes das empresas do arquivo xlsx
for codigo in empresas_df['cod']:
    indice = empresas_df[empresas_df['cod'] == codigo].index
    nome = empresas_df['nome'].values[indice]

    navegador = webdriver.Chrome()

# navegador receberá a string do exporta dados do SOC referente aos eventos do eSocial
    navegador.get("https://ws1.soc.com.br/WebSoc/exportadados?parametro={"
                  + f"'empresa':'{empresa_mae}','codigo':'141273',"
                  f"'chave':'{chave_ed}','tipoSaida':'txt','empresaTrabalho':{str(codigo)}'"
                  f"','dataInicio':'{data_inicio}','dataFim':'{data_fim}',"
                  "'status':'2','layout':'0','unidade':'0','ambiente':'1','cabecalho':'1'}")
    csv = WebDriverWait(navegador, 50).until(ec.element_to_be_clickable((By.XPATH, '/html/body/pre'))).text

# tratativa do texto copiado do exporta dados gerando um xlsx para cada empresa
    plan = csv.split("\n")
    for linha in plan:
        if linha == condicao:
            continue
        coluna = linha.split(';')
        evento.append(coluna)
    navegador.close()

    inconsistencias_df = pd.DataFrame(evento)

    empresa = str(nome)

    print(empresa)
    fim = (len(empresa) - 2)

    nome_empresa = empresa[2:fim]

    inconsistencias_df.to_excel(rf'{criacao}{nome_empresa}.xlsx',
                                index=False)
    evento = []

# Alterando a parte visual das planilhas e removendo as colunas com dados desnecessarios
    wb = load_workbook(rf'{criacao}{nome_empresa}.xlsx')
    ws = wb.active

# removendo as colunas com dados desnecessarios
    ws.delete_cols(19)
    ws.delete_cols(17)
    ws.delete_cols(16)
    ws.delete_cols(15)
    ws.delete_cols(14)
    ws.delete_cols(13)
    ws.delete_cols(11)
    ws.delete_cols(7)
    ws.delete_cols(6)
    ws.delete_cols(5)
    ws.delete_cols(3)
    ws.delete_cols(2)

# dando nome para as colunas
    ws['A1'] = 'EVENTO'
    ws['B1'] = 'NOMEUNIDADE'
    ws['C1'] = 'FUNCIONARIO'
    ws['D1'] = 'NOMEFUNCIONARIO'
    ws['E1'] = 'STATUSEVENTO'
    ws['F1'] = 'NRRECIBO'
    ws['G1'] = 'DATAINICIOCONDICAO'

# alterando a largura das colunas
    ws.column_dimensions['A'].width = 54
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 23
    ws.column_dimensions['G'].width = 21

# salvando as planilhas alteradas
    wb.save(f'{criacao}{nome_empresa}_{data_inicio}_a_{data_fim}.xlsx')
