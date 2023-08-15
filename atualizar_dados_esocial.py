# Importar as bibliotecas necessárias
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from IPython.display import display
import os
from datetime import datetime



# Definir as opções do Microsoft Edge
# Define Microsoft Edge options
edge_options = webdriver.EdgeOptions()
edge_options.add_argument("start-maximized")
edge_options.use_chromium = True

# Criar uma instância do driver do Microsoft Edge
# Create an instance of the Microsoft Edge driver
driver = webdriver.Edge(options=edge_options)

# Navegar para a página do eSocial
# Navigate to the eSocial page
driver.get('https://login.esocial.gov.br/login.aspx')

# Esperar até que o botão de login seja carregado
# Wait until the login button is loaded
login_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, '.br-button.sign-in.large'))
)


# Clicar no botão de login
# Click the login button
login_button.click()


# Esperar até que o link seja carregado
# Wait until the link is loaded
certificado_link = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'button[tabindex="4"]'))
)

# Clicar no botão de login
# Click the login button
certificado_link.click()
wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))


# Abre a planilha de colaboradores
# Open the collaborators spreadsheet
funcionarios = pd.read_excel("colaboradores.xlsx")

# Cria um data frame com os dados da planilha
# Create a DataFrame with spreadsheet data
funcionarios_df = pd.DataFrame(funcionarios)

# Cria uma lista com todos os números de contrato
# Create a list of all contract numbers
lista_contratos = funcionarios_df['contrato'].tolist()

# Url base de acesso à pagina de informações contratuais
# Base URL for accessing contractual information page

url_base = "https://www.esocial.gov.br/portal/Trabalhador/ConsultaCompleto?tsv=False&idContratoSelecionado="

# Url que será a soma da URL principal com o número do contrato de cada funcionário
# URL that combines the base URL with each employee's contract number
url_completa = ""

# List of URLs that the program will access and gather data from
lista_urls = list()

# Para cada número de contrato, cria o url completo e salva na lista
# For each contract number, create the complete URL and save it to the list
for contrato in lista_contratos:
    url_completa = url_base + str(contrato)
    lista_urls.append(url_completa)

infos = {'contrato': [],
              'situacao': [],
              'cpf': [],
              'nome': [],
              'tipo_registro': [],
              'matricula': [],
              'regime_trab': [],
              'categoria': [],
              'regime_prev': [],
              'nome_cargo': [],
              'cbo_cargo': [],
              'nome_funcao': [],
              'cbo_funcao': [],
              'un_pagamento': [],
              'salario_base': [],
              'desc_salario_variavel': [],
              'tipo_contrato_trabalho': [],
              'horas_semanais': [],
              'tipo_jornada': [],
              'tempo_parcial': [],
              'desc_jornada': [],
              'data_adm': [],
              'tipo_adm': [],
              'indicativo_adm': [],
              'regime_jornada': [],
              'natureza_atividade': [],
              'mes_database': [],
              'cnpj_sindicato': [],
              }

# Define uma função para extrair dados com valor de um elemento
# Define a function to extract data with value from an element
def extrair_dados_valor(dado, id):
    infos[dado].append(driver.find_element(By.XPATH, f"//input[@id='{id}']").get_attribute("value"))

# Define uma função para extrair dados com a opção selecionada de um elemento de menu suspenso
# Define a function to extract data with selected option from a dropdown element
def extrair_dados_option(dado, id):
    elemento = driver.find_elements(By.CSS_SELECTOR, f"#{id} option[selected='selected']")
    if len(elemento) > 0:
        valor = elemento[0].text
        infos[dado].append((driver.find_element(By.ID, f"{id}").find_element(By.CSS_SELECTOR, "option[selected='selected']")).text)
    else:
        infos[dado].append(None)

contador = 0

for url in lista_urls:
    driver.get(url)
    ativo = driver.find_elements(By.CSS_SELECTOR, 'td.center[data-col="Situação"]')
    infos['situacao'].append(ativo[0].text)
    dados_botao = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.item-menu-interno.dados-contratuais')))
    dados_botao.click()
    extrair_dados_valor('contrato', 'idContrato')
    extrair_dados_valor('cpf', 'cpfTrabalhador')
    extrair_dados_valor('nome', 'nomeTrabalhador')
    extrair_dados_valor('tipo_registro', 'TipoRegistroAdmissao')
    extrair_dados_valor('matricula', 'InfoVinculo_Matricula')
    extrair_dados_option('regime_trab', 'InfoVinculo_TipoRegimeTrabalhista')
    extrair_dados_option('categoria', 'InfoVinculo_InfoContrato_CodigoCategoria')
    extrair_dados_option('regime_prev', 'InfoVinculo_TipoRegimePrevidenciario')
    extrair_dados_valor('nome_cargo', 'InfoVinculo_InfoContrato_NomeCargo')
    extrair_dados_option('cbo_cargo', 'InfoVinculo_InfoContrato_CBOCargo')
    extrair_dados_valor('nome_funcao', 'InfoVinculo_InfoContrato_NomeFuncao')
    extrair_dados_option('cbo_funcao', 'InfoVinculo_InfoContrato_CBOFuncao')
    extrair_dados_option('un_pagamento', 'InfoVinculo_InfoContrato_Remuneracao_UnidadeSalarioFixo')
    extrair_dados_valor('salario_base', 'InfoVinculo_InfoContrato_Remuneracao_ValorSalarioFixo')
    extrair_dados_valor('desc_salario_variavel', 'InfoVinculo_InfoContrato_Remuneracao_DescricaoSalarioVariavel')
    extrair_dados_option('tipo_contrato_trabalho', 'InfoVinculo_InfoContrato_Duracao_TipoContrato')
    extrair_dados_valor('horas_semanais', 'InfoVinculo_InfoContrato_HorarioContratual_QuantidadeHorasSemanal')
    extrair_dados_option('tipo_jornada', 'InfoVinculo_InfoContrato_HorarioContratual_TipoJornada')
    extrair_dados_option('tempo_parcial', 'InfoVinculo_InfoContrato_HorarioContratual_JornadaTempoParcial')
    extrair_dados_valor('desc_jornada', 'InfoVinculo_InfoContrato_HorarioContratual_DescricaoJornada')
    extrair_dados_valor('data_adm', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_DataAdmissao')
    extrair_dados_option('tipo_adm', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_TipoAdmissao')
    extrair_dados_option('indicativo_adm', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_IndicativoAdmissao')
    extrair_dados_option('regime_jornada', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_TipoRegimeJornada')
    extrair_dados_option('natureza_atividade', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_NaturezaAtividade')
    extrair_dados_option('mes_database', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_DataBase')
    extrair_dados_valor('cnpj_sindicato', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_CnpjSindicatoCategoriaProfissional')
    contador += 1
    print(f'Extraindo dados... [{contador}/{len(lista_urls)}]')
    
infos_df = pd.DataFrame(infos)
display(infos_df)

#########################

# Dataframes
df1 = funcionarios_df
df2 = infos_df
df1['contrato'] = df1['contrato'].astype('int64')
df2['contrato'] = df2['contrato'].astype('int64')



if 'contrato' in df1.columns and 'contrato' in df2.columns:
    # Atualizando as colunas em comum
    # Updating the common columns
    colunas_em_comum = [
    'contrato', 'situacao', 'cpf', 'nome', 'tipo_registro', 'matricula', 'regime_trab', 'categoria',
    'regime_prev', 'nome_cargo', 'cbo_cargo', 'nome_funcao', 'cbo_funcao', 'un_pagamento',
    'salario_base', 'desc_salario_variavel', 'tipo_contrato_trabalho', 'horas_semanais',
    'tipo_jornada', 'tempo_parcial', 'desc_jornada', 'data_adm', 'tipo_adm', 'indicativo_adm',
    'regime_jornada', 'natureza_atividade', 'mes_database', 'cnpj_sindicato'
    ]

    df1.update(df2[colunas_em_comum])


###########################

# Transforma a planilha antiga em arquivo de backup
# Substitua pelo nome do seu arquivo Excel
# Transform the old spreadsheet into a backup file
# Replace with the name of your Excel file

nome_arquivo_antigo = "colaboradores.xlsx"

# Obter a data atual no formato dia_mes_ano
# Get the current date in the format day_month_year
data_atual = datetime.now().strftime("%d_%m_%Y")

# Criar o novo nome do arquivo no formato "colaboradores_backup_dia_mes_ano.xlsx"
# Create the new file name in the format "colaboradores_backup_day_month_year.xlsx"

novo_nome_arquivo = f"colaboradores_backup_{data_atual}.xlsx"

# Renomear o arquivo
# Rename the file
try:
    os.rename(nome_arquivo_antigo, novo_nome_arquivo)
    print(f"O arquivo {nome_arquivo_antigo} foi renomeado para {novo_nome_arquivo} com sucesso!")
except FileExistsError:
    print(f"Já existe um arquivo com o nome {novo_nome_arquivo}. Por favor, escolha outro nome.")
except FileNotFoundError:
    print(f"O arquivo {nome_arquivo_antigo} não foi encontrado.")

# Cria um objeto ExcelWriter usando a biblioteca openpyxl
# Create an ExcelWriter object using the openpyxl library
writer = pd.ExcelWriter('colaboradores.xlsx', engine='openpyxl')

# Converte o DataFrame em uma planilha do Excel e adiciona ao objeto ExcelWriter
# Converts the DataFrame into an Excel sheet and adds it to the ExcelWriter object
df1.to_excel(writer, sheet_name='Funcionarios', index=False)

# Salva o arquivo Excel
# Save the file
writer.close()


# Aguardar entrada do usuário antes de fechar o navegador
# Wait for user input before closing the browser

input("Pressione Enter para sair...")
print("Programa concluído!")

# Fechar o navegador
# Close the browser
driver.quit()