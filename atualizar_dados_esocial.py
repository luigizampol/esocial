# Importar as bibliotecas necessárias
# Import the necessary libraries
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from IPython.display import display
import os
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException

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


# Aguardar o botão de login via certificado digital carregar
# Waiting for the digital certificate login button to load.
login_button_certificate = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'button[tabindex="4"]'))
)

# Clicar no botão de login via certificado digital
# Click the digital certificate login button
login_button_certificate.click()
wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))


# Abre a planilha de colaboradores
# Open the collaborators spreadsheet
employees = pd.read_excel("colaboradores.xlsx") # Inserir o nome do arquivo da planilha aqui. A planilha deve estar na mesma pasta do script.

# Cria um data frame com os dados da planilha
# Create a DataFrame with spreadsheet data
employees_df = pd.DataFrame(employees)

# Cria uma lista com todos os números de contrato
# Create a list of all contract numbers
employees_contract_numbes = employees_df['contrato'].tolist()

# Url base de acesso à pagina de informações contratuais
# Base URL for accessing contractual information page

url_base = "https://www.esocial.gov.br/portal/Trabalhador/ConsultaCompleto?tsv=False&idContratoSelecionado="

# Url que será a soma da URL principal com o número do contrato de cada funcionário
# URL that combines the base URL with each employee's contract number
url_with_contract_number = ""

# List of URLs that the program will access and gather data from
url_list = list()

# Para cada número de contrato, cria o url completo e salva na lista
# For each contract number, create the complete URL and save it to the list
for contract_number in employees_contract_numbes:
    url_with_contract_number = url_base + str(contract_number)
    url_list.append(url_with_contract_number)



# Dicionário onde as informações serão armazenadas
# Dictionary where the information will be stored
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

# Define uma função para extrair dados com valor de um elemento pelo caminho XPATH
# Define a function to extract data with value from an element by XPATH
def extract_data(dado, id):
    try:
        elemento = driver.find_element(By.XPATH, f"//input[@id='{id}']")
        infos[dado].append(elemento.get_attribute("value"))
    except NoSuchElementException:
        elemento = driver.find_elements(By.CSS_SELECTOR, f"#{id} option[selected='selected']")
        if len(elemento) > 0:
            valor = elemento[0].text
            infos[dado].append((driver.find_element(By.ID, f"{id}").find_element(By.CSS_SELECTOR, "option[selected='selected']")).text)
        else:
            infos[dado].append(None)

contador = 0

# Laço para extrair informações de cada url da variável url_list
# Loop to extract information from each URL in the url_list variable
for url in url_list:
    # Navega até a url
    # Navigate to the URL
    driver.get(url) 

    # Verifica se o funcionário está ativo e adiciona a informação ao dicionário
    # Check if the employee is active and add the information to the dictionary
    ativo = driver.find_elements(By.CSS_SELECTOR, 'td.center[data-col="Situação"]')
    infos['situacao'].append(ativo[0].text)
    
    # Aguarda até que o botão para acessar as outras informações esteja disponível e clica nele
    # Wait until the button to access other information is available and click on it
    dados_botao = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.item-menu-interno.dados-contratuais')))
    dados_botao.click()

    # Funções que extraem os dados
    # Functions that extract the data
    data_to_extract = [
    ('contrato', 'idContrato'),
    ('cpf', 'cpfTrabalhador'),
    ('nome', 'nomeTrabalhador'),
    ('tipo_registro', 'TipoRegistroAdmissao'),
    ('matricula', 'InfoVinculo_Matricula'),
    ('regime_trab', 'InfoVinculo_TipoRegimeTrabalhista'),
    ('categoria', 'InfoVinculo_InfoContrato_CodigoCategoria'),
    ('regime_prev', 'InfoVinculo_TipoRegimePrevidenciario'),
    ('nome_cargo', 'InfoVinculo_InfoContrato_NomeCargo'),
    ('cbo_cargo', 'InfoVinculo_InfoContrato_CBOCargo'),
    ('nome_funcao', 'InfoVinculo_InfoContrato_NomeFuncao'),
    ('cbo_funcao', 'InfoVinculo_InfoContrato_CBOFuncao'),
    ('un_pagamento', 'InfoVinculo_InfoContrato_Remuneracao_UnidadeSalarioFixo'),
    ('salario_base', 'InfoVinculo_InfoContrato_Remuneracao_ValorSalarioFixo'),
    ('desc_salario_variavel', 'InfoVinculo_InfoContrato_Remuneracao_DescricaoSalarioVariavel'),
    ('tipo_contrato_trabalho', 'InfoVinculo_InfoContrato_Duracao_TipoContrato'),
    ('horas_semanais', 'InfoVinculo_InfoContrato_HorarioContratual_QuantidadeHorasSemanal'),
    ('tipo_jornada', 'InfoVinculo_InfoContrato_HorarioContratual_TipoJornada'),
    ('tempo_parcial', 'InfoVinculo_InfoContrato_HorarioContratual_JornadaTempoParcial'),
    ('desc_jornada', 'InfoVinculo_InfoContrato_HorarioContratual_DescricaoJornada'),
    ('data_adm', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_DataAdmissao'),
    ('tipo_adm', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_TipoAdmissao'),
    ('indicativo_adm', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_IndicativoAdmissao'),
    ('regime_jornada', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_TipoRegimeJornada'),
    ('natureza_atividade', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_NaturezaAtividade'),
    ('mes_database', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_DataBase'),
    ('cnpj_sindicato', 'InfoVinculo_InformacoesRegimeTrabalhista_InformacoesTrabalhadorCeletista_CnpjSindicatoCategoriaProfissional')]

    for dado, id in data_to_extract:
        extract_data(dado, id)

    # Contador para acompanhar o progresso do script
    # Counter to track the progress of the script
    contador += 1
    print(f'Extraindo dados... [{contador}/{len(url_list)}]')

# Transforma o dicionário em data frame
# Convert the dictionary into a data frame
infos_df = pd.DataFrame(infos)
display(infos_df)


# Converte os dataframes em int64
# Convert the data frames to int64
employees_df['contrato'] = employees_df['contrato'].astype('int64')
infos_df['contrato'] = infos_df['contrato'].astype('int64')


# Compara as informações da planilha e atualiza somente com os dados que foram atualizados
# Compare the information from the spreadsheet and update only with the data that has been updated
if 'contrato' in employees_df.columns and 'contrato' in infos_df.columns:
    # Atualizando as colunas em comum
    # Updating the common columns
    colunas_em_comum = [
    'contrato', 'situacao', 'cpf', 'nome', 'tipo_registro', 'matricula', 'regime_trab', 'categoria',
    'regime_prev', 'nome_cargo', 'cbo_cargo', 'nome_funcao', 'cbo_funcao', 'un_pagamento',
    'salario_base', 'desc_salario_variavel', 'tipo_contrato_trabalho', 'horas_semanais',
    'tipo_jornada', 'tempo_parcial', 'desc_jornada', 'data_adm', 'tipo_adm', 'indicativo_adm',
    'regime_jornada', 'natureza_atividade', 'mes_database', 'cnpj_sindicato'
    ]

    employees_df.update(infos_df[colunas_em_comum])


###########################

# Transforma a planilha antiga em arquivo de backup
# Substitua pelo nome do seu arquivo Excel
# Transform the old spreadsheet into a backup file
# Replace with the name of your Excel file

old_file_name = "colaboradores.xlsx"

# Obter a data atual no formato dia_mes_ano
# Get the current date in the format day_month_year
current_date = datetime.now().strftime("%d_%m_%Y")

# Criar o novo nome do arquivo no formato "colaboradores_backup_dia_mes_ano.xlsx"
# Create the new file name in the format "employees_backup_day_month_year.xlsx"

new_file_name = f"colaboradores_backup_{current_date}.xlsx"

# Renomear o arquivo
# Rename the file
try:
    os.rename(old_file_name, new_file_name)
    print(f"O arquivo {old_file_name} foi renomeado para {new_file_name} com sucesso!") # File {file name} successfully renamed.
except FileExistsError:
    print(f"Já existe um arquivo com o nome {new_file_name}. Por favor, escolha outro nome.") # There is already a file named {new_file_name}. Please choose another name.
except FileNotFoundError:
    print(f"O arquivo {old_file_name} não foi encontrado.") # The file {old_file_name} was not found

# Cria um objeto ExcelWriter usando a biblioteca openpyxl
# Create an ExcelWriter object using the openpyxl library
writer = pd.ExcelWriter('colaboradores.xlsx', engine='openpyxl')

# Converte o DataFrame em uma planilha do Excel e adiciona ao objeto ExcelWriter
# Converts the DataFrame into an Excel sheet and adds it to the ExcelWriter object
employees_df.to_excel(writer, sheet_name='Funcionarios', index=False)

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