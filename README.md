Este projeto apresenta um código Python desenvolvido para automatizar o processo de consulta e atualização dos dados de funcionários de uma empresa a partir de uma planilha de Excel. O código utiliza a biblioteca Selenium para interagir com páginas da web e coletar informações relevantes, visando simplificar e agilizar a atualização de registros no sistema eSocial.

O que é o eSocial?

O eSocial é um sistema do governo brasileiro que unifica a prestação de informações trabalhistas, previdenciárias e tributárias relacionadas aos trabalhadores. Ele tem o objetivo de simplificar e padronizar a transmissão de dados, garantindo maior eficiência na gestão de obrigações fiscais e reduzindo a burocracia para empresas e empregadores.

Motivação e Desenvolvimento do Código

Este projeto foi concebido para resolver um desafio específico: a dificuldade em realizar consultas e atualizações dos dados dos funcionários de maneira eficiente através de APIs convencionais. Dado que a plataforma eSocial oferece uma interface web para acessar e modificar informações de colaboradores, desenvolvemos um código Python automatizado para interagir com a plataforma e agilizar esse processo.

Funcionamento do Código

O código requer uma planilha de Excel preenchida apenas com os números de contrato de cada colaborador, com cada número em uma linha. Além disso, o código foi projetado para acessar o eSocial por meio do site gov.br, utilizando um certificado digital da empresa instalado na máquina. É necessário também ativar a seleção automática de certificado digital nas configurações do navegador.

O processo automatizado inclui os seguintes passos:

Autenticação com Certificado Digital: O código realiza o login na plataforma eSocial por meio do site gov.br, utilizando o certificado digital da empresa para autenticação.

Acesso aos Dados: O programa navega até a seção de consulta e atualização dos dados dos funcionários.

Iteração pela Planilha: O código lê os números de contrato da planilha de Excel e, para cada número, acessa a página correspondente do colaborador na plataforma.

Extração e Atualização: Utilizando recursos de automação do Selenium, o código extrai informações como CPF, nome, cargo, salário, entre outros, da página do colaborador. Esses dados são então inseridos na planilha original, atualizando as informações de cada funcionário.

Salvar Alterações: Após todas as consultas e atualizações serem concluídas, o código salva as modificações na planilha de Excel.

Benefícios do Projeto

Ao automatizar o processo de consulta e atualização de dados dos funcionários, este projeto proporciona economia de tempo e redução de esforços repetitivos. Com a simples inserção dos números de contrato na planilha e o uso do certificado digital da empresa, o usuário é capaz de atualizar com precisão os registros de cada colaborador no sistema eSocial de maneira eficiente e segura.

Em resumo, este código Python com Selenium oferece uma solução prática e automatizada para superar as limitações de consulta e atualização de dados no sistema eSocial, utilizando o certificado digital da empresa como meio de autenticação. Isso contribui para uma gestão mais eficiente e assertiva das informações dos funcionários de uma organização, garantindo conformidade com as obrigações legais.


----------------- ENGLISH --------------------

This project presents a Python code developed to automate the process of querying and updating employee data from a company using an Excel spreadsheet. The code utilizes the Selenium library to interact with web pages and gather relevant information, aiming to simplify and expedite the record update process within the eSocial system.

What is eSocial?

eSocial is a Brazilian government system that consolidates the reporting of labor, social security, and tax information related to employees. Its objective is to simplify and standardize the transmission of data, ensuring greater efficiency in managing fiscal obligations and reducing bureaucracy for companies and employers.

Motivation and Code Development

This project was conceived to address a specific challenge: the difficulty of efficiently performing queries and updates of employee data through conventional APIs. Given that the eSocial platform offers a web interface for accessing and modifying employee information, we developed an automated Python code to interact with the platform and streamline this process.

Code Operation

The code requires an Excel spreadsheet filled only with the contract numbers of each employee, with each number on a separate line. Additionally, the code is designed to access eSocial through the gov.br website, utilizing a company's digital certificate installed on the machine. It is also necessary to enable the automatic selection of the digital certificate in the browser settings.

The automated process includes the following steps:

Digital Certificate Authentication: The code logs into the eSocial platform via the gov.br website, using the company's digital certificate for authentication.

Accessing Data: The program navigates to the section for querying and updating employee data.

Iteration through Spreadsheet: The code reads the contract numbers from the Excel spreadsheet and, for each number, accesses the corresponding employee's page on the platform.

Extraction and Update: Using Selenium's automation features, the code extracts information such as CPF (taxpayer identification number), name, position, salary, and more, from the employee's page. These data are then inserted into the original spreadsheet, updating the information for each employee.

Saving Changes: Once all queries and updates are completed, the code saves the modifications to the Excel spreadsheet.

Project Benefits

By automating the process of querying and updating employee data, this project provides time savings and reduces repetitive efforts. With the simple input of contract numbers into the spreadsheet and the use of the company's digital certificate, users can accurately update each employee's records within the eSocial system efficiently and securely.

In summary, this Python code with Selenium offers a practical and automated solution to overcome the limitations of querying and updating data within the eSocial system, utilizing the company's digital certificate as an authentication method. This contributes to more efficient and accurate management of employee information, ensuring compliance with legal obligations.
