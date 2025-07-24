from openpyxl import load_workbook
from dotenv import load_dotenv
import os
import requests 

load_dotenv()

#url da api
#url = ""

# Carrega o arquivo 
wb = load_workbook('./base envio.xlsx')

# Seleciona a planilha ativa
planilha = wb.active

#Id do departamento
departmentIds = os.getenv("DEPARTMENT_ID")
my_array = departmentIds.split(',')

for departmentId in my_array:
    url = "https://api.zapresponder.com.br/api/whatsapp/message/" + departmentId
    print(url)
    #valor do range
    qtd = int(os.getenv("RANGE"))
    # Nome do template
    template_name = os.getenv("TEMPLATE_NAME")
    print(template_name)

    #loop para enviar mensagens
    for i in range(qtd):
        # Lê o valor da célula
        valor_celula_telefone = planilha['A2'].value
        telefone ="55"+ str(valor_celula_telefone)

        payload = {
            "type": "template",
            "template_name": template_name,
            "number": telefone,
            "language": "pt_BR"
        }

        headers = {
            "authorization": "Bearer "+ os.getenv("TOKEN"),
            "accept": "application/json",
            "content-type": "application/json"
        }

        try:
            response = requests.post(url, json=payload, headers=headers)

            response.raise_for_status()  # Lança exceção para códigos de erro

            if response.status_code == 200:
                # A requisição foi bem-sucedida, faça algo aqui
                print("Requisição bem-sucedida!")
                print(f"Conteúdo da resposta: {response.text}")  # Ou response.json() se for JSON
                
                # Define o índice da linha a ser excluída
                row_index_to_delete = 2  # Excluir a linha 2

                # Define o número de linhas a serem excluídas
                num_rows_to_delete = 1  # Excluir apenas uma linha

                # Remove a linha da planilha
                planilha.delete_rows(row_index_to_delete, num_rows_to_delete)

                # Salva as alterações
                wb.save('./base envio.xlsx')
            else:
                # Tratamento de outros códigos de status
                print(f"Erro na requisição. Código de status: {response.status_code}")

        except requests.exceptions.RequestException as e:
            print(f"Ocorreu um erro durante a requisição: {e}")