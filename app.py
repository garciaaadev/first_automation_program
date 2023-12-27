# Ler dados da Planilha
# Inserir cada célula da planilha em um campo do sistema

import openpyxl
import pyautogui

# A leitura dos dados foi armazenada em uma variavel para que fosse possível Iterar sobre os dados na estrutura FOR
workbook = openpyxl.load_workbook('dados_teste.xlsx')
cadastro_sheet = workbook['Sheet1']


for linha in cadastro_sheet.iter_rows(min_row=2):
    # ["Lucas", "Belo Horizonte", "17/08/1983", "Ensino Medio Completo"]
    # Será acessado de forma semelhante a listas
    # Instalar ferramenta mouseinfo para mapear a tela
    # Basta posicionar o mouse no local correto coletar as coordenadas com a tecla F6 e especificar o tempo que será gasto até chegar ao campo de preenchimento
    pyautogui.click(1369,368,duration=1.5)
    pyautogui.click(1369,368,duration=1.5)
    pyautogui.click(1369,368,duration=1.5)
    # O pyautogui não faz inserção direta de numeros é preciso converter para string
    pyautogui.write(str(linha[2].value))
    pyautogui.click(1298,487,duration=1.5)
    pyautogui.write(linha[3].value)
