import os
import time
import datetime
import pyautogui
import pandas as pd

# Função para baixar o anexo do Outlook
def baixar_anexo_outlook():
    # Minimizar todas as janelas
    pyautogui.hotkey('win', 'd')
    
    # Abrir Outlook
    pyautogui.hotkey('win', 's')
    time.sleep(3)
    pyautogui.write("outlook")
    pyautogui.press("enter")
    time.sleep(30)
    
    # Maximizar a janela do Outlook
    pyautogui.hotkey('win', 'up')
    
    # Definir a data atual no formato dia/mês/ano
    data_atual = datetime.datetime.now().strftime("%d.%m.%Y")
    pesquisa = f"anderson.furlan@panpharma.com.br {data_atual}"
    
    # Clicar na caixa de pesquisa e buscar o e-mail do remetente com a data atual
    pyautogui.hotkey('Ctrl', 'e')
    time.sleep(3)
    pyautogui.write(pesquisa)
    pyautogui.press("enter")
    time.sleep(3)
    
    # Selecionar o primeiro e-mail da pesquisa
    pyautogui.press("down")  # Pressionar 'down' para abrir o e-mail encontrado
    pyautogui.hotkey('Alt', 'h', 'a')  # Ajuste este atalho se necessário
    pyautogui.press('x')
    time.sleep(5)  # Pausa para carregar a visualização do anexo
    pyautogui.press('Enter')
    time.sleep(3)
    
    # Salvar o anexo na pasta de destino
    pyautogui.hotkey('Ctrl', 'l')
    pyautogui.write("C:\\Users\\AS\\Documents\\Arquivo estoque")
    pyautogui.press("enter")
    time.sleep(1)
    for _ in range(10):
        pyautogui.press('tab')
    pyautogui.press("enter")
    time.sleep(1)
    
    # Fechar o Outlook
    pyautogui.hotkey('Alt', 'F4')
    print("Anexo baixado com sucesso.")

# Função para selecionar e abrir o arquivo Excel baixado
def selecionar_abrir_arquivo(pasta_destino):
    data_atual = datetime.datetime.now().strftime("%d.%m")
    nome_arquivo = f"ESTOQUE NACIONAL {data_atual}.xlsb"
    caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)

    if not os.path.exists(caminho_arquivo):
        print(f"Arquivo {nome_arquivo} não encontrado na pasta {pasta_destino}")
        return None

    print(f"Abrindo o arquivo {nome_arquivo} no Excel...")
    os.startfile(caminho_arquivo)
    time.sleep(5)
    return caminho_arquivo

# Função para manipular o arquivo Excel
def manipular_arquivo_estoque(pasta_destino):
    data_atual = datetime.datetime.now().strftime("%d.%m")
    nome_arquivo = f"ESTOQUE NACIONAL {data_atual}.xlsb"
    caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)

    if not os.path.exists(caminho_arquivo):
        print(f"Arquivo {nome_arquivo} não encontrado na pasta {pasta_destino}")
        return

    df = pd.read_excel(caminho_arquivo, engine='pyxlsb')
    colunas_a_manter = [
        'CODIGO', 'CODIGO EAN', 'FORNECEDOR', 'DESCRICAO',
        'CATEGORIA', 'CXA EMB', 'ESTOQUE RN', 'FABRICA RN',
        'ESTOQUE PB', 'FABRICA PB'
    ]
    df_filtrado = df[colunas_a_manter]

    caminho_arquivo_novo = caminho_arquivo.replace(".xlsb", ".xlsx")
    df_filtrado.to_excel(caminho_arquivo_novo, index=False)
    print(f"Arquivo manipulado e salvo com sucesso em {caminho_arquivo_novo}")

# Função para filtrar por fornecedor e salvar o arquivo
def filtrar_fornecedor_e_salvar(pasta_destino, fornecedores, nome_mapa):
    data_atual = datetime.datetime.now().strftime("%d.%m")
    nome_arquivo = f"ESTOQUE NACIONAL {data_atual}.xlsx"
    caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)

    if not os.path.exists(caminho_arquivo):
        print(f"Arquivo {nome_arquivo} não encontrado na pasta {pasta_destino}")
        return

    df = pd.read_excel(caminho_arquivo, engine='openpyxl')
    fornecedores_regex = '|'.join(fornecedores)
    df_filtrado = df[df['FORNECEDOR'].str.contains(fornecedores_regex, case=False, na=False)]

    caminho_arquivo_mapa = os.path.join(pasta_destino, f"{nome_mapa} {data_atual}.xlsx")
    df_filtrado.to_excel(caminho_arquivo_mapa, index=False)
    print(f"Arquivo '{nome_mapa}' manipulado e salvo com sucesso em {caminho_arquivo_mapa}")

# Função para enviar arquivos para grupos do WhatsApp Desktop com abertura do ícone de clipe
def enviar_arquivo_whatsapp(grupo, caminho_arquivo):
    # Abrir o WhatsApp Desktop
    pyautogui.press("win")
    time.sleep(2)
    pyautogui.write("whatsapp")
    pyautogui.press("enter")
    time.sleep(10)

    # Procurar pelo grupo e abrir a conversa
    pyautogui.write(grupo)
    time.sleep(2)
    pyautogui.press('Tab')
    pyautogui.press('Enter')

    # Clicar no ícone de clipe para abrir o menu de anexos
    pyautogui.click(x=497, y=689)  # Ajuste conforme a posição do ícone de clipe no WhatsApp
    pyautogui.press("enter")
    time.sleep(2)

    # Selecionar a opção de enviar documentos
    pyautogui.press('Tab')
    pyautogui.press('Down')  # Ajuste conforme a posição da opção "Documento"
    pyautogui.press('Down')
    time.sleep(3)
    pyautogui.press("enter")
    time.sleep(2)


    # Digitar o caminho do arquivo e enviar
    pyautogui.write(caminho_arquivo)
    time.sleep(1)
    pyautogui.press("enter")  # Confirmar a seleção do arquivo
    time.sleep(2)

    # Confirmar o envio
    pyautogui.press("enter")
    time.sleep(5)
    pyautogui.hotkey('Alt', 'F4')
    
# Função para enviar múltiplos arquivos para múltiplos grupos
def enviar_para_grupos(grupos_por_arquivo, pasta_destino):
    data_atual = datetime.datetime.now().strftime("%d.%m")
    for arquivo_base, grupos in grupos_por_arquivo.items():
        arquivo_com_data = f"{arquivo_base} {data_atual}.xlsx"  # Adiciona a data no nome do arquivo
        caminho_arquivo = os.path.join(pasta_destino, arquivo_com_data)
        
        if os.path.exists(caminho_arquivo):
            for grupo in grupos:
                print(f"Enviando {arquivo_com_data} para o grupo {grupo}...")
                enviar_arquivo_whatsapp(grupo, caminho_arquivo)
        else:
            print(f"Arquivo {arquivo_com_data} não encontrado.")

# Função principal para executar todo o processo
def executar_processos():
    pasta_destino = "C:\\Users\\AS\\Documents\\Arquivo estoque"
    
    baixar_anexo_outlook()
    
    caminho_arquivo = selecionar_abrir_arquivo(pasta_destino)
    
    if caminho_arquivo:
        manipular_arquivo_estoque(pasta_destino)

        fornecedores_mapas = [
            (["ACHE LABORATORIOS FARMACEUTICO", "LABOFARMA PROD FARMACEUTICOS L", "LABOFARMA PRODUTOS FARMACEUTIC"], "Mapa Ache"),
            (["EUROFARMA LABORATORIOS S A"], "Mapa EUROFARMA"),
            (["SANDOZ DO BRASIL IND.FARM.", "SANDOZ DO BRASIL IND FARMACEUT"], "Mapa SANDOZ"),
            (["BIOLAB SANUS FARM LTDA", "BIOLAB FARMA GENERICOS LTDA", "BIOLAB SANUS FARMACEUTICA LTDA"], "Mapa BIOLAB"),
            (["RANBAXY FARMACEUTICA LTDA"], "Mapa RANBAXY"),
            (["SANOFI MEDLEY FARMACEUTICA LTD", "MEDLEY FARMACEUTICA LTDA"], "Mapa MEDLEY"),
            (["MYRALIS INDUSTRIA FARMACEUTICA"], "Mapa MYRALIS"),
            (["ALTHAIA S.A. INDUSTRIA FARMACE"], "Mapa ALTHAIA"),
            (["QUESALON DIST DE PROD FARM LTD"], "Mapa QUESALON"),
            (["SUPERA RX MEDICAMENTOS LTDA"], "Mapa SUPERA")

        ]
        
        for fornecedores, nome_mapa in fornecedores_mapas:
            filtrar_fornecedor_e_salvar(pasta_destino, fornecedores, nome_mapa)

        grupos_por_arquivo = {
            "Mapa Ache": ["Ache Genericos - PB/RN", "Santa Cruz 🤝 Achê"],
            "Mapa EUROFARMA": ["Eurofarma / Santa Cruz PB", "Santa Cruz 🤜🤛 Eurofarma GEN - RN", "Santa Cruz RN & Momenta", "MOMENTA-SC / PB"],
            "Mapa SANDOZ": ["Santa Cruz RN 🤝🏼 Sandoz"],
            "Mapa BIOLAB": ["Biolab Gen X Santa Cruz - PB","Biolab Gen X Sta Cruz RN"],
            "Mapa RANBAXY": ["Equipe Ranbaxy PB/RN"],
            "Mapa MEDLEY": ["Santa Cruz PB X Medley", "Medley & Santa Cruz RN"],
            "Mapa MYRALIS": ["MYRALIS X SANTA CRUZ PB"],
            "Mapa ALTHAIA": ["Althaia & SC PB.RN"],
            "Mapa QUESALON": ["Hebron🤝Santa Cruz RN"],
            "Mapa SUPERA": ["Santa Cruz 🤝 Supera"],
            # Adicione mais arquivos e grupos conforme necessário
        }
        
        enviar_para_grupos(grupos_por_arquivo, pasta_destino)

# Executa o processo completo
if __name__ == "__main__":              
    executar_processos()
