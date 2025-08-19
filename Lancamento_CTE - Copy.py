#! C:\python\venv\Scripts\python.exe

from clicknium import clicknium, ui, locator
from clicknium.common.enums import *
import subprocess
import os
from datetime import datetime, timedelta
import pandas as pd
import win32com.client as client
import glob
import requests
import json
import sys
import subprocess

autoit = client.Dispatch("AutoItX3.Control")
wshell = client.Dispatch("WScript.Shell")

clicknium.config.set_license('idC+m5GXnIGXlqad0MjQxMLF0N7QoZGal5+TpJeA0MjQwNDe0KGZh9DI0KKXgIGdnJOe0qKAnZSXgYGbnZyTntDe0KSTnpuWk4aXtICdn9DI0MDCwMbfwsPfw8Wmw8vIx8XIwsvcy8rGwMbAxqjQ3tCkk56blpOGl6ad0MjQwMLAx9/Cw9/DxabDy8jHxcjCy9zLysbAxsDKqNDe0LSXk4aHgJeB0MipidC8k5+X0MjQv5OKvp2Rk4adgL6bn5uGl9De0KSTnoeX0MjQw8rGxsTFxsbCxcHFwsvHx8PEw8fQj6+P.Hct2yVB5Jd6T65sm7o5u30CK/v4Zf7AwFQJLI+d4oiKavroJCcf0pgHNVgTxfttyzU6JVbDzWz2WQV4u6dN+m8cX5jf617Xhl+DEPEUnbq969nzHuukp9n2J0UkRulwBUz9CG5DNq/LWDCTyfs7ICXKuCGgcas6X7saLy0EmzRA=')

try:
    
    Login = sys.argv[1]
    Senha = sys.argv[2]
    excel_file_base = sys.argv[3]
    excel_Processamento = sys.argv[4]
    folderXML = str(sys.argv[5])
    token = str(sys.argv[6])

except:

    folder = r"C:\python"
    excel_file_base = r"C:\Studio\Process-Studio-TITAN_UAT\process-studio\ps-workspace\CJ_Maxys_TrocaNotaFob\Base.xlsx"
    excel_Processamento = f"{folder}\Lancamento_CTE.xlsx"
    folderXML = r"C:\Temp\XML"
    pass

def kill_Maxys_process():
    try:
        # Executa o comando 'taskkill' para encerrar processos Python
        result = subprocess.run(['taskkill', '/F', '/IM', 'java.exe'], capture_output=True, text=True)
        print(result.stdout)  # Exibe o resultado no terminal
    except Exception as e:
        print(f"Erro ao tentar encerrar o processo: {e}")

def _Login_MaxysErp(empresa):
    
    # Digita Login
    clicknium.wait_appear(locator.java.maxys.text_usuário)
    ui(locator.java.maxys.text_usuário).send_hotkey("{Tab}")

    # Digita Senha
    ui(locator.java.maxys.password_text_senha).send_hotkey("{Tab}")

    # Digita Empresa
    ui(locator.java.maxys.text_empresa).send_hotkey(empresa)

    # Click Ok
    ui(locator.java.maxys.push_button_ok_alt_o).click()

def _abrirMaxysErp():
    
    # Define the path to the executable
    executable_path = r"C:\Users\svc_rpa\AppData\Local\Maxicon Sistemas\Maxys\Maxys.exe"

    # Define the command-line parameter 
    parameter = " -url http://maxys.cjtrade.com.br:7777/forms/frmservlet?config=teste_app"

    # Create the subprocess command as a list   
    command = [executable_path] + parameter.split()

    # Use subprocess.Popen to run the executable with the parameters
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True,shell=True)
    stdout, stderr = process.communicate()
    return_code = process.returncode

def _configJava():
    
    # Get username
    username = os.getenv("USERNAME")

    # Abrir o arquivo para escrita (modo 'w')
    accessibility_properties = rf'C:\Users\{username}\.accessibility.properties'

    with open(accessibility_properties, 'w') as a_p:
        
        # Escrever o novo conteúdo no arquivo
        a_p.write("assistive_technologies=com.clicknium.ClickniumJavaBridge")

def _Executar(tela):
    
    clicknium.wait_appear(locator.java.maxys.Executar_text)
    ui(locator.java.maxys.Executar_text).clear_text("send-hotkey")
    ui(locator.java.maxys.Executar_text).set_text(tela)
    ui(locator.java.maxys.Executar_text).send_hotkey("{ENTER}")

def _FecharSistema():
    
    # Fechar Telas
    autoit.WinActivate("MAXYS","")
    autoit.ControlSend("MAXYS","","","^{F4}")
    
    # Sair Programar
    if clicknium.is_existing(locator.java.maxys.SairDoPrograma):
        ui(locator.java.maxys.SairDoPrograma).click()

    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
        print(Mensagem)

    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
        print(Mensagem)

    #Fechar Maxicon
    if clicknium.is_existing(locator.java.maxys.FecharMaxicon):
        ui(locator.java.maxys.FecharMaxicon).click()
    if clicknium.is_existing(locator.java.maxys.push_button_Sim_FecharSistema):
        ui(locator.java.maxys.push_button_Sim_FecharSistema).click()

def _fechar_Observacao():
    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
        autoit.sleep(500)
        raise Exception(Mensagem)
    
if __name__ == "__main__":

    empresas = [
        [1, "CJ INTERNATIONAL BRASIL COMERCIAL"],
        [2, "CJ INTERNATIONAL BRASIL - PARANAGUA"],
        [3, "CJ INTERNATIONAL BRASIL - RIO VERDE"],
        [4, "CJ INTERNATIONAL BRASIL - LONDRINA"],
        [5, "CJ INTERNATIONAL BRASIL - PASSO FUNDO"],
        [6, "CJ INTERNATIONAL BRASIL - RIO GRANDE"],
        [7, "CJ INTERNATIONAL BRASIL - SORRISO"],
        [8, "CJ INTERNATIONAL BRASIL - SANTOS"],
        [9, "CJ INTERNATIONAL BRASIL - ARAGUARI"],
        [10, "CJ INTERNATIONAL BRASIL - QUERENCIA"],
        [11, "CJ INTERNATIONAL BRASIL - IMBITUBA"]
    ]

    today = datetime.now()
    yesterday = today - timedelta(days=1)
    yesterday_date = yesterday.strftime('%d/%m/%Y')
    yesterday_date = "01/01/2024"

    # Lista para armazenar dados para criar o DataFrame
    data = []
    
    for empresa in empresas:
        cod_empresa = empresa[0]
        nome_empresa = empresa[1]
        Tipos_ctes = ["normal", "complementar", "subustituição"]

        # Abrir o sistema e realizar login
        _abrirMaxysErp()
        _Login_MaxysErp(cod_empresa)

        for tipo_cte in Tipos_ctes:
            resultado_do_processamento = f"{folder}\\{cod_empresa}_{nome_empresa}_{tipo_cte}_{yesterday_date.replace('/', '')}.xlsx"
            print(resultado_do_processamento)
            try:

                # Executar rotina no sistema
                _Executar("TAF117")

                clicknium.wait_appear(locator.java.maxys_TAF117.text_dt_emissão_inicial)

                # Configurar os filtros
                ui(locator.java.maxys_TAF117.text_empresa).set_text(cod_empresa)
                ui(locator.java.maxys_TAF117.text_dt_emissão_inicial).set_text(yesterday_date)
                ui(locator.java.maxys_TAF117.text_dt_emissão_final).set_text("30/01/2024")

                if tipo_cte == "normal":
                    ui(locator.java.maxys_TAF117.radio_button_ct_e_normal).click()
                elif tipo_cte == "complementar":
                    ui(locator.java.maxys_TAF117.radio_button_ct_e_complementar).click()
                elif tipo_cte == "subustituição":
                    ui(locator.java.maxys_TAF117.radio_button_ct_e_de_substituição).click()

                # Consultar
                ui(locator.java.maxys_TAF117.push_button_f3_consultar).click()

                _fechar_Observacao()

                # Marcar tudo e gravar
                ui(locator.java.maxys_TAF117.push_button_marcar_todos).click()
                ui(locator.java.maxys_TAF117.push_button_gravar).click()

                # Capturar os dados
                clicknium.wait_appear(locator.java.maxys_TAF117.Resultado_do_processamento)
                
                json_data = []
                columns = ["CTE", "Série", "Emissão", "Sucesso", "Resultado do lançamento"]
                i = 0

                while True:
                    row_data = {}
                    for column in columns:
                        variables = {"index": i, "name_column": column}
                        text = ui(locator.java.maxys_TAF117.text_Tabela, variables).get_text().strip()
                        row_data[column] = text

                    if not any(row_data.values()):
                        break
                    json_data.append(row_data)
                    i += 1

                df_tabelaExecucao = pd.DataFrame(json_data)
                df_tabelaExecucao.to_excel(resultado_do_processamento, index=False)

                current_date = datetime.now()
                dataLog = current_date.strftime("%d/%m/%Y %H:%M:%S")

                data.append({
                    "Cod. Empresa": cod_empresa,
                    "Empresa": nome_empresa,
                    "Tipo CTE": tipo_cte,
                    "Mensagem_Saida": "Sucesso!",
                    "DataLog_Saida": dataLog
                })

                if tipo_cte == "subustituição":
                    _FecharSistema()
                else:
                    ui(locator.java.maxys_TAF117.push_button_voltar).click()
                    ui(locator.java.maxys_TAF117.push_button_sair_programa).click()

                print("Sucesso!")

            except BaseException as e:
                print(str(e))
                current_date = datetime.now()
                dataLog = current_date.strftime("%d/%m/%Y %H:%M:%S")

                data.append({
                    "Cod. Empresa": cod_empresa,
                    "Empresa": nome_empresa,
                    "Tipo CTE": tipo_cte,
                    "Mensagem_Saida": f"Erro: {str(e)}",
                    "DataLog_Saida": dataLog
                })

                ui(locator.java.maxys_TAF117.push_button_sair_programa).click()

        _FecharSistema()

    # Criar DataFrame final e salvar no Excel
    df = pd.DataFrame(data)
    df.to_excel(excel_Processamento, index=False)