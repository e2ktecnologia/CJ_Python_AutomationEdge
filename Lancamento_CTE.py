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
import time
from dotenv import load_dotenv
import fnmatch

# Carregar variáveis do arquivo .env
load_dotenv()

# Capturar a licença do Clicknium do .env
clicknium_license = os.getenv("CLICKNIUM_LICENSE")
autoit = client.Dispatch("AutoItX3.Control")
wshell = client.Dispatch("WScript.Shell")

clicknium.config.set_license(clicknium_license)

try:
    
    folder = sys.argv[1]
    excel_Processamento =  sys.argv[2]
    EnvMaxyCon = sys.argv[3]
    DataInicial = str(sys.argv[4])
    DataFinal = str(sys.argv[5])

except:

    folder = "C:\\python\\"
    excel_Processamento = f"{folder}\\Lancamento_CTE.xlsx"
    #EnvMaxyCon = " -url http://maxys.cjtrade.com.br:7777/forms/frmservlet?config=teste_app"
    EnvMaxyCon = " -url http://maxys.cjtrade.com.br:7777/forms/frmservlet?config=maxys_prod_app"
    DataInicial = "None"
    DataFinal = "None"
    pass

def SelecionaTabela():
    
    i=1

    CTE_Anterior = ""
    Transportadora_Anterior = ""

    while True:
        
        variables = {"linha": i}
            
        PesoDestino = ui(locator.java.maxys_TAF117.Tabela_pesodestino, variables).get_text().strip()
        CTE = ui(locator.java.maxys_TAF117.Tabela_text_cte, variables).get_text().strip()
        Transportadora = ui(locator.java.maxys_TAF117.Tabela_transp, variables).get_text().strip()

        if (CTE == CTE_Anterior and Transportadora == Transportadora_Anterior):
            break
        elif (CTE == "" and Transportadora == "" and PesoDestino == ""):
            break

        CTE_Anterior = CTE
        Transportadora_Anterior = Transportadora

        if PesoDestino != "" and ui(locator.java.maxys_TAF117.Tabela_check_box, variables).get_text() != "checked":
            ui(locator.java.maxys_TAF117.Tabela_check_box, variables).click()
            
        i+=1
        
        if i == 16:
            
            ui(locator.java.maxys_TAF117.Tabela_pesodestino, variables).click()
            ui(locator.java.maxys_TAF117.Tabela_pesodestino, variables).send_hotkey("{DOWN}")

            CTE = ui(locator.java.maxys_TAF117.Tabela_text_cte, variables).get_text().strip()
            Transportadora = ui(locator.java.maxys_TAF117.Tabela_transp, variables).get_text().strip()
            
            if CTE == CTE_Anterior and Transportadora == Transportadora_Anterior:
                break

            else:
                ui(locator.java.maxys_TAF117.Tabela_pesodestino, variables).send_hotkey("{DOWN 14}")
                i=1

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
    time.sleep(1)
    ui(locator.java.maxys.text_usuário).send_hotkey("{Tab}")
    time.sleep(1)

    # Digita Senha
    ui(locator.java.maxys.password_text_senha).send_hotkey("{Tab}")
    time.sleep(1)

    # Digita Empresa
    ui(locator.java.maxys.text_empresa).send_hotkey(empresa)
    time.sleep(1)

    # Click Ok
    clicknium.wait_appear(locator.java.maxys.push_button_ok_alt_o)
    time.sleep(1)
    ui(locator.java.maxys.push_button_ok_alt_o).click()

def _abrirMaxysErp():
    
    # Define the path to the executable
    executable_path = r"C:\Users\svc_rpa\AppData\Local\Maxicon Sistemas\Maxys\Maxys.exe"

    # Define the command-line parameter 
    parameter = EnvMaxyCon

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
        if Mensagem != "Processo finalizado com sucesso.":
            raise Exception(Mensagem)
        else:
            return Mensagem

def _fechar_Precaucao():
    if clicknium.is_existing(locator.java.maxys_TAF117.PopUp_precaucao):
        Mensagem = ui(locator.java.maxys_TAF117.precaucao_text_area).get_text()
        ui(locator.java.maxys_TAF117.precaucao_ok_alt_o).click()
        autoit.sleep(500)
        raise Exception(Mensagem)

def rename_xlsx(directory: str, filial: str, new_name: str):
    """
    Busca um arquivo Excel no diretório especificado com o padrão 'TAF1170007SRP*.XLSX'
    e o renomeia para o nome fornecido.

    :param directory: Caminho do diretório onde o arquivo está localizado.
    :param new_name: Novo nome para o arquivo (deve incluir a extensão .xlsx).
    """

    # Garante que o novo nome tenha a extensão .xlsx
    if not new_name.lower().endswith(".xlsx"):
        new_name += ".xlsx"

    # Lista todos os arquivos do diretório
    files = os.listdir(directory)

    # Filtra arquivos que correspondem ao padrão
    matching_files = [f for f in files if fnmatch.fnmatch(f, f"TAF11700{filial}SRP*.XLSX")]

    # Exibe os arquivos encontrados para depuração
    if not matching_files:
        print("Nenhum arquivo correspondente encontrado no diretório.")
        return None

    # Ordena os arquivos por nome para pegar o mais recente, se houver múltiplos
    matching_files.sort(reverse=True)  # Pega o último (mais recente) pelo nome

    # Pega o primeiro arquivo encontrado
    old_file = os.path.join(directory, matching_files[0])
    new_file = os.path.join(directory, new_name)

    os.rename(old_file, new_file)

def ExportarExcel(filial, resultado_do_processamento):
    
    # Check Gerar EXECEL
    clicknium.wait_appear(locator.java.maxys_TAF117.Resultado_do_processamento)
    clicknium.wait_appear(locator.java.maxys_ExportExcel.Export_check_box_gerar_excel)
    ui(locator.java.maxys_ExportExcel.Export_check_box_gerar_excel).set_checkbox("uncheck")
    ui(locator.java.maxys_ExportExcel.Export_check_box_gerar_excel).set_checkbox("check")

    # UnCheck 
    clicknium.wait_appear(locator.java.maxys_ExportExcel.Export_check_box_visualizar_planilha)
    ui(locator.java.maxys_ExportExcel.Export_check_box_visualizar_planilha).set_checkbox("uncheck")

    # Check Salvar Automaticamente
    ui(locator.java.maxys_ExportExcel.Export_check_box_salvar_automaticamente).set_checkbox("uncheck")
    ui(locator.java.maxys_ExportExcel.Export_check_box_salvar_automaticamente).set_checkbox("check")

    # Click escolher pasta
    ui(locator.java.maxys_ExportExcel.Export_choose_push_button).click()

    # Digita Folder
    # Espera Janela Selecina uma Pasta
    WinTitle = "Selecione uma pasta"
    WinText = ""
    
    if autoit.WinWait(WinTitle,WinText,15) == 1:
        
        if clicknium.wait_appear(locator.java.maxys_VFS014.Selecione_Uma_Pasta_janela,wait_timeout=5):
            
            while ui(locator.java.maxys_VFS014.Selecione_Uma_Pasta_input_folder_name).get_text().strip() != folder:
                
                ui(locator.java.maxys_VFS014.Selecione_Uma_Pasta_input_folder_name).set_text(folder)
                autoit.Sleep(500)

            ui(locator.java.maxys_VFS014.Selecione_Uma_Pasta_button_ok).click()

        else:
        
            while autoit.ControlGetText(WinTitle, WinText,"Edit1") != folder:
                
                # Set Text caminho pasta xml
                autoit.ControlSend(WinTitle,WinText,"Edit1",folder)
                autoit.ControlSetText(WinTitle,WinText,"Edit1",folder)
                autoit.Sleep(500)
            
            #Click Selecionar
            autoit.ControlClick(WinTitle, WinText,"Button1")

            ui(locator.java.maxys_ExportExcel.Export_select_edit_folder).set_text(folder)
            time.sleep(1)

            # Click Select Folder
            ui(locator.java.maxys_ExportExcel.Export_button_select_folder).set_focus()
            ui(locator.java.maxys_ExportExcel.Export_button_select_folder).click()
    
    else:
        raise Exception(f"Janela 'Selecione uma pasta nao apareceu', para salvar excel.")
    
    # Click Gerar
    clicknium.wait_appear(locator.java.maxys_ExportExcel.Export_push_button_gerar_alt_g)
    ui(locator.java.maxys_ExportExcel.Export_push_button_gerar_alt_g).click()

    # Esperar Popup sucesso
    if clicknium.wait_appear(locator.java.maxys_VFS014.Observacao_Sucesso,wait_timeout=5):
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
    
    # Fechar PopUp
    clicknium.wait_appear(locator.java.maxys_ExportExcel.Export_push_button_fechar_alt_f)
    ui(locator.java.maxys_ExportExcel.Export_push_button_fechar_alt_f).click()

    rename_xlsx(folder, filial, resultado_do_processamento)

if __name__ == "__main__":

    # empresas = [
    #     [5, "CJ INTERNATIONAL BRASIL - PASSO FUNDO"],
    #     [6, "CJ INTERNATIONAL BRASIL - RIO GRANDE"]
    # ]

    empresas = [
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

    # empresas = [
    #     [1, "CJ INTERNATIONAL BRASIL COMERCIAL"],
    #     [2, "CJ INTERNATIONAL BRASIL - PARANAGUA"],
    #     [3, "CJ INTERNATIONAL BRASIL - RIO VERDE"],
    #     [4, "CJ INTERNATIONAL BRASIL - LONDRINA"],
    #     [5, "CJ INTERNATIONAL BRASIL - PASSO FUNDO"],
    #     [6, "CJ INTERNATIONAL BRASIL - RIO GRANDE"],
    #     [7, "CJ INTERNATIONAL BRASIL - SORRISO"],
    #     [8, "CJ INTERNATIONAL BRASIL - SANTOS"],
    #     [9, "CJ INTERNATIONAL BRASIL - ARAGUARI"],
    #     [10, "CJ INTERNATIONAL BRASIL - QUERENCIA"],
    #     [11, "CJ INTERNATIONAL BRASIL - IMBITUBA"],
    #     [12, "CJ INTERNATIONAL BRASIL - PARAGOMINAS"],
    #     [13, "CJ INTERNATIONAL BRASIL - NOVA ODESSA"]
    # ]


    today = datetime.now()

    if DataInicial == "None" and DataFinal == "None":
        DataFinal = (today - timedelta(days=2)).strftime('%d/%m/%Y')
        DataInicial = (today - timedelta(days=10)).strftime('%d/%m/%Y')

    # Lista para armazenar dados para criar o DataFrame
    data = []
    
    for empresa in empresas:
        cod_empresa = empresa[0]
        nome_empresa = empresa[1]
        Tipos_ctes = ["normal", "complementar", "substituição"]

        # Abrir o sistema e realizar login
        _abrirMaxysErp()
        _Login_MaxysErp(cod_empresa)

        for tipo_cte in Tipos_ctes:
            resultado_do_processamento = f"{folder}\\{cod_empresa}_{nome_empresa}_{tipo_cte}_{DataInicial.replace('/', '')}.xlsx"
            print(resultado_do_processamento)
            try:

                # Executar rotina no sistema
                _Executar("TAF117")

                clicknium.wait_appear(locator.java.maxys_TAF117.text_dt_emissão_inicial)

                # Configurar os filtros
                ui(locator.java.maxys_TAF117.text_empresa).set_text(cod_empresa)
                ui(locator.java.maxys_TAF117.text_dt_emissão_inicial).set_text(DataInicial)
                ui(locator.java.maxys_TAF117.text_dt_emissão_final).set_text(DataFinal)

                if tipo_cte == "normal":
                    ui(locator.java.maxys_TAF117.radio_button_ct_e_normal).click()
                elif tipo_cte == "complementar":
                    ui(locator.java.maxys_TAF117.radio_button_ct_e_complementar).click()
                elif tipo_cte == "substituição":
                    ui(locator.java.maxys_TAF117.radio_button_ct_e_de_substituição).click()

                # Consultar
                ui(locator.java.maxys_TAF117.push_button_f3_consultar).click()
                
                _fechar_Observacao()
                                
                # Marcar tudo peso Destino difernete de varia e gravar
                SelecionaTabela()
                
                # CLICK GRAVAR
                ui(locator.java.maxys_TAF117.push_button_gravar).click()
                
                Mensagem = ""
                Mensagem = _fechar_Observacao()
                _fechar_Precaucao()
                
                if Mensagem != "Processo finalizado com sucesso.":
                    
                    ExportarExcel(f"{cod_empresa:02}", resultado_do_processamento)

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
                
                if not "A consulta não retornou dados com base nos filtros" in str(e):
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