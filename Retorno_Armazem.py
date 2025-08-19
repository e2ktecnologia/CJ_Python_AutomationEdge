#! C:\python\venv\Scripts\python.exe

from clicknium import clicknium, ui, locator
from clicknium.common.enums import *
import subprocess
import os
import datetime
import pandas as pd
import win32com.client as client
import sys
from dotenv import load_dotenv
import unicodedata

# Carregar variáveis do arquivo .env
load_dotenv()

# Capturar a licença do Clicknium do .env
clicknium_license = os.getenv("CLICKNIUM_LICENSE")
clicknium.config.set_license(clicknium_license)
autoit = client.Dispatch("AutoItX3.Control")
wshell = client.Dispatch("WScript.Shell")
NrNFE = None

try:
    
    excel_file_base = sys.argv[1]
    excel_Processamento = sys.argv[2]
    folderXML = str(sys.argv[3])
    EnvMaxyCon = sys.argv[4]
    
except:

    excel_file_base = r"C:\Studio\Process-Studio\process-studio\ps-workspace\CJ_TrocaDeNota_Retorno\Base.xlsx"
    excel_Processamento = r"C:\Studio\Process-Studio\process-studio\ps-workspace\CJ_TrocaDeNota_Retorno\Base_Processada.xlsx"
    folderXML = r"C:\Temp\XML"
    EnvMaxyCon = " -url http://maxys.cjtrade.com.br:7777/forms/frmservlet?config=maxys_prod_app"
    pass

df = pd.read_excel(excel_file_base)

def remover_acentos(texto: str) -> str:
    # Normaliza e remove os acentos
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

def capturar_caminhos(pasta):
    """
    Captura os caminhos de arquivos XML e PDF em uma pasta.
    
    :param pasta: Caminho da pasta onde os arquivos serão buscados.
    :return: Dicionário com listas de caminhos para XML e PDF.
    """
    arquivos_xml = "Não Gerado"
    arquivos_pdf = "Não Gerado"
    
    # Verifica se a pasta existe
    if not os.path.exists(pasta):
        raise FileNotFoundError(f"A pasta '{pasta}' não foi encontrada.")
    
    # Percorre os arquivos na pasta
    for arquivo in os.listdir(pasta):
        caminho_completo = os.path.join(pasta, arquivo)
        
        if os.path.isfile(caminho_completo):
            if arquivo.lower().endswith(".xml"):
                arquivos_xml = (caminho_completo)
            elif arquivo.lower().endswith(".pdf"):
                arquivos_pdf = (caminho_completo)
    
    return {"XML": arquivos_xml, "PDF": arquivos_pdf}

def _Login_MaxysErp(empresa):

    try:

        # Digita Login
        clicknium.wait_appear(locator.java.maxys.text_usuário)
        ui(locator.java.maxys.text_usuário).send_hotkey("{Tab}")
        
        # Digita Senha
        ui(locator.java.maxys.password_text_senha).send_hotkey("{Tab}")

        # Digita Empresa
        ui(locator.java.maxys.text_empresa).send_hotkey(empresa)

        # Click Ok
        ui(locator.java.maxys.push_button_ok_alt_o).click()

    except BaseException as e:
        raise Exception(f"Erro login Maxys: {str(e)}.")

def _abrirMaxysErp():
    try:

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

    except BaseException as e:
        raise Exception(f"Erro ao abrir Maxys: {str(e)}.")

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
        #print(Mensagem)

    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
        #print(Mensagem)

    #Fechar Maxicon
    if clicknium.is_existing(locator.java.maxys.FecharMaxicon):
        ui(locator.java.maxys.FecharMaxicon).click()
    if clicknium.is_existing(locator.java.maxys.push_button_Sim_FecharSistema):
        ui(locator.java.maxys.push_button_Sim_FecharSistema).click()

def _fechar_Observacao():
    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        if "obtido através da chave de acesso não está cadastrado" in Mensagem:
            ui(locator.java.maxys_VFS014.Observacao_OK).click()
        else:    
            ui(locator.java.maxys_VFS014.Observacao_OK).click()
            autoit.sleep(500)
            raise Exception(Mensagem)

def _selecionaTabela(peso_saldo,valor_saldo,nota_fiscal):

    #columns = ["Nota Fiscal", "Emissão", "Código\nMoviment", "Contrato", "Peso\nNota Fiscal" , "Valor\nNota Fiscal", "Valor\nUnitário", "Peso\nSaldo", "Valor\nSaldo", "Peso\nUsado", "Valor\nUsado"]
    i = 1

    while True:

        variables = {"index": i, "column": "Nota Fiscal"}
        nota_fiscalMaxys = ui(locator.java.maxys_GEX004.text_cell_table, variables).get_text().strip()

        variables = {"index": i, "column": "Peso\nSaldo"}
        peso_saldoMaxys = ui(locator.java.maxys_GEX004.text_cell_table, variables).get_text().strip()

        variables = {"index": i, "column": "Valor\nSaldo"}
        valor_saldoMaxys = ui(locator.java.maxys_GEX004.text_cell_table, variables).get_text().strip()
        
        if nota_fiscal == nota_fiscalMaxys and valor_saldo == valor_saldoMaxys and peso_saldo == peso_saldoMaxys:
            
            variable = {"index": i+2}
            
            ui(locator.java.maxys_GEX004.check_box_romaneios,variable).click()
            
            break

        else:
            i = i + 1

def _ProcessaGEX004(contrato, placa, uf , motorista,peso_nf, valor,nota_fiscal, chave_de_acesso,cnpj_transportadora, infCpl):
    
    # Wait Devolucao
    clicknium.wait_appear(locator.java.maxys_GEX004.radio_button_devolucao)
    
    # Click Devolucao
    ui(locator.java.maxys_GEX004.radio_button_devolucao).click()

    # Digita Contrato
    ui(locator.java.maxys_GEX004.text_contratopedido_de_graos).set_text(contrato)
    ui(locator.java.maxys_GEX004.text_contratopedido_de_graos).send_hotkey("{TAB 3}")
    
    # Seleciona Tranpostadora
    if not clicknium.is_existing(locator.java.maxys_GEX004.input_Localizador):
        ui(locator.java.maxys_GEX004.text_transportador).send_hotkey("{F9}")
    
    ui(locator.java.maxys_GEX004.input_Localizador).set_text(f"%")
    autoit.sleep(500)
    ui(locator.java.maxys_GEX004.input_Localizador).send_hotkey("{ENTER}")
    autoit.sleep(500)
    ui(locator.java.maxys_GEX004.input_Localizador).set_text(f"{cnpj_transportadora}")
    autoit.sleep(500)
    ui(locator.java.maxys_GEX004.localizar_button_localiza).click()
    autoit.sleep(500)
    ui(locator.java.maxys_GEX004.localizar_button_ok).click()
    autoit.sleep(500)   
    
    peso_nf_f = float(str(peso_nf))
    saldo_atual = float((ui(locator.java.maxys_GEX004.text_saldo).get_text()).replace(".",""))

    if peso_nf_f > saldo_atual:
        raise Exception(f"Saldo atual menor que o peso nota fiscal saldo atual é: {str(saldo_atual)}, peso da nf é: {str(peso_nf)}")

    # Digita Placa
    ui(locator.java.maxys_GEX004.text_placa).set_text(placa)
    _fechar_Observacao()
    
    # Digita UF
    ui(locator.java.maxys_GEX004.text_uf).set_text(uf)
    
    # Digita Nome Motorista
    ui(locator.java.maxys_GEX004.text_motorista).set_text(motorista)

    # Digita Peso NF
    ui(locator.java.maxys_GEX004.text_peso_nf).set_text(str(int(peso_nf)))

    # Digita Valor
    ui(locator.java.maxys_GEX004.text_valor).set_text(str(valor).replace(".",","))
    ui(locator.java.maxys_GEX004.text_valor).send_hotkey("{TAB}")

    # Click Romaneios ou F4
    ui(locator.java.maxys_GEX004.page_tab_romaneios).click()

    # Wait button Impostos
    clicknium.wait_appear(locator.java.maxys_GEX004.push_button_impostos)
    
    # Seleciona Campo com Numero da Nota Fiscal de Referencia
    #_selecionaTabela(peso_nf,valor, nota_fiscal)
    variable = {"index": 3}
    ui(locator.java.maxys_GEX004.check_box_romaneios,variable).click()

    # Salvar 
    ui(locator.java.maxys_GEX004.push_button_salvar).click()

    # Digita Chave de Acesso
    ui(locator.java.maxys_GEX004.text_chave_de_acesso).set_text(chave_de_acesso)
    ui(locator.java.maxys_GEX004.text_chave_de_acesso).send_hotkey("{TAB}")

    # Digita Data Lancamento
    # Obter a data atual
    #data_atual = datetime.datetime.now()
    #data_atual.strftime('%d/%m/%Y')
    #ui(locator.java.maxys_GEX004.text_data_de_lancamento).set_text()

    # Digita Peso de balança
    #ui(locator.java.maxys_GEX004.text_peso_de_balança).set_text(peso_nf)
    if clicknium.is_existing(locator.java.maxys_GEX004.Mensagem_button_ok):
        ui(locator.java.maxys_GEX004.Mensagem_button_ok).click()
    
    # Verifica se deu mensagem de observacao
    _fechar_Observacao()

    # Seleciona Tipo Calculo
    ui(locator.java.maxys_GEX004.combo_box_tipo_cálculo).click()
    autoit.sleep(1000)
    ui(locator.java.maxys_GEX004.combo_box_tipo_cálculo).send_hotkey("{HOME}")
    autoit.sleep(1000)
    ui(locator.java.maxys_GEX004.label_2_preco_x_quantidade).click()

    # Send Tab
    ui(locator.java.maxys_GEX004.combo_box_tipo_cálculo).send_hotkey("{TAB 3}")

    # Click - OK
    ui(locator.java.maxys_GEX004.push_button_ok_alt_o).click()

    # Wait Devolucao
    clicknium.wait_appear(locator.java.maxys_GEX004.vinculada_expedicao_ok)
    
    # Click Ok
    ui(locator.java.maxys_GEX004.vinculada_expedicao_ok).click()
    

def _ProcessaGEX001(contrato_venda,clifor_transportadora,transgenia):
    
    # Espera tela formação de lote click OK
    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso,timeout=15):
        
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        
        if not "A movimentação deste contrato está associada" in Mensagem:
            raise Exception(Mensagem)
        
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
        autoit.Sleep(1500)
        if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_OK):
            ui(locator.java.maxys_VFS014.Observacao_OK).click()

    # Digita Contrato Saida
    if clicknium.is_existing(locator.java.maxys_GEX001.Popup_Contrato_button_cancelar):
        ui(locator.java.maxys_GEX001.Popup_Contrato_button_cancelar).click()

    clicknium.wait_appear(locator.java.maxys_GEX001.Principal_contrato_text)
    
    # Digita Contrato
    contrato = ui(locator.java.maxys_GEX001.Principal_contrato_text)
    if contrato.get_text().replace(".","") != str(contrato_venda):
        try:
            contrato.clear_text("send-hotkey")
            contrato.set_text(contrato_venda)
            contrato.send_hotkey("{TAB}")
        except:
            Mensagem = "Erro ao tentar digitar contrato venda"
            #_pendencia(Mensagem,number)
            raise Exception(Mensagem)
   
    # Se Existir Popup Observacao
    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        if "paga frete" in Mensagem: 
            ui(locator.java.maxys_VFS014.Observacao_OK).click()
        elif "A movimentação deste contrato está associada a CMI" in Mensagem: 
            ui(locator.java.maxys_VFS014.Observacao_OK).click()
        else:
            _fechar_Observacao()

    # Digita Transportador
    if clicknium.is_existing(locator.java.maxys_GEX001.Principal_transportador_text):
        principal_transportador = ui(locator.java.maxys_GEX001.Principal_transportador_text)

        if principal_transportador.get_text().replace(".","") != clifor_transportadora:
            principal_transportador.set_text(clifor_transportadora)

    # Digitar Tela Trasportador a transportadora
    if clicknium.is_existing(locator.java.maxys_GEX001.Transportador_localizar_text):
        transportador_text = ui(locator.java.maxys_GEX001.Transportador_localizar_text)
        transportador_text.clear_text("send-hotkey")
        transportador_text.set_text(f"%{clifor_transportadora}")
        transportador_text.send_hotkey("{ENTER}")

        if ui(locator.java.maxys_GEX001.Transportadora_list).get_text()=='':
            
            cancelar = ui(locator.java.maxys_GEX001.Transportadora_Cancelar_btn)
            cancelar.click()
            autoit.Sleep(1000)
            #_pendencia("Transportador não vinculado ao contrato de venda",number)
            
            raise Exception("Transportador não vinculado ao contrato de venda")

        # Click Localizar
        localizar = ui(locator.java.maxys_GEX001.Transportador_localizar_text)
        localizar.click()

        # Click OK
        OK_Transportador = ui(locator.java.maxys_GEX001.Transportador_push_button_ok)
        OK_Transportador.click()
    
    
    # Embarque
    #wshell.popup("Teste",0)
    #embarque = ui(locator.java.maxys_GEX001.Principal_local_de_embarque_text)

    #if embarque.get_text()!=str(codLocalEmbarque):
        
    #    while embarque.get_text() != '':
    #        embarque.clear_text("send-hotkey","HSED")

    #   embarque.set_text(codLocalEmbarque)
    #    embarque.send_hotkey("{TAB}")

    # Valor Liquido
    valorLiquido = ui(locator.java.maxys_GEX001.Principal_liquido_text).get_text()
    
    if valorLiquido=="":
        Mensagem =f"Valor do campo liquido em branco valor capturado:{valorLiquido}"
        raise Exception(Mensagem)
    
    # Click Salvar
    ui(locator.java.maxys.Salvar).click()
    
    # FinsExportacao
    if clicknium.is_existing(locator.java.Venda_C_FinsExportacao.Popup_nao):
        ui(locator.java.Venda_C_FinsExportacao.Popup_nao).click()
    
    if clicknium.wait_appear(locator.java.maxys_GEX001.SelecionaAmostra_OK,wait_timeout=15):
        ui(locator.java.maxys_GEX001.SelecionaAmostra_localizar_text).clear_text("send-hotkey")
        ui(locator.java.maxys_GEX001.SelecionaAmostra_localizar_text).set_text("%" + str(transgenia).strip())
        ui(locator.java.maxys_GEX001.SelecionaAmostra_localizar_text).send_hotkey("{ENTER}")
        ui(locator.java.maxys_GEX001.SelecionaAmostra_OK).click()
    
    if clicknium.wait_appear(locator.java.maxys_VFS014.Observacao_Sucesso,wait_timeout=5):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
        raise Exception(Mensagem)
    
    #Espera Tela
    clicknium.wait_appear(locator.java.maxys_VFS014.NFE_consultarSefaz)
    
    # Get NrNFE
    NrNFE = ui(locator.java.maxys_VFS014.NFE_nr_nfempr).get_text()

    i = 0
    while ui(locator.java.maxys_VFS014.NFE_combobox_result_Sefaz).child(0).get_text() == "Pendente":
        
        autoit.sleep(3000)
        
        # Click Consultar
        ui(locator.java.maxys_VFS014.NFE_consultarSefaz).click()

        # Verifica Status
        if i >10 or ui(locator.java.maxys_VFS014.NFE_combobox_result_Sefaz).child(0).get_text() != "Pendente":
        
            break
        
        # Incremente
        i=i+1
    
    # Get Status Sefaz
    status_Sefaz = ui(locator.java.maxys_VFS014.NFE_combobox_result_Sefaz).child(0).get_text()

    if status_Sefaz.strip() != "Aprovada":
    
        Mensagem=  f"Status da nota no Sefaz diferente de Aprovada: {status_Sefaz}, Numero NFE: {str(NrNFE)}"
        raise Exception(f"Status da nota no Sefaz diferente de Aprovada: {status_Sefaz}")
    
    else:

        #Click Checkbox
        ui(locator.java.maxys_VFS014.NFE_check_box_nfe).click()

        # Exportar NFE
        ui(locator.java.maxys_VFS014.NFE_exportar_nfe).click()

        # Seleciona Tipo de arquivo
        ui(locator.java.maxys_VFS014.Exportacao_tipo_de_arquivo).click()
        ui(locator.java.maxys_VFS014.Exportacao_ambos_label).click()
        
        # Seleciona Folder
        ui(locator.java.maxys_VFS014.Exportacao_changeFolder).click()

        # Espera Janela Selecina uma Pasta
        WinTitle = "Selecione uma pasta"
        WinText = ""
        
        if autoit.WinWait(WinTitle,WinText,15) == 1:
            
            if clicknium.wait_appear(locator.java.maxys_VFS014.Selecione_Uma_Pasta_janela,wait_timeout=5):
                
                while ui(locator.java.maxys_VFS014.Selecione_Uma_Pasta_input_folder_name).get_text().strip() != folderXML:
                    
                    ui(locator.java.maxys_VFS014.Selecione_Uma_Pasta_input_folder_name).set_text(folderXML)
                    autoit.Sleep(500)

                ui(locator.java.maxys_VFS014.Selecione_Uma_Pasta_button_ok).click()

            else:
                while autoit.ControlGetText(WinTitle, WinText,"Edit1") != folderXML:
                    
                    # Set Text caminho pasta xml
                    autoit.ControlSend(WinTitle,WinText,"Edit1",folderXML)
                    autoit.ControlSetText(WinTitle,WinText,"Edit1",folderXML)
                    autoit.Sleep(500)
                
                #Click Selecionar
                autoit.ControlClick(WinTitle, WinText,"Button1")
        
        else:
            raise Exception(f"Janela 'Selecione uma pasta nao apareceu', para salvar xml.")

    # Click Exportar
    ui(locator.java.maxys_VFS014.Exportacao_exportar_enviar).click()
    
    # Espera Popup Observacao
    clicknium.wait_appear(locator.java.maxys_VFS014.Observacao_Sucesso)
    
    # Se Existir Popup Observacao
    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
        #print(Mensagem)

    # Exportacao Voltar
    ui(locator.java.maxys_VFS014.Exportacao_Voltar).click()

    return NrNFE

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
        [11, "CJ INTERNATIONAL BRASIL - IMBITUBA"],
        [12, "CJ INTERNATIONAL BRASIL - PARAGOMINAS"],
        [13, "CJ INTERNATIONAL BRASIL - NOVA ODESSA"]
    ]
      
    for index, row in df.iterrows():
        
        try:

            for i in range(0,len(empresas)):
                
                if "MATRIZ SÃO PAULO" in df["Local"].iloc[0].upper():
                    empresa = 1
                    break

                elif df["Local"].iloc[0].upper() in empresas[i][1]:
                    empresa = empresas[i][0]
                    break
            else:
                raise Exception("Não foi possivel definir empresa para login!")

            # for i in range(0,len(empresas)):
                
            #     if remover_acentos(df["Grupo"].iloc[0].split('-')[0].strip().upper()) in remover_acentos(str(empresas[i][1]).upper()):
            #         empresa = empresas[i][0]
            #         break
            # else:
            #     raise Exception(f'Não foi possivel definir empresa ({df["Grupo"].iloc[0].split("-")[0].strip().upper()}) para login!')

            # Login
            _abrirMaxysErp()
            _Login_MaxysErp(empresa)

            # Abrir tela GEX004
            _Executar("GEX004")

            # Processa GEX004
            _ProcessaGEX004(str(row['numero_do_contrato']), row['placa'].replace("-",""), row['uf'] , row['nome_do_motorista'],row["qCom"], row["vNF"], row["contrato_de_venda"], str(row['chave_de_acesso_nf_compra'].replace(" ","")), row["cnpj_da_transportadora"], row["infCpl"])
            
            current_date = datetime.datetime.now()
            dataLog = current_date.strftime("%d/%m/%Y %H:%M:%S")
                
            # Escreve mensagem e data log
            df.loc[index, 'Mensagem_Entrada'] = "Nota de Entrada Concluida"
            df.loc[index, 'DataLog_Entrada'] = dataLog
            
            # Processa GEX001
            NrNFE=_ProcessaGEX001(row['contrato_de_venda'],str(row['clifor_transportadora']),row["transgenia"])

            current_date = datetime.datetime.now()
            dataLog = current_date.strftime("%d/%m/%Y %H:%M:%S")
                
            # Escreve mensagem e data log
            df.loc[index, 'Nota Fiscal Saida'] = NrNFE
            df.loc[index, 'Mensagem_Saida'] = "Nota de Saida Concluida"
            df.loc[index, 'DataLog_Saida'] = dataLog

            # Sair Sistema
            if clicknium.is_existing(locator.java.maxys.SairDoPrograma):
                ui(locator.java.maxys.SairDoPrograma).click()
            
            clicknium.wait_appear(locator.java.maxys_VFS014.Observacao_Sucesso)
            
            if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
                Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
                ui(locator.java.maxys_VFS014.Observacao_OK).click()
            
            clicknium.wait_appear(locator.java.maxys_VFS014.Observacao_Sucesso)

            if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
                Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
                ui(locator.java.maxys_VFS014.Observacao_OK).click()

            # Sair Sistema
            if clicknium.is_existing(locator.java.maxys.SairDoPrograma):
                ui(locator.java.maxys.SairDoPrograma).click()
            
            # Sair Sistema
            if clicknium.is_existing(locator.java.maxys.SairDoPrograma):
                ui(locator.java.maxys.SairDoPrograma).click()
                
            Mensagem = "Efetuado com sucesso."

        except BaseException as e:
            
            #print(str(e))
            Mensagem = str(e)
            #wshell.Popup(str(e),0,"Error!")
            
            if "Element can not be found" in Mensagem or "Set focus failed" in Mensagem:
                Mensagem = "O lançamento automático não foi concluído com sucesso. Por favor, verifique se há algum problema com os dados inseridos e realize o lançamento manualmente."

            current_date = datetime.datetime.now()
            dataLog = current_date.strftime("%d/%m/%Y %H:%M:%S")
            
            df.loc[index, 'Mensagem_Saida'] = f"Erro:{str(e)}"
            df.loc[index, 'DataLog_Saida'] = dataLog

            raise Exception(str(e))
        
        finally:

            # Fechar Sistema
            _FecharSistema()
            #result = capturar_caminhos(folderXML)
                        
            print(Mensagem)

            # Gerar Excel Final de Processamento
            df.to_excel(excel_Processamento, index=False)