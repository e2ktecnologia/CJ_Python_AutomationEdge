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
    
    excel_file_base = r"C:\Studio\Process-Studio_prod\process-studio\ps-workspace\CJ_Maxys_TrocaDeNota_Retorno\Base.xlsx"
    excel_Processamento = r"C:\Studio\Process-Studio_prod\process-studio\ps-workspace\CJ_Maxys_TrocaDeNota_Retorno\Base_Processada.xlsx"
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

def mask_cnpj(cnpj: str) -> str:
    """
    Aplica máscara de CNPJ no formato 00.000.000/0000-00
    :param cnpj: string contendo apenas os números do CNPJ
    :return: string formatada
    """
    cnpj = ''.join(filter(str.isdigit, cnpj))  # remove qualquer caractere não numérico
    if len(cnpj) != 14:
        raise ValueError("CNPJ deve conter 14 dígitos.")
    
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

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
    
def _ProcessaGRE001(NFE,novo_contrato,nome_do_motorista,placa,uf,analise,transgenia,cnpj_da_transportadora,vNF):
    
    # Digita contrato
    clicknium.wait_appear(locator.java.maxys_GRE001.text_chave_acesso_nf_e)
    chNFE = ui(locator.java.maxys_GRE001.text_chave_acesso_nf_e)
    chNFE.clear_text("send-hotkey")
    chNFE.set_text(NFE)
    chNFE.send_hotkey("{TAB}")
    
    _fechar_Observacao()
    
    # Get preco unitario
    contrato = ui(locator.java.maxys_GRE001.text_contrato)

    contrato.click()
    
    if clicknium.is_existing(locator.java.maxys_GRE001.Janela_SelecaoDeContratos):
        ui(locator.java.maxys_GRE001.SelecaoDeContratos_btn_cancel).click()

    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
    
    chNFE.click()

    while contrato.get_text() != '':
        contrato.clear_text("send-hotkey","HSED")

    contrato.set_text(novo_contrato.strip())
    contrato.send_hotkey("{TAB}")

    if clicknium.wait_appear(locator.java.maxys_GRE001.Transportador.pesquisa_transportador,wait_timeout=5):
        
        p_transportador = ui(locator.java.maxys_GRE001.Transportador.pesquisa_transportador)
        p_transportador.clear_text("send-hotkey","HSED")
        p_transportador.set_text(f"%")
        autoit.sleep(500)
        p_transportador.send_hotkey("{ENTER}")
        autoit.sleep(500)
        p_transportador.set_text(f"{cnpj_da_transportadora}")
        autoit.sleep(500)
        p_transportador.send_hotkey("{ENTER}")
        autoit.sleep(500)
        ui(locator.java.maxys_GRE001.Transportador.pesquisa_click_localizar).click()
        resultado_transportadora=ui(locator.java.maxys_GRE001.Transportadora_list).get_text()
            
        if resultado_transportadora == "":
            raise Exception("Não foi possivel encontrar transportadora na pesquisa Maxys, verifique se transportadora vinculada ao contrato!")
        
        ui(locator.java.maxys_GRE001.Transportador.pesquisa_click_ok).click()

    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        if "paga frete" in Mensagem: 
            ui(locator.java.maxys_VFS014.Observacao_OK).click()
        else:
            _fechar_Observacao()
    
    fornecedor = ui(locator.java.maxys_GRE001.text_fornecedor).get_text()
            
    # Digita Localizar Embarque/Desembarque
    while clicknium.is_existing(locator.java.maxys_GRE001.E_D_push_button_cancelar_alt_c):
        ui(locator.java.maxys_GRE001.E_D_push_button_cancelar_alt_c).click()
    
    # Pop up Atencao
    if clicknium.wait_appear(locator.java.maxys_GRE001.Atencao_TrocaNota_Sim,wait_timeout=10):
        ui(locator.java.maxys_GRE001.Atencao_TrocaNota_Sim).click()

    # Digitar Transportador
    transportador = ui(locator.java.maxys_GRE001.text_transportador)

    if ui(locator.java.maxys_GRE001.text_transportador).get_text() == "":
        
        transportador.send_hotkey("{F9}")

        # Espera Janela Pesquisa
        if clicknium.wait_appear(locator.java.maxys_GRE001.Transportador.pesquisa_transportador,wait_timeout=10):
            
            p_transportador = ui(locator.java.maxys_GRE001.Transportador.pesquisa_transportador)
            p_transportador.set_text(f"%")
            autoit.sleep(500)
            p_transportador.send_hotkey("{ENTER}")
            autoit.sleep(500)
            p_transportador.set_text(f"{cnpj_da_transportadora}")
            autoit.sleep(500)
            p_transportador.send_hotkey("{ENTER}")
            autoit.sleep(500)
            ui(locator.java.maxys_GRE001.Transportador.pesquisa_click_localizar).click()
            resultado_transportadora=ui(locator.java.maxys_GRE001.Transportadora_list).get_text()
            
            if resultado_transportadora == "":
                raise Exception("Não foi possivel encontrar transportadora na pesquisa Maxys, verifique se transportadora vinculada ao contrato!")
            
            ui(locator.java.maxys_GRE001.Transportador.pesquisa_click_ok).click()

        else:
        
            raise Exception("Tela de Pesquisar Transportador nao encotrado.")
        
    _fechar_Observacao()
        
    # Digitar Motorista
    ui(locator.java.maxys_GRE001.text_motorista).clear_text("send-hotkey")
    ui(locator.java.maxys_GRE001.text_motorista).set_text(nome_do_motorista)
    ui(locator.java.maxys_GRE001.text_motorista).send_hotkey("{TAB}")

    # Digitar Placa
    if ui(locator.java.maxys_GRE001.text_placa).get_text() == "":
        ui(locator.java.maxys_GRE001.text_placa).clear_text("send-hotkey")
        ui(locator.java.maxys_GRE001.text_placa).set_text(placa)
        ui(locator.java.maxys_GRE001.text_placa).send_hotkey("{TAB}")
    else:
        ui(locator.java.maxys_GRE001.text_placa).send_hotkey("{TAB}")

    _fechar_Observacao()
    
    # Digitar UF
    if ui(locator.java.maxys_GRE001.text_uf).get_text().strip() == "":

        ui(locator.java.maxys_GRE001.text_uf).clear_text("send-hotkey")
        ui(locator.java.maxys_GRE001.text_uf).set_text(uf)
        ui(locator.java.maxys_GRE001.text_uf).send_hotkey("{TAB}")

    else:
        ui(locator.java.maxys_GRE001.text_uf).send_hotkey("{TAB}")

    # Click Tab Analise
    ui(locator.java.maxys_GRE001.page_tab_análise).click()

    i = 16
    while True:
        
        dict = {"index":i}
        
        if ui(locator.java.maxys_GRE001.text_table_analise,dict).get_text() == "" :
            break

        # Get Values Analise
        value_analise = analise[ui(locator.java.maxys_GRE001.text_table_analise,dict).get_text()]

        # Set Value
        dict = {"index":i-15}
        ui(locator.java.maxys_GRE001.text_resultado_deorigem,dict).click()
        ui(locator.java.maxys_GRE001.text_resultado_deorigem,dict).clear_text("send-hotkey")
        ui(locator.java.maxys_GRE001.text_resultado_deorigem,dict).set_text(value_analise)
        ui(locator.java.maxys_GRE001.text_resultado_deorigem,dict).send_hotkey("{TAB}")

        i=i+1
    
    # Click Tab Fornecedor
    ui(locator.java.maxys_GRE001.page_tab_fornecedores).click()

    if clicknium.is_existing(locator.java.maxys_GRE001.Precaucao_popup):
        Mensagem = ui(locator.java.maxys_GRE001.Precaucao_text_mensagem).get_text()
        ui(locator.java.maxys_GRE001.Precaucao_push_button_ok).click()
        if "Não é permitido recebimento de peso" in Mensagem:
            raise Exception(Mensagem)
        
    if clicknium.is_existing(locator.java.maxys_GRE001.popup_observacao):        
        ui(locator.java.maxys_GRE001.push_button_ok_alt_o).click()
    
    if clicknium.is_existing(locator.java.maxys_GRE001.Janela_Atencao):
        ui(locator.java.maxys_GRE001.JanelaAtencao_btn_sim).click()

    if clicknium.is_existing(locator.java.maxys_GRE001.popup_observacao):
        ui(locator.java.maxys_GRE001.push_button_ok_alt_o).click()
    
    # Get Valor Total Origem
    copiaVTO = ui(locator.java.maxys_GRE001.For_text_valor_total_de_origem).get_text()
    
    if copiaVTO.replace(".","") != f"{vNF:.2f}".replace(".", ","):
        raise Exception("Error Valor Total Origem diferente do que está na nota!")
    else:
        ui(locator.java.maxys_GRE001.text_preço_unitário).send_hotkey("{TAB}")
        
        if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
            ui(locator.java.maxys_VFS014.Observacao_OK).click()
        
        ui(locator.java.maxys_GRE001.combo_box_tipo_de_cálculo_do_preço).send_hotkey("{TAB 2}")

        ui(locator.java.maxys_GRE001.For_text_valor_total_de_origem).clear_text("send-hotkey")
        ui(locator.java.maxys_GRE001.For_text_valor_total_de_origem).set_text(copiaVTO)
        ui(locator.java.maxys_GRE001.For_text_valor_total_de_origem).send_hotkey("{TAB}")
        
        #wshell.Popup("Text",0,"Error!")

    # Gravar
    ui(locator.java.maxys_GRE001.Gravar).click()
    
    if clicknium.is_existing(locator.java.maxys_VFS014.Observacao_Sucesso):
        Mensagem = ui(locator.java.maxys_VFS014.Observacao_Mensagem).get_text()
        if "paga frete" in Mensagem or "O valor unitário informado" in Mensagem: 
            ui(locator.java.maxys_VFS014.Observacao_OK).click()

    if clicknium.is_existing(locator.java.maxys_GRE001.Janela_Atencao):
        ui(locator.java.maxys_GRE001.JanelaAtencao_btn_sim).click()

    if clicknium.wait_appear(locator.java.maxys_GEX001.Principal_contrato_text,wait_timeout=5):
        return "OK"
    else:
        if clicknium.is_existing(locator.java.maxys_GRE001.SelecionaAmostra_OK):
            ui(locator.java.maxys_GRE001.SelecionaAmostra_Search).clear_text("send-hotkey")
            ui(locator.java.maxys_GRE001.SelecionaAmostra_Search).set_text("%" + str(transgenia).strip())
            ui(locator.java.maxys_GRE001.SelecionaAmostra_Search).send_hotkey("{ENTER}")
            ui(locator.java.maxys_GRE001.SelecionaAmostra_OK).click()
        # Digita Contrato Saida
        if clicknium.is_existing(locator.java.maxys_GEX001.Popup_Contrato_button_cancelar):
            ui(locator.java.maxys_GEX001.Popup_Contrato_button_cancelar).click()

def _ProcessaGEX001(contrato_venda,clifor_transportadora,transgenia,emitCNPJ):
    
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
        if "paga frete" in Mensagem or "A movimentação deste contrato está associada" in Mensagem:
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
    if empresa == 6:

        embarque = ui(locator.java.maxys_GEX001.Principal_local_de_embarque_text)
        embarque.send_hotkey("{F9}")
        
        # Limpa campo Localizar e settext CNPJ Emitente
        localizar = ui(locator.java.maxys_GEX001.text_localizar)
        localizar.clear_text("send-hotkey","HSED")
        autoit.sleep(500)
        localizar.set_text("%")
        autoit.sleep(500)
        localizar.send_hotkey("{ENTER}")
        autoit.sleep(500)
        localizar.set_text(f"{mask_cnpj(str(emitCNPJ).zfill(14))}")
        
        # Click localizar
        ui(locator.java.maxys_GEX001.push_button_localizar_alt_l).click()
        
        autoit.sleep(600)

        # Click OK
        ui(locator.java.maxys_GEX001.push_button_ok_alt_o).click()

        print("A")
        print("A")

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
    
    if clicknium.wait_appear(locator.java.maxys_GEX001.SelecionaAmostra_OK,wait_timeout=5):
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

def _ConsultaEmbarque(contrato_saida, Cidade, UF):
     # Abrir tela GRE001
    _Executar("GPE001")

    # Digita Contrato Saida
    ui(locator.java.maxys_GPE001.text_contrato).set_text(contrato_saida)
    ui(locator.java.maxys_GPE001.text_contrato).send_hotkey("{ENTER}")

    # Click Local de Embarque
    ui(locator.java.maxys_GPE001.page_tab_local_embarque).click()

    i=1
    while True:
        variables = {"index": i}
        Cidade_UF = str(Cidade+"-"+UF)
        
        if ui(locator.java.maxys_GPE001.text_cidade_uf,variables).get_text()==Cidade_UF:
            
            # Get Codigo Embarque
            codigoEmbarque= ui(locator.java.maxys_GPE001.text_cód_local_ed_index,variables).get_text()            
            
            # Sair Tela
            ui(locator.java.maxys_GPE001.SairTela).click()
            
            return codigoEmbarque.replace(".","")
        else:
            i=i+1

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

            analise = {
                "ARDIDO" : 0,
                "AVARIADO" : 0,
                "IMPUREZA" : 0,
                "UMIDADE" : 0,
                "GRÃOS QUEBRADO" : 0,
                "GRAOS VERDES" :0,
                "ESVERDEADOS" : 0,
                "QUEIMADO" : 0,
                "MOLHADO" : 0,
                "FERMENTADO" : 0,
                "PICADO" : 0,
                "CHOCO" : 0,
                "IMATURO" : 0,
                "MOFADO": 0,
                "VOMITOXINA (DON)": 0,
                "FN":0,
                "PROTEINA":0,
                "DANIFICADO INSETOS":0,
                "BROTADOS":0,
                "PH":0
            }

            # Abrir tela GRE001
            _Executar("GRE001")

            # Processa GRE001 
            _ProcessaGRE001(str(row['chave_de_acesso_nf_compra'].replace(" ","")).zfill(14),str(row['numero_do_contrato']),row['nome_do_motorista'],row['placa'].replace("-",""),row["uf"],analise,row["transgenia"],row["cnpj_da_transportadora"],row["vNF"])
            
            current_date = datetime.datetime.now()
            dataLog = current_date.strftime("%d/%m/%Y %H:%M:%S")
                
            # Escreve mensagem e data log
            df.loc[index, 'Mensagem_Entrada'] = "Nota de Entrada Concluida"
            df.loc[index, 'DataLog_Entrada'] = dataLog
            
            # Processa GEX001
            NrNFE=_ProcessaGEX001(row['contrato_de_venda'],str(row['clifor_transportadora']),row["transgenia"],row["emitCNPJ"])

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