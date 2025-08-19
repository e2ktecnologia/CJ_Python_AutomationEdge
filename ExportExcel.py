#! C:\python\venv\Scripts\python.exe

from clicknium.common.enums import *
import pandas as pd
from clicknium import clicknium as cc, locator, ui

# Configurar licença
cc.config.set_license('idC+m5GXnIGXlqad0MjQwMTLxtDe0KGRmpefk6SXgNDI0MDQ3tChmYfQyNCil4CBnZyTntKigJ2Ul4GBm52ck57Q3tCkk56blpOGl7SAnZ/QyNDAwsDH38LD38PFpsPKyMbGyMPL3MDAxcXHx8eo0N7QpJOem5aThpemndDI0MDCwMTfwsPfw8Wmw8rIxsbIw8vcwMDFxcfEwajQ3tC0l5OGh4CXgdDIqYnQvJOfl9DI0L+Tir6dkZOGnYC+m5+bhpfQ3tCkk56Hl9DI0MPKxsbExcbGwsXBxcLLx8fDxMPH0I+vjw==.CkD4OnvRakP0KT1zN1O8eLCEy2/iOVBryGXXcBUXxGOkOCcjLaWvQ/6BgNJVh5cdVJjJaj8t6T9NV9w5N5DZf/GhzfgMmsR1CzNdPqcdIkbCD/PFd2DboPuIHrHxY3YJ2zsc6nTzV6tW90q8fJ0KNH/zLnWVi+J6Gpc4xY9vOUA=')

# Capturar os dados
cc.wait_appear(locator.java.maxys_TAF117.Resultado_do_processamento)

def ExportarExcel():
    
    # Check Gerar EXECEL
    cc.wait_appear(locator.java.maxys_TAF117.Resultado_do_processamento)
    cc.wait_appear(locator.java.maxys_ExportExcel.Export_check_box_gerar_excel)
    ui(locator.java.maxys_ExportExcel.Export_check_box_gerar_excel).set_checkbox("uncheck")
    ui(locator.java.maxys_ExportExcel.Export_check_box_gerar_excel).set_checkbox("check")

    # UnCheck 
    cc.wait_appear(locator.java.maxys_ExportExcel.Export_check_box_visualizar_planilha)
    ui(locator.java.maxys_ExportExcel.Export_check_box_visualizar_planilha).set_checkbox("uncheck")

    # Check Salvar Automaticamente
    ui(locator.java.maxys_ExportExcel.Export_check_box_salvar_automaticamente).set_checkbox("uncheck")
    ui(locator.java.maxys_ExportExcel.Export_check_box_salvar_automaticamente).set_checkbox("check")

    # Click escolher pasta
    ui(locator.java.maxys_ExportExcel.Export_choose_push_button).click()

    # Digita Folder
    ui(locator.java.maxys_ExportExcel.Export_select_edit_folder).send_hotkey("C:\\temp\\{ESC}")
    
    # Click Select Folder
    ui(locator.java.maxys_ExportExcel.Export_button_select_folder).click()
    
    # Click Gerar
    cc.wait_appear(locator.java.maxys_ExportExcel.Export_push_button_gerar_alt_g)
    ui(locator.java.maxys_ExportExcel.Export_push_button_gerar_alt_g).click()

    # Esperar Popup sucesso
    if cc.wait_appear(locator.java.maxys_VFS014.Observacao_Sucesso,wait_timeout=5):
        ui(locator.java.maxys_VFS014.Observacao_OK).click()
    
    # Fechar PopUp
    cc.wait_appear(locator.java.maxys_ExportExcel.Export_push_button_fechar_alt_f)
    ui(locator.java.maxys_ExportExcel.Export_push_button_fechar_alt_f).click()
    
ExportarExcel()