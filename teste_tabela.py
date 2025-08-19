#! C:\python\venv\Scripts\python.exe

from clicknium.common.enums import *
import pandas as pd
from clicknium import clicknium as cc, locator, ui
from clicknium import clicknium, ui, locator
from clicknium.common.enums import *
import subprocess
import os
import datetime
import pandas as pd
import win32com.client as client
import sys
from dotenv import load_dotenv

# Carregar variáveis do arquivo .env
load_dotenv()

# Capturar a licença do Clicknium do .env
clicknium_license = os.getenv("CLICKNIUM_LICENSE")
clicknium.config.set_license(clicknium_license)
autoit = client.Dispatch("AutoItX3.Control")
wshell = client.Dispatch("WScript.Shell")

def _Executar(tela):
    
    clicknium.wait_appear(locator.java.maxys.Executar_text)
    ui(locator.java.maxys.Executar_text).clear_text("send-hotkey")
    ui(locator.java.maxys.Executar_text).set_text(tela)
    ui(locator.java.maxys.Executar_text).send_hotkey("{ENTER}")


# Capturar os dados
ui(locator.java.maxys_GEX004.text_cell_table, variables).get_text().strip()