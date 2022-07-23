# -*- coding: utf-8 -*-
"""
Created on Thu Nov  4 10:57:46 2021

@author: I7770871
"""
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from datetime import datetime, timedelta
import pandas as pd


"""
CRIANDO VARIÁVEIS DATA EM VÁRIOS FORMATOS P CONTROLE
"""




vardata = datetime.today().strftime('%Y%m%d')
i = [0,1,2] 

for dia in i:
    
    ini = datetime.today() - timedelta(dia)
    ini = ini.strftime('%d/%m/%Y')
    
    """
    ABRINDO O SITE
    """
    
    driver = webdriver.Chrome(executable_path=r"C:\CHROME DRIVER\chromedriver.exe")
    url = f'https://p.dfe.mastersaf.com.br/mvc/receptorNfe/lista?dateIni={ini}%2000:00&dateFim={ini}%2023:59'
    driver.get(url)
    
    """
    FAZENDO O PROCESSO DE LOGAR
    """
    
    driver.find_element(By.CLASS_NAME, "logarSaml").click()
    Dominio = driver.find_element(By.ID, "dominio")
    Dominio.send_keys("saint-gobain")
    driver.find_element(By.ID, "enterDominio").click()
    time.sleep(8)
    driver.find_element(By.ID, "login.submit.id").click()
    time.sleep(8)
    
    
    """
    FAZENDO O PROCESSO DE FILTRAR POR FILIAL
    """
    
    driver.find_element(By.CLASS_NAME, "fs-arrow").click()
    driver.find_element_by_xpath("//div[@data-value='4028481a6e474b81016e4ae607670000']").click()
    time.sleep(5)
    
    """
    FAZENDO O PROCESSO DE ABRIR POR 500 NFes
    """
    
    driver.find_element(By.CLASS_NAME, "ui-pg-selbox").click()
    driver.find_element_by_xpath("//option[@value='500']").click()
    time.sleep(2)
    
    """
    FILTRAGEM POR DATA
    """
    
    
    dataini = driver.find_element_by_name("consultaDataInicial")
    dataini.clear()
    dataini.send_keys(f"{ini}")
    
    datafim = driver.find_element_by_name("consultaDataFinal")
    datafim.clear()
    datafim.send_keys(f"{ini}")
    time.sleep(3)
    
    driver.find_element(By.ID, "listagem_atualiza").click()
    
    """
    FAZENDO O PROCESSO DE CHECKBOX
    """
    
    driver.find_element(By.ID, "listagem_checkBox").click()
    time.sleep(3)
    driver.find_element(By.CLASS_NAME, "botao-img-pp-download-XLS").click()
    time.sleep(2)
    
    """
    BAIXANDO O ARQUIVO
    """
    driver.find_element(By.ID, "downloadEmMassaXls").click()
    time.sleep(3)
    
    """
    PEGANDO O ARQUIVO NO PANDAS E SALVANDO NO DIRETÓRIO
    """
    df = pd.read_excel(fr"C:\Users\I7770871\Downloads\{vardata}_RECEBIMENTO_NFE.xlsx")
    df.to_excel(f'//Qsbrprd/qliksense/BRASIL/OUTRAS FONTES/PCR_INDUSTRIAL/Finance/CONCILIAÇÃO SIR/Automatização/PPL_SCS_MASTERSAF/{vardata}_RECEBIMENTO_NFE.xlsx', index = False)
    
    """
    CARREGANDO A BASE TOTAL - PPL SCS
    """
    dfbase = pd.read_excel(r"\\Qsbrprd\qliksense\BRASIL\OUTRAS FONTES\PCR_INDUSTRIAL\Finance\CONCILIAÇÃO SIR\Automatização\PPL_SCS_MASTERSAF\BD\BD_PPL_SCS.xlsx")
    
    """
    CARREGANDO A CARGA INCREMENTAL
    """
    dfcarga = pd.read_excel(fr"\\Qsbrprd\qliksense\BRASIL\OUTRAS FONTES\PCR_INDUSTRIAL\Finance\CONCILIAÇÃO SIR\Automatização\PPL_SCS_MASTERSAF\{vardata}_RECEBIMENTO_NFE.xlsx")
    
    """
    CONCATENANDO BASE COM CARGA INCREMENTAL, REMOVENDO VALORES DUPLICADOS (PARA MANTER LOG DE NFs CANCELADAS E ATIVAS COM RANGE DE 48h) e sobrepõe a BASE TOTAL
    """
    dffinal = pd.concat([dfbase, dfcarga])
    dffinal = dffinal.drop_duplicates()
    dffinal.to_excel(r"\\Qsbrprd\qliksense\BRASIL\OUTRAS FONTES\PCR_INDUSTRIAL\Finance\CONCILIAÇÃO SIR\Automatização\PPL_SCS_MASTERSAF\BD\BD_PPL_SCS.xlsx", index = False)
    
    """
    MERGE COM BASE MÃE USADA NO QLIKSENSE, RETIRADA DE LINHAS DUPLICADAS E SALVAMENTO AUTOMÁTICO
    """
    
    dfmae = pd.read_excel(r"\\Qsbrprd\qliksense\BRASIL\OUTRAS FONTES\PCR_INDUSTRIAL\Finance\CONCILIAÇÃO SIR\Automatização\MASTERSAF_TODAS_AS_FILIAIS_BD_TOTAL\BD_MASTERSAF.xlsx")
    
    dfmae = pd.concat([dfmae, dffinal])
    dfmae = dfmae.drop_duplicates()
    
    dfmae.to_excel(r"\\Qsbrprd\qliksense\BRASIL\OUTRAS FONTES\PCR_INDUSTRIAL\Finance\CONCILIAÇÃO SIR\Automatização\MASTERSAF_TODAS_AS_FILIAIS_BD_TOTAL\BD_MASTERSAF.xlsx", index = False)
    
    """
    EXCLUSÃO DE BASE DE CARGA INCREMENTAL NO USER PARA LAÇO PARA i DIAS
    """
    os.remove(fr"C:\Users\I7770871\Downloads\{vardata}_RECEBIMENTO_NFE.xlsx")
    
    
    
    
