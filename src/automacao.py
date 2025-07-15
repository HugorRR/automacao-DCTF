import os
import logging
import time
from pathlib import Path
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import pyautogui
import pandas as pd
from .utils import reconhecimento, clique, clique2, lentidao

# Função para renomear arquivo

def renomear_arquivo_recente(codigo, competencia, pasta_competencia):
    try:
        arquivos = list(Path(pasta_competencia).glob("*"))
        arquivo_recente = max(arquivos, key=os.path.getctime)
        novo_nome = Path(pasta_competencia) / f"{codigo} DARFWEB {competencia}.pdf"
        arquivo_recente.rename(novo_nome)
        logging.info(f"Arquivo renomeado para: {novo_nome}")
    except Exception as e:
        logging.error(f"Erro ao renomear o arquivo: {e}")

# Configurar driver Chrome

def configurar_driver(perfil_path, pasta_competencia):
    options = uc.ChromeOptions()
    options.add_argument("--user-data-dir=" + str(perfil_path))
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-infobars")
    options.add_experimental_option("prefs", {
        "download.default_directory": str(pasta_competencia),
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
    })
    driver = uc.Chrome(options=options)
    driver.get('https://cav.receita.fazenda.gov.br/autenticacao/login')
    driver.maximize_window()
    driver.implicitly_wait(10)
    logging.info("Driver configurado com sucesso.")
    return driver

# Login manual

def login(driver):
    try:
        logging.info("Iniciando processo de login manual.")
        print("==== ATENÇÃO ====")
        print("O login precisa ser realizado manualmente.")
        print("Por favor, faça o login e pressione ENTER quando estiver na tela principal.")
        input("Pressione ENTER quando o login estiver concluído...")
        logging.info("Usuário confirmou que o login foi concluído.")
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]')))
            logging.info("Login realizado com sucesso. Página principal identificada.")
            return True
        except:
            logging.warning("Não foi possível confirmar se o login foi bem-sucedido.")
            input("Confirme que está na página principal e pressione ENTER...")
            return True
    except Exception as e:
        logging.error(f"Erro durante o processo de login: {e}")
        print("Ocorreu um erro durante o login. Verifique o arquivo de log para mais detalhes.")
        print("Tente novamente manualmente.")
        input("Pressione ENTER quando o login estiver concluído...")
        return True

# Navegação

def navegacao(driver, IMAGEM_DIR, data_inicial, data_final):
    try:
        logging.info("Iniciando navegação no sistema.")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))).click() # Botão Home 
        lentidao()
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//li[@id="btn214"]'))).click() # Botão Declarações e Demonstrativos
        lentidao()
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="containerServicos214"]/div[2]/ul/li[1]/a'))).click() # Assinar e transmitir DCTF
        lentidao()
        lentidao()
        bt_captcha1 = os.path.join(IMAGEM_DIR, 'bt_captcha.png') # Procura o Quadrado "Sou Humano"
        lentidao()
        clique([bt_captcha1], 30, confidence=0.8) # Clica no Quadrado "Sou Humano"
        lentidao()
        bt_prosseguir1 = os.path.join(IMAGEM_DIR, 'bt_prosseguir.png') # Procura o Botão Prosseguir
        lentidao()
        clique([bt_prosseguir1], 30, confidence=0.8) # Clica no Botão Prosseguir
        lentidao()
        iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
        driver.switch_to.frame(iframe)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_chkListarOutorgantes"]'))).click() # Check box Sou procurador
        time.sleep(10)
        data_inicio = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtDataInicio"]'))) # Campo de data inicial
        data_inicio.clear() # Limpa o campo de data inicial
        data_inicio.send_keys(data_inicial) # Escreve a data inicial
        lentidao()
        data_fim = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtDataFinal"]'))) # Campo de data final
        data_fim.clear() # Limpa o campo de data final
        data_fim.send_keys(data_final) # Escreve a data final
        lentidao()
        driver.switch_to.default_content()

    except Exception as e:
        logging.error(f'Erro de navegacao: {e}')
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))).click()
        except:
            logging.error("Não foi possível retornar à página inicial após erro de navegação.")
            print("Erro na navegação. Tente navegar manualmente para a página correta.")
            input("Pressione ENTER quando estiver pronto para continuar...")

# Transmissão

def transmissao(cnpjs, codigos, df, driver, IMAGEM_DIR, competencia, pasta_competencia):
    for cnpj, codigo in zip(cnpjs, codigos):
        status = df.loc[df['CNPJ'] == cnpj, 'STATUS'].values
        if len(status) > 0 and ('Guia baixada' in str(status[0]) or pd.isna(status[0])):
            logging.info(f"CNPJ {cnpj} já processado ou sem status. Pulando...")
            continue
        tentativas = 3
        while tentativas > 0:
            try:
                logging.info(f'Iniciando a transmissão da empresa: {cnpj}')

                iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
                driver.switch_to.frame(iframe)

                bt_ortogante = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/button')))
                bt_ortogante.click() # Campo ortogante onde tem entrada do cnpj do cliente


                bt_nenhum = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/div[2]/div/button[2]')))
                bt_nenhum.click() # remover clientes selecionados antes de inserir o novo cnpj para pesquisar

                campo_cnpj = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/div[1]/input')))
                campo_cnpj.send_keys(cnpj)
                lentidao()

                selecionar_cnpj = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/ul')))
                selecionar_cnpj.click()

                bt_pesquisar = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_btnFiltar"]')))
                bt_pesquisar.click()

                driver.switch_to.default_content()
               
                erro_sem_declaracoes = os.path.join(IMAGEM_DIR, 'rc_sem_declaracoes.png')
                time.sleep(3)
                try:
                    reconhecimento([erro_sem_declaracoes], 30, confidence=0.8)
                    logging.info("Nenhuma declaração encontrada.")
                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Nenhuma declaração encontrada'
                    df.to_excel('Clientes.xlsx', index=False)
                    break
                except Exception as e:
                    logging.info('Declaracao encontrada')
                    iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
                    driver.switch_to.frame(iframe)
                    bt_visualizar = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs_ctl02_lbkVisualizarDctf"]')))
                    bt_visualizar.click() # Clica no botão visualizar
                    lentidao()
                    driver.switch_to.default_content()
                try:
                    lentidao()
                    iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
                    driver.switch_to.frame(iframe)
                    bt_emitir_darf = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LinkEmitirDARFIntegral"]')))
                    bt_emitir_darf.click() # Clica no botão emitir DARF
                    driver.switch_to.default_content()
                    lentidao()
                    time.sleep(5)
                    renomear_arquivo_recente(codigo, competencia, pasta_competencia)
                    logging.info(f"Download successful for {cnpj}")
                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'
                    df.to_excel('Clientes.xlsx', index=False)
                    iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
                    driver.switch_to.frame(iframe)
                    bt_ok = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//button[text()='OK']")))
                    bt_ok.click()
                    bt_voltar = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="cphConteudo_LinkRetornar"]')))
                    bt_voltar.click()
                    driver.switch_to.default_content()
                    bt_captcha2 = os.path.join(IMAGEM_DIR, 'bt_captcha.png')
                    lentidao()
                    clique([bt_captcha2], 15, confidence=0.8)
                    lentidao()
                    bt_prossegui2 = os.path.join(IMAGEM_DIR, 'bt_prosseguir.png')
                    lentidao()
                    clique([bt_prossegui2], 15, confidence=0.8)
                    lentidao()
                    break
                except Exception as e:
                    logging.error("Nenhuma declaração encontrada.")
                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Possivel erro no e-CAC, verificar cliente manualmente'
                    df.to_excel('Clientes.xlsx', index=False)
                bt_voltar = os.path.join(IMAGEM_DIR, 'bt_voltar.png')
                lentidao()
                clique([bt_voltar], 30, confidence=0.8)
                lentidao()
                bt_captcha2 = os.path.join(IMAGEM_DIR, 'bt_captcha.png')
                lentidao()
                clique([bt_captcha2], 30, confidence=0.8)
                lentidao()
                bt_prossegui2 = os.path.join(IMAGEM_DIR, 'bt_prosseguir.png')
                lentidao()
                clique([bt_prossegui2], 30, confidence=0.8)
                lentidao()
                df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'
                df.to_excel('Clientes.xlsx', index=False)
                break
            except Exception as e:
                logging.error(f"Erro no processamento do cliente {cnpj}: {e}")
                df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Erro no download'
                df.to_excel('Clientes.xlsx', index=False)
                tentativas -= 1
                if tentativas > 0:
                    logging.info(f"Tentando novamente ({tentativas} tentativas restantes)")
                    time.sleep(3)
                    navegacao(driver, IMAGEM_DIR, '', '')
                else:
                    logging.error(f"Falha após {3} tentativas para o cliente {cnpj}") 