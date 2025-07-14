import pyautogui
import time
import pandas as pd
import os
from pathlib import Path
import logging
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

logging.basicConfig(filename='AUTOMACAO-DCTF.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

data_inicial = ('01062025')
data_final = ('30062025')
competencia = ('06 2025')   
pyautogui.PAUSE = 1

PASTA_COMPETENCIA = Path(__file__).parent / "Competencias executadas"
PASTA_DOWNLOAD = PASTA_COMPETENCIA / f"{competencia}"
IMAGEM_DIR = Path(__file__).parent / "img"
PLANILHA = Path(__file__).parent / "Clientes.xlsx"

PASTA_DOWNLOAD2 = r'I:\Drives compartilhados\BANCO DE DADOS T.I\Automacoes\automacao-DCTF\Competencias executadas\01 2025'

pasta_competencia = os.path.join(PASTA_COMPETENCIA, competencia)
if not os.path.exists(pasta_competencia):
    os.mkdir(pasta_competencia)

# Diretório para armazenar o perfil do Chrome
CACHE = Path(__file__).parent / 'perfil-path'
if not CACHE.exists():
    CACHE.mkdir()

def renomear_arquivo_recente(codigo, competencia):
    try:
        arquivos = list(PASTA_COMPETENCIA.glob("*"))
        arquivo_recente = max(arquivos, key=os.path.getctime)
        novo_nome = PASTA_COMPETENCIA / f"{codigo} DARFWEB {competencia}.pdf"
        arquivo_recente.rename(novo_nome)
        logging.info(f"Arquivo renomeado para: {novo_nome}")
    except Exception as e:
        logging.error(f"Erro ao renomear o arquivo: {e}")

def reconhecimento(imagens_referencia, tempo_limite, confidence=1.0):
    """Espera uma das imagens de referência aparecer na tela dentro de um tempo limite."""
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicao = pyautogui.locateCenterOnScreen(imagem_referencia, confidence=confidence)
            if posicao is not None:
                logging.info(f"Imagem encontrada: {imagem_referencia}")
                # Move o mouse até a posição da imagem de forma gradual, simulando comportamento humano
                pyautogui.moveTo(posicao.x, posicao.y, duration=0.5)
                return True
            logging.info(f"Imagem não encontrada: {imagem_referencia}")
        # Adiciona tempo de espera aleatório para simular comportamento humano
        tempo_espera = 1.0 + (time.time() % 0.5)  # Entre 1.0 e 1.5 segundos
        time.sleep(tempo_espera)
    logging.info("Nenhuma imagem encontrada dentro do tempo limite.")
    return False

def clique(imagens_referencia, tempo_limite, confidence=1.0):
    """Tenta clicar em uma das imagens de referência na tela dentro de um tempo limite."""
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicao = pyautogui.locateCenterOnScreen(imagem_referencia, confidence=confidence)
            if posicao is not None:
                # Move o mouse até a posição da imagem de forma gradual e natural
                pyautogui.moveTo(posicao.x, posicao.y, duration=0.6)
                # Pequena pausa antes do clique para simular comportamento humano
                time.sleep(0.1 + (time.time() % 0.3))  # Entre 0.1 e 0.4 segundos
                pyautogui.click()
                logging.info(f"Clique na imagem: {imagem_referencia}")
                # Pausa após o clique para simular comportamento humano
                time.sleep(0.2 + (time.time() % 0.3))  # Entre 0.2 e 0.5 segundos
                return True
            logging.info(f"Imagem não reconhecida: {imagem_referencia}")
        # Tempo de espera variável entre tentativas
        tempo_espera = 1.0 + (time.time() % 0.8)  # Entre 1.0 e 1.8 segundos
        time.sleep(tempo_espera)
    return False

def clique2(imagens_referencia, tempo_limite, confidence=1.0, ocorrencia=1):
    """Tenta clicar em uma das imagens de referência na tela dentro de um tempo limite."""
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicoes = list(pyautogui.locateAllOnScreen(imagem_referencia, confidence=confidence))
            logging.info(f"Número de ocorrências da imagem '{imagem_referencia}' encontradas na tela: {len(posicoes)}")
            if len(posicoes) >= ocorrencia:
                posicao_target = posicoes[ocorrencia-1]
                centro_x = posicao_target.left + posicao_target.width / 2
                centro_y = posicao_target.top + posicao_target.height / 2
                
                # Movimento de mouse mais natural com pequena variação na velocidade
                pyautogui.moveTo(centro_x, centro_y, duration=0.4 + (time.time() % 0.3))  # Entre 0.4 e 0.7 segundos
                
                # Pequena pausa antes do clique, simulando hesitação humana
                time.sleep(0.1 + (time.time() % 0.2))  # Entre 0.1 e 0.3 segundos
                
                pyautogui.click()
                logging.info(f"Clique na imagem: {imagem_referencia} (ocorrência {ocorrencia})")
                
                # Pausa após o clique
                time.sleep(0.2 + (time.time() % 0.3))  # Entre 0.2 e 0.5 segundos
                return True
        logging.info(f"Imagem não reconhecida: {imagem_referencia}")
        # Tempo de espera variável entre tentativas
        tempo_espera = 1.0 + (time.time() % 0.5)  # Entre 1.0 e 1.5 segundos
        time.sleep(tempo_espera)
    return False

def ler_planilha():
    df = pd.read_excel(PLANILHA)
    cnpjs = df['CNPJ'].astype(str).tolist()
    codigos = df['COD'].astype(str).tolist()
    df['STATUS'] = ''
    return cnpjs, codigos, df

def lentidao():
    """Função para adicionar uma pequena pausa com tempo aleatório para simular comportamento humano"""
    time.sleep(1.5)

def configurar_driver():
    """Configura o driver do Chrome com opções personalizadas para evitar detecção de automação."""
    perfil_path = str(CACHE)
    options = uc.ChromeOptions()
    # Configurações para evitar detecção de automação
    options.add_argument("--user-data-dir=" + perfil_path)
    options.add_argument("--disable-blink-features=AutomationControlled")  # Importante para evitar detecção
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-infobars")
    # Configurações para download automático
    options.add_experimental_option("prefs", {
        "download.default_directory": str(PASTA_COMPETENCIA),
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
    })
    
    driver = uc.Chrome(options=options)
    driver.get('https://cav.receita.fazenda.gov.br/autenticacao/login')
    driver.maximize_window()
    driver.implicitly_wait(10)
    logging.info("Driver configurado com sucesso.")
    return driver

def login(driver):
    """Realiza login manual no sistema, aguardando intervenção do usuário."""
    try:
        logging.info("Iniciando processo de login manual.")
        print("==== ATENÇÃO ====")
        print("O login precisa ser realizado manualmente.")
        print("Por favor, faça o login e pressione ENTER quando estiver na tela principal.")
        try:
            pass
        except:
            logging.warning("Não foi possível clicar no botão de acesso gov.br automaticamente.")
        
        # Aguarda intervenção manual do usuário
        input("Pressione ENTER quando o login estiver concluído...")
        logging.info("Usuário confirmou que o login foi concluído.")
        
        # Verifica se está na página principal
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

def navegacao(driver):
    try:
        logging.info("Iniciando navegação no sistema.")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))).click()
        lentidao()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//li[@id="btn214"]'))).click()
        lentidao()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="containerServicos214"]/div[2]/ul/li[1]/a'))).click()
        lentidao()

        bt_captcha1 = os.path.join(IMAGEM_DIR, 'bt_captcha.png')
        lentidao()
        clique([bt_captcha1], 30, confidence=0.8)
        lentidao()

        bt_prosseguir1 = os.path.join(IMAGEM_DIR, 'bt_prosseguir.png')
        lentidao()
        clique([bt_prosseguir1], 30, confidence=0.8)
        lentidao()

        bt_souprocurador = os.path.join(IMAGEM_DIR, 'bt_souprocurador.png')
        lentidao()
        clique([bt_souprocurador], 30, confidence=0.8)
        lentidao()

        bt_calendario = os.path.join(IMAGEM_DIR, 'bt_calendario.png')
        lentidao()
        clique([bt_calendario], 30, confidence=0.8)
        lentidao()
        
        # Apaga o conteúdo do campo com comportamento mais humanizado
        pyautogui.hotkey('ctrl', 'a')
        lentidao()
        pyautogui.press('delete')
        lentidao()

        # Digita o valor com velocidade mais humanizada
        for caractere in data_inicial:
            pyautogui.write(caractere)
            time.sleep(0.1)  # Pequena pausa entre cada caractere
        lentidao()

        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('delete')
        lentidao()
        
        # Digita o valor com velocidade mais humanizada
        for caractere in data_final:
            pyautogui.write(caractere)
            time.sleep(0.1)  # Pequena pausa entre cada caractere
        lentidao()
        

    except Exception as e:
        logging.error(f'Erro de navegacao: {e}')
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))).click()
        except:
            logging.error("Não foi possível retornar à página inicial após erro de navegação.")
            print("Erro na navegação. Tente navegar manualmente para a página correta.")
            input("Pressione ENTER quando estiver pronto para continuar...")

def transmissao(cnpjs, codigos, df, driver):
    for cnpj, codigo in zip(cnpjs, codigos):

        # Verifica se o CNPJ já foi processado
        status = df.loc[df['CNPJ'] == cnpj, 'STATUS'].values
        
        # Verifica se o status contém 'Guia baixada' ou está vazio
        if len(status) > 0 and ('Guia baixada' in str(status[0]) or pd.isna(status[0])):
            logging.info(f"CNPJ {cnpj} já processado ou sem status. Pulando...")
            continue

        tentativas = 3
        while tentativas > 0:
            try:
                logging.info(f'Iniciando a transmissão da empresa: {cnpj}')

                bt_seta = os.path.join(IMAGEM_DIR, 'bt_seta.png')
                lentidao()
                clique2([bt_seta], 30, confidence=0.8)
                lentidao()

                bt_cancelar = os.path.join(IMAGEM_DIR, 'bt_cancelar.png')
                lentidao()
                clique([bt_cancelar], 30, confidence=0.8)
                lentidao()

                # Digita o CNPJ com velocidade mais humanizada
                for caractere in cnpj:
                    pyautogui.write(caractere)
                    time.sleep(0.1)  # Pequena pausa entre cada caractere
                lentidao()

                pyautogui.press('enter')
                lentidao()

                pyautogui.click(x=6, y=728)
                lentidao()

                bt_pesquisar = os.path.join(IMAGEM_DIR, 'bt_pesquisar.png')
                lentidao()
                clique([bt_pesquisar], 30, confidence=0.8)
                lentidao()

                erro_sem_declaracoes = os.path.join(IMAGEM_DIR, 'rc_sem_declaracoes.png')
                # Tempo de espera mais humanizado
                time.sleep(3)

                try:
                    reconhecimento([erro_sem_declaracoes], 30, confidence=0.8)

                    logging.info("Nenhuma declaração encontrada.")

                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Nenhuma declaração encontrada'

                    df.to_excel('Clientes.xlsx', index=False)

                    break

                except Exception as e:
                    logging.info('Declaracao encontrada')

                    bt_visualizar = os.path.join(IMAGEM_DIR, 'bt_visualizar.png')
                    lentidao()
                    clique2([bt_visualizar], 30, confidence=0.8)
                    lentidao()


                try:
                    lentidao()
                    bt_emitir_darf = os.path.join(IMAGEM_DIR, 'bt_emitir_darf.png')
                    lentidao()
                    clique([bt_emitir_darf], 30, confidence=0.8)
                    lentidao()

                    # Aguarda o download com tempo variável
                    time.sleep(5)

                    renomear_arquivo_recente(codigo, competencia)

                    logging.info(f"Download successful for {cnpj}")
                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'
                    df.to_excel('Clientes.xlsx', index=False)

                    bt_ok = os.path.join(IMAGEM_DIR, 'bt_okk.png')
                    lentidao()
                    clique([bt_ok], 30, confidence=0.8)
                    lentidao()

                    bt_voltar = os.path.join(IMAGEM_DIR, 'bt_voltar.png')
                    lentidao()
                    clique([bt_voltar], 15, confidence=0.8)
                    lentidao()

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
                    # Adicionando um tempo de pausa aleatório entre tentativas
                    time.sleep(3)
                    navegacao(driver)
                else:
                    logging.error(f"Falha após {3} tentativas para o cliente {cnpj}")


def main():
    tentativas_gerais = 3
    
    while tentativas_gerais > 0:
        try:
            logging.info("Iniciando automação DCTF")
            print("="*50)
            print("AUTOMAÇÃO DCTF - INICIANDO")
            print("="*50)
            
            # Criar pasta para competência se não existir
            if not os.path.exists(pasta_competencia):
                os.mkdir(pasta_competencia)
                logging.info(f"Pasta de competência criada: {pasta_competencia}")
            
            # Configurar driver com proteção anti-detecção
            driver = configurar_driver()
            
            # Carregar dados da planilha
            cnpjs, codigos, df = ler_planilha()
            
            # Realizar login manual
            print("Aguardando login manual...")
            login(driver)
            
            # Aguardar confirmação do usuário antes de prosseguir
            input("Login concluído? Pressione ENTER para continuar com a navegação e processamento...")
            
            # Iniciar navegação no sistema
            navegacao(driver)
            
            # Processar cada CNPJ da lista
            transmissao(cnpjs, codigos, df, driver)
            
            # Salvar planilha final
            df.to_excel('Clientes.xlsx', index=False)
            
            print("="*50)
            print("AUTOMAÇÃO CONCLUÍDA COM SUCESSO!")
            print("="*50)
            
            logging.info("Automação concluída com sucesso")
            
            # Fechar o driver
            try:
                driver.quit()
            except:
                pass
                
            break  # Sai do loop de tentativas se tudo correr bem
            
        except Exception as e:
            tentativas_gerais -= 1
            logging.error(f"Erro geral na execução: {e}")
            print(f"Ocorreu um erro: {e}")
            
            if tentativas_gerais > 0:
                print(f"Tentando novamente. Restam {tentativas_gerais} tentativas.")
                
                # Fechar driver se estiver aberto
                try:
                    driver.quit()
                except:
                    pass
                    
                # Pausa antes de nova tentativa
                tempo_espera = 5
                print(f"Aguardando {tempo_espera} segundos antes da próxima tentativa...")
                time.sleep(tempo_espera)
            else:
                print("Número máximo de tentativas excedido. Encerrando programa.")
                logging.error("Número máximo de tentativas excedido. Programa finalizado com erro.")
                
    # Salvamento final da planilha independente do resultado
    try:
        df.to_excel('Clientes.xlsx', index=False)
        print("Planilha salva com status final dos processamentos.")
    except:
        print("Não foi possível salvar a planilha final.")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\nPrograma interrompido pelo usuário.")
        logging.info("Programa interrompido pelo usuário (KeyboardInterrupt)")
    except Exception as e:
        print(f"\nErro não tratado: {e}")
        logging.critical(f"Erro crítico não tratado: {e}")
    finally:
        print("\nFinalizando...")
        logging.info("Programa finalizado")
