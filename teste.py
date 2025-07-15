import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
import time

URL = 'https://cav.receita.fazenda.gov.br/autenticacao/login'

if __name__ == '__main__':
    options = uc.ChromeOptions()
    # Você pode adicionar argumentos se quiser usar o mesmo perfil do projeto:
    # options.add_argument('--user-data-dir=perfil-path')
    driver = uc.Chrome(options=options)
    driver.get(URL)
    driver.maximize_window()
    print('Navegador aberto. Faça login manualmente e use o DevTools (F12) para inspecionar XPaths.')
    try:
        while True:
            time.sleep(8000000000)
    except KeyboardInterrupt:
        driver.quit()
        print('Navegador fechado.') 