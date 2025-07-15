# Automação DCTF

Automação para download e organização de DARFs via e-CAC, utilizando Selenium, undetected-chromedriver, PyAutoGUI e manipulação de planilhas Excel.

## EM PROCESSO DE DESENVOLVIMENTO

## Funcionalidades

- Automatiza o acesso ao portal e-CAC da Receita Federal.
- Realiza login manual com proteção anti-bloqueio.
- Pesquisa, baixa e renomeia arquivos DARF conforme planilha de clientes.
- Atualiza o status de cada cliente em uma planilha Excel.
- Utiliza reconhecimento de imagens para interagir com elementos visuais do sistema.

## Pré-requisitos

- Python 3.8 ou superior (recomendado Python 3.10+)
- Google Chrome instalado
- [ChromeDriver compatível](https://chromedriver.chromium.org/downloads) (ou utilize undetected-chromedriver)
- Pacotes Python:
  - pandas
  - openpyxl
  - pyautogui
  - selenium
  - undetected-chromedriver
  - pathlib

## Instalação

1. Clone este repositório:
   ```sh
   git clone https://github.com/seu-usuario/automacao-DCTF.git
   cd automacao-DCTF
   ```

2. Instale as dependências:
   ```sh
   pip install -r requirements.txt
   ```
   > Se não existir um arquivo `requirements.txt`, instale manualmente:
   ```sh
   pip install pandas openpyxl pyautogui selenium undetected-chromedriver
   ```

3. Certifique-se de que a planilha `Clientes.xlsx` está na raiz do projeto, com as colunas `CNPJ` e `COD`.

4. Ajuste o caminho das pastas e imagens conforme necessário no código.

## Como usar

1. Execute o script principal:
   ```sh
   python new.py
   ```
2. Faça o login manualmente no e-CAC quando solicitado.
3. Aguarde a automação processar todos os clientes da planilha.
4. Os arquivos baixados serão salvos na pasta `Competencias executadas/<competencia>` e a planilha será atualizada com o status de cada cliente.

## Estrutura de Pastas

- `img/` — Imagens de referência para reconhecimento visual (botões, etc).
- `Competencias executadas/` — Onde os DARFs baixados são salvos.
- `Clientes.xlsx` — Planilha de entrada e saída dos dados dos clientes.
- `perfil-path/` — Perfil do Chrome para evitar bloqueios.

## Observações

- O login no e-CAC é manual por questões de segurança.
- O script utiliza reconhecimento de imagens, então a resolução da tela e o zoom do navegador devem ser mantidos conforme configurado durante a criação das imagens.
- O log de execução é salvo em `AUTOMACAO-DCTF.log`.

## Contribuição

Pull requests são bem-vindos! Para grandes mudanças, abra uma issue primeiro para discutir o que você gostaria de modificar.

## Licença

[MIT](LICENSE) 