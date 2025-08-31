from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from IPython.display import display
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import urllib
import pandas as pd
import time

# Busca na planilha os dados a serem enviados utilizando o pandas
contatos_df = pd.read_excel("Devedores.xlsx")

# Imprime na tela os dados obtidos na planilha
display(contatos_df)

# Crie o perfil no Firefox com o comando Firefox -p, ou pelo próprio navegador na aba correspondente
# Após criado o perfil no Firefox, acesse a URL about:profiles e identifique o caminho raiz do perfil, atribuindo a variável a seguir
profile_path = r'C:\Users\user\AppData\Roaming\Mozilla\Firefox\Profiles\opc2ukv0.wpp'

# Declara a variável de construção e adiciona em duas linhas a opção Profile com seu respectivo caminho
options=Options()
options.add_argument ("-profile")
options.add_argument (profile_path)

# Atribui a variável navegador a aplicação Firefox com as configurações declaradas
navegador = Firefox(options=options)

# Inicia o navegador com o link declarado, só será necessário este comando na primeira vez em que for sincronizar o WhatsApp
# navegador.get(https://web.whatsapp.com/)

# Procura o elemento "side", disponível somente quando o WhatsApp está carregado
# while len(navegador.find_elements(By.ID, "side")) < 1:
    # time.sleep(1)
    
# Declara uma variável com a função enumerate, utilizado pra criar índices
contatos=enumerate(contatos_df['Link'])

# Com o WhatsApp Web logado, o bloco de repetição a seguir enviará uma mensagem por vez, utilizando a planilha
for i, mensagem in contatos:
    nome = contatos_df.loc[i, "Nome"]
    numero = contatos_df.loc[i, "Número"]
    link = contatos_df.loc[i,"Link"]
    # texto = urllib.parse.quote(f"Fala {nome}, beleza? {mensagem}")
    # link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
    navegador.get(link)
    
    # O while se torna importante aqui para que a página seja carregada antes de seguir o próximo item do índice
    while len(navegador.find_elements(By.ID, "main")) < 1:
        time.sleep(1)
    
    # Tempo necessário pra possibilitar a próxima função
    time.sleep(3)
   
    # Realiza o clique que enviará a mensagem e após aguarda 10s para enviar a próxima
    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
    time.sleep(2)
    
    # Cria uma instância da função ActionsChains
    actions = ActionChains(navegador)
    
    # Pressiona Ctrl + Alt + Shift + E, que vai arquivar o chat no Web WhatsApp
    actions.key_down(Keys.CONTROL).key_down(Keys.ALT).key_down(Keys.SHIFT).send_keys('e').perform()

    # Libere as teclas pressionadas
    actions.key_up(Keys.CONTROL).key_up(Keys.ALT).key_up(Keys.SHIFT).perform()
    
    # Delay pra não ter problemas com o Wpp
    time.sleep (10)
