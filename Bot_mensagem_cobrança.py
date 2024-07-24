from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import win32com.client as client
import pandas as pd
import datetime as dt

driver = webdriver.Chrome()

driver.get("https://web.whatsapp.com/")
    
def contatos_excel():
    login = input("Confima o login no Whatsapp com um Sim: ")
    tabela = pd.read_excel('numeros_whastzapp.xlsx')
    print(tabela)
    tabela.info()

    hoje = dt.datetime.now()
    print(hoje)

    tabela_devedores = tabela.loc[tabela['Status']=='Em aberto']
    print(tabela_devedores)
    tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento']<hoje]
    print(tabela_devedores)

    sleep(5)
 
    dados= tabela_devedores[['Nome','Valor','Data Prevista para pagamento','Status', 'Contato']].values.tolist()

    for dado in dados:
        nome = dado[0]
        contato = dado[4]
        valor = dado[1]
        mensagem = f"""

        Olá {nome},
        %0A%0A
        Espero que esteja tudo bem. Estamos entrando em contato para informar que o boleto referente ao valor de R${valor:.2f} esta pendente e estaremos enviado o boleto em breve aqui pelo Whatzapp e também para o seu e-mail cadastrado.
        %0A%0A
        Fique à vontade para entrar em contato conosco se precisar de qualquer informação adicional.
        %0A%0A
        Atenciosamente,
        Mw Tech

        """

        # Trocar para a nova aba
        driver.execute_script("window.open('about:blank', 'tab2');")
        driver.switch_to.window("tab2")
        driver.get(f'https://web.whatsapp.com/send/?phone={contato}&text={mensagem}&type=phone_number&app_absent=0')
        print("Nova aba:", driver.current_url)


        sleep(10)
        
        driver.find_element(by=By.XPATH, value='//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
        print(f"Mensagem enviada para {nome}: ")

        sleep(20)



sleep(20)
    

contatos_excel()    

driver.quit()