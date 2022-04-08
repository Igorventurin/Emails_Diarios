#!/usr/bin/env python
# coding: utf-8

# # ENVIO DE E-MAIL DIÁRIO
# Código criado para enviar e-mails diariamente, respeitando o limite diário de 500 e-mail do gmail para não ser cair no spam.
# #### DEV: Igor Matheus Lial Venturin

# #### README:
# Para o funcionamento do código é necessario uma planilha com a primeira célula escrita "email" e abaixo todos os e-mails que você pretende enviar a mensagem e que o e-mail seja configurado de forma a permitir o uso de aplicativos de baixa integridade, tal configuração deve ser feita entrando no gerenciamento da conta google na sessão de segurança.

# In[1]:


import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import pandas as pd
import pyautogui


# #### LENDO A PLANILHA COM O EMAILS

# In[2]:


emails = pd.read_excel("emailsgeral.xlsx") # <---- Coloque dentro das ""(aspas) o nome do arquivo que contém a lista de e-mails
emails.set_index('email')


# #### EXTRAINDO O TEXTO DO E-MAIL
# OBS: O texto irá exatamente da forma como foi escrita no Google Doc, com imagens, formatação e etc

# In[3]:


driver = webdriver.Chrome()
driver.get('DOC GOOGLE') #<-- Link do Doc com o texto
driver.maximize_window()
actions=ActionChains(driver)
time.sleep(15)

#Selecionando todo o texto do documento e copiando
actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
time.sleep(3)
actions.key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()


# #### CADA BLOCO POSSUI UM CONTA QUE VOCÊ DESEJA UTILIZAR PARA OS ENVIOS
# Neste código foram utilizadas apenas duas contas

# #### CONTA 1

# In[4]:


#CONTA DO EIXO UNIVERSIDADE

#ABRINDO O NAVEGADOR
driver.get('https://mail.google.com/mail/u/0/#inbox')
time.sleep(15)
driver.maximize_window()
time.sleep(5)

#FAZENDO LOGIN NO E-MAIL E ABRINDO A CAIXA PARA ESCREVER O E-MAIL
driver.find_element(By.XPATH, '/html/body/header/div/div/div/a[2]').click() #clicando pra iniciar sessão
time.sleep(10)
driver.find_element(By.XPATH, '//*[@id="identifierId"]').click() #clicando pra fazer login
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id="identifierId"]').send_keys("LOGIN") # <-- Login da conta
driver.find_element(By.XPATH, '//*[@id="identifierId"]').send_keys(Keys.ENTER)
time.sleep(10)
driver.find_element(By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input').send_keys("SENHA") # <-- Senha da conta
driver.find_element(By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input').send_keys(Keys.ENTER)
time.sleep(30)
driver.find_element(By.XPATH, '/html/body/div[7]/div[3]/div/div[2]/div[1]/div[1]/div[1]/div/div').click()
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id=":mm"]').click()
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id=":q0"]').click()
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id=":q0"]').click()
time.sleep(4)

#COLOCANDO OS E-MAILS
for i in range(0,499):
    email_bruto = emails.iloc[ i:(i+1) , 0]
    email_tratado = (email_bruto[i].split(" "))
    driver.find_element(By.XPATH, '//*[@id=":q0"]').send_keys(email_tratado,", ")
time.sleep(5)

#COLOCANDO O ASSUNTO
actions.key_down(Keys.TAB).key_up(Keys.TAB).perform()
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id=":pg"]').send_keys("PORROGAÇÃO do Edital Centelha 2!!!") # <---- ASSUNTO DO E-MAIl
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id=":ql"]/div[1]').click()
time.sleep(3)
actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
time.sleep(10)
actions.key_down(Keys.CONTROL).key_down(Keys.ENTER).key_up(Keys.CONTROL).key_up(Keys.ENTER).perform()
time.sleep(10)
driver.quit()


# #### CONTA 2

# In[5]:


#CONTA 2 DO EIXO UNIVERSIDADE

#ABRINDO O NAVEGADOR
time.sleep(5)
driver = webdriver.Chrome()
actions=ActionChains(driver)
time.sleep(5)
driver.get('https://mail.google.com/mail/u/0/#inbox')
time.sleep(5)
driver.maximize_window()
time.sleep(5)

#FAZENDO LOGIN NO E-MAIL E ABRINDO A CAIXA PARA ESCREVER O E-MAIL
driver.find_element(By.XPATH, '//*[@id="identifierId"]').click() #clicando pra fazer login
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id="identifierId"]').send_keys("LOGIN") # <-- Login da conta
driver.find_element(By.XPATH, '//*[@id="identifierId"]').send_keys(Keys.ENTER)
time.sleep(10)
driver.find_element(By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input').send_keys("SENHA") # <-- Senha da conta
driver.find_element(By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input').send_keys(Keys.ENTER)
time.sleep(10)

#Clicar em escrever novo e-mail
driver.find_element(By.XPATH, '/html/body/div[7]/div[3]/div/div[2]/div[1]/div[1]/div[1]/div/div').click()
time.sleep(3)
#Colocando envio CCO
actions.key_down(Keys.CONTROL).key_down(Keys.SHIFT).send_keys('b').key_up(Keys.CONTROL).key_up(Keys.SHIFT).perform()
time.sleep(3)
driver.find_element(By.XPATH, '//*[@id=":q0"]').click()
time.sleep(4)

#COLOCANDO OS E-MAILS
for i in range(500,1000):
    email_bruto = emails.iloc[ i:(i+1) , 0]
    email_tratado = (email_bruto[i].split(" "))
    driver.find_element(By.XPATH, '//*[@id=":q0"]').send_keys(email_tratado,", ")
time.sleep(5)

#COLOCANDO O ASSUNTO
actions.key_down(Keys.TAB).key_up(Keys.TAB).perform()
time.sleep(3)
driver.find_element(By.NAME, 'subjectbox').send_keys("PORROGAÇÃO do Edital Centelha 2!!!") # <---- ASSUNTO DO E-MAIl
time.sleep(3)
actions.key_down(Keys.TAB).key_up(Keys.TAB).perform()
time.sleep(3)
actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
time.sleep(10)
actions.key_down(Keys.CONTROL).key_down(Keys.ENTER).key_up(Keys.CONTROL).key_up(Keys.ENTER).perform()
time.sleep(10)
driver.quit()


# #### EXCLUINDO OS E-MAILS UTILIZADOS PARA QUE NÃO ENVIE NOVAMENTE O MESMO E-MAIL

# In[6]:


for i in range(0,1000):
    emails.drop( [i] , inplace = True)

emails.to_excel("emailsgeral.xlsx", index = False)

