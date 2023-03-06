#!/usr/bin/env python
# coding: utf-8

# Criar um navegador
# Importar base de dados
# Para cada produto da base de dados:
# Procurar o produto no Google Shopping
# Procurar o produto no Google Buscapé
# Verificar se algum dos produtos está na faixa de preço, para ambos os sites.
# Salvar as ofertas em uma tabela e exportar em xlsx.
# Enviar arquivo por e-mail

# In[95]:


#Criar Navegador
#Importar base de dados

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time

driver = webdriver.Chrome()

df_produtos = pd.read_excel('buscas.xlsx')
display(df_produtos)


# Definição das functions de busca no Google Shopping e Buscapé

# In[96]:



def busca_gshopping(driver, produto, termos_banidos, preco_min, preco_max):

    driver.get('https://www.google.com.br/')

    #tratando os valores da tabela
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ") 
    lista_termos_produto = produto.split(" ")

    #pesquisando o nome do produto no google
    driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto)
    driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

    #Clicar na aba Shopping
    elementos = driver.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for elemento in elementos:
        if 'Shopping' in elemento.text:
            elemento.click()
            break

    #pegar a lista de pesquisa no shopping
    resultados = driver.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')   
    #print(resultados)
    
    
    #para cada resultado, verificando nome, precoi e link
    ofertas_gshopping = [] #lista com resultados filtrados

    for resultado in resultados:
    
        nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        nome = nome.lower()
        #verificacao do nome
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                    tem_termos_banidos = True
        #verificacao dos produtos
        tem_produtos_banidos  = True
        for palavra in lista_termos_produto:
            if palavra not in nome:
                tem_produtos_banidos = False

        if tem_termos_banidos == False and tem_produtos_banidos == True: 
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco = float(preco)

                preco_min = float(preco_min)
                preco_max = float(preco_max)
                #verificando se o preço está nas condições impostas
                if preco_min <= preco <= preco_max:

                    elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
                    elemento_parent = elemento_link.find_element(By.XPATH, '..')
                    link = elemento_parent.get_attribute('href')
                    ofertas_gshopping.append((nome, preco, link))
            except:
                continue

                
    return ofertas_gshopping


def busca_buscape(driver, produto, termos_banidos, preco_min, preco_max):
    #tratar os valores da função
    preco_min = float(preco_min)
    preco_max = float(preco_max)
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ") 
    lista_termos_produto = produto.split(" ")
    
    
    
    #entrar no buscapé
    driver.get('https://www.buscape.com.br/')
    
    #pesquisar pelo produto no buscapé
    driver.find_element(By.XPATH,'//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)

    #pegar a lista no resultado no buscapé
    time.sleep(5)
    resultados = driver.find_elements(By.CLASS_NAME, 'SearchCard_ProductCard_Inner__7JhKb')
    
    #olhar se tem termo banido
    ofertas_buscape = []
    
    for resultado in resultados:
        try:
            preco = resultado.find_element(By.XPATH, '//*[@id="__next"]/div/div/div[2]/div[3]/div[1]/a/div[2]/div[2]/div[2]/p[1]').text
           
        
            nome = resultado.find_element(By.CLASS_NAME, 'Text_Text__h_AF6').text
            nome = nome.lower()
            #verificacao do nome
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                        tem_termos_banidos = True
                        
            tem_produtos_banidos  = True
            for palavra in lista_termos_produto:
                if palavra not in nome:
                    tem_produtos_banidos = False
                        
            if tem_termos_banidos == False and tem_produtos_banidos == True: 
                preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco = float(preco)
                if preco_min <= preco <= preco_max:
                    ofertas_buscape.append((preco, nome, link))
                
            
            link = resultado.get_attribute('href')
            #print(preco, nome, link)
        except:
            pass
    
    return ofertas_buscape
    #olhar se tem todos os termos do produto
    
        
    
    #conferir a faixa de preço
    
    #retornar a lista de ofertas
    


# Construção da lista de ofertas encontradas

# In[97]:


#Para cada produto da base de dados: Procurar o produto no Google Shopping

#Procurando produto no Google
tabela_ofertas = pd.DataFrame()
for linha in df_produtos.index:
    
    produto = df_produtos.loc[linha, 'Nome']
    termos_banidos = df_produtos.loc[linha, 'Termos banidos']
    preco_min = df_produtos.loc[linha, 'Preço mínimo']
    preco_max = df_produtos.loc[linha, 'Preço máximo']
    ofertas_gshopping = busca_gshopping(driver, produto, termos_banidos, preco_min, preco_max)
    if ofertas_gshopping:
        tabela_gshopping = pd.DataFrame(ofertas_gshopping, columns=['Produto','Preço', 'Link'])
        tabela_ofertas = tabela_ofertas.append(tabela_gshopping)
    else:
        tabela_gshopping = None
        
    ofertas_buscape = busca_buscape(driver, produto, termos_banidos, preco_min, preco_max)
    if ofertas_buscape:
        tabela_buscape = pd.DataFrame(ofertas_buscape, columns=['Produto','Preço', 'Link'])
        tabela_ofertas = tabela_ofertas.append(tabela_buscape)
    else:
        tabela_buscape = None

        
display(tabela_ofertas)


# #### Exportar a base de ofertas para Excel

# In[98]:


tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)


# #### Enviando o e-mail

# In[99]:


import win32com.client as win32

if len(tabela_ofertas.index) > 0:
    
    outlook = win32.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)
    mail.To = 'thomas.montolive@gmail.com'
    mail.Subject = 'Produtos Encontrados na faixa de preço desejada'
    mail.HTMLBody = f'''
    <p>Prezados </p>
    <p>Encontramos alguns produtos em oferta para a faixa de preço desejada. Segue a tabela com detalhes. </p>
    {tabela_ofertas.to_html()}
    
    <p> Qualquer dúvida, estou a disposição </p>
    
    <p> Atte. </p>
    '''

    mail.Send()
    

    
driver.quit()


# In[ ]:




