from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd

nav = webdriver.Chrome()
tabela_produtos = pd.read_excel('buscas.xlsx')

import time

def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    #1. Entrar no google
    nav.get("https://www.google.com.br/")

    #1.1 tratar os valores que vieram da tabela
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)

    #2. Pesquisar o nome do produto no google
    nav.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto)
    nav.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

    #3. Clicar na aba shopping
    elementos = nav.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in elementos:
        if 'Shopping' in item.text:
            item.click()
            break

    #4. Pegar a lista de resultados da busca no google shopping
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')

    #5. Para cada resultado, ele vai verificar se o resultado corresponde a todas as nossas condições
    lista_ofertas = [] # lista que a função vai me dar como resposta
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME,'Xjkr3b').text
        nome = nome.lower()

        #Verificação do nome - se o nome tem algum termo banido
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
        
        #Verificar se no nome tem todos os termos do nome do produto
        tem_todos_termos_produto = True
        for palavra in lista_termos_produto:
            if palavra not in nome:
                tem_todos_termos_produto = False
        
        
        #Se tem_termos_banidos = False e o tem_todos_termos_produto = True, então executa análise 
        if not tem_termos_banidos and tem_todos_termos_produto: # Verificando o nome
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                preco = preco.replace('R$', '').replace(' ','').replace('.','').replace(',','.')
                preco = float(preco)
                #verificando se o preco está dentro do minimo e maximo
                if preco_minimo <= preco <= preco_maximo:
                    #pegando o elemento que está abaixo do que queremos obter
                    elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
                    #subindo para pegar o elemento que queremos
                    elementos_pai = elemento_link.find_element(By.XPATH, '..')
                    #Pegando de fato o link de cada produto
                    link = elementos_pai.get_attribute('href')
                    lista_ofertas.append((nome, preco, link))
            except:
                continue
    return lista_ofertas

def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    
    #1. Tratar os valores da função
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    
    
    #2. Entrar no buscapé
    nav.get('https://www.buscape.com.br/')

    #3. Pesquistar pelo produto no buscape
    nav.find_element(By.CLASS_NAME, 'search-bar__text-box').send_keys(produto, Keys.ENTER)
   
    #4. Pegar a lista de resultados da busca do buscape
    time.sleep(10)
    lista_resultados = nav.find_elements(By.CLASS_NAME,'Cell_Content__1630r')

    #5. Para cada resultado
    lista_ofertas = []
    for resultado in lista_resultados:
        try:
            preco = resultado.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
            nome = resultado.get_attribute('title')
            nome = nome.lower()
            link = resultado.get_attribute('href')

            #Verificação do nome - se o nome tem algum termo banido
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True
            
            #Verificar se no nome tem todos os termos do nome do produto
            tem_todos_termos_produto = True
            for palavra in lista_termos_produto:
                if palavra not in nome:
                    tem_todos_termos_produto = False
            
            if not tem_termos_banidos and tem_todos_termos_produto:
                preco = preco.replace('R$', '').replace(' ','').replace('.','').replace(',','.')
                preco = float(preco)
                if preco_minimo <= preco <= preco_maximo:
                    lista_ofertas.append((nome, preco, link))
        except:
            pass
    return lista_ofertas

tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preco', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_google_shopping)
    else:
        tabela_google_shopping = None

    lista_ofertas_buscape = busca_buscape(nav,produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_buscape)
    else:
        tabela_buscape = None

#exportar por excel
tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)

#enviar por e-mail o resultado da tabela
import win32com.client as win32


#Verificando se existe alguma oferta dentro da tabela de ofertas
if len(tabela_ofertas.index) > 0 :
    # vou enviar email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'joyceamorim051996@gmail.com'
    mail.Subject = f'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p> Prezados,</p>
    <p> Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <Att.,</p>
    """

    mail.Send()

nav.quit() 