import pandas as pd
import win32com.client as win32
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
# Bloco de codigo usado para atualizar automaticamente
# o webdriver sem necessidade de dowloads constantes
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# Passo 1 - Pesquisar a cotação do dólar
navegador.get('https://www.google.com/')
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação do dolar')
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

# Passo 2 - Pegar a cotação do dólar
cot_dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print('Dólar = ', cot_dolar)

# Passo 3 - Pegar a cotação do euro
navegador.find_element(By.XPATH, '//*[@id="logo"]/img').click() # voltar p google
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação do euro')
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cot_euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print('Euro = ', cot_euro)

# Passo 4 - Pegar a cotação do ouro
navegador.get('https://www.melhorcambio.com/ouro-hoje')
cot_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
cot_ouro = cot_ouro.replace(',', '.')
print('Ouro = ', cot_ouro)

# Importar e atualizar o Banco de Dados
tabela = pd.read_excel('Produtos.xlsx')
pd.set_option('display.max_columns', None)
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(cot_dolar)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(cot_euro)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(cot_ouro)

# atualizar preço de compra e venda
tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']
tabela['Preço de Compra'] = tabela['Preço de Compra'].map('R${:.2f}'.format)
tabela['Preço de Venda'] = tabela['Preço de Venda'].map('R${:.2f}'.format)

print('-' * 25, 'DADOS ATUALIZADOS', '-' * 25)
print(tabela)

# exportar uma nova tabela atualizada para o Excel
tabela.to_excel('Produtos Atualizados.xlsx', index=False)

# Enviando email com os Dados atualizados
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'teste25896@gmail.com'
mail.Subject = 'Produtos Atualizados'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue os dados atualizados,por cada produto e seu valor final para venda.</p>

<p>Dados Atualizados:</p>
{tabela.to_html(index=False)}

<p>Qualquer dúvida estou à disposição.</p>
<p>Att., Jaqueline</p>

'''
mail.Send()
print('Email enviado')
navegador.quit()
