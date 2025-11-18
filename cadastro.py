import getpass
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import UnexpectedAlertPresentException
import time

df = pd.read_csv('adjunto.csv')
driver = webdriver.Chrome()
driver.get('https://sipac.rn.gov.br/sipac')


print('#####  Dados de acesso ao SIPAC #####')
usuario = input('Usuário: ')
campo_usuario = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[3]/form/table/tbody/tr[1]/td/input')
campo_usuario.send_keys(usuario)
senha = getpass.getpass('Senha: ')
campo_senha = driver.find_element(By.XPATH,'/html/body/div[1]/div[2]/div[3]/form/table/tbody/tr[2]/td/input')
campo_senha.send_keys(senha)
campo_usuario.submit()

modulos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[1]/div[2]/div[1]/ul/li[1]/span/span/a')))
modulos.click()

patrimonio_movel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[7]/div[2]/div/div/div[1]/ul[1]/li[23]/a')))
patrimonio_movel.click()

gerencia = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/form/div/div[1]/div/table/tbody/tr/td[2]/a/span/em/span')))
gerencia.click()

cadastro = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/form/div/div[2]/div[5]/ul/li[1]/ul/li/a')))
cadastro.click()

for index, row in df.iterrows():
    try:
        actions = ActionChains(driver)
        tombamento_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[3]/td/input')))
        tombamento_input.clear()
        tombamento_input.send_keys(row['Tombamento'])
        actions.send_keys(Keys.TAB).perform()


        material_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[4]/td/input[1]')))
        material_input.clear()
        material_input.send_keys(row['Material'])
        time.sleep(2)
        actions.send_keys(Keys.TAB).perform()
        actions.send_keys(Keys.TAB).perform()
        time.sleep(4) # foi necessário adicionar um tempo de espera para que o campo seguinte fosse identificado


        marca_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'nome')))
        marca_input.send_keys(row['Marca'])
        time.sleep(2)
        actions.send_keys(Keys.TAB).perform()

        valor_estimado_input = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[8]/td/input')
        valor_estimado_input.clear()
        valor_estimado_input.send_keys(row['Valor Estimado'])
        # time.sleep(3)
        actions.send_keys(Keys.TAB).perform()

        estado_bem_select = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[9]/td/select')
        (Select(estado_bem_select)).select_by_visible_text('EM USO')

        finalidade_select = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[10]/td/select')
        (Select(finalidade_select)).select_by_visible_text('ACERVO')

        # unidade_responsavel_input = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[11]/td/input')
        unidade_responsavel_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[11]/td/input')))
        unidade_responsavel_input.clear()
        unidade_responsavel_input.send_keys(row['Unidade Responsável'])
        time.sleep(2)
        actions.send_keys(Keys.TAB).perform()

        t1_input = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[13]/td/input[1]')
        t1_input.clear()
        t1_input.send_keys(row['T1'])
        actions.send_keys(Keys.TAB).perform()

        t2_input = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[13]/td/input[2]')
        t2_input.clear()
        t2_input.send_keys(row['T2'])
        actions.send_keys(Keys.TAB).perform()

        etiquetavel_select = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[1]/tbody/tr[17]/td/select')
        (Select(etiquetavel_select)).select_by_visible_text('Sim (seu custo total de aquisição foi MAIOR do que R$ 458,00).')

        registrar_bem_button = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/form/table[2]/tfoot/tr/td/center/input[1]')
        registrar_bem_button.click()
        
        alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
        time.sleep(1)
        alert.accept()

        actions.send_keys(Keys.ENTER).perform()


        patrimonio_movel_link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/center[2]/a')))
        patrimonio_movel_link.click()

        cadastro_link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[2]/form/div/div[2]/div[5]/ul/li[1]/ul/li/a')))
        cadastro_link.click()


    except Exception as e:
        print(e)
        print(f'Item {row["Tombamento"]} não foi cadastrado. Revisar os dados e tentar novamente.')
        time.sleep(5)
        break

    print(f'Item {row["Tombamento"]} cadastrado com sucesso.')
    time.sleep(2)

time.sleep(2)
driver.quit()
