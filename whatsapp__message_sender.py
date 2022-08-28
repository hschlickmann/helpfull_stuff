'''
Web Whats App Message Sender

This project was designed to automate the message sending in Whatsapp.
It uses Selenium and Pandas to consume a Excel file and send messages to a list of numbers. 


The excel file need to be prepared with the phone numbers and messages to be send.
'''


import pandas as pd
import numpy as np
from selenium import webdriver
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from time import localtime , strftime
import time
from webdriver_manager.chrome import ChromeDriverManager
import ait
import easygui


def element_presence(by, xpath, time):
    element_present = EC.element_to_be_clickable((By.XPATH, xpath))
    WebDriverWait(driver, time).until(element_present)
    return

def send_message(url, image_path, atraso):
    driver.get(url)
    # aceita alertas
    try: 
        alert = alert = driver.switch_to.alert
        alert.accept()
    except Exception:
        time.sleep(0.5)

    # enviar imagem

    element_presence(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span', 30)
    anexo = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()

    element_presence(By.XPATH, '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[1]/button/span', 15)
    img = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[1]/button/span').click()

    time.sleep(1+atraso)
    ait.write(image_path)
    time.sleep(0.5)
    ait.press('\n')
    # enviar texto
    element_presence(By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div', 15)
    envio = driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()

    msg = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').clear()
    time.sleep(1+atraso)
    return

def send_message_only_text(url, atraso):
    driver.get(url)
    # aceita alertas
    try: 
        alert = alert = driver.switch_to.alert
        alert.accept()
    except Exception:
        time.sleep(0.5)

    # enviar texto
    element_presence(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span', 30)
    envio = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()

    msg = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').clear()
    time.sleep(0.5+atraso)
    return


easygui.msgbox(r'Before using this script prepare XLSX (excel) file. It need to habe a tab called "list" and a collumns called "link_wpp" with this pattern of link: https://web.whatsapp.com/send?phone=FULLPHONENUMBER&text=MESSAGE ', 'MESSAGE SENDER')


# validar se tem imagem ou não

tem_imagem = easygui.ynbox('Do you want to attach a image?', 'MESSAGE SENDER', ('Attach image', 'Send Just text'))

if tem_imagem == True:
    imagem_unica = easygui.ynbox('All the messages will receive the same image?', 'MESSAGE SENDER', ('I will send just one image', 'The messages have diferent images'))
    if imagem_unica == True:  
        # perguntar caminho da imagem 
        image_path = easygui.enterbox(r"What is the path of the image file? You need to paste the full path: C:\Users\bla\Downloads\imagem.jpg", 'MESSAGE SENDER')
    else: 
        easygui.msgbox(r'In this case include a columns called "image_path" in you excel file. It need contains the full path of the image you want to attach to each message: C:\Users\bla\Downloads\imagem.jpg', 'MESSAGE SENDER')


db_path = easygui.enterbox(r"What is the path of the excel file? C:\Users\bla\Downloads\base_dados.xlsx", 'MESSAGE SENDER')
# perguntar caminho do csv
try:
    db = pd.read_excel(db_path,sheet_name='list')
except Exception:
    easygui.msgbox('Erro no arquivo, reinicie o programa', 'MESSAGE SENDER')
    time.sleep(3)
    sys.exit()

# determina um atraso padrão
atraso = easygui.enterbox(r"Insert a delay to the bot work with. The default value is 1.", 'MESSAGE SENDER')
try:
    atraso = float(atraso)
except Exception:
    atraso = 1.0


options = Options()
options.add_argument("--user-data-dir-Session")
options.add_argument("--profile-directory=Default")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--disable-notifications")
options.add_argument("--disable-popup-blocking")
options.add_argument("chrome.switches")
options.add_argument("--disable-extensions --disable-extensions-file-access-check --disable-extensions-http-throttling --disable-infobars --enable-automation --start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),chrome_options=options)
driver.get('https://web.whatsapp.com/')

val_final = easygui.ynbox('Now you need to log in with your Whats App Account.\n\n Click OK when you ready.', 'MESSAGE SENDER', ('OK', 'Cancel'))

if val_final == False:
    driver.quit()
    sys.exit()

time.sleep(10)

db['send_log'] = np.nan

if tem_imagem == True:
    if imagem_unica == True:  
        for i in range(len(db)):
            try:
                send_message(db['link_wpp'][i].strip("\u202a"), image_path, atraso)
                db['send_log'][i] = strftime("%Y-%m-%d %H:%M:%S", localtime())
            except Exception:
                print("Erro in line: "+str(i))
            if i <= 1:
                time.sleep(5)
    else: 
        for i in range(len(db)):
            if len(str(db['image_path'][i])) <= 5:
                try:
                    send_message_only_text(db['link_wpp'][i].strip("\u202a"), atraso)
                    db['send_log'][i] = strftime("%Y-%m-%d %H:%M:%S", localtime())
                except Exception:
                    print("Error in line: "+str(i))
            else:
                try:
                    send_message(db['link_wpp'][i].strip("\u202a"), db['image_path'][i].strip("\u202a"), atraso)
                    db['send_log'][i] = strftime("%Y-%m-%d %H:%M:%S", localtime())
                except Exception:
                    print("Error in line: "+str(i)) 
                if i <= 1:
                    time.sleep(5)
else: 
    for i in range(len(db)):
        try:
            send_message_only_text(db['link_wpp'][i].strip("\u202a"), atraso)
            db['send_log'][i] = strftime("%Y-%m-%d %H:%M:%S", localtime())
        except Exception:
            print("Error in line: "+str(i))
        if i <= 1:
            time.sleep(5)


driver.quit()

db.to_excel('Result File'+strftime("%Y-%m-%d", localtime())+" Lines "+str(len(db))+".xlsx", index = False, header=True)
easygui.msgbox('Well done, we sent '+str(len(db))+' messages', 'MESSAGE SENDER')

