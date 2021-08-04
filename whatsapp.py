from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import datetime
import time
import openpyxl as excel
from urllib.parse import quote
import pyautogui as pg
from selenium.common.exceptions import TimeoutException
from io import BytesIO
from PIL import Image
import pyperclip as pc
import re

driver = webdriver.Chrome('./chromedriver')
count = 1

def readContacts(fileName):
    lst = []
    msgs = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    secondCol = sheet['B']
    for cell in range(len(firstCol)):
        if firstCol[cell].value != None:
            contact = (str(firstCol[cell].value))
            lst.append(contact)
        else:
            break
    for cell in range(len(secondCol)):
        print(secondCol[cell].value)
        if secondCol[cell].value != None:
            msg = str(secondCol[cell].value)
            msgs.append(msg)
        else:
            break
    return lst, msgs

def isAlertExists():
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present(), '')
        alert = driver.switch_to.alert
        alert.accept()
    except TimeoutException:
        pass

def sendChat(msg):
    useImage = False
    fileName = ''
    if re.match(r"[0-9]+.", msg.split(' ')[0]):
        useImage = True
        fileName = msg.split('.')[0]
        msg = msg[2+len(fileName):]

    newLine = str(msg).split('|')

    if len(newLine) > 1:
        for m in newLine:
            pc.copy(m.strip())
            pg.hotkey('ctrl', 'v')
            pg.hotkey('shift', 'enter')
            time.sleep(1)
    else:
        pc.copy(newLine[0])
        pg.hotkey('ctrl', 'v')

    if useImage:
        import win32clipboard
        image = ''
        try:
            image = Image.open('./'+ fileName +'.JPEG')
        except:
            image = Image.open('./'+ fileName +'.JPEG')
        output = BytesIO()
        image.convert('RGB').save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        #win32clipboard.SetClipboardText('test copy paste')
        win32clipboard.CloseClipboard()

        time.sleep(0.5)
        inputField = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[1]/div/div[2]')
        inputField.click()
        pg.hotkey('ctrl', 'v')
        time.sleep(1)
    
    try:   
        sendbutton = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]/button')
        if useImage:
            sendbutton = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/span/div/div')
        sendbutton.click()
    except:
        print('chat ke '+ target +' gagal')
    print('chat ke ' + target + ' sukses')

targets, msgs = readContacts("tt.xlsx")

#print(targets)

driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 600)
wait5 = WebDriverWait(driver, 5)
# msg = 'Test wa dari bot buatan nanda'
#targets =['Mona 1','School']
# targets = ['6281372280096', '6281212138860']
isMultipleChat = False

if len(msgs) > 1:
    isMultipleChat = True

# parsed_message = quote(msgs[0])
for target in targets:
    # driver.get('https://web.whatsapp.com/send?phone=' + target + '&text=' + parsed_message)
    driver.get('https://web.whatsapp.com/send?phone=' + target)
    time.sleep(5)
    isAlertExists()
    try:
        noWa = driver.find_element_by_css_selector('#app > div._1ADa8._3Nsgw.app-wrapper-web.font-fix.os-mac > span:nth-child(2) > div._2q3FD > span > div:nth-child(1) > div > div > div > div > div._2i3w0 > div > div > div')
        noWa.click()
        time.sleep(5)
        print(target + ' ke ' + str(count) + ' invalid')
        count += 1
        continue
    except:
        pass
    width, height = pg.size()
    #pg.click(width / 2, height / 2)
    time.sleep(5)
    # x_arg = '//span[contains(@title,' +'"' +target.decode('utf-8') + '"' +')]' #.decode('utf-8') 
    # group_title = wait.until(EC.presence_of_element_located((
    #     By.XPATH, x_arg)))
    # group_title.click()
    # time.sleep(5)
    # pg.press('enter')

    # message = driver.find_element_by_css_selector('#main > footer > div.vR1LG._3wXwX.copyable-area > div._2A8P4 > div > div._2_1wd.copyable-text.selectable-text')
    # message.send_keys(msg)
    # wait.until(EC.presence_of_element_located((
    #     By.CSS_SELECTOR, '#main > footer > div.vR1LG._3wXwX.copyable-area > div:nth-child(3) > button > span')))
    if isMultipleChat:
        for msg in msgs:
            sendChat(msg)
    else:
        sendChat(msgs[0])
    count += 1
    time.sleep(3)
driver.close()