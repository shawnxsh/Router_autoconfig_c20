from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import requests
import time
import re
import pandas as pd
import openpyxl
import openpyxl.cell._writer


class FileHandler:
    def __init__(self, fileName):
        self.fileName = fileName
    
    def getNewConfigureInfo(self):
        df = pd.read_excel(self.fileName, engine="openpyxl", sheet_name='Sheet1')
        if len(df) == 0:
            print("\n'configurationInfo.xlsx' has no information to use, please update it...\nThis program will be shut down...")
            time.sleep(5)
            quit()
        firstRow = df.iloc[0]           
        return firstRow

    # def appendConfiguredInfo(self, configuredInfo):
    #     wb = openpyxl.load_workbook(self.fileName)
    #     sheet = wb['Sheet2']
    #     sheet.append(configuredInfo)
    #     wb.save(self.fileName)   
    #     return
    

    # def deleteConfiguredInfo(self):
    #     wb = openpyxl.load_workbook(self.fileName)
    #     sheet = wb['Sheet1']
    #     sheet.delete_rows(2, 1)
    #     wb.save(self.fileName)
    #     return
    
    def appendConfiguredInfo(self, configuredInfo):
        df = pd.read_excel(self.fileName, engine="openpyxl", sheet_name='Sheet2')
        df.loc[len(df.index)] = configuredInfo
        
        with pd.ExcelWriter(self.fileName, engine="openpyxl", mode='a', if_sheet_exists='replace' ) as writer:
            df.to_excel(writer, sheet_name='Sheet2', index=False)
            
        return
    

    def deleteConfiguredInfo(self):
        df = pd.read_excel(self.fileName, engine="openpyxl", sheet_name='Sheet1')
        df = df.iloc[1: , :]
        
        with pd.ExcelWriter(self.fileName, engine="openpyxl", mode='a', if_sheet_exists='replace' ) as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            
        return


class Router:
    def __init__(self, routerID, routerModel, pppoeUsername, pppoePassword, loginPassword):
        self.routerID = routerID
        self.routerModel = routerModel
        self.pppoeUsername = pppoeUsername
        self.pppoePassword = pppoePassword
        self.loginPassword = loginPassword
        
    def handleInputItem(self, id, content):
        inputItem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, id)))
        inputItem.clear()
        inputItem.send_keys(content)
        return

    def handleButtonItem(self, xpath):
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath))).click()   
        return
    
    def switchToFrame(self, id):
        driver.switch_to.default_content()
        WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID,id)))
        return
    
    def createAccount(self):
        self.handleInputItem('change-pcPassword', self.loginPassword)
        self.handleInputItem('confirm-pcPassword', self.loginPassword)
        self.handleButtonItem('//*[@id="createBtn"]')
        return
    
    def loginAccount(self):
        self.handleInputItem('pcPassword', LoginPassword)
        self.handleButtonItem('//*[@id="loginBtn"]')
        return

    def replaceInputItem(self, id, oldPattern, newPattern): 
        inputItem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, id)))
        oldValue = inputItem.get_attribute('value')
        newValue = re.sub(oldPattern, newPattern, oldValue)
        inputItem.clear()
        inputItem.send_keys(newValue)
        return newValue

    def changePPPoEInfo(self):
        self.switchToFrame("frame1")
        self.handleButtonItem('//*[@id="menu_network"]')
        self.handleButtonItem('//*[@id="menu_wan"]')

        time.sleep(1)
        
        self.switchToFrame("frame2")
        self.handleButtonItem('//*[@id="link_type"]')
        self.handleButtonItem('//*[@id="pppoe"]')
        
        time.sleep(1)
        
        self.handleInputItem('username', self.pppoeUsername)
        PPPPassword = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'pwd')))
        PPPPassword.clear()
        PPPPassword2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID,'pwd2')))
        PPPPassword2.clear()

        PPPPassword.send_keys(self.pppoePassword)
        PPPPassword2.send_keys(self.pppoePassword)

        self.handleButtonItem('//*[@id="saveBtn"]')
        WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By. XPATH, '//*[@id="_load"]')))
        return

    def changeWirelessNetworkName(self, menuItemXpath, BasicSettingXpath, oldPattern, newPattern):
        self.switchToFrame("frame1")
        self.handleButtonItem(menuItemXpath)
        self.handleButtonItem(BasicSettingXpath)

        time.sleep(1)
        
        self.switchToFrame("frame2")
        newValue = self.replaceInputItem('ssid', oldPattern, newPattern)

        self.handleButtonItem('//*[@id="tail"]/input')
        WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="_load"]')))

        return newValue

    def getRouterInfoByXpath(self, xpath, elementType):
        selectedItem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))
        if elementType == 'INPUT':
            value = selectedItem.get_attribute('value')
        elif elementType == 'SPAN':
            value = selectedItem.get_attribute('innerHTML')
        return value

    def configuration(self):
        # set wan username & password    
        self.changePPPoEInfo()
        
        # change Wireless Network Name 2.4GHz & 5GHz
        WiFiSSID2G = self.changeWirelessNetworkName('//*[@id="menu_wl2g"]', '//*[@id="menu_wlbasic"]', '^TP-Link', 'Occom')
        WiFiSSID5G = self.changeWirelessNetworkName('//*[@id="menu_wl5g"]', '//*[@id="menu_wlbasic5g"]', '^TP-Link', 'Occom')

        self.switchToFrame("frame1")
        self.handleButtonItem('//*[@id="menu_wlsec5g"]')
        time.sleep(1)
        self.switchToFrame("frame2")
        WiFiPassword = self.getRouterInfoByXpath('//*[@id="pskSecret"]', 'INPUT')


        self.switchToFrame("frame1")
        self.handleButtonItem('//*[@id="menu_status"]')
        time.sleep(1)
        self.switchToFrame("frame2")
        MACAddress = self.getRouterInfoByXpath('//*[@id="lanmac"]', 'SPAN')

        TimeConfigured = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
        
        configInfoFileHandler.appendConfiguredInfo([self.routerID, self.routerModel, self.pppoeUsername, self.pppoePassword, self.loginPassword,MACAddress, WiFiSSID2G, WiFiSSID5G, WiFiPassword, TimeConfigured])
        
        configInfoFileHandler.deleteConfiguredInfo()
                    
        driver.close()
        
        return



PATH = 'chromedriver.exe' 
IPAddress = 'http://192.168.0.1/'

configInfoFileHandler = FileHandler('configurationInfo.xlsx')

while True:
    
    OccomRouterID, RouterModel, PPPoEUsername, PPPoEPassword, LoginPassword = configInfoFileHandler.getNewConfigureInfo()
    
    routerConfig = Router(OccomRouterID, RouterModel, PPPoEUsername, PPPoEPassword, LoginPassword)
    
    try:
        response = requests.get(IPAddress)
        try: 
            options = webdriver.ChromeOptions()
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_argument('--headless')
            driver = webdriver.Chrome(PATH, options=options)
            driver.get(IPAddress)
            
            routerConfig.handleInputItem('change-pcPassword', LoginPassword)

            try:
                print('\nThe router starts configuring...')
                options.headless = False
                driver = webdriver.Chrome(PATH, options=options)
                driver.get(IPAddress)
                
                routerConfig.createAccount()
                routerConfig.configuration()
                
                print('This router is configured successfully !!! Unplug it and connect a new one !!!')
                time.sleep(5)
            except:
                print('\nWoops...something went wrong, shut down the program and try again...')  

        except:
            print('\nThe current plugged router has been configured. Please disconnect it and process the next one.')
            command = input('Enter "Y" if you want to reconfigure the current router with new parameters: ')
            if command == 'Y':
                try:
                    response = requests.get(IPAddress)
                    
                    options = webdriver.ChromeOptions()
                    options.add_experimental_option('excludeSwitches', ['enable-logging'])
                    options.add_argument('--headless')
                    driver = webdriver.Chrome(PATH, options=options)
                    driver.get(IPAddress)
                    
                    routerConfig.handleInputItem('pcPassword', LoginPassword)
                    
                    try:
                        print('\nThe router starts reconfiguring...')
                        options.headless = False
                        driver = webdriver.Chrome(PATH, options=options)
                        driver.get(IPAddress)
                        
                        routerConfig.loginAccount()
                        routerConfig.configuration()
                        
                        print('This router is configured successfully !!! Unplug it and connect a new one !!!')
                        time.sleep(5)
                    except:
                        print('\nWoops...something went wrong, shut down the program and try again...')
                except: 
                    continue  
            else:
                print('Please connect a new router...')
                time.sleep(5)    
    except :
        print('No router is connected yet, please connect a router...')