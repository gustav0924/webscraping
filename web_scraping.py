import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import EdgeOptions
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from datetime import datetime
import os
import glob
from selenium.webdriver import DesiredCapabilities

###### Global Variables
xlsx_to_read = pd.read_excel(r'C:\Users\gusta\Documents\LocalCode\Python_Scripts\WebScrap\datasets\ListPart.xlsx')
columns = ['Part #','Manufacturer','Description','Stock','Price','Buy','Distributor']
webpages = ['oemstrade','findchips']
options = EdgeOptions()
service = Service("MicrosoftWebDriver.exe")
###### Selenium ################### 
cap = DesiredCapabilities().EDGE.copy()
cap['acceptSslCerts'] = True
cap['acceptInsecureCerts'] = True
options.add_argument('--headless')
options.add_argument('disable-gpu')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument('log-level=3')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument('--disable-extensions')
options.add_experimental_option("excludeSwitches", ['enable-automation'])
driver = webdriver.Edge(service=service,options=options)
#################################################



#Web scrapping from oemstrade
def newoemstrade(key):
    nopart = key
    driver.get('https://oemstrade.com')
    element = driver.find_element(By.ID,'part')
    element.send_keys(nopart)
    element.send_keys(Keys.RETURN)
    sleep(1)
    try: 
        driver.find_element(By.XPATH,"//div[contains(text(), 'No results')]")
        return pd.DataFrame()
    except NoSuchElementException: 
        try: 
            driver.find_element(By.CLASS_NAME,"no-results-results-title")
            return pd.DataFrame()
        except NoSuchElementException:
            while True:
                sleep(0.5)
                try:
                    button_to_click = driver.find_element(By.CLASS_NAME,"show-more")
                    button_to_click.click()
                except:
                    try:
                        button_to_click = driver.find_element(By.XPATH, "//a[text()='Show All']")
                        driver.execute_script("arguments[0].click();", button_to_click)
                    except:
                        break
            topd = []
            # Wait for the elements to load
            wait = WebDriverWait(driver, 5)
            
            tables = wait.until(EC.presence_of_all_elements_located((By.XPATH,'//div[contains(@class, "distributor-results")]')))
            # Loop through each table and extract its content
            for index, table in enumerate(tables):
            # Convert the table to a Pandas DataFrame
                minidf = pd.read_html(table.get_attribute("outerHTML"))
                minidf = pd.concat(minidf)
                minidf = minidf.dropna(how='all')
                h=table.find_element(By.CLASS_NAME,'distributor-title')
                hpd = h.text
                minidf['Distributor'] = hpd  
                topd.append(minidf)
            df = pd.concat(topd)
            df.columns = columns
            df['WEBSITE'] = webpages[0]
            df['Stock']=df['Stock'].astype(str)
            df['Stock'] = df['Stock'].str.extract('([0-9]+)')
            return df

#Web scrapping from findchips
def findchips(nopart):    
    driver.get('https://www.findchips.com')
    sleep(3)
    element = driver.find_element(By.XPATH, "//input[contains(@placeholder, 'Enter a part number')]")
    element.send_keys(nopart)
    element.send_keys(Keys.RETURN)
    try: 
        driver.find_element(By.CLASS_NAME,'no-results')
        return pd.DataFrame()
    except NoSuchElementException:
        while True:
            try:
                button_to_click = driver.find_element(By.XPATH,"//ul[@class='price-list']//button[@class='hyperlink']")
                driver.execute_script("arguments[0].click();",button_to_click)
            except NoSuchElementException or TimeoutError:
                break
        # while True:
        topd = []
        distributors = driver.find_elements(By.CLASS_NAME,'distributor-results')
        result_items = distributors
        for index,result_item in enumerate(result_items):
            item_class = result_item.get_attribute("outerHTML")
            minidf = pd.read_html(item_class)
            minidf = pd.concat(minidf)
            minidf = minidf.dropna(how='all')
            h=result_item.find_element(By.CLASS_NAME,'distributor-title')
            hpd = h.text
            minidf['Distributor'] = hpd
            topd.append(minidf)
        df = pd.concat(topd,axis=0,ignore_index=True)
        df.drop(df.columns[df.columns.str.contains('Price Range',case = False)],axis = 1, inplace = True)
        df['Part #'] = nopart
        df.columns = columns
        df['WEBSITE'] = webpages[1]
        df['Stock']=df['Stock'].astype(str)
        df['Stock'] = df['Stock'].str.extract('([0-9]+)') 
        return df
    
def join(nopart):
    ''' 'join' Join both websites and deletes "Buy" column attached by web scrapping each webpage
                        ---Parameters---
                        nopart: str, Part number to search in both websites
                        
                        return Pandas Dataframe''' 

    o = newoemstrade(nopart)
    f = findchips(nopart)
    if not o.empty and not f.empty: 
        df = pd.concat([o,f])
        del df['Buy'], df['Description']
        return df 
    else:
        return pd.DataFrame()  

def main():
    report = pd.DataFrame(xlsx_to_read) 
    listofdf=[]
    for i in range(len(report)):
        Buyer = report.loc[i].dropna(axis=0) 
        print(Buyer)
        for j in range(len(Buyer.iloc[2:])):
            p = join(Buyer.iloc[2:][j])
            if not p.empty:
                p[['BPN','Buyer']] = Buyer.iloc[:2]
                listofdf.append(p)
                print(f'\n Results found on websites:{len(p)}')
            else:
                print('No Results Found')
                pass
    driver.close()
    df = pd.concat(listofdf,axis=0)
    splitdf = df['Price'].str.extractall(r'(\d+) \$(\d+\.\d+)')
    del df['Price']
    splitdf = splitdf.reset_index(level=1,drop=True)
    splitdf.columns = ['Lot Qty', 'PRICE']
    df = df.join(splitdf)
    df.rename(columns= {'BPN':'Material'},inplace=True)
    df =df.dropna(subset=['Stock'])
    df= df[(df['Stock'] != 0) | (~df['Stock'].isnull()) | (df['Stock'] != 0.0) | (df['Stock']!='0')]
    df['Stock'] = df['Stock'].astype('int64')
    df['Stock'] = df[['Stock']!=0]
    df.to_csv(datetime.now().strftime("%m%d%Y")+'_webdata.csv', index=False, encoding='utf-8')
if __name__ == '__main__':
    main()