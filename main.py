import urllib.request
import pandas as pd
import os
import re
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import *
import time
import urllib


df = pd.read_excel(
                    '/home/hymar/Pulpit/projekty/docu_downloader/odczynniki.xls',
                    engine='xlrd'
                    )

dostawcy = {'thermo':'https://www.thermofisher.com/order/catalog/product/',
            'sigma':'https://www.sigmaaldrich.com/PL/pl',
            'vwr':'https://pl.vwr.com/store/'}

driver = webdriver.Chrome()

#wyłączenie okna z wyborem wyszukiwarki

chrome_options = Options()
chrome_options.add_argument("--disable-search-engine-choice-screen")
driver = webdriver.Chrome(options=chrome_options)

#funkcja 'table_search' wyszukuje klucze 'keys' ze słownika 'dostawcy' iterując po data frame, zwraca listy do funkcji data_pass
#funkcja data_pass zwraca listę list -> [link (słownik "dostawcy"), cat nr, lot nr (z data frame)]
def data_pass():

    web_cat_lot=[]

    def table_search(key, link):
        for index, row in df.iterrows():
            if len(re.findall(key,row['web']))==1:
                web_cat_lot.append([link, str(row['cat nr']), str(row['lot nr'])])
                
    for item in dostawcy.items():
        table_search(item[0], item[1])

    return web_cat_lot

#funkcja 'shadow_element' wyciąga element html z shadow-root
def shadow_element(entry_selector: str, element_selector: str):
    
    shadow_entry = driver.find_element(By.CSS_SELECTOR, entry_selector)
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", shadow_entry)
    element = shadow_root.find_element(By.CSS_SELECTOR, element_selector)

    return element


def shadow_elements(entry_selector: str, elements_selector: str):

    shadow_entry = driver.find_element(By.CSS_SELECTOR, entry_selector)
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", shadow_entry)
    elements = shadow_root.find_elements(By.CSS_SELECTOR, elements_selector)

    return elements

#funkcja parser nawiguje po stronie internetowej w zależności od producenta (zrobiono wersję dla thermoscientific)
def parser():

    web_cat_lot = data_pass()


    for i in web_cat_lot:
        
        if len(re.findall('thermo', i[0]))==1:
            adress = i[0]+i[1]
            

            driver.get(adress)
            driver.maximize_window()
            

            #zatwierdzenie "cookie"
            try:
                cookie = driver.find_element(
                    By.ID,'''truste-consent-button''').click()

            except NoSuchElementException:
                pass           
            
            #zapisanie do zmiennej "header_text" nazwy odczynnika
            try:
                header = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div[1]/div[2]/div[2]/div[1]/span''')
                header_text = header.text
                print('pobranie nazwy A')
                      
            except NoSuchElementException:
                print('pobranie nazwy B')
                header = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div/div[4]/div/div/div/div[3]/div[1]/div/div[2]/h1''')
                header_text = header.accessible_name  
            time.sleep(3)
            

            try:
                selector = '''#certificates > div.pdp-certificates-search > 
                        div.pdp-certificates-search__inputs > div > core-search'''
                iframe = driver.find_element(By.CSS_SELECTOR,selector)
                ActionChains(driver).scroll_to_element(iframe).perform()
                print('wyszukanie searchbar A')

            except NoSuchElementException:
                selector = '''#root > div > div > div.p-tabs > div.p-tabs__content > div:nth-child(7) > div > div > div:nth-child(1) > div:nth-child(2) > div.pdp-documents__document-section > div > div.pdp-documents__search > div.pdp-documents__search-inputs > div > div.c-search-bar.pdp-documents__search-bar.pdp-documents__search-bar--desktop > input'''
            
                iframe = driver.find_element(
                    By.CSS_SELECTOR, selector)
                ActionChains(driver).scroll_to_element(iframe).perform()
                print('wyszukanie searchbar B')
                             
            #wyszukanie search bar'u i wprowadzenie nr lot
            
            try:

                element = shadow_element(
                            entry_selector = selector,
                            element_selector = 'div > div > input'
                                        )
                print('drugi blok A')
                
            except AttributeError:
                print('drugi blok B')
                element = driver.find_element(
                    By.XPATH, '''//*[@id="root"]/div/div/div[5]/div[2]/div[7]/div/div/div[1]/div[1]/div[2]/div/div[1]/div[2]/div/div[2]/input''')
            element.click()
            element.send_keys(i[2])
            element.send_keys(Keys.RETURN)
            time.sleep(2)
            

            #pobranie linka pliku CoA
            try:
                link = driver.find_element(
                    By.XPATH, '''//*[@id="certificates"]
                    /div[2]/div[1]/div[2]/span[1]''')
                
                print('trzeci blok A')
                
            except NoSuchElementException:
                print('trzeci blok B')
                link = driver.find_element(By.XPATH, 
                        '''//*[@id="root"]/div/div/div[5]
                        /div[2]/div[7]/div/div/div[1]/div[1]
                        /div[2]/div/div[2]/div[2]/div/span[1]
                        /a/span[2]''')
            
            window_before = driver.window_handles[0]
            link.click()
            window_after = driver.window_handles[1]
            driver.switch_to.window(window_after)
            String_url = driver.current_url
            driver.close()
            driver.switch_to.window(window_before)

            
            new_dir = 'output/' + header_text + '_' + 'nr_cat_' + i[1] + '/' + 'CoA/'
            
            try:
                os.makedirs(new_dir)

            except FileExistsError:
                pass
            
           
            url_path = new_dir + 'CoA_nr_lot_' + i[2]
            urllib.request.urlretrieve(String_url, url_path)
            
            try:

                element = driver.find_element(By.XPATH,'//*[@id="sds"]/div')
                ActionChains(driver).scroll_to_element(element).perform()
                print('znalezienie SDS A')

            except NoSuchElementException:
                element = driver.find_element(By.XPATH, '''//*[@id="root"]/div/div/div[5]/div[2]/div[7]/div/div/div[1]/div[2]/div[2]/div/div[1]/div/a/span[2]''')
                ActionChains(driver).scroll_to_element(element).perform()
                print('znalezienie SDS B')     

            element.click()

            try:
                element = driver.find_element(
                    By.XPATH,'''//*[@id="modal"]/div/div/div[3]
                    /div/div/div[2]/div/div[1]''')
                print('pobranie SDS A')

            except NoSuchElementException:

                try:

                    element = shadow_element(
                                entry_selector = '''#modal > core-modal > div.pdp-sds-modal
                                > div > div.sds-modal-selectors-catalog-number > core-dropdown''',
                                element_selector = 'div > div > label > div > div > span > span'
                                        )
                    print('pobranie SDS B')

                except NoSuchElementException:
                
                    element = driver.find_element(
                        By.XPATH,'''//*[@id="sds"]/div''')
                    
                    print('pobranie SDS C') 


            element.click()
            time.sleep(2)

            try:
                elements = shadow_elements(
                            entry_selector = '''#modal > core-modal > div.pdp-sds-modal
                            > div > div.sds-modal-selectors-catalog-number > core-dropdown''',
                            elements_selector = 'div > div > div > div > *'
                                    )
                print(len(elements))
                print('zatwierdzenie cat A')

            except NoSuchElementException:
                try:

                    element = shadow_element(
                            entry_selector = '''#modal > core-modal''',
                            element_selector = '''#modal > core-modal > div.pdp-sds-modal
                            > div > div.sds-modal-selectors-catalog-number > core-dropdown'''
                                    )
                    
                    elements = shadow_elements(
                            entry_selector = element,
                            elements_selector = '''div > div > div > div > *'''
                                    )

                    print(len(elements))
                    print('zatwierdzenie cat B')
               
                except NoSuchElementException: 
                    try:                                                  
                        elements = driver.find_elements(By.CSS_SELECTOR, '''#modal > div > div > div.c-modal__content > div > div > div.c-dropdown > div > div.c-dropdown__options-container.c-dropdown--active > div > *''')
                        print(len(elements))
                        print('zatwierdzenie cat C')
                    except NoSuchElementException:
                        elements = driver.find_elements(By.CSS_SELECTOR, '''div > div > div > div > *''')
                        print(len(elements))
                        print('zatwierdzenie cat D')
            

            for a in elements:
                
                if len(re.findall(i[1], a.accessible_name))==1:
                    a.click()
                else:
                    pass
                print(f'''brak SDS dla nr cat {i[1]} lub ambiwalentny wynik.
                    Usuń nr cat z tabelki, usuń outputs, uruchom program ponownie''')
                
               

            time.sleep(3)
            try:
                element = driver.find_element(By.CSS_SELECTOR, 'div > div > label > div > div > span > span')
                
                print('language drop list A')
            except NoSuchElementException:
                try:
                    element = driver.find_element(By.CSS_SELECTOR, '#modal > div > div > div.c-modal__content > div > div:nth-child(2) > div.c-dropdown > div > div.c-dropdown__selected-option')
                    
                    print('language drop list B')

                except NoSuchElementException:

                    element = shadow_element(
                                entry_selector = '''#modal > core-modal > div.pdp-sds-modal
                                > div > div.sds-modal-selectors-language > core-dropdown''',
                                element_selector = 'div > div > label > div > div > span > span'
                                        )
                    
                    print('language drop list C')
            element.click()

            try:
                elements = driver.find_elements(By.CSS_SELECTOR,'''#modal > div > div > div.c-modal__content > div > div:nth-child(2) > div.c-dropdown > div > div.c-dropdown__options-container > div > *''')
                dupa = print(len(elements))
                print(f'blok A ma {dupa} elementów')
                    
            except NoSuchElementException:
                print('blok A jest spierdolony')

            if len(elements)==0:

                try:
                    elements = driver.find_elements(By.CSS_SELECTOR,'''#modal > div > div > div.c-modal__content > div > div:nth-child(2) > div.c-dropdown > div > div.c-dropdown__options-container.c-dropdown--active > div > *''')
                    dupa = print(len(elements))
                    print(f'blok A.2 ma {dupa} elementów')

                except NoSuchElementException:
                    print('blok A.2 jest spierdolony')
            
            if len(elements)==0:
                try:
                    elements = shadow_elements(
                                entry_selector = '''#modal > core-modal > div.pdp-sds-modal
                                    > div > div.sds-modal-selectors-language > core-dropdown''',
                                elements_selector = 'div > div > div > div > *'
                                        )
                    dupa = print(len(elements))
                    print(f'blok B ma {dupa} elementów')

                except NoSuchElementException:
                    print('blok B jest spierdolony')
                    pass
            else:
                pass

            if len(elements)==0:
                try:
                    element = shadow_element(
                                entry_selector = '''#modal > core-modal''',
                                element_selector = '''#modal > core-modal > div.pdp-sds-modal > div > div.sds-modal-selectors-language > core-dropdown'''
                                        )

                    elements = shadow_elements(
                                entry_selector = element,
                                elements_selector = '''div > div > div > div > *'''
                                        )
                    dupa = print(len(elements))
                    print(f'blok C ma {dupa} elementów')

                except NoSuchElementException:
                    print('blok C jest spierdolony')
                    pass
            else:
                pass

            if len(elements) == 0:
                print('dalej jest coś spierdolone')
                break

            else:
                for e in elements:
                    if len(re.findall('Polish', e.accessible_name))==1:
                        e.click()

            window_before = driver.window_handles[0]

            try:

                element = driver.find_element(By.CSS_SELECTOR, '''#modal > core-modal > div.pdp-sds-modal-footer
                                > div > core-button:nth-child(2)''')
            except NoSuchElementException:

                try:
                    element = shadow_element(
                            entry_selector = '''#modal > core-modal > div.pdp-sds-modal-footer
                              > div > core-button:nth-child(2)''',
                            element_selector = 'button'
                                    )
                except NoSuchElementException:

                    element = driver.find_element(By.CSS_SELECTOR, '''#modal > div > div > div.c-modal__buttons > button.c-btn.c-btn--outline''')
               
            
            
            driver.execute_script("arguments[0].click();", element)
            
            window_after = driver.window_handles[1]
            driver.switch_to.window(window_after)
            string_url = driver.current_url
            driver.close()
            driver.switch_to.window(window_before)
            
            path = 'output/' + header_text + '_' + 'nr_cat_' + i[1] + '/'
            
            url_path = path + 'SDS_' + i[1]
            urllib.request.urlretrieve(string_url, url_path)
                       
  

parser()

    

