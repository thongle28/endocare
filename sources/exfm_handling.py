import sources.logins as lg
from time import sleep
import warnings
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium import webdriver

from datetime import datetime as dt
warnings.filterwarnings('ignore')


class action():

	def __init__(self):
		pass

	def create_RMA(self,driver,last_rma_tb):
	    # click Create RMA
	    driver.find_element_by_xpath('//*[@id="RPR_EDIT_LINK"]').click()
	    
	    # processing last_rma_tb
	    rma = last_rma_tb['RMA No.'][0]
	    sn = last_rma_tb['SERIAL_NO'][0]
	    contact_by = last_rma_tb['Contacted by'][0]
	    contact_name = last_rma_tb['Contact Name'][0]
	    request = last_rma_tb['Customer Request'][0]
	    if contact_name == '': contact_name = '.'
	    if str(request) == 'None': request = '.'
	    receive = dt.strptime(last_rma_tb['RECIEVE_DATE'][0],'%Y-%m-%d %H:%M:%S').strftime('%d-%b-%y')
	    inspect = dt.strptime(last_rma_tb['IN_INSPECT_DATE'][0],'%Y-%m-%d %H:%M:%S').strftime('%d-%b-%y')
	    ship_out = dt.strptime(last_rma_tb['Shipped'][0],'%Y-%m-%d %H:%M:%S').strftime('%d-%b-%y')
	    note = last_rma_tb['Internal Note'][0]
	    # Search Serial
	    sn_field = driver.find_element_by_xpath('//*[@id="sidEQP_BODY_NO"]')
	    sn_field.clear()
	    sn_field.send_keys(sn)
	    sn_field.send_keys(Keys.RETURN)

	    #Get message
	    try:
	        msg = driver.find_element_by_xpath('//*[@id="messageLine"]/td[2]/div/span').get_attribute('innerHTML')
	    except:
	        msg =''
	    customer_name = driver.find_element_by_xpath('//*[@id="sidACC_NAME"]').get_attribute('value')
	    print(customer_name)

	    if customer_name !='':
	        if msg =='':
	            driver.find_element_by_xpath('//*[@id="sidINFORMATION_GET_MEDIUM_CODE"]').send_keys(contact_by)
	            driver.find_element_by_xpath('//*[@id="sidCUSTOMER_PERSON_NAME"]').send_keys(contact_name)
	            driver.find_element_by_xpath('//*[@id="sidNOTE_CUSTOMER_COMPLAINT"]').send_keys(request)
	            if str(note) !='None':
	                history = f'''{note}\n-----History Last RMA# {rma}------------\nReceived:{receive}\nInspection:{inspect}\nShip out: {ship_out}'''
	                driver.find_element_by_xpath('//*[@id="sidNOTE_INTERNAL"]').send_keys(history)
	            else:
	                history = f'''-----History Last RMA# {rma}------------\nReceived:{receive}\nInspection:{inspect}\nShip out: {ship_out}'''
	                driver.find_element_by_xpath('//*[@id="sidNOTE_INTERNAL"]').send_keys(history)
	        else:
	            print(msg)

	    # save and printout details
	    try:
	        bt_save = driver.find_element_by_xpath('//*[@id="sidADD_BUTTON_IMAGE"]')
	        bt_save.click()
	    except:
	        pass
	    sleep(0.5)
	    # check rma and model
	    rma = driver.find_element_by_xpath('//*[@id="sidREPAIR_ID"]').get_attribute('value')
	    model = driver.find_element_by_xpath('//*[@id="sidEQP_NAME"]').get_attribute('value')
	    create_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_1"]').get_attribute('value')
	    print('-'*50)
	    print('RMA:       ', rma)
	    print('Customer:  ',customer_name)
	    print('Model:     ',model)
	    print('Serial:    ',sn)
	    print('Received:  ',create_date)
	    print('Request:   ', request)
	    print('Note:      ', note,'\n')

	def received(self,drive,sn):
	    
	        
	    #click recieved button
	    driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_2"]').click()
	    sleep(0.5)
	    try:
	        driver.find_element_by_xpath('//*[@id="sidCOMPLETE_BUTTON_IMAGE"]').click()
	    except:
	        print(f'Already completed Receive {sn}')
	        driver.find_element_by_xpath('//*[@id="sidBACK_BUTTON_IMAGE"]').click()
	        
	def basic_page(self,diver,sn):
	    #loading data inspect
	    driver.get('https://exfm-asia-app.fujifilm.co.jp/')
	    driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]').click() #click RMA Search
	    sn_field = driver.find_element_by_xpath('//*[@id="sidIN_BODY_NO"]')
	    sn_field.clear()
	    sn_field.send_keys(sn)
	    driver.find_element_by_xpath('//*[@id="sidSEARCH_BUTTON_IMAGE"]').click()

	    #move to each job page
	    driver.find_element_by_xpath('//*[@id="search_result[0].sidRS_REPAIR_ID_data"]').click()

	    return sn