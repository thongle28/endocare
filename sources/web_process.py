from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import os
import sources.logins as lg
def auto_import_data_tb(driver,web_link,table_name,file_name):
	table_dict={
		'exfm_tb':'/admin/home/exfm/import',
		'part_tb':'/admin/home/part/import',
		'rcode_tb':'/admin/home/rcode/import',
		'pending_tb':'/admin/home/pending/import',
		'comment_tb':'/admin/whiteboard/comment/import',
	}
	ext = 'csv'
	# file =lg.file_select('.'+ext,'exports')
	file = file_name[table_name]

	driver.get(web_link+str(table_dict[table_name]))
	upload_file = driver.find_element_by_xpath('//*[@id="id_import_file"]')
	upload_file.send_keys(file)
	file_format = Select(driver.find_element_by_name('input_format'))
	file_format.select_by_visible_text(ext)
	driver.find_element_by_xpath("//input[@type='submit']").click()
	sleep(2)
	driver.find_element_by_xpath("//input[@type='submit']").click()
	messages = driver.find_element_by_class_name('success').text
	print(messages)


def import_data_tb(driver,web_link,table_name,file_type='csv'):
	table_dict={
		'exfm_tb':'/admin/home/exfm/import',
		'part_tb':'/admin/home/part/import',
		'rcode_tb':'/admin/home/rcode/import',
		'pending_tb':'/admin/home/pending/import',
		'comment_tb':'/admin/whiteboard/comment/import',
	}
	ext = file_type
	file =lg.file_select('.'+ext,'exports')
	driver.get(web_link+str(table_dict[table_name]))
	upload_file = driver.find_element_by_xpath('//*[@id="id_import_file"]')
	upload_file.send_keys(file)
	file_format = Select(driver.find_element_by_name('input_format'))
	file_format.select_by_visible_text(ext)
	driver.find_element_by_xpath("//input[@type='submit']").click()
	sleep(2)
	driver.find_element_by_xpath("//input[@type='submit']").click()
	messages = driver.find_element_by_class_name('success').text
	print(messages)


def delete_all_table(driver,table_name='exfm'):
	select_all = driver.find_element_by_xpath('//*[@id="action-toggle"]')
	select_all.click()
	#--------------try to select all----------
	try:
		select_all = driver.find_element_by_xpath('//*[@id="changelist-form"]/div[1]/span[3]/a')
		select_all.click()
	except:
		pass
	#-----------------select action and go-----------------
	action = Select(driver.find_element_by_xpath('//*[@id="changelist-form"]/div[1]/label/select'))
	action.select_by_visible_text('Delete selected '+table_name+'s')
	driver.find_element_by_xpath('//*[@id="changelist-form"]/div[1]/button').click() #Go

	#----------Yes, I'm sure-------------------------------
	sleep(0.5)
	driver.find_element_by_xpath("//input[@type='submit']").click() #Yes, I'm sure

def login_web(driver,web_link,user_id,user_pw):

	driver.get(web_link+'/accounts')
	#--------logins webbase-----------
	user = driver.find_element_by_xpath('//*[@id="main"]/div/form/div[1]/input')
	user.send_keys(user_id)
	pw = driver.find_element_by_xpath('//*[@id="main"]/div/form/div[2]/input')
	pw.send_keys(user_pw)
	bt_submit = driver.find_element_by_xpath('//*[@id="main"]/div/form/input[2]')
	bt_submit.click()
	sleep(1)
	driver.get(web_link+'/admin')

def web_export(driver,file_format):
	Select(driver.find_element_by_name('file_format')).select_by_visible_text(file_format)
	driver.find_element_by_xpath("//input[@type='submit']").click()
	