import sources.logins as lg 
import sources.pending_process as pp
import sources.web_process as wp
import os
import pathlib
from time import sleep
import warnings
warnings.simplefilter("ignore")

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

# v1.1.0 hidden username

#----------- Select Link
link_dict={
	1:('noisoifujifilm','http://noisoifujifilm.vn'),
	2:('localhost','http://127.0.0.1:8000')
}
print('\nSelect link to import')
for i in range(1,len(link_dict)+1):
	print(f'{i}   {link_dict[i][0]}')



while True:
	s_link = input('Select link to import or type other link(Default 1):')
	if s_link=='': s_link = '1'
	if s_link.isnumeric():
		s_link = int(s_link)
		if s_link <= len(link_dict):
			web_link = link_dict[s_link][1]
			break
		else:
			print('Type wrong number\n')
	else:
		web_link = s_link
		break

folder_name='exports'
driver = lg.check_driver()
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : str(os.path.join(pathlib.Path().absolute(),folder_name))}
chromeOptions.add_experimental_option("prefs",prefs)
if not os.path.exists(folder_name):
	os.makedirs(folder_name)
	print('\nCreate new folder: {}'.format(folder_name))

driver = webdriver.Chrome(options = chromeOptions)
driver.maximize_window()


driver.get(web_link)

wp.login_web(driver,web_link,user_id='robocon2021',user_pw='endocare2021')
#------backup data pending----------
driver.get(web_link+'/admin/home/pending/export/?')
wp.web_export(driver,'xlsx')
sleep(0.5)
file,ctime = lg.file_latest('exports')
print (f"download success {file}")

#------backup data comment----------
driver.get(web_link+'/admin/whiteboard/comment/export/?')
wp.web_export(driver,'xlsx')
sleep(0.5)
file,ctime = lg.file_latest('exports')
print (f"download success {file}")

#-------delete exfm-----------------
driver.get(web_link+'/admin/home/exfm/')
try:
	wp.delete_all_table(driver)
	messages = driver.find_element_by_class_name('success').text
	print(messages)
except:
	print('Every thing is empty.')

#------auto import------------------------------------
folder_list = os.listdir(folder_name)
path = os.path.join(pathlib.Path().absolute(),folder_name)
count_file=0
for i in range(len(folder_list)):
	if folder_list[i].startswith('exfm_web'):
		ex = folder_list[i]
		ex = os.path.join(path,ex)
		count_file +=1
		continue
	if folder_list[i].startswith('part_name'):
		part = folder_list[i]
		part = os.path.join(path,part)
		count_file +=1
		continue
	if folder_list[i].startswith('r_code'):
		r_code = folder_list[i]
		r_code = os.path.join(path,r_code)
		count_file +=1
		continue
	if folder_list[i].startswith('pending'):
		pen = folder_list[i]
		pen = os.path.join(path,pen)
		count_file +=1
		continue
	if count_file == 4: break
file_name={
	'exfm_tb':ex,
	'part_tb':part,
	'rcode_tb':r_code,
	'pending_tb':pen
}
#--------import exfm-----------------

wp.auto_import_data_tb(driver,web_link,'exfm_tb',file_name)
wp.auto_import_data_tb(driver,web_link,'part_tb',file_name)
wp.auto_import_data_tb(driver,web_link,'rcode_tb',file_name)
wp.auto_import_data_tb(driver,web_link,'pending_tb',file_name)
try:
	log_out = driver.find_element_by_xpath('//*[@id="user-tools"]/a[3]')
	log_out.click()
	print('\nLogout successful.')
except:
	pass
# print('\n_select_comment_tb')
# wp.import_data_tb(driver,web_link,'comment_tb','xlsx')
