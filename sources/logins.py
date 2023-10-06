from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

#---import alert-----------
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#-----------------------------------------

from time import sleep
import pathlib
import os
from pathlib import Path
import datetime
from datetime import datetime as dt
import shutil
# import base64
import sources.chrome_version as cv
from getpass import getpass
from IPython.display import display

from cryptography.fernet import Fernet

#  ---Version 3.4.2 14-Feb-2023
# none encrypte username
# modify encode and decode using fernet

#  ---Version 3.4.1 02-Feb-2023
# Add border to file select and file filter


#  ---Version 3.3.1 02-Dec-2022
# Add contains to file select

#  ---Version 3.2.1 21-Sep-2022
# Add time to file select
# Add download consolidated

#  ---Version 3.1.0
# Date: 220823
# modify save_id for first user
# encrypt password 220817


def backups(folder_name,dest,ext='.xlsx'):
	for file in os.listdir(folder_name):
		if file.endswith(ext):
			file_name = os.path.join(folder_name,file)
			# ctime = datetime.datetime.fromtimestamp(pathlib.Path(os.path.join(folder_name,file_name)).stat().st_ctime).strftime('%d-%b-%y %H:%M')
			ctime = datetime.datetime.fromtimestamp(pathlib.Path(file_name).stat().st_ctime).strftime('%d-%b-%y %H:%M')
			mtime = dt.fromtimestamp(os.path.getmtime(file_name)).strftime('%d-%b-%y %H:%M')
			dest_name = os.path.join(dest,file)
			print(f'{file}    {mtime} was backup to {dest}')
			shutil.copy(file_name,dest_name)

			
def pw_encode(message,key):
	fernet = Fernet(key.encode())
	return fernet.encrypt(message.encode()).decode()


def pw_decode(encrypted_message,key):
	fernet = Fernet(key.encode())
	return fernet.decrypt(encrypted_message.encode()).decode()

def save_id(file,key):
	if not os.path.exists(file): # chua co file saved.txt
		username = input('username: ')
		pw = getpass('Password: ')
		u1 = str(username)
		u2 = str(pw_encode(pw,key))
		L = f'{u1},{u2}\n' #u1 + ',' + u2
		file1 = open(file,'w')
		file1.write(L)
		file1.close()
		return username,pw

	else:
		file1 = open(file,'r')
		data = file1.read()
		file1.close()
		lines = data.split('\n')
		i_line = 0
		print('\nSelect account:')
		for line in lines[:-1]:

			i_line += 1
			uid = str(line.split(',')[0])
			pw = line.split(',')[1]
			print(f'{i_line}   |   {uid}')
		print ('0   |   CREATE NEW ACCOUNT')
		while True:
			try:
				uid = int(input('Select account by number: '))
				if uid < len(lines):
					break
				else:
					print(f'Vui long nhap so nho hon {len(lines)}')
			except:
				print(f'Chi duoc nhap so.')
		if uid == 0:
			print(f'\nCreate new account to {file}')
			while True:
				username = input('username: ')
				if username !='':
					break
					print('Checkpoint')
				else:
					print(f'{username} cannot empty.')

			pw = getpass('Password: ')
			u1 = str(username)
			u2 = str(pw_encode(pw,key))
			L = u1 + ',' + u2
			file1 = open(file,'a')
			file1.write(f'{L}\n')
			file1.close()
			print(f'Saved and selected {username} to login')
			return username,pw
		else:
			name = str(lines[uid-1].split(',')[0])
			print(f'{name} is selected to login')
			return name, pw_decode(lines[uid-1].split(',')[1],key)



def hint():
	print("\nfile_latest(folder_name='')")
	print("\nfile_select(end_with='',folder_name='')")
	print("\nlogin_url(target,credential,folder_name='')")
	print("\nfile_check(end_src,path='')")
	print("\nfile_filter(end_with='',path='',printer = True)")
	print("\nduration_process(s_time ='' )")
	# print("\n")
	
def check_driver(driver_name = 'chromedriver.exe'):
	while True:
		try:
			driver = webdriver.Chrome(driver_name)
			print('Chrome Driver matched.')
		#     driver.close()
			return driver
			break
		except:
			chrome_version = str(cv.get_chrome_version()).split('.')[0]
			print(f'Access https://sites.google.com/chromium.org/driver/downloads?authuser=0 and download version {chrome_version}.x.xxxx.xx')
			
			sleep(3)

def file_latest(folder_name=''):
	created_file ={}
	path = os.path.join(pathlib.Path().absolute(),folder_name)
	for file in os.listdir(path): # Read all file in folder_name
		fname = pathlib.Path(os.path.join(folder_name,file))
		ctime = fname.stat().st_ctime 
		created_file.update({ctime:file})
	file_name = created_file[max(created_file.keys())] # select latest file
	# formate datetime
	ctime = datetime.datetime.fromtimestamp(pathlib.Path(os.path.join(folder_name,file_name)).stat().st_ctime).strftime('%d-%b-%y %H:%M')
	return file_name,ctime

def file_filter(start_with='',end_with='',path='',printer = True):
	paths = pathlib.Path().absolute()
	if path =='':
		path = paths
	else:
		path = Path(os.path.join(paths,path))
	
	filepaths={}
	i = 0
	if printer:
		# print('Index',' | ','Names')
		# print('--------------------------')
		#Table Header
		print(f'\n{"_"*50}')
		print(f'{"|  No.|  File Name": <49}|')
		print(f'|{"_"*48}|')

	for file in os.listdir(path):
		if file.startswith(start_with) and file.endswith(end_with):
			i += 1
			filepaths.update({i:file})
			if printer:
				# print('  ',i,'  | ',file)
				print(f'|{i: >3}  |  {file: <40}|')
	if printer: print(f'|{"_"*48}|') #bottom border
	return filepaths


def file_select(start_with='',end_with='',contains='',not_contains='fdsfdsfasdrews',path='',folder_name=''):

	if path =='': path = pathlib.Path().absolute()
	if folder_name !='':
		path = os.path.join(path,folder_name)
	filepaths={}
	i = 0

	#Table Header
	print(f'\n{"_"*70}')
	print(f'{"|  No.|  File Name": <49}|  {"Modified Time": <17}|')
	print(f'|{"_"*68}|')
	for file in os.listdir(path):
		if file.startswith(start_with) and file.endswith(end_with):
			if contains in file and not_contains not in file:
				i += 1
				ctime = datetime.datetime.fromtimestamp(pathlib.Path(os.path.join(folder_name,file)).stat().st_ctime).strftime('%d-%b-%y %H:%M')
				mtime = dt.fromtimestamp(os.path.getmtime(os.path.join(folder_name,file))).strftime('%d-%b-%y %H:%M')
				
				filepaths.update({i:file})
				space = 50- len(file)
				# print(i,' ',file,' '*space,mtime)
				print(f'|{i: >3}  |  {file: <40}|  {mtime: <17}|')
	print(f'|{"_"*68}|') #bottom border
	
	if len(filepaths)>0:
		file = str(input('Select file (Default 1): ') or '1')
		try:
			file = int(file)
		except:
			pass
		if isinstance(file,int):
			file = filepaths[file]
			
		print(str('"') + str(file) + str('" is selected'))
		
	else:
		print('\nNo file to select.')
	if folder_name !='':
		# print(path)
		file = os.path.join(path,file)

	return file

def login_url(target,user_name,pass_word,folder_name=''):    

	#ver 3.0.0 edit with encrypt
	# add urls, xpath for id, password


	url_dict= {'facebook':('https://facebook.com','//*[@id="email"]','//*[@id="pass"]','send'),
		  'esp':('https://es-portal-jp.fujifilm.co.jp/login/login.html','//*[@id="item_email"]','//*[@id="item_password"]','//*[@id="wrap"]/div/div[1]/div[2]/form/div/button'),
		  'exfm':('https://exfm-asia-app.fujifilm.co.jp/page/exfmLogin.jsp','//*[@id="user"]','//*[@id="pass"]','button1')
		  }

	
	
	
	global driver
	url = url_dict[target][0]
	xpath_id = url_dict[target][1]
	xpath_pw = url_dict[target][2]
	bt_login = url_dict[target][3]

	
	if folder_name =='': folder_name = 'Downloads' 

	#-----------------------
	# del_ans = str(input('Do you want to remove folder "{}"?(Y/N): '.format(folder_name) or 'Y'))

	# -------- modify only for auto quotation
	del_ans = 'Y'

	if del_ans.upper() =='Y':
		shutil.rmtree(os.path.join(pathlib.Path().absolute(),folder_name),ignore_errors = True)
		print ('Removed {}\n'.format(folder_name))
		sleep(0.5)
	else:
		print('Still keep folder "{}"\n'.format(folder_name))

	chromeOptions = webdriver.ChromeOptions()
	prefs = {"download.default_directory" : str(os.path.join(pathlib.Path().absolute(),folder_name)),}
			
	chromeOptions.add_experimental_option("prefs",prefs)
	
	if not os.path.exists(folder_name):
		os.makedirs(folder_name)
		print('\nCreate new folder: {}'.format(folder_name))
	
	print('Select Chrome Driver and press enter for authorize Certificate...\n')
	driver = webdriver.Chrome(options = chromeOptions)
	driver.maximize_window()
	driver.get(url)

	sleep(0.5)
	# driver.find_element_by_tag_name('body').send_keys(Keys.RETURN)
	user = driver.find_element_by_xpath(xpath_id) 
	
	user.send_keys(user_name)
	print ('Typing your ID...')
	sleep(0.5)

	pw = driver.find_element_by_xpath(xpath_pw)  # '//*[@id="item_user_pw"]' #ESP
	pw.send_keys(pass_word)
	print ('Typing your password...')
	sleep(0.5)
	pw.send_keys(Keys.RETURN)
	try:
		sleep(2)
		driver.find_element_by_class_name(bt_login).click()
		print ('You Missed ID or Password' + '\n')

	except:
		print('Login sucessful!\n')
	return driver

def file_check(end_src,path=''):
	if path == '':
		path = pathlib.Path().absolute()
	else:
		path = pathlib.Path().absolute()  + path
	filecheck = False
	
	for file in os.listdir(path):
		if file.endswith(end_src):
			if file == 'chromedriver.exe':
				filecheck = True
				print ('Chrome Driver is ready to use.\n')
				break

	if filecheck == False:
		print('Please access to download Web Driver https://sites.google.com/a/chromium.org/chromedriver/downloads')
		input('Press any key to continue...')





def duration_process(s_time ='' ):
	from datetime import datetime
	
	if s_time =='':
		now = datetime.now()
		start_time = now.strftime("%H:%M:%S")
		print('Start Time: ',start_time)
		return start_time
	
	now = datetime.now()
	end_time = now.strftime("%H:%M:%S")
#    print('Start Time: ',s_time)
	print('End time: ',end_time)

	tdelta = datetime.strptime(end_time, '%H:%M:%S') - datetime.strptime(s_time, '%H:%M:%S')
	m_time = int(tdelta.seconds/60)
	s_time = tdelta.seconds % 60
	print('Duration: {}m {}s'.format(m_time,s_time))

def download_consolidated(status='incompleted'):
	# generate key using fernet
	# key = Fernet.generate_key().decode()
	key = 'cLAiAdZU1U0sxWWK8sxhF0IWIBP4KrR3xhdu3x1EDQI='

	# login ExFM and download latest incompleted data
	name,pw = save_id('exfm.txt',key)
	driver = login_url('exfm',name,pw)
	while True:
		try:
			driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]').click() # select RMA Search	
			break
		except Exception as e:
			print ('Wrong Username or Password. Try again...')
			name,pw = save_id('exfm.txt',key)
			driver = login_url('exfm',name,pw)
			
	# driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]').click() # select RMA Search
	if status =='all':
		Select(driver.find_element_by_xpath('//*[@id="sidIN_REPAIR_STATUS"]')).select_by_index(0)
		print('Select Status "All" and download.')
	
	driver.find_element_by_xpath('//*[@id="sidEXPORT_CONSOLIDATED_BUTTON_IMAGE"]').click() #click download
	
	# check latest file
	i_wait = 0
	while True:
		try: # check file already
			file, ctime = file_latest(folder_name='Downloads')
			if file.endswith('xls'):
				print (f'\nDonwload file {file} succesful at {ctime}')
				break
			else:
				print(f'Please wait for compressing file {file}...{i_wait}s')
				i_wait +=3
				sleep(3)
		except Exception as e: #waiting for download
			print(f'Please wait for dowloading...{i_wait}s')
			i_wait +=3
			sleep(3)
	#finish close driver
	driver.close()
	sleep(2)

def login_exfm(uname,driver,driver_on = True):
	# generate key using fernet
	# key = Fernet.generate_key().decode()
	key = 'cLAiAdZU1U0sxWWK8sxhF0IWIBP4KrR3xhdu3x1EDQI='
	
	if uname == '':
		# login ExFM and download latest incompleted data
		name,pw = save_id('exfm.txt',key)
		driver = login_url('exfm',name,pw)
		
		
		while True:
			try:
				driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]').click() # select RMA Search	
				break
			except Exception as e:
				print ('Wrong Username or Password. Try again...')
				name,pw = save_id('exfm.txt',key)
				driver = login_url('exfm',name,pw)
	else:
		print(f'Already login with user name {uname}')
		# driver = ''

	d_type_menu = [
					'Incompleted',
					'History',
					'Equipments',
					'Customers',
					'Do not Download'
					]
	#border table
	print(f'\n{"_"*50}')
	print(f'{"|  No.|  Function": <49}|')
	print(f'|{"_"*48}|')
	for i in range(1,1+len(d_type_menu)):
		print(f'|{i: >3}  |  {d_type_menu[i-1]: <40}|')
	print(f'|{"_"*48}|') #bottom border
	while True:
		ind = str(input('Select Dataset to download: '))
		try:
			ind = int(ind)
		except:
			if ind.upper() == 'Q' or ind.upper() == 'QUIT':
				break
			else:
				print('Only accept number')
				continue
		if 0<ind<=len(d_type_menu):
			d_type = d_type_menu[ind-1]
			break
		else:
			print(f'Input number must be less than {len(d_type_menu)}\n')
	
	
	
	# driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]').click() # select RMA Search
	
	if d_type == d_type_menu[0]: # Incompleted
		driver.find_element_by_xpath('//*[@id="sidEXPORT_CONSOLIDATED_BUTTON_IMAGE"]').click() #click download
	
	if d_type == d_type_menu[1]: # History
		Select(driver.find_element_by_xpath('//*[@id="sidIN_REPAIR_STATUS"]')).select_by_index(0)
		print('Select Status "All" and download.')
		driver.find_element_by_xpath('//*[@id="sidEXPORT_CONSOLIDATED_BUTTON_IMAGE"]').click() #click download
	
	if d_type == d_type_menu[2]: # Equipments
		driver.find_element_by_xpath('//*[@id="EQP_MGT_LINK"]').click() # Equipments
		sleep(0.5)
		driver.find_element_by_xpath('//*[@id="sidEXPORT_BUTTON_IMAGE"]').click()
	
	if d_type == d_type_menu[3]: # Customers
		driver.find_element_by_xpath('//*[@id="ACC_MGT_LINK"]').click() 
		sleep(0.5)
		driver.find_element_by_xpath('//*[@id="sidEXPORT_BUTTON_IMAGE"]').click()
	
	if d_type != d_type_menu[4]: # Do not download
		print(d_type)
		# check latest file
		i_wait = 0
		while True:
			try: # check file already
				file, ctime = file_latest(folder_name='Downloads')
				if file.endswith('xls'):
					print (f'\nDonwload file {file} succesful at {ctime}')
					break
				else:
					print(f'Please wait for compressing file {file}...{i_wait}s')
					i_wait +=3
					sleep(3)
			except Exception as e: #waiting for download
				print(f'Please wait for dowloading...{i_wait}s')
				i_wait +=3
				sleep(3)
	else:
		# no download
		pass 
		
	if driver_on:
		return driver,d_type
	else:
		#finish close driver
		driver.close()
