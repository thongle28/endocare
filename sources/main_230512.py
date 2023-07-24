import sources.logins as lg
import sources.databases as db 
import os

from sqlite3 import connect
import pandas as pd
from IPython.display import display

global conn, driver,db_name,ver 

ver = '1.0.2'

db_name =''
uname =''
class chel():

	def __init__(self):
		
		self.menu_list =[
					f'Login ExFM ({uname})',
					f'Select Database ({db_name})',
					'Update Master List',
					'Auto QC',
					'GDKT/Trouble Report',
					'Add Image Report',
					'Run SQL',
					'Exit',
		]
		# can not change order
		# self.menu_list.sort()




	def menu(self): #return index
		menu_list = self.menu_list
		#border table
		print(f'\n{"_"*50}')
		print(f'{"|  No.|  Function": <49}|')
		print(f'|{"_"*48}|')
		for i in range(1,1+len(menu_list)):
			print(f'|{i: >3}  |  {menu_list[i-1]: <40}|')
		print(f'|{"_"*48}|') #bottom border
		while True:
			ind = str(input('Select Function by Index: '))
			try:
				ind = int(ind)
			except:
				if ind.upper() == 'Q' or ind.upper() == 'QUIT':
					break
				else:
					print('Only accept number')
					continue
			if 0<ind<=len(menu_list):
				return menu_list[ind-1]
				break
			else:
				print(f'Input number must be less than {len(menu_list)}\n')

	def processing(self,function):
		menu_list = self.menu_list
		
		if function == menu_list[0]: # Login ExFM

			driver,uname = functions().DownloadExFM()
			return	driver,uname	

		if function == menu_list[1]: # Select Database
			return functions().SelectDatabase()

		if function == 'Run SQL':
			functions().run_sql()

		if function == menu_list[-1]:
			pass

		else:
			print(function)

	def update_ml(self,driver):
		pass

class functions():
	def __init__(self):
		
		self.d_type_menu = [
						'Incompleted',
						'History',
						'Equipments',
						'Customers',
						]


	def DownloadExFM(self):
		
		# check driver
		driver = lg.check_driver()
		try:
			driver.close()
		except Exception as e:
			pass

		d_type_menu = self.d_type_menu

		
		driver,d_type = lg.login_exfm()
		uname = driver.find_element_by_xpath('//*[@id="username"]').get_attribute('innerHTML')


		# Rename for Download All
		if d_type.lower() =='history':
			folder_name = 'Downloads'
			file,ctime = lg.file_latest(folder_name)
			new_name = f'{file.split(".")[0]}_All.xls'
			fname = os.path.join(folder_name,file)
			n_name =  os.path.join(folder_name,new_name)
			try:
				os.rename(fname,n_name)
			except Exception as e:
				print (e)
		try:
			folder_name = 'files'
			os.mkdir(folder_name)
			print(f'folder {folder_name} was created.')
		except:
			print(f'Folder {folder_name} exists')
		lg.backups('Downloads',folder_name,'.xls')
		return driver,uname

	def SelectDatabase(self):
		global db_name
		db_name = db.databases().select_db_name()
		conn = db.databases().update_db(db_name)
		return conn
		# if not update_type == None: print(update_type)

	def run_sql(self):
		try :
			print(conn)
			
		except:
			print('Select Database first.(#2)')
			return

		while True:
			try:
				print('\nWrite Query here. End with ";"')
				lines = []

				while True:
					user_input = input('>>>')

					# 👇️ if user pressed Enter without a value, break out of loop
					lines.append(user_input + '\n')
					if ';' in user_input: break
					if 'quit()' in user_input: break  # inner loop
				# 👇️ join list into a string
				# print(''.join(lines))
				q = ''.join(lines)
				print (q.upper())
				if 'ALL TABLES' in q.upper():
					print('SELECT name from sqlite_master where type= "table";')
					display(pd.read_sql('SELECT name from sqlite_master where type= "table";',conn))
				else:
					display(pd.read_sql(q,conn))
			except Exception as e:
				print(e)
				q ='quit()'

			if 'quit()' in q:
				break # outer loop


# chel = chel()
# conn = None
while True:
	print(db_name)
	name = chel().menu()
	if name == None: 
		print('Thank you')
		break
	else:
		if not name == None: print(f'Select function "{name}"')
		
		print('-'*30)
		if name == 'Login ExFM ()':
			driver,uname = chel().processing(name)
		elif name =='Select Database ()':
			conn = chel().processing(name)
			# print(conn)
			q = 'SELECT * FROM consolidated'
			# display(pd.read_sql(q,conn))
		elif name == 'Exit': 
			print('Thank you!!! End of Game.')
			break
		elif name == 'Update Master List':
			print(conn)

		else:
			chel().processing(name)
	

