import sources.logins as lg
import sources.databases as db
import sources.reports as rp
import sources.new_web_231125 as nw
import sources.update_ml as uml
import sources.qr_code as qrc

import os,sys
import pathlib
import shutil
from time import sleep
from datetime import datetime as dt

from sqlite3 import connect
import pandas as pd
from IPython.display import display

# pyinstaller -i neverstop.ico -n Endocare --onefile main.py



class chel():

	def __init__(self):
		
		self.menu_list =[
					f'Login ExFM ({uname})',
					f'Select Database ({db_name})',
					'Update Master List',
					'Quotation',
					'GDKT/Trouble Report',

					'Weekly Report',
					'Run SQL',
					'Print QR Code',
					'Exit',
		]


	def menu(self): #return index
		while True:
			global driver, uname
			menu_list = self.menu_list
			#border table
			print(f'\n{"_"*50}')
			print(f'{"|  No.|  Function": <49}|')
			print(f'|{"_"*48}|')
			for i in range(1,1+len(menu_list)):
				print(f'|{i: >3}  |  {menu_list[i-1]: <40}|')
			print(f'|{"_"*48}|') #bottom border
		
			# next_step = False
			ind = str(input('Select Function by Index: '))
			web_key = ['noisoifujifilm',
						'website',
						'noisoi',
						'endo',
						'endocare',

					]
			login_key = ['login',
						'sign-in',
						]
			try:
				ind = int(ind)
			except:
				if ind.upper() == 'Q' or ind.upper() == 'QUIT':
					break
				elif ind.lower() in login_key:
					# global driver, uname
					driver,uname = functions().DownloadExFM()
					continue

				elif ind.lower() in web_key:
					print('Add data to website: noisoifujifilm.vn')
					try:
						if conn:
							pass
					except:
						conn = connect('quotation.db')
						print('Auto connect to history.db')
					run_all = nw.data_process(conn)
					# run_all.rma_list()
					run_all.exfm_web()
					run_all.pending_file()
					run_all.parts_name()
					run_all.r_code()
					# try:
					# 	run_all.export_files()
					# except Exception as e:
						# print(e)
					run_all.export_csv()
					import sources.auto_import

					continue
				elif ind.lower() =='adminstrator':
					uname = 'Admin'
					print('Unlock sucessfull Admin mode')
					continue

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
		global conn,driver,uname
		# if function == menu_list[0]: # Login ExFM
		if 'login' in function.lower():

			driver,uname = functions().DownloadExFM()
			return	driver,uname	

		# elif function == menu_list[1]: # Select Database
		elif 'database' in function.lower(): # Select Database
			return functions().SelectDatabase()

		elif function == 'Run SQL':
			functions().run_sql()

		elif function == 'GDKT/Trouble Report':
			try:
				rp.main(conn)
			except Exception as e:
				print(e,'\nSelect Database (Step 2 first)')
		
		elif function == 'Update Master List':
			
			try:
				monday = uml.update_ml(conn)
				monday.new_job()
				monday.update_job()
				monday.empty_status()
				monday.export()
			except Exception as e:
				print(e,'\nSelect Database (Step 2 first)')
		
		elif function == 'Quotation':
			try:
				rp.quotation(conn)
			except Exception as e:
				print(e,'\nSelect Database (Step 2 first)')

		elif function == 'Weekly Report':
			try:
				swr = rp.weekly_report(conn)
				swr.receive()
				swr.inspection()
				swr.export()
			except Exception as e:
				print(e,'\nSelect Database (Step 2 first)')

		elif function == 'Print QR Code':
			# try:
			abba = qrc.main(folder_name='images',wb_name = 'templates/QR Template.xlsx')
			abba.write_template()


		elif function == menu_list[-1]:
			pass

		else:
			print(function)


class functions():
	def __init__(self):
		
		self.uname = uname

	def DownloadExFM(self):
		# check driver
		global driver
		try:
			driver_bk = driver
			driver,d_type = lg.login_exfm(uname = self.uname,driver = driver_bk)

		except:
			driver,d_type = lg.login_exfm(uname = self.uname,driver='')

		

		
		
		if driver =='':
			driver = driver_bk
		uname = driver.find_element_by_xpath('//*[@id="username"]').get_attribute('innerHTML')
		

		# Rename for Download All
		if d_type.lower() =='history':
			folder_name = 'Downloads'
			file,ctime = lg.file_latest(folder_name)
			today = dt.now().strftime('%y%m%d')
			new_name = f'{file.split(".")[0]}_{today}.xls'
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

		# remove download folder
		try:
			shutil.rmtree(os.path.join(pathlib.Path().absolute(),'Downloads'),ignore_errors = True)
		except Exception as e:
			print(f'\n{e}')
			
		return driver,uname

	def SelectDatabase(self):
		global db_name,conn
		call_db = db.databases(uname)
		if uname in ('Le Quang Thong','Admin'):
			db_name = call_db.select_db_name()
		else:
			db_name = 'quotation.db'
		conn = call_db.update_db(db_name)
		return conn
		
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

					lines.append(user_input + '\n')
					if ';' in user_input: break
					if 'quit()' in user_input: break  # inner loop
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

if __name__ == "__main__":
	global conn, driver,db_name,ver,uname
# pyinstaller -i neverstop.ico -n Endocare --onefile main.py

	ver = '3.1.2'
	db_name =''
	uname =''


	print (f'{"-"*17}Endo Care v{ver}{"-"*17}')
	while True:
		print(db_name)
		name = chel().menu()
		print(name,uname)
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
				
				
			elif name == 'Exit': 
				print('Thank you!!! End of Game.')
				sleep(1)
				break
			
			else:
				chel().processing(name)
		
	sys.exit()


