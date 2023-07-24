import os
from time import sleep
import warnings
import pandas as  pd
from sqlite3 import connect
import datetime
import openpyxl
import requests
from io import BytesIO
from IPython.display import display
import pathlib
warnings.simplefilter("ignore")

ver = '2.1.0'
#2.1.1 storages db in google sheets
# 2.0.0 all databases
# 1.0.2 add customers
# 1.0.1 exact '.db' add database
def version():
	print(f'Create Database Version {ver}')

def hint():
	print("\ncreate_db(db_name,file,type='consolidated/installation/database')")
def create_db(db_name,file='',f_type='consolidated'):
	print(db_name)
	if db_name[-3:] != '.db': # change from  1.0.1
		conn = connect(db_name +'.db')
	else: conn = connect(db_name)

	# 	print(f'database {db_name} was created.')
	con = file
	ins = file
	
	if f_type.upper() == 'CONSOLIDATED':
		try:
			df = pd.read_excel(con,sheet_name=None)
			c = df['Consolidated'].drop(['OWNERSHIP'],axis=1)
			c.to_sql('consolidated',conn,index=False,if_exists='replace')
			(df['PF-Code']).to_sql('pf_code',conn,index=False,if_exists='replace')
			(df['CD-Code']).to_sql('cd_code',conn,index=False,if_exists='replace')
			(df['R-Code']).to_sql('repair_code',conn,index=False,if_exists='replace')
			(df['Parts']).to_sql('parts',conn,index=False,if_exists='replace')
			print(f'\nConsolidated file stored in {db_name}')
		except Exception as e:
			print(f'can not find file {con} and import as consolidated')
			print(e)
	elif f_type.upper() == 'INSTALLATION':
		try:
			df1 = pd.read_excel(ins,sheet_name=None)
			(df1['csvdata']).to_sql('install',conn,index=False,if_exists='replace')
			print(f'\nInstallation file stored in {db_name}')
		except:
			print(f'can not find file {ins} and import as installation data')
	
	elif f_type.upper() =='DATABASE': #2.1.0
		try:
			spreadsheetId = '1bT4W0CiLVD_B_ddRcVkS3MEXbSjvmGb4'
			url = "https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=" + spreadsheetId
			res = requests.get(url)
			data = BytesIO(res.content)
			xlsx = openpyxl.load_workbook(filename=data)
			for name in xlsx.sheetnames:
				values = pd.read_excel(data, sheet_name=name)
				values.to_sql(name,conn,index=False,if_exists='replace')
			print(f'\nDatabase file stored in {db_name}')
		except:
			print(f'can not find file {ins} and import as database data')
	
	

	elif f_type.upper() == 'M_LIST': 
		filename = file
		# m_list
		m_list = pd.read_excel(filename,sheet_name='1.MasterPendingList',skiprows=range(1,3))
		new_header = m_list.iloc[0] #grab the first row for the header
		m_list = m_list[1:] #take the data less the header row
		m_list.columns = new_header

		#completed
		completed = pd.read_excel(filename,sheet_name='2. Completed',skiprows=range(1,1))
		new_header = completed.iloc[0] #grab the first row for the header
		completed = completed[1:] #take the data less the header row
		completed.columns = new_header

		#transfer
		transfer = pd.read_excel(filename,sheet_name='3. Transfer to sales',skiprows=range(1,1))
		new_header = transfer.iloc[0] #grab the first row for the header
		transfer = transfer[1:] #take the data less the header row
		transfer.columns = new_header

		m_list.to_sql('m_list',conn,index=False,if_exists='replace')
		completed.to_sql('completed',conn,index=False,if_exists='replace')
		transfer.to_sql('transfers',conn,index=False,if_exists='replace')
		return conn

	else:
		print('\nWrongf_type. Please select f_type as below:')
		print('Consolidated,installation')
	# file to SQL



	# check table in memory
	# display(pd.read_sql('SELECT name from sqlite_master where type= "table";',conn))
	conn.close()
