import sources.create_db as cd
import sources.logins as lg
import pandas as pd
from datetime import datetime as dt
import xlwings as xw
from sqlite3 import connect
from IPython.display import display
import codecs # Python standard library
import xlwings as xw

import os
import pathlib
from math import ceil
import qrcode


class parts_list():
	'''Tao part list cho UI 1.0.0'''

	def __init__(self):
		pass

	def update(self,db_name,file_name,folder_name='files',):
		# add database file
		print('\nSelect database.xlsx file:')
		# db_file = lg.file_select(end_with = '.xlsx',folder_name = 'files')
		cd.create_db(db_name,'filein Google sheets',f_type = 'database')

		# add price files .xlsx
		print('\nSelect Price List file:')
		conn = connect(db_name)
		# price_file = lg.file_select(end_with = '.xlsx',folder_name = 'files')
		price_file = 'files\\New Price -FFVN-pricelist-01sep21 rev.xlsx'
		pr = pd.read_excel(price_file,sheet_name = 'Subsidary-06aug21')
		pr.to_sql('prices',conn,index = False, if_exists = 'replace')

		# consolidated
		print('\nSelect SearchResultConsolidated.xls file:')
		# file_name = lg.file_select(end_with ='.xls',folder_name = 'files')
		# file_name = f'{folder_name}\\{file_name}'
		cd.create_db(db_name,file_name,f_type = 'consolidated')

		return conn

	def update_time(self,conn):
		q='''
			SELECT strftime('%d/%m/%Y  %H:%M',max([UPDATE TIME])) as [latest], strftime('%d/%m/%Y  %H:%M',min([UPDATE TIME])) as oldest,
        max([rma no.]) as rma_max FROM consolidated
			'''
		udt = pd.read_sql(q,conn)
		
		return udt
	def check_db(self,db_name):
		conn = connect(db_name)
		return conn
				