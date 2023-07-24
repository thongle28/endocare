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
from PIL import Image

class parts_list():
	'''Tao part list tu RMA ver 3.1.0
	Them hang print tilte GDKT
	Them GDKT tu dong
	Them def list_rma'''


	def __init__(self):
		self.rma = '' 
		self.db_name = f"quotation_{dt.now().strftime('%y%m%d')}.db"
		# print(f'auto create db {self.db_name}')
		self.ans_update = ''
		self.service_fee = ''
		# global conn


	def list_rma(self,conn,sn_list,delimeter='\n'):
		sn_list =sn_list.split(delimeter)
		
		for sn in sn_list:
			if sn!='':
				q=f'''

					SELECT c.[rma no.] AS rma,c.customer_name,c.serial_no,c.model,c.approval,c.repair_status,c.in_inspect_user_name

					FROM consolidated c
					WHERE upper(c.serial_no) = '{sn.upper()}'
					ORDER BY rma DESC
					LIMIT 1
				'''
				if sn == sn_list[0]:
					list_rma = pd.read_sql(q,conn)
				else:
					list_rma = list_rma.append(pd.read_sql(q,conn),ignore_index=True)
		return list_rma

	def select_db(self):
		print('\nSelect Database: ')
		db_name = lg.file_select(end_with='.db')
		if db_name[-3:] != '.db': # change from  1.0.1
			conn = connect(db_name +'.db')
		else: conn = connect(db_name)
		pd.read_sql('select name from sqlite_master where type="table"',conn)
		self.conn = conn
		return conn,db_name

	def question(self):
		while True:
			db_name = self.db_name
			ans_update = str(input('Update data(y/N): ') or 'N')
			if ans_update.upper()=='MASTER':

				# add database file
				print('\nSelect database.xlsx file:')
				# db_file = lg.file_select(end_with = '.xlsx',folder_name = 'files')
				cd.create_db(db_name,f_type = 'database')

				# add price files .xlsx
				print('\nSelect Price List file:')
				conn = connect(db_name)
				price_file = lg.file_select(end_with = '.xlsx',folder_name = 'files')
				xl = pd.ExcelFile(price_file)
				i=1
				for sh in xl.sheet_names:  # see all sheet names
					print(i,'  |  ',sh)
					i+=1
				sh_index = int(input('Select file index: '))-1
				pr = pd.read_excel(price_file,sheet_name = xl.sheet_names[sh_index])
				pr.to_sql('prices',conn,index = False, if_exists = 'replace')

				# consolidated
				print('\nSelect SearchResultConsolidated.xls file:')
				file_name = lg.file_select(end_with ='.xls',folder_name = 'files')
				cd.create_db(db_name,file_name,f_type = 'consolidated')
				break

			elif ans_update.upper() =='DATABASE' or ans_update.upper() =='DB':
				# add database file
				# print('\nSelect database.xlsx file:')
				# db_file = lg.file_select(end_with = '.xlsx',folder_name = 'files')
				print('\nSelect database file from Google Sheet')
				cd.create_db(db_name,f_type = 'database')
				break

			elif ans_update.upper() == 'PRICE':
				# add price files .xlsx
				print('\nSelect Price List file:')
				conn = connect(db_name)
				price_file = lg.file_select(end_with = '.xlsx',folder_name = 'files')
				pr = pd.read_excel(price_file,sheet_name = 'Subsidary-06aug21')
				pr.to_sql('prices',conn,index = False, if_exists = 'replace')
				break

			elif ans_update.upper() == 'M_LIST' or ans_update.upper() =='ML':
				# Master List
				conn,db_name = self.select_db()
				print('\nSelect Master List Endo.xlsm file: ')
				file_name = lg.file_select(end_with='.xlsm',folder_name = 'files')
				cd.create_db(db_name,file_name,f_type = 'm_list')
				break
				
			elif ans_update.upper() == 'Y':
				# consolidated
				conn,db_name = self.select_db()
				print('\nSelect SearchResultConsolidated.xls file:')
				file_name = lg.file_select(end_with ='.xls',folder_name = 'files')
				cd.create_db(db_name,file_name,f_type = 'consolidated')
				print('\nSelect database file from Google Sheet')
				cd.create_db(db_name,f_type = 'database')
				break

			elif ans_update.upper() =='N':
				conn,db_name = self.select_db()
				print('\n Used exists data.')
				break

			else:
				print('Typing wrong answer.')
		self.conn = conn
		return conn

	def replace_part_number(self,part_list,conn):
		for i in range(len(part_list)):
			a = part_list.loc[i:i]
		#     display(a)
			if list(a['FFVN Price'].isnull())[0]:
				print(list(a['part_num'])[0])
				part_original = list(a['part_num'])[0]
				new_part = str(input(f'\nSelect new part number for "{part_original}": '))
				
			# find out family
				q=f'''
					SELECT [Article Code],Family
					
					FROM prices
					WHERE [Article Code] = '{new_part}'
					
				'''
				family = pd.read_sql(q,conn)
				
				# list of family
				q=f'''
					SELECT [Parts Name(EN)] AS [PART_DESCRIPTION],PN.VIE,
					CASE 
						WHEN substr(p.[Parts Name(EN)],4,1) IN ('Y','N','S') THEN 'F'
						WHEN substr(p.[Parts Name(EN)],1,1) = 'J' THEN 'F'
						ELSE 'FW12G' END AS SAP,
					P.[Article Code] as part_num,p.[Cost of Goods sold] as [FFVN Price],
					p.[Dealer Standard Price (wiout VAT)] as [Dealer Price]
					
					
					
					FROM prices P
					LEFT JOIN part_name pn ON TRIM(p.[Parts Name(EN)]) = pn.part_description
					WHERE Family = '{family['Family'][0]}'
					
				'''
				part_replace = pd.read_sql(q,conn)
				try:
					display(part_replace)
				except:
					pass

				while True:
					select_ind = input('Select Replace part number: ')
					if select_ind == '':
						select_ind = 0
						break
					
					elif int(select_ind) >len(part_list):
						print('Select wrong number')
					else:
						try:
							select_ind = int(select_ind)
							break
						except:
							print('Only number.')
				part_list['part_num'] = part_list['part_num'].replace([part_original],[part_replace['part_num'][select_ind]])
				part_list['FFVN Price'] = part_list['FFVN Price'].fillna(part_replace['FFVN Price'][select_ind])
				part_list['Dealer Price'] = part_list['Dealer Price'].fillna(part_replace['Dealer Price'][select_ind])
				part_list['SAP'] = part_replace['SAP'][select_ind]
		return part_list

	def part_list_final(self,rma,conn):
		rma = rma.upper()
		# query read part number from rma
		q=f'''
			SELECT DISTINCT p.part_description,pn.vie,
			
			CASE 
				WHEN substr(p.part_no,4,1) IN ('Y','N','S') THEN 'F'
				WHEN substr(p.part_no,1,1) = 'J' THEN 'F'
				ELSE 'FW12G' END AS SAP,
			CASE 
				WHEN substr(p.part_no,-1,1) = '_' THEN substr(p.part_no,-1,-20)
				ELSE p.part_no END as part_num,
			p.quantity,
			pr.[Cost of Goods sold] AS [FFVN Price],p.quantity*pr.[Cost of Goods sold] as [FFVN_AMOUNT],
			pr.[Dealer Standard Price (wiout VAT)] AS [Dealer Price],p.quantity*pr.[Dealer Standard Price (wiout VAT)] as [Dealer_AMOUNT],
			c.[RMA NO.] AS rma
			FROM (consolidated c LEFT JOIN parts p ON C.[RMA NO.] = p.[rma no.]
			LEFT JOIN part_name pn ON p.part_description = pn.part_description)
			LEFT JOIN prices pr ON part_num = pr.[Article Code]
			WHERE rma = '{rma}'

			ORDER BY [DEALER price] DESC
		'''
		part_list_1 = pd.read_sql(q,conn)
		if len(part_list_1) >0:
			# if part_list_1[part_list_1['PART_DESCRIPTION'].isnull()]>0:
		
			# create list empty price
			empty_price = part_list_1[part_list_1['Dealer Price'].isnull()]['part_num']
			empty_price_str = str(list(empty_price))[1:-1]
			try:
				display(part_list_1)
			except:
				pass
			# query looking for part name price
			q = f'''
				SELECT DISTINCT 
				pr.[article code],
				CASE 
					WHEN substr(p.part_no,4,1) IN ('Y','N','S') THEN 'F'
					WHEN substr(p.part_no,1,1) = 'J' THEN 'F'
					ELSE 'FW12G' END AS SAP,
				CASE 
					WHEN substr(p.part_no,-1,1) = '_' THEN substr(p.part_no,-1,-20)
					ELSE p.part_no END as part_num,
				CASE 
					WHEN substr(p.part_no,-1,1) = '_' THEN substr(p.part_no,-2,-20)
					ELSE substr(p.part_no,-1,-20) END as part_family,
				p.quantity,
				pr.[Cost of Goods sold] AS [FFVN Price],p.quantity*pr.[Cost of Goods sold] as [FFVN_AMOUNT],
				pr.[Dealer Standard Price (wiout VAT)] AS [Dealer Price],p.quantity*pr.[Dealer Standard Price (wiout VAT)] as [Dealer_AMOUNT],
				
				pr.[parts name(EN)],
				c.[RMA NO.] AS rma
				
				FROM (consolidated c LEFT JOIN parts p ON C.[RMA NO.] = p.[rma no.]
				LEFT JOIN part_name pn ON p.part_description = pn.part_description)
				LEFT JOIN prices pr ON part_family = pr.family
				WHERE rma='{rma}'
				AND part_num IN ({empty_price_str})
			  
				ORDER BY [Dealer price] DESC
			'''
			part_list_2 = pd.read_sql(q,conn)
			print(str(len(part_list_2))+ ' part(s) no price')
			
			if len(part_list_2) >0:
				for pn in empty_price:
					
					a = part_list_2[part_list_2['part_num']==pn]
					
					try:
						display(a.reset_index(drop=True))

					except:
						pass

					if len(a)==1:
						select_ind = 0
						# print('Case 1')
					
					else:
						# print('Case 3')
						select_ind = str(input('Select correct part number: ' or '0'))
						
					if select_ind =='': 
						try:
							select_ind = len(a) - 1
							print(len(a))
						except:
							select_ind = 0
					else: select_ind = int(select_ind)
					
					if pn == list(empty_price)[0]:
						
						part_list_3 = pd.DataFrame(a.iloc[[int(select_ind)]])
					else:
						
						part_list_3 = part_list_3.append(a.iloc[[int(select_ind)]])
				
				try:
					display(part_list_3)
				except:
					pass

				part_list_1.to_sql('exfm_part_list',conn,index=False,if_exists='replace')
				part_list_3.to_sql('parts_replace',conn,index = False,if_exists='replace')

				q='''
					SELECT epl.part_description,
					epl.vie,
					CASE 
						WHEN epl.sap is null then pr.sap
						ELSE epl.sap END AS SAP,
					CASE
						WHEN pr.[Article Code] IS NULL THEN epl.[part_num] 
						ELSE pr.[Article Code] END AS part_num,epl.quantity, 
					CASE 
						WHEN epl.[ffvn price] is null then pr.[FFVN pRICE]
						ELSE epl.[ffvn price] END AS [FFVN Price],
					CASE
						WHEN epl.[Dealer Price] is null THEN pr.[dealer price]
						ELSE epl.[dealer price] END AS [Dealer Price]
					
					FROM exfm_part_list epl
					LEFT JOIN parts_replace pr ON epl.part_num = pr.part_num
					ORDER BY [dealer price] DESC
					'''
				part_list_final = pd.read_sql(q,conn)
				try:
					display(part_list_final)
				except:
					pass
			else:
				part_list_final = part_list_1
			part_list_final = self.replace_part_number(part_list_final,conn)
			part_list_final = part_list_final.sort_values(by='FFVN Price',ascending=False)
		return part_list_final	

	def create_parts_list(self,conn,rma,service_fee,vnd=23700,folder_name='exports'):
		rma = rma.upper()
		# query read part number from rma
		q=f'''
			SELECT DISTINCT p.part_description,pn.vie,
			
			CASE 
				WHEN substr(p.part_no,4,1) IN ('Y','N','S') THEN 'F'
				WHEN substr(p.part_no,1,1) = 'J' THEN 'F'
				ELSE 'FW12G' END AS SAP,
			CASE 
				WHEN substr(p.part_no,-1,1) = '_' THEN substr(p.part_no,-1,-20)
				ELSE p.part_no END as part_num,
			p.quantity,
			pr.[Cost of Goods sold] AS [FFVN Price],p.quantity*pr.[Cost of Goods sold] as [FFVN_AMOUNT],
			pr.[Dealer Standard Price (wiout VAT)] AS [Dealer Price],p.quantity*pr.[Dealer Standard Price (wiout VAT)] as [Dealer_AMOUNT],
			c.[RMA NO.] AS rma
			FROM (consolidated c LEFT JOIN parts p ON C.[RMA NO.] = p.[rma no.]
			LEFT JOIN part_name pn ON p.part_description = pn.part_description)
			LEFT JOIN prices pr ON part_num = pr.[Article Code]
			WHERE rma = '{rma}'

			ORDER BY [DEALER price] DESC
		'''
		part_list_1 = pd.read_sql(q,conn)
		if len(part_list_1) >0:
			# if part_list_1[part_list_1['PART_DESCRIPTION'].isnull()]>0:
		
			# create list empty price
			empty_price = part_list_1[part_list_1['Dealer Price'].isnull()]['part_num']
			empty_price_str = str(list(empty_price))[1:-1]
			try:
				display(part_list_1)
			except:
				pass
			# query looking for part name price
			q = f'''
				SELECT DISTINCT 
				pr.[article code],
				CASE 
					WHEN substr(p.part_no,4,1) IN ('Y','N','S') THEN 'F'
					WHEN substr(p.part_no,1,1) = 'J' THEN 'F'
					ELSE 'FW12G' END AS SAP,
				CASE 
					WHEN substr(p.part_no,-1,1) = '_' THEN substr(p.part_no,-1,-20)
					ELSE p.part_no END as part_num,
				CASE 
					WHEN substr(p.part_no,-1,1) = '_' THEN substr(p.part_no,-2,-20)
					ELSE substr(p.part_no,-1,-20) END as part_family,
				p.quantity,
				pr.[Cost of Goods sold] AS [FFVN Price],p.quantity*pr.[Cost of Goods sold] as [FFVN_AMOUNT],
				pr.[Dealer Standard Price (wiout VAT)] AS [Dealer Price],p.quantity*pr.[Dealer Standard Price (wiout VAT)] as [Dealer_AMOUNT],
				
				pr.[parts name(EN)],
				c.[RMA NO.] AS rma
				
				FROM (consolidated c LEFT JOIN parts p ON C.[RMA NO.] = p.[rma no.]
				LEFT JOIN part_name pn ON p.part_description = pn.part_description)
				LEFT JOIN prices pr ON part_family = pr.family
				WHERE rma='{rma}'
				AND part_num IN ({empty_price_str})
			  
				ORDER BY [Dealer price] DESC
			'''
			part_list_2 = pd.read_sql(q,conn)
			print(str(len(part_list_2))+ ' part(s) no price')
			
			if len(part_list_2) >0:
				for pn in empty_price:
					
					a = part_list_2[part_list_2['part_num']==pn]
					
					try:
						display(a.reset_index(drop=True))

					except:
						pass

					if len(a)==1:
						select_ind = 0
						# print('Case 1')
					
					else:
						# print('Case 3')
						select_ind = str(input('Select correct part number: ' or '0'))
						
					if select_ind =='': 
						try:
							select_ind = len(a) - 1
							print(len(a))
						except:
							select_ind = 0
					else: select_ind = int(select_ind)
					
					if pn == list(empty_price)[0]:
						
						part_list_3 = pd.DataFrame(a.iloc[[int(select_ind)]])
					else:
						
						part_list_3 = part_list_3.append(a.iloc[[int(select_ind)]])
				
				try:
					display(part_list_3)
				except:
					pass

				part_list_1.to_sql('exfm_part_list',conn,index=False,if_exists='replace')
				part_list_3.to_sql('parts_replace',conn,index = False,if_exists='replace')

				q='''
					SELECT epl.part_description,
					epl.vie,
					CASE 
						WHEN epl.sap is null then pr.sap
						ELSE epl.sap END AS SAP,
					CASE
						WHEN pr.[Article Code] IS NULL THEN epl.[part_num] 
						ELSE pr.[Article Code] END AS part_num,epl.quantity, 
					CASE 
						WHEN epl.[ffvn price] is null then pr.[FFVN pRICE]
						ELSE epl.[ffvn price] END AS [FFVN Price],
					CASE
						WHEN epl.[Dealer Price] is null THEN pr.[dealer price]
						ELSE epl.[dealer price] END AS [Dealer Price]
					
					FROM exfm_part_list epl
					LEFT JOIN parts_replace pr ON epl.part_num = pr.part_num
					ORDER BY [dealer price] DESC
					'''
				part_list_final = pd.read_sql(q,conn)
				try:
					display(part_list_final)
				except:
					pass
			else:
				part_list_final = part_list_1
			part_list_final = self.replace_part_number(part_list_final,conn)
			part_list_final = part_list_final.sort_values(by='FFVN Price',ascending=False)
		
		# part_list_final = self.part_list_final(rma,conn)
			#--------------xlwings--------------------
			wb = xw.Book('templates\\Baogia_Template.xlsx')
			parts = wb.sheets('Parts')
			quotation = wb.sheets('Quotation')

			# rma information
			q = f'''
					SELECT DISTINCT c.[rma no.],c.model,c.[serial_no],sc.vie,cu.web_name,
					c.customer_name,cu.address,cu.addr1,
					e.ks,strftime('%d/%m/%Y',c.in_inspect_date) as inspect_date
					from ((consolidated c LEFT JOIN scopes sc ON c.model = sc.model)
					LEFT JOIN customers cu ON cu.[no.] = c.customer_code)
					LEFT JOIN engineers e ON c.in_inspect_user_name =e.exfm_name
					WHERE c.[rma no.] = '{rma}'
				'''
			info = pd.read_sql(q,conn)
			try:
				display(info)
			except:
				pass

			#return parts template
			for i in range(50):
				if parts.range('A' + str(i+6)).value == 'Total': break
				last_row = i+6
			if parts.range('A7').value != 'Total':
				delete_rows = f'7:{last_row}'
				parts.range(delete_rows).delete()
			parts.range('2:4').value =''
			parts.range('6:6').value =''
			parts.range('I6').value=''

			# fill in general infomation
			price_type = str(input('Price for FFVN/[DEALER]: ') or 'Dealer')
			print(f'Select price list for {price_type}\n')
			kieu_may = str(info['Vie'][0]).title()
			model = str(info['MODEL'][0]).upper()
			model = model.replace(' V2','')
			sn = str(info['SERIAL_NO'][0]).upper()
			web_name = str(info['web_name'][0]).title()
			engineer = str(info['ks'][0]).title()
			hospital_name = str(info['web_name'][0]).capitalize()
			inspect_date = info['inspect_date'][0]

			parts.range('B2').value = f'Báo Giá Sửa Chữa {kieu_may}_{model}_SN:{sn}'
			parts.range('B3').value = web_name
			parts.range('B4').value = engineer
			parts.range('C4').value = f'Date: {inspect_date}'
			parts.range('G4').value = price_type.capitalize()
			parts.range('I6').value = rma
			parts.range('F4').value = service_fee
			parts.range('h4').value = '1 USD ='
			parts.range('I4').value = vnd

			#part_list final
			if parts.range('A7').value == 'Total': #check empty table
				#insert new rows
				for i in range(len(part_list_final)-1):
					parts.range('7:7').insert('down')
				total_price = sum(part_list_final['Dealer Price'])
				# fill in parts informtion
				for i in range(len(part_list_1)):
					parts.range('A' + str(i+6)).value = i+1
					parts.range('B' + str(i+6)).value = part_list_final['PART_DESCRIPTION'][i] 
					parts.range('C' + str(i+6)).value = part_list_final['Vie'][i]
					
					parts.range('D' + str(i+6)).value = part_list_final['SAP'][i] + part_list_final['part_num'][i]
					parts.range('E' + str(i+6)).value = part_list_final['QUANTITY'][i]
					parts.range('F' + str(i+6)).value = f'=IF($F$4="","",$F$4*G{str(i+6)}/{total_price})'
					parts.range('H' + str(i+6)).value = f'=E{str(i+6)}*G{str(i+6)}+F{str(i+6)}'
					if price_type.title() == 'Dealer':
						parts.range('G' + str(i+6)).value = part_list_final['Dealer Price'][i]
						
					elif price_type.capitalize() == 'Ffvn':
						parts.range('G' + str(i+6)).value = part_list_final['FFVN Price'][i]
						parts.range('F4').value = 0
				parts.range('H'+str(i+7)).value = f'=sum(H6:H{str(i+6)})'
				price_cell = i+9
			else:
				print('table not empty')
				
			#save workbook
			customer_name=pd.read_sql(f'''SELECT customer_name from consolidated where [rma no.]='{rma}' ''',conn)
			customer_name = customer_name['CUSTOMER_NAME'][0].replace(' ','_')
			try:
				os.makedirs(folder_name)
			except:
				print('Directory exsists.')
			title =f'{folder_name}\\{rma}_BaoGia_{customer_name}_{model.replace("/","")}_{sn}.xlsx'
			wb.save(title)
			title_print = title.split('\\')[1]
			print(f'{title_print} export successful')
			wb.close()

		else:
			print(f'Cannot find part list for {rma}')

		return part_list_final, info,'I'+ str(price_cell), title


class quotation:
	'''Create Quotation from Parts List'''
	def __init__(self, part_list_final, info,price_addr, title):
		self.wb = xw.Book(title)
		self.info = info
		self.part_list_final = part_list_final
		self.price_addr = price_addr
		print(f'Open WorkBook {title}')

	def input_data(self):
		# try:
		info =self.info
		parts = self.part_list_final
		quotation = self.wb.sheets('Quotation')
		quotation.range('D6').value = dt.now()
		quotation.range('D7').value = f'''FFVN-{dt.now().strftime('%m')}.{dt.now().strftime('%Y')}/'''
		model = info['MODEL'][0]
		quotation.range('D8').value = model.replace(' V2','')
		quotation.range('D9').value = info['SERIAL_NO'][0]
		quotation.range('D10').value = info['web_name'][0]
		quotation.range('D11').value = info['CUSTOMER_NAME'][0]
		quotation.range('D12').value = info['address'][0]
		quotation.range('D13').value = info['Addr1'][0]
		quotation.range('F20').value =f'=CEILING.MATH(Parts!{self.price_addr}/1000000)*1000000'
		quotation.range('F24').value ='=F20'
		
		if quotation.range('A24').value == 'Tổng cộng/Total amount': #check empty table
			#insert new rows
			
			for i in range(len(parts)-1):
				quotation.range('24:24').insert('down')
				quotation.range('B24:C24').merge()
				quotation.range('F24:G24').merge()
				quotation.range('D23:E23').copy()
				quotation.range('D24:E24').paste('formats')

			for i in range(len(parts)):
				quotation.range('A' + str(i+23)).value = i+1
				quotation.range('B' + str(i+23)).value = str(parts['Vie'][i]) + '/' + parts['PART_DESCRIPTION'][i] 
				quotation.range('D' + str(i+23)).value = f''''0{parts['QUANTITY'][i]}'''
				quotation.range('E' + str(i+23)).value = 'Cái/Pcs'
				quotation.range('F' + str(i+23)).value = '(bao gồm/included)'			
		else:
			print('table not empty')

		# except:
			print('Can not find Quotation sheets.')

	def clear_quotations(self):
		try:
			quotation = self.wb.sheets('Quotation')
			for i in range(50):
				if quotation.range('A' + str(i+24)).value == 'Tổng cộng/Total amount': break
			#     print(i+6,parts.range('A' + str(i+6)).value)
				last_row = i+24
			if quotation.range('A24').value != 'Tổng cộng/Total amount':
				delete_rows = f'24:{last_row}'
				quotation.range(delete_rows).delete()
			quotation.range('23:23').value =''

		except:
			print('Can not find Quotation sheets.')

	def save_and_close(self):
		self.wb.save()
		self.wb.close()

class technical_report():
	'''
		Model width 180 ->200

	'''

	def __init__(self): #,info=[],report_num='',part_list=[]):
		# self.part_list = part_list
		# self.info = info
		# self.report_num = report_num
		pass

	def resize_height_image(self,sh,image_name,width_demand):
			ratio = sh.pictures(image_name).height/sh.pictures(image_name).width
			sh.pictures(image_name).height = ratio * width_demand
			sh.pictures(image_name).width = width_demand
			return ratio * width_demand

	def create_qr_image(self,info,image_folder='images'):
		Logo_link = 'logo\\FF_logo_border.png'
 
		logo = Image.open(Logo_link)
		 
		# taking base width
		basewidth = 350
		 
		# adjust image size
		wpercent = (basewidth/float(logo.size[0]))
		hsize = int((float(logo.size[1])*float(wpercent)))
		logo = logo.resize((basewidth, hsize), Image.ANTIALIAS)
		qr = qrcode.QRCode(
			version=1,
			error_correction=qrcode.constants.ERROR_CORRECT_H,
			box_size=16,
			border=0.1,
		)
		qr.add_data(f"https://noisoifujifilm.vn/quick_search/{info['RMA No.'][0]}/{info['SERIAL_NO'][0]}")
		qr.make(fit=True)
		img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
		pos = ((img.size[0] - logo.size[0]) // 2,
			   (img.size[1] - logo.size[1]) // 2)
		img.paste(logo, pos)
		img.save(f"{image_folder}\\{info['RMA No.'][0]}_{info['SERIAL_NO'][0]}.png")
		
	def signatures(self,tp,signatures_folder='signatures'):
		path = pathlib.Path().absolute()
		if tp.range('A50').value == 'Nguyễn Khắc Thắng':
			tp.pictures.add(os.path.join(path,signatures_folder,'Thang.png'),name = 'Thang',top = tp.range('G44').offset(1,0).top+7,left = tp.range('C44').offset(1,0).left-20)
		if tp.range('A50').value == 'Nguyễn Thái Nguyên':
			tp.pictures.add(os.path.join(path,signatures_folder,'Nguyen.png'),name = 'Nguyen',top = tp.range('G44').offset(1,0).top+7,left = tp.range('C44').offset(1,0).left-20)
		if tp.range('I50').value == 'Nguyễn Khắc Thắng':
			tp.pictures.add(os.path.join(path,signatures_folder,'Thang.png'),name = 'Thang',top = tp.range('G44').offset(1,0).top+7,left = tp.range('J44').offset(1,0).left-20)
		elif tp.range('I50').value == 'Lê Văn Hoàn':
			tp.pictures.add(os.path.join(path,signatures_folder,'Hoanle.png'),name = 'Hoanle',top = tp.range('G44').offset(1,0).top+7,left = tp.range('J44').offset(1,0).left-20)
		elif tp.range('I50').value == 'Nguyễn Tuấn Minh':
			tp.pictures.add(os.path.join(path,signatures_folder,'Minh.png'),name = 'Minh',top = tp.range('G44').offset(1,0).top+7,left = tp.range('J44').offset(1,0).left-20)
		elif tp.range('I50').value == 'Lê Quang Thông':
			tp.pictures.add(os.path.join(path,signatures_folder,'Thong.png'),name = 'Thong',top = tp.range('G44').offset(1,0).top+7,left = tp.range('J44').offset(1,0).left-20)
		elif tp.range('I50').value == 'Huỳnh Minh Hoàng':
			tp.pictures.add(os.path.join(path,signatures_folder,'Thang.png'),name = 'Hoang',top = tp.range('G44').offset(1,0).top+7,left = tp.range('J44').offset(1,0).left-20)
		elif tp.range('I50').value == 'Nguyễn Thái Nguyên':
			tp.pictures.add(os.path.join(path,signatures_folder,'Nguyen.png'),name = 'Nguyen',top = tp.range('G44').offset(1,0).top+7,left = tp.range('J44').offset(1,0).left-20)
		elif tp.range('I50').value == 'Trần Minh Đức':
			tp.pictures.add(os.path.join(path,signatures_folder,'Duc.png'),name = 'Nguyen',top = tp.range('G44').offset(1,0).top+7,left = tp.range('J44').offset(1,0).left-20)
		
	def report(self,conn,info,part_list,report_num,folder='files'):
		
		# wb = xw.Book(os.path.join(folder,'GDKT_Template.xlsx'))
		wb = xw.Book(os.path.join(folder,'GDKT_Template_0705.xlsx'))
		
		try:
			tp = wb.sheets('Template')
		except:
			tp = wb.sheets('BCGDKT')

		#clear old data
		tp.range('H5').value =''
		try:
			tp.pictures('QR_Code').delete()
			tp.pictures('Model').delete()
		except:
			pass

		tp.range('C10').value = ''
		tp.range('D16:D25').value=''

		path = pathlib.Path().absolute()
		# add QR Code AND MODEL
		rng_qr = tp.range('J7')
		try:
			tp.pictures('QR_Code').delete()
		except:
			pass
		tp.pictures.add(os.path.join(path,'images',f"{info['RMA No.'][0]}_{info['SERIAL_NO'][0]}.png"),name='QR_Code',top = rng_qr.top-10,left = rng_qr.left)
		tp.pictures('QR_Code').width = 90
		tp.pictures('QR_Code').height = 90

		# replace title
		rng = ['A12','B34','B40']
		rma = info['RMA No.'][0]
		try:
			for i in rng:
				kinh_thua = tp.range(i).value
				tp.range(i).value = kinh_thua.replace('PK/BV',info['Title'][0])
		except:
			print(f'Check title for {rma}')
		
		# inspection info
		# rma = info['RMA No.'][0]
		q = f'''
				SELECT c.[rma no.] ,strftime('%d/%m/%Y',c.[date installed]) AS [date_installed],
				c.[Scope Connect Count] as [scope_count],c.[update user],strftime('%d/%m/%Y',
				c.[Last Repair(Shipping)]) as last_repair,e.location,e.leader,
				strftime('%d/%m/%Y',c.recieve_date) as receive,[defect note]
				FROM consolidated c
				LEFT JOIN engineers e ON C.[IN_INSPECT_USER_NAME] = e.exfm_name
				WHERE C.[RMA NO.]='{rma}'
			'''
		add_info = pd.read_sql(q,conn)
		

		today = dt.now()
		hnay = f"         Ngày {today.strftime('%d')} tháng {today.strftime('%m')} năm {today.strftime('%Y')}"
		tp.range('H5').value = hnay
		tp.range('C10').value = str(info['web_name'][0]).upper()
		tp.range('D16').value = str(info['Vie'][0])
		tp.range('G16').value = rma
		model = str(info['MODEL'][0])
		tp.range('D17').value = model.replace(' V2','').upper()
		tp.range('D18').value = str(info['SERIAL_NO'][0]).upper()
		tp.range('D19').value = str(info['ks'][0])
		tp.range('D20').value = str("'") + str(info['inspect_date'][0])
		tp.range('A50').value = str(add_info['leader'][0]) # thay doi chu ky
		tp.range('D22').value = str("'") + str(add_info['date_installed'][0])
		tp.range('D23').value = str(add_info['scope_count'][0])
		tp.range('D24').value = str("'") + str(add_info['last_repair'][0])
		tp.range('D25').value = str("'") + str(add_info['receive'][0])
		tp.range('A8').value = f"Số:{report_num}/BCGĐKT"

		# add signatures
		self.signatures(tp)
		# code
		rma=info['RMA No.'][0]
		q = f'''
				SELECT distinct 
				cd.[rma no.],pf.P_DESCRIPTION,pf.F_DESCRIPTION,cd.[c/d-Detail]
				FROM pf_code pf
				inner JOIN  cd_code cd ON (cd.[rma no.] = pf.[rma no.] AND cd.[line_no] = pf.[line_no])
				WHERE cd.[rma no.] ='{rma}'
			'''
		ins_code = pd.read_sql(q,conn)
		ins_code

		top_model = tp.range('H16').top
		left_model = tp.range('H16').left
		try:
			tp.pictures('Model').delete()
		except:
			pass
		model = info['MODEL'][0]
		model = model.replace(' V2','')
		sn = info['SERIAL_NO'][0]

		model_folder = 'models'

		if ('XL' in model) or('VP' in model):
			print('Please add pictrue label model.')
		elif ('EB' in model):
			tp.pictures.add(os.path.join(path,model_folder,"eb.jpg"),name='Model',top = top_model,left = left_model)
		elif ('201' in model) or ('250' in model):
			tp.pictures.add(os.path.join(path,model_folder,"250.jpg"),name='Model',top = top_model,left = left_model)
		elif ('450' in model) or ('470' in model):
			tp.pictures.add(os.path.join(path,model_folder,"400.jpg"),name='Model',top = top_model,left = left_model)
		elif ('530' in model) or ('580' in model) or ('590' in model) or ('600' in model):
			tp.pictures.add(os.path.join(path,model_folder,"500.jpg"),name='Model',top = top_model,left = left_model)
		elif ('720' in model) or ('740' in model) or ('760' in model):
			tp.pictures.add(os.path.join(path,model_folder,"700.jpg"),name='Model',top = top_model,left = left_model)
		
		try: 
			self.resize_height_image(tp,'Model',180)
		except:
			print(f'No picture for model {model}')
		
		# part_list = self.part_list
		# add parts
		for i in range(len(part_list)-1):
			tp.range('35:35').offset(1,0).insert('down')
		for i in range(len(part_list)):
			tp.range('B'+str(i+34)).offset(1,0).value = "'- " + str(part_list['Vie'][i])
			if str(part_list['Vie'][i]) =='None': # add english name in F column
				tp.range('F'+str(i+35)).value = str(part_list['PART_DESCRIPTION'][i])
		
		#format evulation
		if len(ins_code)>1:
			for i in range(len(ins_code)-1):
				tp.range('32:32').offset(-1,0).insert('down')
				tp.range('B32:C32').offset(-1,0).merge()
				tp.range('D32:E32').offset(-1,0).merge()
				tp.range('F32:G32').offset(-1,0).merge()
				tp.range('H32:J32').offset(-1,0).merge()

			tp.range('A31:J31').offset(-1,0).copy()
			tp.range('A32:J'+str(len(ins_code)+30)).offset(-1,0).paste('formats')   

		#add code
		for i in range(len(ins_code)):
			tp.range('A'+str(i+30)).value = i+1
			tp.range('B'+str(i+30)).value = ins_code['F_DESCRIPTION'][i]
			tp.range('D'+str(i+30)).value = ins_code['P_DESCRIPTION'][i]
			tp.range('F'+str(i+30)).value = ins_code['C/D-Detail'][i]
			tp.name = 'BCGDKT'
		return wb

	def save_and_close(self,wb,i_report,model,sn,folder_name='files'):
		model = model.replace('-','')
		model = model.replace(' V2','')
		file_name = f"FFVN-GDKT-{i_report}-{model.replace('/','')},{sn}.xlsx"
		wb.save(os.path.join(folder_name,file_name))
		print(f'Exported {file_name}')
		wb.close()

	def report_info(self,rma,conn):
		q = f'''
				SELECT c.[RMA No.],c.MODEL,C.SERIAL_NO,
				sc.vie,cs.web_name,cs.title,e.ks,
				c.in_inspect_user_name,strftime('%d/%m/%Y',c.in_inspect_date) as inspect_date,
				c.customer_name

				FROM ((consolidated c
				LEFT JOIN scopes sc ON c.model = sc.model)
				LEFT JOIN engineers e ON c.in_inspect_user_name = e.exfm_name)
				LEFT JOIN customers cs ON c.customer_code = cs.[no.]
				
				WHERE c.[rma no.] = '{rma}'

			'''
		info = pd.read_sql(q,conn)
		
		q=f'''
				SELECT p.PART_DESCRIPTION,
				pn.vie,pn.vie_gd,pn.priority,
				p.part_no,p.quantity

				FROM parts p
				LEFT JOIN part_name pn ON p.PART_DESCRIPTION = pn.PART_DESCRIPTION
				WHERE p.[rma no.] = '{rma}'
				ORDER BY priority

			'''
		part_list = pd.read_sql(q,conn)
		return info,part_list

class summary_report():

	def __init__(self):
		pass

	def issues(self,conn,rma):
		q =f'''
				SELECT r.[rma no.],r.r_code,r.r_description
				FROM repair_code r
				WHERE r.[rma no.]='{rma}'
			'''
		issues = pd.read_sql(q,conn)
		issue_str =[]
		for issue in issues['R_DESCRIPTION']:
			issue = issue.split('.')[0]
			if issue not in issue_str: issue_str.append(issue)
		issues = str(issue_str).replace("', '",", ")[2:-2]
		return issues

	def summary_info(self,conn,rma):
		q = f'''
				SELECT c.[rma no.],
						e.ks,
						substr(substr(e.ks,instr(e.ks,' ')+1),instr(substr(e.ks,instr(e.ks,' ')+1),' ')+1) AS first_name,
						strftime('%d/%m/%Y',c.[date installed]) AS [date_installed],

						strftime('%d/%m/%Y',c.recieve_date) as receive,strftime('%d/%m/%Y',c.in_inspect_date) as inspect_date
						FROM consolidated c
						LEFT JOIN engineers e ON C.[IN_INSPECT_USER_NAME] = e.exfm_name
						WHERE C.[RMA NO.]='{rma}'
					'''
		sum_info = pd.read_sql(q,conn)
		return sum_info

	def read_summary_data(self,file_name,mode = 'r'):
		summary = xw.Book(file_name) #'2022_FFVN Service Endo_Inspection, WTY & TR .xlsx')
		
		gdkt = summary.sheets('BCGDKT')
		gdkt_lr = gdkt.range('A' + str(gdkt.cells.last_cell.row)).end('up').row
		try:
			gdkt_id = int(gdkt.range('A' + str(gdkt_lr)).value)
		except Exception as e:
			print(e)
			gdkt_id = 0

		wty = summary.sheets('WTY')
		wty_lr = wty.range('A' + str(wty.cells.last_cell.row)).end('up').row
		try:
			wty_id = int(wty.range('A' + str(wty_lr)).value)
		except Exception as e:
			print(e)
			wty_id = 0

		tr = summary.sheets('TR')
		tr_lr = tr.range('A' + str(tr.cells.last_cell.row)).end('up').row
		try:
			tr_id = int(tr.range('A' + str(tr_lr)).value)
		except Exception as e:
			print(e)
			tr_id = 0

		summary_data = {}
		
		summary_data.update({'gdkt':(gdkt_lr,gdkt_id)})
		summary_data.update({'wty':(wty_lr,wty_id)})
		summary_data.update({'tr':(tr_lr,tr_id)})

		if mode=='r': 
			summary.close()
			return summary_data
		elif mode == 'gdkt':
			return gdkt
		elif mode == 'wty':
			return wty
		elif mode == 'tr':
			return tr
		else:
			summary.close()
			print('Select mode r/gdkt/wty/tr')
	def many_gdkt(self):
		pass

	def task_review(self,conn):
		q='''
			SELECT DISTINCT c.[rma no.],c.customer_name,c.model,c.serial_no,
				c.recieve_date as [Receive],pf.[Start time] as [Start Time],
				pf.[end time] as [End Time],c.repair_status,
				CASE
					WHEN c.ship_user_name NOT NULL THEN c.ship_user_name
					WHEN c.qc_user_name NOT NULL THEN c.qc_user_name
					WHEN c.[repair user name] NOT NULL THEN c.[repair user name]
					WHEN c.authorized_user_name NOT NULL THEN c.authorized_user_name
					WHEN c.part_select_user_name NOT NULL THEN c.part_select_user_name
					WHEN c.in_inspect_user_name NOT NULL THEN c.in_inspect_user_name
					WHEN c.recieve_user_name NOT NULL THEN c.recieve_user_name
					WHEN c.create_user_name NOT NULL THEN c.create_user_name
					
				ELSE c.[update user] END AS [update_user],e.location,
				CASE
					WHEN c.repair_status = 'Create' THEN 0
					WHEN c.repair_status = 'Receive' THEN 1
					WHEN c.repair_status = 'Inspection' THEN 2
					WHEN c.repair_status = 'Parts Selection' THEN 3
					WHEN c.repair_status = 'Authorization' THEN 4
					WHEN c.repair_status = 'Repair' THEN 5
					WHEN c.repair_status = 'QC' THEN 6
					WHEN c.repair_status = 'Shipped' THEN 7
				ELSE 8 END AS repair_id
					
			FROM ((consolidated c LEFT JOIN engineers e ON e.[exfm_name] = [update_user])
				LEFT JOIN repair_code pf on c.[rma no.] =pf.[rma no.])
				WHERE (c.repair_status in ('Create','Receive','Authorization','Repair','QC'))
				OR [START TIME] NOT NULL
				ORDER BY location,repair_id
			'''
		task = pd.read_sql(q,conn)
		a = task[(task['REPAIR_STATUS']=='Receive') & (task['location']=='HCM')]
		hcm_receive = len(a)

		b = task[(task['REPAIR_STATUS']=='Receive') & (task['location']=='Hanoi')]
		hanoi_receive = len(b)

		c = task[(~task['Start Time'].isnull()) & (task['End Time'].isnull()) & (task['location']=='HCM')]
		hcm_repair = len(c)

		d = task[(~task['Start Time'].isnull()) & (task['End Time'].isnull()) & (task['location']=='Hanoi')]
		hanoi_repair = len(d)

		e = task[(task['REPAIR_STATUS'].isin(['QC','Repair'])) & (task['location']=='HCM')]
		hcm_ship = len(e)

		f = task[(task['REPAIR_STATUS'].isin(['QC','Repair'])) & (task['location']=='Hanoi')]
		hanoi_ship = len(f)

		dt={'Task':['HCM','Hanoi'],
			'Receive':[hcm_receive,hanoi_receive],
			'Repair':[hcm_repair,hanoi_repair],
			'Ship':[hcm_ship,hanoi_ship]}

		abcdef = pd.DataFrame(data=dt)
		print(abcdef)
		return task
