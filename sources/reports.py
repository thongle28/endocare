import os
import pandas as pd
from sqlite3 import connect
import sources.parts_list_0705 as pl
import xlwings as xw 
from datetime import datetime as dt
import pathlib
import qrcode
from IPython.display import Image, display

def gdkt_report(rma,conn):
	report_num = str(rma.split('FMSV')[1])

	folder_name = 'reports_' + dt.now().strftime('%y%m%d')
	try:
		os.mkdir(folder_name)
		print(f'folder {folder_name} was created.')
	except:
		print(f'Folder {folder_name} exists')

	i_report = 0
	q_gdkt = {}
	report_id = {}
	d = pl.summary_report()



	# for rma in list_rma['rma']:

	i_report +=1
	
	report_num_str = str(report_num)
	# not number
	#     report_num_str = 'xxx'

	info = []
	part_list= []
	c = pl.technical_report()
	info,part_list = c.report_info(rma,conn)
	model = info['MODEL'][0]
	sn = info['SERIAL_NO'][0]
	#     c = pl.technical_report(info,report_num,part_list)
	c.create_qr_image(info)
	tp = c.report(conn,info,part_list,report_num_str,folder = 'templates')
	c.save_and_close(tp,report_num_str,model,sn,folder_name)
	q_gdkt.update({rma:(info,part_list)})    

	
	print('Done!!!')

# Trouble Report
def tr_report(rma,conn):
	r_type ='tr'
	d = pl.summary_report()
	issue = str(d.issues(conn,rma))
	c = pl.technical_report()
	info,part_list = c.report_info(rma,conn)
	part_list
	info
	q = f'''
			SELECT c.[rma no.] ,strftime('%d/%m/%Y',c.[date installed]) AS [date_installed],
			c.[Scope Connect Count] as [scope_count],c.[update user],strftime('%d/%m/%Y',
			c.[Last Repair(Shipping)]) as last_repair,e.location,e.leader,
			strftime('%d/%m/%Y',c.recieve_date) as receive,[defect note],[Last RMA No.],strftime('%d/%m/%Y',[create]) as create_date
			FROM consolidated c
			LEFT JOIN engineers e ON C.[IN_INSPECT_USER_NAME] = e.exfm_name
			WHERE C.[RMA NO.]='{rma}'
		'''
	add_info = pd.read_sql(q,conn)
	report_num_str = str(rma.split('FMSV')[1])


	wb = xw.Book('templates\\Report_Template.xlsx')
	tp = wb.sheets('Report')
	folder_name = 'reports_' + dt.now().strftime('%y%m%d')

	path = pathlib.Path().absolute()
	try:
		os.mkdir(folder_name)
		print(f'folder {folder_name} was created.')
	except:
		print(f'Folder {folder_name} exists')
	if r_type == 'tr':
	#     ref_no = f"{dt.now().strftime('FFVN-%y%m')}{report_num_str}TR"
		ref_no = f"FFVN-TR-{report_num_str}"
	elif r_type =='wty':
		ref_no = f"{dt.now().strftime('FFVN-%y%m')}{report_num_str}"

	tp.range('C5').value = ref_no
	tp.range('C6').value = dt.now().strftime('%d-%b-%y')
	tp.range('E6').value = info['RMA No.'][0]
	tp.range('C7').value = info['CUSTOMER_NAME'][0]
	tp.range('C8').value = 'VIETNAM'
	if r_type == 'wty':
		tp.range('C13').value = 'YES'
	elif r_type == 'tr':
		tp.range('C13').value = 'NO'
	tp.range('C10').value = info['MODEL'][0]
	tp.range('C11').value = str(info['SERIAL_NO'][0]).upper()
	try:
		tp.range('C12').value = dt.strptime(add_info['date_installed'][0],'%d/%m/%Y')
	except Exception as e:
		print('None Installation Date ')
	tp.range('E13').value = add_info['Last RMA No.'][0]
	tp.range('B15').value = add_info['Defect Note'][0]
	tp.range('C25').value = dt.strptime(add_info['create_date'][0],'%d/%m/%Y')

	try:
		tp.range('B27').value = add_info['Defect Note'][0] + '\n' +'Used case: ' + str(int(add_info['scope_count'][0]))
	except:
		tp.range('B27').value = add_info['Defect Note'][0]

	tp.range('C28').value = info['IN_INSPECT_USER_NAME'][0]
	try:
		tp.range('C29').value = dt.strptime(info['inspect_date'][0],'%d/%m/%Y')
	except:
		print('Not inspection completed')
	plf = pl.parts_list().part_list_final(rma,conn)

	if tp.range('B37').value == 'Name / Date :': #check empty table
		#insert new rows
		for i in range(len(plf)-1):
			tp.range('33:33').insert('down')
			tp.range('32:32').copy()
			tp.range('33:33').paste('formats')
	#     total_price = sum(part_list_final['Dealer Price'])
		# fill in parts informtion
		for i in range(len(plf)):

			tp.range('B' + str(i+32)).value = plf['part_num'][i]
			tp.range('C' + str(i+32)).value = plf['PART_DESCRIPTION'][i] 

			tp.range('D' + str(i+32)).value = plf['QUANTITY'][i]
			tp.range('E' + str(i+32)).value = plf['FFVN Price'][i]

		tp.range('E'+str(i+33)).value = f'=SUMPRODUCT(D32:D{str(i+32)},E32:E{str(i+32)})'
		tp.name = ref_no
		model_t = str(info['MODEL'][0]).replace('-','')
		model_t= model_t.replace('/','')
		sn_t = str(str(info['SERIAL_NO'][0]).upper()).replace('-','')
		sn_t = sn_t.replace('/','')
		issue = issue.replace(' ','')
		wb.save(f'{folder_name}\\{ref_no}-{model_t},{sn_t}-{issue}.xlsx')
		print(f'Export {wb.name} completed')
		print(f'{wb.name}')
		wb.close()
	else:
		print('table not empty')
	report_num = int(report_num) +1
	print('Done')
print('Done')


def main(conn):
	esc = True
	while esc:

		#Select RMA or SN:
		sn = str(input('Search by Serial or RMA: '))
		if sn.upper().strip() == 'UPDATE':
			pass 
		else:
			q=f'''
					SELECT c.[rma no.] AS rma,c.customer_name,c.serial_no,c.model,c.approval,c.repair_status,c.in_inspect_user_name
					FROM consolidated c
					WHERE upper(c.serial_no) like '%{sn.upper()}%' or  upper(c.[rma no.]) like '%{sn.upper()}%'
					ORDER BY rma DESC
				'''
			results = pd.read_sql(q,conn)
		if len(results)>0:
			display(pd.read_sql(q,conn))
		elif sn.upper().strip() == 'QUIT' or sn.upper().strip() == 'EXIT' or sn.upper().strip() == 'DONE':
			break
		else:
			print(f'\nCan not search with key "{sn}"')
			continue

		#Select Index
		while True:
			try:
				ind = int(input('\nSelect RMA by index (Default 0): ') or 0)
				break
			except:
				print('Only Accept number') 
		for i in range(results.shape[1]):
			print("{0:40} {1}".format(results.columns[i],results.iloc[ind][i]))


		# Confirm RMA
		confirm = str(input(f'\nConfirm Select RMA"{results.iloc[ind][0]}"? Y/[N]') or 'N')
		if confirm.upper() =='Y':
			rma = results.iloc[ind][0]
		else:
			rma=''
			continue
		chosen = str(input('[GDKT] or Trouble Report(tr): '))
		try:
			if chosen =='': gdkt_report(rma,conn)
			if chosen.upper() =='TR': tr_report(rma,conn)
		except Exception as e:
			print(e,rma)

def quotation(conn):
	# loop for multiple quotation
	# filter Awating Parts Billing
	q = '''
			SELECT [rma no.],customer_name,model,serial_no,repair_status
			FROM consolidated
			WHERE [status info] LIKE "%Parts Billing"

		'''
	display(pd.read_sql(q,conn))
	esc = True

	while esc:

		#Select RMA or SN:
		sn = str(input('\nSearch by Serial or RMA: '))
		if sn.upper().strip() == 'UPDATE':
			pass 
		else:
			q=f'''
					SELECT c.[rma no.] AS rma,c.customer_name,c.serial_no,c.model,c.approval,c.repair_status,c.in_inspect_user_name
					FROM consolidated c
					WHERE upper(c.serial_no) like '%{sn.upper()}%' or  upper(c.[rma no.]) like '%{sn.upper()}%'
					ORDER BY rma DESC
				'''
			results = pd.read_sql(q,conn)
		if len(results)>0:
			display(pd.read_sql(q,conn))
		elif sn.upper().strip() == 'QUIT' or sn.upper().strip() == 'EXIT' or sn.upper().strip() == 'DONE':
			break
		else:
			print(f'\nCan not search with key "{sn}"')
			continue

		#Select Index
		while True:
			try:
				ind = int(input('\nSelect RMA by index (Default 0): ') or 0)
				break
			except:
				print('Only Accept number') 
		for i in range(results.shape[1]):
			print("{0:40} {1}".format(results.columns[i],results.iloc[ind][i]))


		# Confirm RMA
		confirm = str(input(f'\nConfirm Select RMA"{results.iloc[ind][0]}"? Y/[N]') or 'N')
		if confirm.upper() =='Y':
			rma = results.iloc[ind][0]
		else:
			rma=''
			continue


		# Export Excel File
		if rma !='':
			service_fee = 450
			currency = 24882
			q_report={}
			rma_no_part_list=[]
			folder_name = 'quotation_' + dt.now().strftime('%y%m%d')
			image_folder = 'images'
			path = pathlib.Path().absolute()
			try:
				os.mkdir(folder_name)
				print(f'folder {folder_name} was created.')
			except:
				print(f'Folder {folder_name} exists')
			 #     create rma
			qr = qrcode.QRCode(
			version=1,
			error_correction=qrcode.constants.ERROR_CORRECT_H,
			box_size=4,
			border=0.1,
			)
			qr.add_data(rma)
			qr.make(fit=True)
			img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
			img_name = f"{image_folder}\\{rma}.png"
			img.save(img_name)
			# a = pl.parts_list()
			try:

				part_list,info,price_add,title = pl.parts_list().create_parts_list(conn,rma,service_fee,currency,folder_name)
				#create quotation
				b = pl.quotation(part_list,info,price_add,title)
			   
				qr_top = b.wb.sheets('Parts').range("I7").top
				qr_left = b.wb.sheets('Parts').range("I7").left
				b.wb.sheets('Parts').pictures.add(os.path.join(path,img_name),name='img_name',top = qr_top+5,left = qr_left+22)
			
				# add contact
				try:
					q =f'SELECT [contacted by],[contact name] FROM consolidated WHERE [RMA NO.]="{rma}"'
					contact = pd.read_sql(q,conn)
					b.wb.sheets('Quotation').range('D14').value = f"{str(contact['Contacted by'][0])}/{str(contact['Contact Name'][0])}"
				except Exception as e:
					print(e)

				b.input_data()
				b.save_and_close()
				q_report.update({rma:(part_list,info,price_add,title)})
			except Exception as e:
				print(e)
				rma_no_part_list.append(rma)
				print('Can not export')
			print(f'Done for {rma}!')