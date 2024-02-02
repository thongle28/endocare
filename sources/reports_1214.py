import os
import pandas as pd
from sqlite3 import connect
import sources.parts_list_0802 as pl
import xlwings as xw 
from datetime import datetime as dt
from datetime import timedelta
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
		# choosen = str(input('[GDKT] or Trouble Report(tr): '))
		choosen_list = ['Quotation',
					'GDKT (Default)',
					'Trouble Report',
					'Warranty Report',
				]

		#border table
		print(f'\n{"_"*50}')
		print(f'{"|  No.|  Function": <49}|')
		print(f'|{"_"*48}|')
		for i in range(1,1+len(choosen_list)):
			print(f'|{i: >3}  |  {choosen_list[i-1]: <40}|')
		print(f'|{"_"*48}|') #bottom border

		choosen = str(input('Select Function by Index: '))

		try:
			if choosen ==choosen_list[1]: 
				gdkt_report(rma,conn)
				# change_stt_info = input('Change Status Info to "Waiting for Next Process Avaiable? ([Y]/N)"')
			if choosen.upper() =='TR': 
				tr_report(rma,conn)

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
				# try:
				q =f'SELECT [contacted by],[contact name] FROM consolidated WHERE [RMA NO.]="{rma}"'
				contact = pd.read_sql(q,conn)
				b.wb.sheets('Quotation').range('D14').value = f"{str(contact['Contacted by'][0])}/{str(contact['Contact Name'][0])}"
				# except Exception as e:
				# 	print(e)

				b.input_data()
				b.save_and_close()
				q_report.update({rma:(part_list,info,price_add,title)})
			except Exception as e:
				print(e)
				# rma_no_part_list.append(rma)
				print('Can not export')
			print(f'Done for {rma}!')

class weekly_report():

	def __init__(self,conn):
		self.conn = conn
		today = dt.today()
		current_weekday = today.weekday()  # 0 for Monday, 1 for Tuesday, ..., 6 for Sunday
		monday_of_this_week = today - timedelta(days=current_weekday)

		# print("Monday of this week:", monday_of_this_week.date())
		ans = str(input(f'Start Date of Report(YYYY-MM-DD): {monday_of_this_week.date()}?[Y]/N') or monday_of_this_week.date())
		
		self.report_start_date = ans


	def receive(self):
		report_start_date = self.report_start_date
		conn = self.conn
		q = f'''
				SELECT [rma no.],customer_name,model,serial_no,
				strftime('%Y-%m-%d',recieve_date)AS [receive date],
				recieve_user_name,repair_status,e.location
				
				FROM consolidated c
				LEFT JOIN engineers e ON c.recieve_user_name = e.exfm_name
				WHERE recieve_date >= '{report_start_date}'
				ORDER BY location,customer_name
				
			'''
		receive = pd.read_sql(q,conn)
		self.receive = receive 

	def inspection(self):

		# inspection and repair by location
		report_start_date = self.report_start_date
		conn = self.conn
		
		q = f'''
					SELECT 
						DISTINCT c.[rma no.],customer_name,model,serial_no,recieve_date,in_inspect_date,
						r.[start time],r.[end time],'Inspection' AS [repair size],
						in_inspect_user_name AS PIC,e.location

						FROM (consolidated c
						LEFT JOIN engineers e ON c.recieve_user_name = e.exfm_name)
						LEFT JOIN repair_code r ON c.[rma no.] = r.[rma no.]
						WHERE in_inspect_date >= '{report_start_date}'
						
						
					UNION ALL
						SELECT DISTINCT c.[rma no.],customer_name,model,serial_no,recieve_date,in_inspect_date,
							r.[start time],r.[end time],c.[repair size],
							CASE
								WHEN r.[start user] NOT NULL THEN r.[start user]
								ELSE c.[repair user name] END AS [Repair User],
							e.location


							FROM (consolidated c
							LEFT JOIN engineers e ON c.recieve_user_name = e.exfm_name)
							LEFT JOIN repair_code r ON c.[rma no.] = r.[rma no.]
							WHERE r.[start time] >= '{report_start_date}'
							
			'''
		wr = pd.read_sql(q,conn)
		wr.to_sql('weekly_report',conn,if_exists = 'replace',index = False)

		q = f'''
				SELECT [rma no.],customer_name,model,serial_no,recieve_date,[repair size],location,
				CASE
					WHEN [repair size] = 'Inspection' AND PIC ='Nguyen Thai' THEN 'Inspection'
					WHEN [repair size] = 'Minor' AND PIC ='Nguyen Thai' THEN 'Minor'
					WHEN [repair size] = 'Other' AND PIC ='Nguyen Thai' THEN 'Minor'
					WHEN [repair size] = 'Major' AND PIC ='Nguyen Thai' THEN 'Major'
					
					ELSE '-' END AS 'Nguyen',
				CASE
					WHEN [repair size] = 'Inspection' AND PIC ='hoang' THEN 'Inspection'
					WHEN [repair size] = 'Minor' AND PIC ='hoang' THEN 'Minor'
					WHEN [repair size] = 'Other' AND PIC ='hoang' THEN 'Minor'
					WHEN [repair size] = 'Major' AND PIC ='hoang' THEN 'Major'
					ELSE '-' END AS 'Hoang',
				CASE
					WHEN [repair size] = 'Inspection' AND PIC ='Le Quang Thong' THEN 'Inspection'
					WHEN [repair size] = 'Minor' AND PIC ='Le Quang Thong' THEN 'Minor'
					WHEN [repair size] = 'Other' AND PIC ='Le Quang Thong' THEN 'Minor'
					WHEN [repair size] = 'Major' AND PIC ='Le Quang Thong' THEN 'Major'
					ELSE '-' END AS 'Thong',
				CASE
					WHEN [repair size] = 'Inspection' AND PIC ='Nguyen Khac Thang (Hanoi)' THEN 'Inspection'
					WHEN [repair size] = 'Minor' AND PIC ='Nguyen Khac Thang (Hanoi)' THEN 'Minor'
					WHEN [repair size] = 'Other' AND PIC ='Nguyen Khac Thang (Hanoi)' THEN 'Minor'
					WHEN [repair size] = 'Major' AND PIC ='Nguyen Khac Thang (Hanoi)' THEN 'Major'
					ELSE '-' END AS 'Thang',
				CASE
					WHEN [repair size] = 'Inspection' AND PIC ='Le Van Hoan' THEN 'Inspection'
					WHEN [repair size] = 'Minor' AND PIC ='Le Van Hoan' THEN 'Minor'
					WHEN [repair size] = 'Other' AND PIC ='Le Van Hoan' THEN 'Minor'
					WHEN [repair size] = 'Major' AND PIC ='Le Van Hoan' THEN 'Major'
					ELSE '-' END AS 'Hoanle',
				CASE
					WHEN [repair size] = 'Inspection' AND PIC ='Nguyen Tuan Minh' THEN 'Inspection'
					WHEN [repair size] = 'Minor' AND PIC ='Nguyen Tuan Minh' THEN 'Minor'
					WHEN [repair size] = 'Other' AND PIC ='Nguyen Tuan Minh' THEN 'Minor'
					WHEN [repair size] = 'Major' AND PIC ='Nguyen Tuan Minh' THEN 'Major'
					ELSE '-' END AS 'Minh'
					
					
					
				FROM weekly_report
				ORDER BY location,[repair size]
			'''
		completed = pd.read_sql(q,conn)

		self.completed = completed

	def export(self):
		# open template
		wb = xw.Book('templates/weekly_report.xlsx')
		sh_in = wb.sheets('In')
		sh_out = wb.sheets('Out')

		#initial IN
		sh_in.range('A2').value = 'Customer'
		sh_in.range('B2').value = 'Model'
		sh_in.range('C2').value = 'Serial'
		sh_in.range('D2').value = 'Recieve'
		sh_in.range('E2').value = 'Location'
		sh_in.range('F2').value = 'Status'

		# import data
		receive = self.receive
		completed = self.completed

		for i in range(len(receive)):
			sh_in.range('A3').offset(i,0).value = receive['CUSTOMER_NAME'][i]
			sh_in.range('B3').offset(i,0).value = receive['MODEL'][i]
			sh_in.range('C3').offset(i,0).value = receive['SERIAL_NO'][i]
			sh_in.range('D3').offset(i,0).value = receive['receive date'][i]
			sh_in.range('E3').offset(i,0).value = receive['location'][i]
			sh_in.range('F3').offset(i,0).value = receive['REPAIR_STATUS'][i]
			
		#initial OUT
		sh_out.range('A2').value = 'Customer'
		sh_out.range('B2').value = 'Model'
		sh_out.range('C2').value = 'Serial'
		sh_out.range('D2').value = 'Receive'
		sh_out.range('E2').value = 'Nguyên'
		sh_out.range('F2').value = 'Hoàng'
		sh_out.range('G2').value = 'Thông'

		hcm = completed[completed['location']=='HCM']
		hcm = hcm.reset_index(drop=True)

		for i in range(len(hcm)):
			sh_out.range('A3').offset(i,0).value = hcm['CUSTOMER_NAME'][i]
			sh_out.range('B3').offset(i,0).value = hcm['MODEL'][i]
			sh_out.range('C3').offset(i,0).value = hcm['SERIAL_NO'][i]
			sh_out.range('D3').offset(i,0).value = hcm['RECIEVE_DATE'][i]
			sh_out.range('E3').offset(i,0).value = hcm['Nguyen'][i]
			sh_out.range('F3').offset(i,0).value = hcm['Hoang'][i]
			sh_out.range('G3').offset(i,0).value = hcm['Thong'][i]
		 
		k = len(hcm)+3

		sh_out.range('A2').offset(k,0).value = 'Customer'
		sh_out.range('B2').offset(k,0).value = 'Model'
		sh_out.range('C2').offset(k,0).value = 'Serial'
		sh_out.range('D2').offset(k,0).value = 'Receive'
		sh_out.range('E2').offset(k,0).value = 'Thắng'
		sh_out.range('F2').offset(k,0).value = 'Hoàn'
		sh_out.range('G2').offset(k,0).value = 'Minh'

		hanoi = completed[completed['location']=='Hanoi']
		hanoi = hanoi.reset_index(drop=True)

		for i in range(len(hanoi)):
			sh_out.range('A3').offset(i+k,0).value = hanoi['CUSTOMER_NAME'][i]
			sh_out.range('B3').offset(i+k,0).value = hanoi['MODEL'][i]
			sh_out.range('C3').offset(i+k,0).value = hanoi['SERIAL_NO'][i]
			sh_out.range('D3').offset(i+k,0).value = hanoi['RECIEVE_DATE'][i]
			sh_out.range('E3').offset(i+k,0).value = hanoi['Thang'][i]
			sh_out.range('F3').offset(i+k,0).value = hanoi['Hoanle'][i]
			sh_out.range('G3').offset(i+k,0).value = hanoi['Minh'][i]
		print('Done')

		today = dt.now().strftime('%y%m%d')
		folder_name = 'weekly_report'
		try:
			os.mkdir(folder_name)
			print(f'folder {folder_name} was created.')
			
		except:
			print(f'Folder {folder_name} exists')
		try:
			wb.save(f'{folder_name}/weekly_report_{today}.xlsx')
			print(f'{wb.name} saved and close.')
			wb.close()
		except Exception as e:
			print(eS)