import pandas as pd
from sqlite3 import connect
import datetime
import pathlib
import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import openpyxl
import requests
from io import BytesIO

ver = '1.1.0'

# change COMPLETE TABLE to clear 
# row 319-354

# -------------backup on 220224:
# avoid empty row in repair code
#  edit consolidtaed_sql
def r_code(rma_list,check_r =False):
	#create rcode tbale
	query = '''SELECT pf.[rma no.] as rma,pf.r_code as r_code,pf.r_description as r_description,r.vie_name

					FROM pf_code pf
					LEFT JOIN r_code r ON pf.r_code = r.r_code
				UNION ALL
				SELECT pf.[rma no.] as rma,pf.r_code as r_code,pf.r_description as r_description,r.vie_name
					FROM C_pf_code pf
					LEFT JOIN r_code r ON pf.r_code = r.r_code
					WHERE rma IN ({})
				
				ORDER BY rma
			'''.format(rma_list)
	check_r_code = '''
				SELECT DISTINCT r_code,r_description FROM ({})
				WHERE vie_name IS NULL
				'''.format(query)	
	if check_r: r_code = pd.read_sql(check_r_code,conn)
	else: r_code = pd.read_sql(query,conn)
	return r_code

def part_name(rma_list,check_part=False):
	query = '''
			SELECT 
				DISTINCT p.[RMA No.] AS rma,p.PART_NO AS part_no,
				p.Part_description AS part_description,pn.vie as vie_name
				FROM parts p
				LEFT JOIN part_name pn ON p.part_description = pn.part_description
			UNION ALL
			SELECT
				DISTINCT p.[RMA No.] AS rma,p.PART_NO AS part_no,
				p.Part_description AS part_description,pn.vie as vie_name
				FROM c_parts p
				LEFT JOIN part_name pn ON p.part_description = pn.part_description
				WHERE rma IN ({}) 		   
			ORDER BY rma
		'''.format(rma_list)
	check_part_empty = '''
						SELECT DISTINCT part_description
						FROM ({})
						WHERE vie_name 	IS NULL
						'''.format(query)
	if check_part:part_name = pd.read_sql(check_part_empty,conn)
	else: part_name = pd.read_sql(query,conn)
	return part_name

def date_name():
	a = datetime.datetime.now()
	b =str(a.year)[2:4]
	if a.month<10:c ='0'+str(a.month)
	else: c = str(a.month)
	if a.day<10:d ='0'+str(a.day)
	else: d = str(a.day)
	return b+c+d

def consolidated_sql(folder_name,file_name,is_complete,printer=False):
	# read file COMPLETED consolidated from ExFM
	xls = os.path.join(folder_name,file_name)
	df = pd.read_excel(xls,sheet_name=None)
	# try:
	c = df['Consolidated'].drop(['OWNERSHIP'],axis=1)
	
	# except:
	# 	pass

	# process pf_table

	p = df['Parts']
	pf = df['PF-Code']
	cd = df['CD-Code']
	r = df['R-Code']
	# c.to_sql('consolidated',conn,index=False, if_exists='replace')
	pf.to_sql('position_code',conn,index=False, if_exists='replace')
	cd.to_sql('cd_code',conn,index=False, if_exists='replace')
	# p.to_sql('parts',conn,index=False, if_exists='replace')
	r.to_sql('repair_code',conn,index=False, if_exists='replace')
	df_r = pd.DataFrame(r,columns=['RMA No.','LINE_NO','R_CODE','R_DESCRIPTION','SERVICE_TYPE','Start Time','End Time','Start User','End User'])

	issue=pd.DataFrame(r['RMA No.'])
	issue['Issue']=''
	df_r['Issue']=''
	for i in range(issue.shape[0]):
		issue['Issue'][i]= r['R_DESCRIPTION'].str.split('.')[i][0]
	gb = issue.groupby(['RMA No.'])
	result = gb['Issue'].unique()

	# gán giá trị issue vào bảng tổng
	for i in range(df_r.shape[0]):
		df_r['Issue'][i]=str(result[df_r['RMA No.'][i]])
	
	df_r.to_sql('repair_code',conn,index=False, if_exists='replace')
	# print (pd.read_sql('SELECT name from sqlite_master where type= "table";',conn))
	q ='''
		SELECT DISTINCT pf.[rma no.],pf.p_code,pf.p_description,pf.f_code,pf.f_description,
		cd.[c/d-code],cd.[c/d-detail],
		r.r_code,r_description,r.service_type,r.[start time],r.[end time],r.[start user],r.[end user],r.issue
		
		
		FROM (
		position_code pf LEFT JOIN cd_code cd ON (pf.[rma no.] = cd.[rma no.] AND pf.LINE_NO = CD.LINE_NO)
		LEFT JOIN repair_code r ON (pf.[rma no.] = r.[rma no.] AND pf.LINE_NO = r.LINE_NO))
		
	   

	'''
	pfcr_code = pd.read_sql(q,conn)
	
	#--keep old code
	if is_complete:
		c.to_sql('c_consolidated',conn,index=False, if_exists='replace') #create table consolidated
		pfcr_code.to_sql('c_pf_code',conn,index=False, if_exists='replace')
		p.to_sql('c_parts',conn,index=False, if_exists='replace')  # create table parts
	else:
		c.to_sql('consolidated',conn,index=False, if_exists='replace') #create table consolidated
		pfcr_code.to_sql('pf_code',conn,index=False, if_exists='replace')
		p.to_sql('parts',conn,index=False, if_exists='replace')  # create table parts
	check_table = pd.read_sql('SELECT name from sqlite_master where type= "table";',conn)
	# check_table =pd.read_sql('SELECT * from repair_code;',conn)
	if printer: print (check_table)
	
def new_table(is_completed,printer=False):
	if is_completed:
		var_consolidated = 'c_consolidated'
		var_pf_code = 'c_pf_code'
		new_table = 'c_new_table'
	else:
		var_consolidated = 'consolidated'
		var_pf_code = 'pf_code'
		new_table = 'new_table'
	query = '''SELECT DISTINCT
			c.[RMA No.] as rma,c.[OTHER ID] as wr_report,c.customer_code AS customer,c.model AS model_d,c.serial_no AS sn,c.[Date Installed] AS install_date,
			c.WARRANTY_END_DATE AS warranty_date,c.repair_status AS repair_status,c.approval,
			CASE
				WHEN pf.[End Time] IS NOT NULL AND c.repair_status NOT IN ('Completed','Shipped') THEN 'QC'
				WHEN pf.[Start Time] IS NOT NULL AND c.repair_status NOT IN ('Completed','Shipped','QC') THEN 'Under Repair'
				ELSE c.repair_status
			END as change_stt,
			pf.[Start Time] as start_time,pf.[End Time] as end_time,c.[create] as create_date,
			c.RECIEVE_DATE as receive_date,c.IN_INSPECT_DATE as inspection_date, 
			c.[Part Select Date] as part_list_date,c.[Repair Size] as repair_size,c.[update user] as pic,
			c.IN_INSPECT_USER_NAME,pf.issue,c.[update time],c.note1 as Note
		FROM {} c
		LEFT JOIN {} pf ON c.[rma no.] = pf.[rma no.]
		'''.format(var_consolidated,var_pf_code)
	new_query = pd.read_sql(query,conn)
	new_query.to_sql(new_table,conn,index=False)
	asd =pd.read_sql('SELECT name from sqlite_master where type= "table";',conn)
	
	
	if printer: print(asd)


def exfm_web(rma_list):
	query = '''SELECT 
				c.rma,c.customer,
				CASE
				WHEN e.location IS NULL THEN (SELECT e.location FROM new_table c LEFT JOIN engineers e ON c.IN_INSPECT_USER_NAME = e.exfm_name)
				ELSE e.location END AS location,
				c.model_d,c.sn,c.install_date,c.warranty_date,
				c.receive_date,c.inspection_date,c.start_time,c.end_time,
				CASE
				WHEN c.repair_size = 'Major' THEN 'Nặng'
				WHEN C.repair_size = 'Minor' THEN 'Nhẹ'
				WHEN c.repair_size IS NOT NULL THEN 'Khác' 
				ELSE c.repair_size END AS repair_size,
				REPLACE(SUBSTR(c.issue,3,LENGTH(c.issue)-4),"' '",". ") as issue_part,
				e.ks,c.wr_report,c.change_stt,c.approval,
				CASE
				WHEN c.approval  IN ('Decline','Cancel') THEN 'Chờ xác nhận (quá 3 tháng)'
				WHEN c.create_date < c.warranty_date AND c.repair_status ='Inspection' AND c.wr_report IS NULL THEN 'Chờ duyệt Bảo hành'
				WHEN c.wr_report IS NOT NULL AND c.repair_status ='Inspection' THEN 'Chuẩn bị linh kiện'
				ELSE s.vie_status END AS repair_status,
				CASE
				WHEN c.approval  IN ('Decline','Cancel') THEN 9
				ELSE s.stt_score END  AS e_score,
				CASE
				WHEN wr_report NOT NULL AND UPPER(wr_report) NOT LIKE '%REJECTED%' THEN 0.5
				ELSE 0 END AS w_score,c.[update time] AS update_time,c.note as return_date
				
				FROM 
				(new_table c LEFT JOIN engineers e ON c.pic = e.exfm_code)
				LEFT JOIN status s ON c.change_stt = s.repair_status
				
				UNION ALL 
				
				SELECT 
				c.rma,c.customer,
				CASE
				WHEN e.location IS NULL THEN (SELECT e.location FROM new_table c LEFT JOIN engineers e ON c.IN_INSPECT_USER_NAME = e.exfm_name)
				ELSE e.location END AS location,
				c.model_d,c.sn,c.install_date,c.warranty_date,
				c.receive_date,c.inspection_date,c.start_time,c.end_time,
				CASE
				WHEN c.repair_size = 'Major' THEN 'Nặng'
				WHEN C.repair_size = 'Minor' THEN 'Nhẹ'
				WHEN c.repair_size IS NOT NULL THEN 'Khác' 
				ELSE c.repair_size END AS repair_size,
				REPLACE(SUBSTR(c.issue,3,LENGTH(c.issue)-4),"' '",". ") as issue_part,
				e.ks,c.wr_report,c.change_stt,c.approval,
				CASE
				WHEN c.approval  IN ('Decline','Cancel') THEN 'Chờ xác nhận (quá 3 tháng)'
				WHEN c.create_date < c.warranty_date AND c.repair_status ='Inspection' AND c.wr_report IS NULL THEN 'Chờ duyệt Bảo hành'
				WHEN c.wr_report IS NOT NULL AND c.repair_status ='Inspection' THEN 'Chuẩn bị linh kiện'
				ELSE s.vie_status END AS repair_status,
				CASE
				WHEN c.approval  IN ('Decline','Cancel') THEN 9
				ELSE s.stt_score END  AS e_score,
				CASE
				WHEN wr_report NOT NULL AND UPPER(wr_report) NOT LIKE '%REJECTED%' THEN 0.5
				ELSE 0 END AS w_score,c.[update time] as update_time,c.note as return_date
				
				FROM 
				(c_new_table c LEFT JOIN engineers e ON c.pic = e.exfm_name)
				LEFT JOIN status s ON c.change_stt = s.repair_status
				
				WHERE RMA  IN({})
				
				

			'''.format(rma_list)

	
	
	
	web_base = pd.read_sql(query,conn)
	
	return web_base

#---------pending Master List

def pending_sql(filename):
	# pending-completed-transfer to sql memory
	

	global conn
	conn = connect(':memory:')

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

def database_process(folder_name,file):
#------------------------database
	# path = pathlib.Path().absolute()
	# path = os.path.join(path,folder_name,file)
	# # xls = 'database.xlsx'
	# xls = path 
	# db=pd.read_excel(xls,sheet_name = None)
	# db['engineers'].to_sql('engineers',conn,index=False,if_exists='replace')
	# db['status'].to_sql('status',conn,index=False,if_exists='replace')
	# db['part_name'].to_sql('part_name',conn,index=False,if_exists='replace')
	# db['r_code'].to_sql('r_code',conn,index=False,if_exists='replace')
	# db['scopes'].to_sql('scopes',conn,index=False,if_exists='replace')
	# db['dealers'].to_sql('dealers',conn,index=False,if_exists='replace')
	
	db_name ='memory'
	ins = 'memory'
	try:
		spreadsheetId = '1bT4W0CiLVD_B_ddRcVkS3MEXbSjvmGb4'
		url = "https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=" + spreadsheetId
		res = requests.get(url)
		data = BytesIO(res.content)
		sleep(5)
		xlsx = openpyxl.load_workbook(filename=data)
		for name in xlsx.sheetnames:
			values = pd.read_excel(data, sheet_name=name)
			values.to_sql(name,conn,index=False,if_exists='replace')
			print(f'check table {name}')
		print(f'\nDatabase file stored in {db_name}')
	except:
		print(f'can not find file {ins} and import as database data')

def pending_file():
	query = '''
			SELECT DISTINCT 
				m.rma AS rma_id,m.customer,m.model,m.sn,d.dealer,scopes.vie AS type_s,
				STRFTIME('%Y-%m-%d',m.technical_report) AS tr_date,
				STRFTIME('%Y-%m-%d',m.quotation) AS quotation_date,
				STRFTIME('%Y-%m-%d',m.confirmation) AS confirm_date,
				m.note as repair_note,m.status,status.vie_status as p_status,status.stt_score as p_score,
				m.part_list as return_date
				FROM ((m_list m
					LEFT JOIN scopes ON m.model = scopes.model)
					LEFT JOIN status ON upper(m.status) = upper(status.repair_status))
					LEFT JOIN dealers d ON m.customer = d.customer

			UNION ALL
			
				SELECT DISTINCT
				rma,tf.customer,tf.model,sn,d.dealer,
				scopes.vie as type_s,
				STRFTIME('%Y-%m-%d',tf.tr_date) as tr_date,
				STRFTIME('%Y-%m-%d',tf.quotation_date) as quotation_date,
				STRFTIME('%Y-%m-%d',tf.confirmation) as confirm_date,
				tf.note as repair_note,
				CASE
				WHEN tf.clear IS NULL THEN 'transfer table'
				ELSE tf.clear END as status,
				CASE
					WHEN tf.return is null THEN
					(CASE
						WHEN s.vie_status ="Chờ xác nhận" THEN "Chờ xác nhận (quá 3 tháng)"

						ELSE s.vie_status END)
					ELSE 'Đã trả hàng (không sửa)' END AS p_status,
				 CASE
					WHEN tf.return is null THEN	
					(CASE
						WHEN s.stt_score = "3.0" THEN 9
						ELSE s.stt_score END )
					ELSE 10 END AS p_score,
				(strftime('%Y-%m-%d',tf.return)) as return_date

				FROM ((transfers tf LEFT JOIN scopes ON tf.model = scopes.model)
					LEFT JOIN dealers d ON tf.customer = d.customer)
					LEFT JOIN status s ON upper(tf.old_status) = upper(s.repair_status)

				WHERE  RMA NOT NULL
				AND (strftime('%Y/%m/%d',tf.return) > strftime('%Y/%m/%d',date('now','-7 days'))
				OR tf.return IS NULL)
			
			UNION ALL

				SELECT DISTINCT
					rma,tf.customer,tf.model,sn,d.dealer,
					scopes.vie as type_s,
					STRFTIME('%Y-%m-%d',tf.tr_date) as tr_date,
					STRFTIME('%Y-%m-%d',tf.quotation_date) as quotation_date,
					STRFTIME('%Y-%m-%d',tf.confirmation) as confirm_date,
					tf.note as repair_note,
					CASE
					WHEN tf.clear IS NULL THEN 'complete table'
					ELSE tf.clear END as status,
					CASE
						WHEN tf.return is null THEN 'Đang giao hàng'
						ELSE 'Hoàn tất giao hàng' END AS p_status,
					CASE
						WHEN tf.return is null THEN 7
						ELSE 8 END AS p_score,
					(strftime('%Y-%m-%d',tf.return)) as return_date

					FROM (completed tf LEFT JOIN scopes ON tf.model = scopes.model)
						LEFT JOIN dealers d ON tf.customer = d.customer
					WHERE  RMA NOT NULL
					AND (strftime('%Y/%m/%d',tf.return) > strftime('%Y/%m/%d',date('now','-7 days'))
					OR tf.return IS NULL)

			UNION ALL

				SELECT DISTINCT rma as rma_id,ew.customer,model_d as model,sn,d.dealer,
				s.vie,strftime('%Y-%m-%d',inspection_date) as tr_date,
				CASE
				WHEN rma IS NULL THEN NULL
				ELSE NULL END AS quotation_date,
				CASE
				WHEN rma IS NULL THEN NULL
				ELSE NULL END AS confirm_date,
				CASE
				WHEN rma IS NULL THEN NULL
				ELSE NULL END AS note,
				change_stt as status,st.vie_status as p_status,st.stt_score as p_score,(strftime('%Y-%m-%d',ew.return_date)) as return_date
				
				
				FROM ((exfm_web ew
				LEFT JOIN dealers d ON d.customer = ew.customer)
				LEFT JOIN scopes s ON s.model = ew.model_d)
				LEFT JOIN status st ON st.repair_status = ew.change_stt
				
				WHERE rma > (SELECT max(rma) FROM m_list)

			ORDER BY rma_id
		'''
	abc = pd.read_sql(query,conn)
	
	return abc
def min_completed(printer=False):
	query='''
		SELECT rma
			FROM completed
			WHERE  RMA NOT NULL
					AND (strftime('%Y/%m/%d',return) > strftime('%Y/%m/%d',date('now','-7 days'))
					OR return IS NULL)
		UNION ALL
			SELECT rma
				FROM transfers
				   WHERE  RMA NOT NULL
						AND (strftime('%Y/%m/%d',return) > strftime('%Y/%m/%d',date('now','-7 days'))
						OR return IS NULL)
		UNION ALL
			SELECT rma
			FROM m_list
		ORDER BY RMA
		'''
	min_compl = pd.read_sql(query,conn)

	min_rma = min_compl['RMA'][0]
	yy = min_rma[4:8]
	mm = min_rma[8:10]
	receive_date = yy + '-' + mm + '-01'
	receive_date
	list_rma =min_compl['RMA']
	rma_list=''
	for i in list_rma:
		rma_list += "','"+ i
	rma_list=rma_list[2:]+"'"
	print (f'num of completed {len(min_compl)}')
	if printer: print('\n'+rma_list)
	return rma_list,receive_date

#-------access to ExFM
def export_table(destination,table_name,file_name):
	path_file = os.path.join(destination,file_name)
	table_name.to_excel(path_file + '_' + date_name() + '.xls',index=False)

def download_incomplete(driver):
	Search_RMA = driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]')
	Search_RMA.click()
	sleep(0.5)
	repair_status = Select(driver.find_element_by_xpath('//*[@id="sidIN_REPAIR_STATUS"]'))
	repair_status.select_by_index(1)
	sleep(0.5)
	bt_export_con = driver.find_element_by_xpath('//*[@id="sidEXPORT_CONSOLIDATED_BUTTON_IMAGE"]')
	bt_export_con.click()

def download_complete(driver,receive_date):
	bt_clear = driver.find_element_by_xpath('//*[@id="sidCLEAR_BUTTON_IMAGE"]')
	bt_clear.click()
	Search_RMA = driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]')
	Search_RMA.click()
	sleep(0.5)
	repair_status = Select(driver.find_element_by_xpath('//*[@id="sidIN_REPAIR_STATUS"]'))
	repair_status.select_by_index(10)
	sleep(0.5)
	received_date = driver.find_element_by_xpath('//*[@id="sidIN_FROM_RECEIVE_DATE"]')
	received_date.send_keys(str(receive_date))
	bt_export_con = driver.find_element_by_xpath('//*[@id="sidEXPORT_CONSOLIDATED_BUTTON_IMAGE"]')
	bt_export_con.click()


def task_review():
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
	        LEFT JOIN pf_code pf on c.[rma no.] =pf.[rma no.])
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
	
