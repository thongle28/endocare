import sources.logins as lg
from sqlite3 import connect
import sources.parts_list_0802 as pl
import pandas as pd
from datetime import datetime as dt
import xlwings as xw
from IPython.display import Image, display


class update_ml():

	def __init__(self,conn):
		self.conn = conn

	def new_job(self):
		conn = self.conn
				# new job no empty
		q = f'''
				SELECT DISTINCT
						c.[rma no.],cu.customer_type,customer_name,cu.dealer,ac.territory,e.location,
						c.[create] as [Create Date],c.[Date Installed] AS [Installed Date],
						c.Model,c.Serial_no as Serial,sc.type,sc.mo,
						c.warranty_end_date as [Warranty End],c.recieve_date as [Received Date],
						c.in_inspect_date as [Inspection Date],c.[repair size] as [Size],c.[update user] as PIC,
						CASE
							WHEN sc.status IS NULL THEN c.repair_status
							ELSE sc.status || c.repair_status END AS [Status]

				FROM ((((consolidated c
					 LEFT JOIN pf_code pf on c.[rma no.] = pf.[rma no.])
					 LEFT JOIN customerS cu ON c.[customer_code] = cu.[No.])
					 LEFT JOIN engineers e ON c.[update user] = e.exfm_code)
					 LEFT JOIN scopes sc ON c.model = sc.model)
					 LEFT JOIN acc_tbl_exp ac ON cu.[No.] = ac.[No.]
				WHERE c.[rma no.] > (select max(rma) from new_ml)

				ORDER BY c.[rma no.]

			'''
		newjob = pd.read_sql(q,conn)

				# add new job into ML

		m_list = pd.read_sql('SELECT * FROM new_ml',conn)
		bb = pl.summary_report()
		cur = conn.cursor()
		for i in range(len(newjob)):
			rma = newjob.iloc[i]['RMA No.']
			issue = bb.issues(conn,rma)
			i_no = i + len(m_list) + 1
			for j in range(len(newjob.iloc[i])):
				if newjob.iloc[i][j] == None:newjob.iloc[i][j] =''

			cur.execute(f"""INSERT INTO new_ml(rma,cus_type,Customer,Dealer,Location,installed_date,
											   model,sn,scope_type,model2,wty_end_date,receive,inspection,repair_size,issue,
											   pic,[status]) 
							VALUES('{newjob.iloc[i]['RMA No.']}','{newjob.iloc[i]['customer_type']}',"{newjob.iloc[i]['CUSTOMER_NAME']}",
							'{newjob.iloc[i]['dealer']}','{newjob.iloc[i]['location']}',
							'{str(newjob.iloc[i]['Installed Date'])}','{newjob.iloc[i]['MODEL']}','{newjob.iloc[i]['Serial']}',
							'{newjob.iloc[i]['Type']}','{newjob.iloc[i]['mo']}','{newjob.iloc[i]['Warranty End']}',
							'{newjob.iloc[i]['Received Date']}','{str(newjob.iloc[i]['Inspection Date'])}','{newjob.iloc[i]['Size']}',
							"{issue}","{newjob.iloc[i]['PIC']}","{newjob.iloc[i]['Status']}")
			""")

		# save conn
		conn.commit()
		print('Done')

	def update_job(self):
		conn = self.conn

		q = '''select max([UPDATE_time]) AS max_update FROM new_ml'''
		udt = pd.read_sql(q,conn)
		update_time = udt['max_update'][0]
		print(update_time)

		# filter new update in consolidated 
		cur = conn.cursor()
		q = f'''

				SELECT c.[rma no.],c.customer_name,c.model,c.serial_no,

						c.recieve_date,c.[repair size],c.[update user],c.in_inspect_date,c.repair_status,c.[status info],
						c.[other id],c.last_update_time,c.approval,c.[repair user name],
						c.[update time]
				FROM consolidated c
				WHERE c.[update time] > '{update_time}'
				AND C.[rma no.] IN (SELECT rma FROM new_ml)

				ORDER BY 1

			'''
		update_job = pd.read_sql(q,conn)
		print(len(update_job))
		# update_job

				# update all
		bb = pl.summary_report() # create issue
		cur = conn.cursor()
		m_list = pd.read_sql('SELECT * FROM new_ml',conn)
		for i in range(len(update_job)):
		# i = 0

			rma = update_job.iloc[i]['RMA No.']
			issue = bb.issues(conn,rma)
			print(issue)
			#replace none to empty
			for j in range(len(update_job.iloc[i])):
				if update_job.iloc[i][j] == None:update_job.iloc[i][j] =''

			# UPDATE GENERAL
			q = f""" UPDATE new_ml 
						SET receive = '{update_job.iloc[i]["RECIEVE_DATE"]}',
						inspection = '{update_job.iloc[i]['IN_INSPECT_DATE']}',
						exfm_status = '{update_job.iloc[i]['REPAIR_STATUS']}',
						approval = '{update_job.iloc[i]['Approval']}',
						repair_user = '{update_job.iloc[i]['Repair User Name']}',
						update_time = '{update_job.iloc[i]['Update Time']}',
						update_user = '{update_job.iloc[i]['Update User']}',
						exfm_info = '{update_job.iloc[i]['Status Info']}' 

						WHERE RMA = '{rma}'"""
			cur.execute(q)
			#update fixed item
			fixed_items = ['receive','report_id','inspection','repair_size']
			update_items = ['RECIEVE_DATE','Other ID','IN_INSPECT_DATE','Repair Size']
		#     update_items[fixed_items.index('Report_ID')]       # compare each point

			for item in fixed_items:
				if str(m_list[m_list['rma']==rma][item])=='':
					cur.execute(f'''
									UPDATE new_ml SET '{item}' = '{updatejob.iloc[i][update_items[fixed_items.index(item)]]}'
									WHERE RMA = '{rma}'
								''')
					print(f'Update {rma}: {item} values {updatejob.iloc[i][update_items[fixed_items.index(item)]]}')
		conn.commit()


	def empty_status(self):
		conn = self.conn
				# update empty status
		bb = pl.summary_report() # create issue
		self.bb = bb
		cur = conn.cursor()
		m_list = pd.read_sql('SELECT * FROM new_ml',conn)

		q = f'''
				SELECT m.rma,m.customer,m.model,m.sn,C.[repair_status],c.[status info],m.exfm_status,m.exfm_info,
				c.approval,c.[update time],c.[update user],c.[repair user name],c.recieve_date,c.in_inspect_date,
				c.[repair size]

				FROM new_ml m
				LEFT JOIN consolidated c ON m.rma = c.[rma no.]

			'''
		stt_info = pd.read_sql(q,conn)
		# display(stt_info)

		for i in range(len(stt_info)):
		# i = 0
		#     print(i)
			rma = m_list.iloc[i]['rma']
			issue = bb.issues(conn,rma)
		#     print(rma,issue)

			#replace none to empty
			for j in range(len(stt_info.iloc[i])):
				if stt_info.iloc[i][j] == None:stt_info.iloc[i][j] =''

			# UPDATE GENERAL
			cur.execute(f"""UPDATE new_ml SET issue = '{issue}',
			exfm_status = '{stt_info.iloc[i]['REPAIR_STATUS']}',
			exfm_info = '{stt_info.iloc[i]['Status Info']}',
			approval = '{stt_info.iloc[i]['Approval']}',
			update_time = '{stt_info.iloc[i]['Update Time']}',
			update_user = '{stt_info.iloc[i]['Update User']}',
			receive = '{stt_info.iloc[i]['RECIEVE_DATE']}',
			inspection = '{stt_info.iloc[i]['IN_INSPECT_DATE']}',
			repair_size = '{stt_info.iloc[i]['Repair Size']}'

			WHERE rma = '{rma}'""")

		conn.commit()
		print('Done')

	def export(self,folder_name='files',end_with='.xlsm'):
		# open lastest Master list
		file_name = lg.file_select(folder_name = folder_name, end_with = end_with)
		xw.Book(file_name)

		# final
		conn = self.conn
		q = '''SELECT * FROM new_ml'''
		m_list_final = pd.read_sql(q,conn)
		for i in range(len(m_list_final)):
			rma = m_list_final.iloc[i]['rma']
			issue = self.bb.issues(conn,rma)

			#replace none to empty
			for j in range(len(m_list_final.iloc[i])):
		#         if m_list_final.iloc[i][j] == None: m_list_final.iloc[i][j] =''
				if m_list_final.iloc[i][j] == 'None': m_list_final.iloc[i][j] =''

		m_list_final = m_list_final.drop(['TAT_receive','TAT_PO','TAT_PART'],axis=1)
		dtmp = dt.now().strftime('%y%m%d')
		file_name = f'ML_{dtmp}.xlsx'
		m_list_final.to_excel(file_name,index = False)
		xw.Book(file_name)

