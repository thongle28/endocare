import pandas as pd
from sqlite3 import connect
from datetime import datetime as dt
from IPython.display import Image, display
import os
import openpyxl

class data_process():

	def __init__(self,conn):
		self.conn = conn
		folder_name = 'New Web'
		try:
			os.mkdir(folder_name)
			print(f'folder {folder_name} was created.')
		except:
			print(f'Folder {folder_name} exists')
		self.folder_name = folder_name

	
	def exfm_web(self):
		conn = self.conn
		# rma_list = self.rma_list
		q = f'''
				SELECT DISTINCT
					c.[rma no.] AS rma,c.customer_name,c.customer_code AS customer,
					CASE
						WHEN e.location IS NULL THEN (SELECT e.location FROM consolidated c LEFT JOIN engineers e ON c.IN_INSPECT_USER_NAME = e.exfm_name)
						ELSE e.location END AS location,
					c.model AS model_d,c.serial_no AS sn,c.[Date Installed] AS install_date,
					c.[WARRANTY_END_DATE] AS warranty_date,
					c.[RECIEVE_DATE] AS receive_date,c.[IN_INSPECT_DATE] AS inspection_date,
					r.[start time] AS start_time,r.[end time] AS end_time,
					CASE
						WHEN c.[repair size] = 'Major' THEN 'Nặng'
						WHEN C.[repair size] = 'Minor' THEN 'Nhẹ'
						WHEN c.[repair size] IS NOT NULL THEN 'Khác' 
						ELSE c.[repair size] END AS repair_size,
					'Other' as issue_part,
					e.ks,c.[other id] AS wr_report,c.REPAIR_STATUS AS change_stt,c.approval,
					CASE
						WHEN c.approval  IN ('Decline','Cancel') THEN 'Chờ xác nhận (quá 3 tháng)'
						WHEN c.[create] < c.warranty_end_date AND c.repair_status ='Inspection' AND c.[other id] IS NULL THEN 'Chờ duyệt Bảo hành'
						WHEN c.[other id] IS NOT NULL AND c.repair_status ='Inspection' THEN 'Chuẩn bị linh kiện'
						ELSE s.vie_status END AS repair_status,
					CASE
					WHEN c.approval  IN ('Decline','Cancel') THEN 9
					ELSE s.stt_score END  AS e_score,
					CASE
						WHEN c.[other id] NOT NULL AND UPPER(c.[other id]) NOT LIKE '%REJECTED%' THEN 0.5
						ELSE 0 END AS w_score,
					
					c.[update time] AS update_time,c.Shipped as return_date

					FROM 
					((consolidated c LEFT JOIN engineers e ON c.[update user] = e.exfm_code)
					LEFT JOIN status s ON c.[repair_status] = s.repair_status)
					LEFT JOIN repair_code r ON c.[rma no.] = r.[rma no.]
				
			'''
		try:
			exfm_web = pd.read_sql(q,conn)
			exfm_web.to_sql('exfm_web',conn,index=False,if_exists = 'replace')
			self.exfm_web = exfm_web
		#     display(exfm_web)
		except Exception as e:
			print(e)

	def pending_file(self):
		conn = self.conn
		# rma_list = self.rma_list

				# pending_file on consolidated
		q = '''
				SELECT DISTINCT
					[rma no.] AS rma_id,customer_name as customer,c.model, serial_no AS sn,
					Territory, s.vie as type_s,
					STRFTIME('%Y-%m-%d',c.in_inspect_date) AS tr_date,
					STRFTIME('%Y-%m-%d',c.[part select date]) AS quotation_date,
					STRFTIME('%Y-%m-%d',c.[authorized_date]) AS confirm_date,
					'' as repair_note,c.repair_status as status,status.vie_status as p_status,status.stt_score as p_score,
					STRFTIME('%Y-%m-%d',c.shipped) as return_date

					FROM ((consolidated c
					LEFT JOIN acc_tbl_exp a ON c.customer_code = a.[no.])
					LEFT JOIN scopes s ON c.model = s.model)
					LEFT JOIN status ON c.repair_status = status.repair_status

			'''
			
		try:
			pending_file = pd.read_sql(q,conn)
			display(pending_file)
			self.pending_file = pending_file
		except Exception as e:
			print(e)

	def parts_name(self):
		conn = self.conn
		# rma_list = self.rma_list
		q = f'''
				SELECT 
					DISTINCT p.[RMA No.] AS rma,p.PART_NO AS part_no,
					p.Part_description AS part_description,pn.vie as vie_name
					FROM parts p
					LEFT JOIN part_name pn ON p.part_description = pn.part_description

					
				ORDER BY rma
			'''
		try:
			parts = pd.read_sql(q,conn)
			display(parts)
			self.parts = parts
		except Exception as e:
			print(e)

	def r_code(self):
		conn = self.conn
		# rma_list = self.rma_list
		q = f'''
				SELECT r.[rma no.] as rma,pf.r_code as r_code,
				SUBSTRING(pf.r_description,1,50) as r_description,
				pf.vie_name

				FROM r_code pf
				RIGHT JOIN repair_code r ON pf.r_code = r.r_code
				
			'''
		try:
			r_code = pd.read_sql(q,conn)
			display(r_code)
			self.r_code = r_code
		except Exception as e:
			print(e)

	def export_files(self):
		today = dt.now().strftime('%y%m%d')
		folder_name = self.folder_name
		pending_name = f'pending_{today}.xls'
		exfm_name = f'exfm_web_{today}.xls'
		parts_name = f'part_name_{today}.xls'
		r_code_name = f'r_code_{today}.xls'

		self.pending_file.to_excel(f'New Web/{pending_name}',index = False)
		self.exfm_web.to_excel(f'New Web/{exfm_name}',index = False)
		self.parts.to_excel(f'New Web/{parts_name}',index = False)
		self.r_code.to_excel(f'New Web/{r_code_name}',index = False)
		# writer = pd.ExcelWriter(pending_name)
		# self.pending_file.to_excel(writer,sheet_name='Sheet1',index=False)
		# writer = pd.ExcelWriter(exfm_name)
		# self.exfm_web.to_excel(writer,sheet_name='Sheet1',index=False)
		# writer = pd.ExcelWriter(parts_name)
		# self.parts.to_excel(writer,sheet_name='Sheet1',index=False)
		# writer = pd.ExcelWriter(r_code_name)
		# self.r_code.to_excel(writer,sheet_name='Sheet1',index=False)

	def export_csv(self):
		today = dt.now().strftime('%y%m%d')
		folder_name = self.folder_name
		pending_name = f'pending_{today}.csv'
		exfm_name = f'exfm_web_{today}.csv'
		parts_name = f'part_name_{today}.csv'
		r_code_name = f'r_code_{today}.csv'
		
		self.pending_file.to_csv(f'New Web/{pending_name}',index = False)
		self.exfm_web.to_csv(f'New Web/{exfm_name}',index = False)
		self.parts.to_csv(f'New Web/{parts_name}',index = False)
		self.r_code.to_csv	(f'New Web/{r_code_name}',index = False)
		

if __name__ == "__main__":
	pass
# else:
# 	run_all = data_process(conn)
# 	run_all.rma_list()
# 	run_all.exfm_web()
# 	run_all.pending_file()
# 	run_all.parts_name()
# 	run_all.r_code()
# 	run_all.export_files()

