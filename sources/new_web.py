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

	def rma_list(self):
		conn = self.conn
		rma_list = []

		q = '''
				SELECT rma FROM new_ml
			'''
		try:
			rl = pd.read_sql(q,conn)
		#     display(rl)
			for rma in rl['rma']:
				rma_list.append(rma)
		except Exception as e:
			
			print(e)
		print(len(rma_list))
		
		q = '''
				SELECT rma,[return] FROM transfers
				
				WHERE (strftime('%Y/%m/%d',[return]) > strftime('%Y/%m/%d',date('now','-7 days'))
					OR [return] IS NULL)
			'''
		try:
			rl = pd.read_sql(q,conn)
		#     display(rl)
			for rma in rl['RMA']:
				rma_list.append(rma)
		except Exception as e:
			print(e)
		len(rma_list)

		# RMA LIST FROM TRANSFER
		q = '''
				SELECT rma,[return] FROM completed
				
				WHERE (strftime('%Y/%m/%d',[return]) > strftime('%Y/%m/%d',date('now','-7 days'))
					OR [return] IS NULL)
			'''
		try:
			rl = pd.read_sql(q,conn)
			display(rl)
			for rma in rl['RMA']:
				rma_list.append(rma)
		except Exception as e:
			print(e)
		len(rma_list)

		# return rma_list
		self.rma_list = rma_list

	def exfm_web(self):
		conn = self.conn
		rma_list = self.rma_list
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
		        WHERE c.[rma no.] IN {tuple(rma_list)}
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
		rma_list = self.rma_list

		q = '''
		        SELECT DISTINCT 
		            m.rma AS rma_id,m.customer,m.model,m.sn,d.dealer,scopes.vie AS type_s,
		            STRFTIME('%Y-%m-%d',m.inspection) AS tr_date,
		            STRFTIME('%Y-%m-%d',m.quotation) AS quotation_date,
		            STRFTIME('%Y-%m-%d',m.confirmation) AS confirm_date,
		            m.note as repair_note,m.status,status.vie_status as p_status,status.stt_score as p_score,
		            STRFTIME('%Y-%m-%d',m.return_date) AS return_date
		            FROM ((new_ml m
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
		            STRFTIME('%Y-%m-%d',tf.quotation_date) as tr_date,
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

		            WHERE rma > (SELECT max(rma) FROM new_ml)
		    '''
		try:
		    pending_file = pd.read_sql(q,conn)
		    display(pending_file)
		    self.pending_file = pending_file
		except Exception as e:
		    print(e)

	def parts_name(self):
		conn = self.conn
		rma_list = self.rma_list
		q = f'''
		        SELECT 
		            DISTINCT p.[RMA No.] AS rma,p.PART_NO AS part_no,
		            p.Part_description AS part_description,pn.vie as vie_name
		            FROM parts p
		            LEFT JOIN part_name pn ON p.part_description = pn.part_description

		            WHERE rma IN {tuple(rma_list)}
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
		rma_list = self.rma_list
		q = f'''
		        SELECT r.[rma no.] as rma,pf.r_code as r_code,
		        SUBSTRING(pf.r_description,1,50) as r_description,
		        pf.vie_name

		        FROM r_code pf
		        LEFT JOIN repair_code r ON pf.r_code = r.r_code
		        WHERE rma IN {tuple(rma_list)}
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

