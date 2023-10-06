from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from time import sleep

class Auto_QC():


	def __init__(self):
		pass 

	def auto_accept(self,driver):
		try:
			WebDriverWait(driver, 10).until(EC.alert_is_present())
			driver.switch_to.alert.accept()
			print('Bypass Alert!!!')
		except:
			pass

	def basic_inspection(self,driver,sn,scope_count,repair_size,defect_note,qc_pass):
		inspection_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_3"]').get_attribute('value')

		if inspection_date =='':
			bt_inspection = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_3"]')
			bt_inspection.click()

			try:
				if str(int(scope_count)).isnumeric():
					sc = driver.find_element_by_xpath('//*[@id="sidSCOPE_CONNECT_COUNT"]')
					sc.clear()
					sc.send_keys(int(scope_count))
					print(f'Scope count {scope_count}')
			except:
				print('Scope count empty')
			#answer 3 questions
			ans_1 = Select(driver.find_element_by_xpath('//*[@id="sidFDA_QUESTION_RESULT_CODE_1"]'))
			ans_1.select_by_index(2)
			ans_2 = Select(driver.find_element_by_xpath('//*[@id="sidFDA_QUESTION_RESULT_CODE_2"]'))
			ans_2.select_by_index(2)
			ans_3 = Select(driver.find_element_by_xpath('//*[@id="sidFDA_QUESTION_RESULT_CODE_3"]'))
			ans_3.select_by_index(2)
			print('\nSay "No" for 3 questions.')
			driver.find_element_by_xpath('//*[@id="sidREPAIR_SIZE_NAME"]').send_keys(repair_size) #repair type
			print(f'Input repair size {repair_size}')
			# //*[@id="inspection[0].sidDEF_POS_CODE"]
			driver.find_element_by_xpath('//*[@id="sidNOTE_DEFECT"]').clear()
			driver.find_element_by_xpath('//*[@id="sidNOTE_DEFECT"]').send_keys(defect_note)


			 #repair type

			#fill QC pass code
			p_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidDEF_POS_CODE"]')
			f_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidDEF_CASE_CODE"]')
			cd_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidDEF_CAUSE_CODE"]')
			r_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidLBR_CODE"]')
			if qc_pass =='OK' or qc_pass == 'Yes':
				p_code.send_keys('P165')
				f_code.send_keys('F103')
				cd_code.send_keys('CD07')
				r_code.send_keys('R025')
				driver.find_element_by_xpath('//*[@id="sidREPAIR_SIZE_NAME"]').send_keys('Maintain')
				print('Apply code for QC pass')
			else:
				print('Inspection for broken item')
			bt_save = driver.find_element_by_xpath('//*[@id="sidUPDATE_BUTTON_IMAGE"]')
			bt_save.click()
	#         bt_return = driver.find_element_by_xpath('//*[@id="sidBACK_BUTTON_IMAGE"]')
	#         bt_return.click()
			try:
				auto_accept(driver)
				print ('Auto accept')
			except:
				pass

	# return Create RMA
	def return_unrepair(self,driver,sn,scope_count='',approval_index='',evaluation='inspection OK'):
		print(f'Start {sn}')
		print('-------------------')
		driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]').click() #click RMA Search
		sn_field = driver.find_element_by_xpath('//*[@id="sidIN_BODY_NO"]')
		sn_field.clear()
		sn_field.send_keys(sn)
		driver.find_element_by_xpath('//*[@id="sidSEARCH_BUTTON_IMAGE"]').click()

		#move to each job page
		driver.find_element_by_xpath('//*[@id="search_result[0].sidRS_REPAIR_ID_data"]').click()



		# receive
		receive_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_2"]').get_attribute('value')
		if receive_date =='':
			bt_receive = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_2"]')
			bt_receive.click()
			driver.find_element_by_xpath('//*[@id="sidCOMPLETE_BUTTON_IMAGE"]').click()
		else:
			print( 'Receive completed.')

		#step in inspection
		inspection_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_3"]').get_attribute('value')
		if inspection_date =='':
			bt_inspection = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_3"]')
			bt_inspection.click()
			if scope_count !='':
				scope_count = driver.find_element_by_xpath('//*[@id="sidSCOPE_CONNECT_COUNT"]')
				scope_count.clear()
				scope_count.send_keys(int(scope_count))
			#answer 3 questions
			ans_1 = Select(driver.find_element_by_xpath('//*[@id="sidFDA_QUESTION_RESULT_CODE_1"]'))
			ans_1.select_by_index(2)
			ans_2 = Select(driver.find_element_by_xpath('//*[@id="sidFDA_QUESTION_RESULT_CODE_2"]'))
			ans_2.select_by_index(2)
			ans_3 = Select(driver.find_element_by_xpath('//*[@id="sidFDA_QUESTION_RESULT_CODE_3"]'))
			ans_3.select_by_index(2)
			print('\nSay "No" for 3 questions.')

			driver.find_element_by_xpath('//*[@id="sidREPAIR_SIZE_NAME"]').send_keys('Maintain') #repair type
			# //*[@id="inspection[0].sidDEF_POS_CODE"]
			driver.find_element_by_xpath('//*[@id="sidNOTE_DEFECT"]').send_keys(evaluation)

			# clear first row
			driver.find_element_by_xpath('//*[@id="TH_inspection"]/tbody/tr[1]/td[1]/div/label/input').click()
			driver.find_element_by_xpath('//*[@id="_eventremoveline_inspection"]').click()

			#fill QC pass code
			p_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidDEF_POS_CODE"]')
			f_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidDEF_CASE_CODE"]')
			cd_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidDEF_CAUSE_CODE"]')
			r_code = driver.find_element_by_xpath('//*[@id="inspection[0].sidLBR_CODE"]')

			p_code.send_keys('P165')
			f_code.send_keys('F103')
			cd_code.send_keys('CD07')
			r_code.send_keys('R025')

			sleep(2)

			driver.find_element_by_xpath('//*[@id="sidUPDATE_BUTTON_IMAGE"]').click()
			driver.find_element_by_xpath('//*[@id="sidCOMPLETE_BUTTON_IMAGE"]').click()
		else:
			print( 'Inspection completed.')

		# part select
		part_select_date =  driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_4"]').get_attribute('value')
		if part_select_date =='':
			bt_part_select = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_4"]')
			bt_part_select.click()
			bt_complete = driver.find_element_by_xpath('//*[@id="sidCOMPLETE_BUTTON_IMAGE"]')
			bt_complete.click()
			self.auto_accept(driver)
		else:
			print( 'Part Select completed.')

		# Change status
		stt = Select(driver.find_element_by_xpath('//*[@id="sidHOLD_CODE"]'))
		stt.select_by_index(0)
		bt_save = driver.find_element_by_xpath('//*[@id="sidUPDATE_BUTTON_IMAGE"]')
		# bt_save.click()

		#Change status Info
		stt_info = Select(driver.find_element_by_xpath('//*[@id="sidTEMPORARY_STATUS_CODE"]'))
		stt_info.select_by_index(6) # Awaiting contract Billing

		bt_save.click()
		
		# authorized inspection
		authorized_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_5"]').get_attribute('value')
		if authorized_date =='':
			bt_authorized = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_5"]')
			bt_authorized.click()
			approval = Select(driver.find_element_by_xpath('//*[@id="sidRPR_COMP_TYPE_CODE"]'))

			approval_list ={'--':0,
						'Approval':1,
						'Decline':2,
						'Inspection':3,
						'Transfer':4,
						'Cancel':5,
						'No Fault Found':6}
			approval.select_by_index(approval_list[approval_index]) # Ispection
			bt_complete = driver.find_element_by_xpath('//*[@id="sidAUTHORIZATION_COMPLETE_BUTTON_IMAGE"]')
			bt_complete.click()

		else:
			print( 'Authorized completed.')

		# ship
		shipped_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_10"]').get_attribute('value')
		if shipped_date =='':
			bt_shipped = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_10"]')
			bt_shipped.click()
			bt_complete = driver.find_element_by_xpath('//*[@id="sidCOMPLETE_BUTTON_IMAGE"]')
			bt_complete.click()
		else:
			print( 'Ship completed.')
		print('------------------------')
		print(f'End job for {sn}\n')
	# QC and ship repair-job
	def qc_and_ship_out(self,driver,sn,ship_out=False):
		driver.find_element_by_xpath('//*[@id="RPR_SEARCH_LINK"]').click() #click RMA Search
		try: 
			auto_accept(driver)
		except:
			pass

		sn_field = driver.find_element_by_xpath('//*[@id="sidIN_BODY_NO"]')
		sn_field.clear()
		# sn = df['Serial'][0]
		sn_field.send_keys(sn)
		driver.find_element_by_xpath('//*[@id="sidSEARCH_BUTTON_IMAGE"]').click()
		stt = driver.find_element_by_xpath('//*[@id="search_result[0].sidRS_CURRENT_STATUS_NAME"]').get_attribute('innerHTML')
		print(f'Status: {stt}')
		#move to each job page
		driver.find_element_by_xpath('//*[@id="search_result[0].sidRS_REPAIR_ID_data"]').click()

		#check status
		receive_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_2"]').get_attribute('value')
		inspection_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_3"]').get_attribute('value')
		part_select_date =  driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_4"]').get_attribute('value')
		authorized_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_5"]').get_attribute('value')
		repair_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_8"]').get_attribute('value')
		qc_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_9"]').get_attribute('value')
		shipped_date = driver.find_element_by_xpath('//*[@id="sidDISP_RPR_STATUS_DATE_10"]').get_attribute('value')

		# QC step
		if repair_date=='':
			print('Completed repair first')
		else:

			bt_qc = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_9"]')
			bt_qc.click()
			driver.find_element_by_xpath('//*[@id="out_inspections[0].sidINS_AIR_CHECK_RESULT"]').send_keys('OK')
			driver.find_element_by_xpath('//*[@id="out_inspections[0].sidINS_VIDEO_CHECK_RESULT"]').send_keys('OK')
			driver.find_element_by_xpath('//*[@id="out_inspections[0].sidINS_ISO86001_CHECK_RESULT_CODE"]').click()
			try:
				driver.find_element_by_xpath('//*[@id="sidUPDATE_BUTTON_IMAGE"]').click()
				driver.find_element_by_xpath('//*[@id="sidCOMPLETE_BUTTON_IMAGE"]').click()
				print(f'QC {sn} Done!')
			except:
				print('Already QC')

		# ship out
		if ship_out:
			bt_shipped = driver.find_element_by_xpath('//*[@id="sidRPR_STATUS_BUTTON_10"]')
			bt_shipped.click()
			driver.find_element_by_xpath('//*[@id="sidUPDATE_BUTTON_IMAGE"]').click()
			driver.find_element_by_xpath('//*[@id="sidCOMPLETE_BUTTON_IMAGE"]').click()
			print(f'Ship out {sn} done!')