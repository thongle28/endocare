# read file consolidated from ExFM
import sources.logins as lg 
import sources.pending_process as pp
import sources.web_process as wp
import os
from time import sleep
import warnings
warnings.simplefilter("ignore")

#Select source master list
global ver
ver='v 2.0.1'
'''
	2.0.1 change login  by file encrypts
	1.1.2 fix waiting for long time download completed
	1.0.2 remove export empty rcode and part name
'''

print('convert Data from ExFM to EndoCare.')
print(f'------------{ver}------------\n')
print('Select file Master List from folder "files"')
file = lg.file_select(end_with='.xlsm',folder_name='files')

# read file master list 
conn = pp.pending_sql(file)
# def database_process(folder_name,file):
pp.database_process('files','database.xlsx')




#Select credential file
print("")
print("Select Credential File")
key = 'cLAiAdZU1U0sxWWK8sxhF0IWIBP4KrR3xhdu3x1EDQI='
name,pw = lg.save_id('exfm.txt',key)
# def login_url(target,credential,download_folder=''):

driver = lg.login_url('exfm',name,pw,folder_name='exports')

#-------download incomplete file------------
res_time = 5
pp.download_incomplete(driver)
sleep(res_time)
# incomp,ctime =  lg.file_latest('exports') # 1.1.2
incomp,ctime =  lg.file_latest('exports') #2.0.0
try:
	while incomp.split('.')[2]=='crdownload':
		incomp,ctime =  lg.file_latest('exports')
		print('\nPlease wait for downloading...')
		sleep(1)
except:
	incomp,ctime =  lg.file_latest('exports')

sleep(res_time)
print (f"incomplete {incomp}")
date_name = pp.date_name()

incomplete = 'Consolidated_InCompleted_' + date_name + '.xls'
sleep(2)
try:
	os.remove(os.path.join('exports',incomplete))
except:
	pass
os.rename(os.path.join('exports',incomp),os.path.join('exports',incomplete))
print (f"Download incomplete file and rename {incomplete} to folder 'exports'.")


#-------download complete file------------
rma_list,receive_date = pp.min_completed()
pp.download_complete(driver,receive_date)
sleep(res_time+5)
complete,ctime =  lg.file_latest('exports') #incompleted

# ver 1.0.2
# try:
# 	while complete.split('.')[2]=='crdownload':
# 		complete,ctime =  lg.file_latest('exports')
# 		print('\nPlease wait for downloading...')
# 		sleep(res_time+1)
# except:
# 	complete,ctime =  lg.file_latest('exports')

# ver 1.1.2
try:
	while complete.startswith('Consolidated_InCompleted') or complete.split('.')[2]=='crdownload':
		print('Please wait for downloading...')
		sleep(2)
		complete,ctime =  lg.file_latest('exports')
except:
	print(complete)
completed = 'Consolidated_Completed_' + date_name + '.xls'
sleep(res_time)
try:
	os.remove(os.path.join('exports',completed))
except:
	pass
#ver 1.0.2
# os.rename(os.path.join('exports',complete),os.path.join('exports',completed))

# ver 1.1.2
if complete.startswith('Consolidated_InCompleted'):
	print('check point change name')
	complete,ctime =  lg.file_latest('exports')
	sleep(2)

os.rename(os.path.join('exports',complete),os.path.join('exports',completed))

print (f"Download completed file and rename {completed} to folder 'exports'.")

#--------add incomplete to sql-------------
pp.consolidated_sql('exports',incomplete,False)

#--------add complete to sql-------------
pp.consolidated_sql('exports',completed,True)

#--------create new table---------------------
pp.new_table(False) #new table for incomplete
pp.new_table(True) # new table for completed

#----export exfm_web--------------------------
exfm_web = pp.exfm_web(rma_list)
pp.export_table('exports',exfm_web,'exfm_web')
exfm_web.to_sql('exfm_web',conn,index=False,if_exists='replace')
# create file pending
pending_tb = pp.pending_file()
if not os.path.exists('exports'):
	os.makedirs('exports')
# def export_table(destination,table_name,file_name):
pp.export_table('exports',pending_tb,'pending')


#-----create part name table------------------
part_name = pp.part_name(rma_list)
pp.export_table('exports',part_name,'part_name')

#-----create r_code table------------------
check_rcode = pp.r_code(rma_list,check_r=True)
a = check_rcode.shape[0]
if a>0:
	print (f'please update {a} rows empty vie_name for rcode')
	# pp.export_table('exports',check_rcode,'r_code_empty_name')
r_code = pp.r_code(rma_list)
pp.export_table('exports',r_code,'r_code')


#-----create part_name table------------------
check_part_name = pp.part_name(rma_list,check_part=True)
a = check_part_name.shape[0]
if a>0:
	print (f'please update {a} rows empty vie_name for part name')
	# pp.export_table('exports',check_part_name,'empty_part_name')
part_name = pp.part_name(rma_list)
pp.export_table('exports',part_name,'part_name')

task_list = pp.task_review()
task_list.to_excel(os.path.join('exports','task_list_')+pp.date_name()+'.xls',index=False)

driver.close()


