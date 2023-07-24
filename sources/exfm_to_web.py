#check version
ver = '2.0.1'
'''
------------------Release-----------------
2.0.1: 16/11/2022 change login to file encypts
1.0.2: add option dont download new files
1.0.1: add log out 

'''
print('Welcome to Endo Service')
print('Software export data from ExFM and import to web noisoifujifilm.vn')
print('-----Version {}-----'.format(ver))

# run export_needed_file

while True:
	# asking export new files
	ans = input('Do you want to exports new files (Y/n)?')
	if ans.upper() != 'N': 

		import sources.logins as lg 
		lg.backups('exports','backups')
		import sources.export_files

	# run import_file
	import sources.auto_import

	run_again = input('Do you want to run again (y/N)?')
	if run_again.upper() !='Y':
		break # for while

print('Good Bye!!!')
