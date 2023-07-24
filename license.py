from cryptography.fernet import Fernet
import os
# Generate a key
# key = Fernet.generate_key()
key = str('HfOfd2sE7RO8TQ21qckawMG82uTchV_kQC4o7EZYTPE=').encode()


def pw_encode(message,key):
	fernet = Fernet(key)
#     fernet = key
	message = message.encode()
	b_msg = fernet.encrypt(message)
	return b_msg.decode()


def pw_decode(encrypted_message,key):
	
	fernet = Fernet(key)
#     fernet = key
	encrypted_message = encrypted_message.encode()
	b_encrypt_msg = fernet.decrypt(encrypted_message)
	return b_encrypt_msg.decode()

file = 'license\\log.txt'
while True:
	if os.path.exists(file): # read data key
		logs = open(file,'r')
		data = logs.read()
		logs.close()
		key_date = data.split('\n')[0]
		try:
			# check Date code
			dt_date_code = dt.strptime(pw_decode(key_date,key),'%Y-%m-%d')
		except:
			print('Wrong key. Please update new key.')
			break
		if dt.now() <= dt_date_code:
			
			print(f"License key will expire on {dt_date_code.strftime('%Y-%m-%d')}")
			break
			sources.main_230512
		else:
			print('License key was expired. Please update new key.')
			str_key = input('Date Key: ')
			files = open(file,'w')
			files.write(str_key)
			files.close()
	else:
		str_key = input('Date Key: ')
		files = open(file,'w')
		files.write(str_key)
		files.close()
  