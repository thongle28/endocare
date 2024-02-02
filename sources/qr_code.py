import qrcode
import xlwings as xw
from sqlite3 import connect
import pandas as pd
import os
import pathlib

class main():

	def __init__(self,folder_name,wb_name):

		self.folder_name = folder_name
		self.wb_name = wb_name

	def qr_printer(self,folder_name,in_name,out_name=''):

		folder_name = self.folder_name

		# Generate QR code
		qr = qrcode.QRCode(
			version=1,
			error_correction=qrcode.constants.ERROR_CORRECT_H,
			box_size=10,
			border=0.2,
		)
		qr.add_data(in_name)
		qr.make(fit=True)

		# Create an image from the QR code
		qr_img = qr.make_image(fill_color="black", back_color="white")

		# Save the QR code as an image file
		if out_name =='': out_name = in_name
		qr_img.save(f"{folder_name}/{out_name}.png")


	def write_template(self):
		
		#open template
		# wb = xw.Book('templates/QR Template.xlsx')
		wb = xw.Book(self.wb_name)
		folder_name = self.folder_name

		zb = wb.sheets('Zebra')
		li = wb.sheets('List')
		zb.copy()
		for sh in wb.sheets:
			if sh.name not in ('List','Zebra') and not sh.api.Visible:
				sh.api.Visible = True
				print(sh.name)
				break

		#read from list
		lrow = li.range('B' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
		print(f'Have {lrow} items in this list')

				#create rma
		
		# i = 0
		margin = zb.range('1:1').api.RowHeight
		try:
			margin = float(input(f'Modify margin? {margin}: ' or int(margin)))
			zb.range('1:1').api.RowHeight = margin
		except Exception as e:
			print(e)

		for k in range(3,lrow+1):
			i = k - 3

			# format margin
			sh['1:1'].offset(i*7,0).api.RowHeight = margin
			rma = li.range('B3').offset(i,0).value
			sn = li.range('C3').offset(i,0).value
			# RMA QR Code
			self.qr_printer(folder_name,rma)

			#SN QR Code
			self.qr_printer(folder_name,sn)

			#Link web
			link = f'https://noisoifujifilm.vn/quick_search/{rma}/{sn}'
			self.qr_printer(folder_name,link,f'{rma}_{sn}')
			
			#add qr image
			#B2 E2 H2 offset 7
			rma_img = 'B2'
			sn_img = 'E2'
			link_img = 'H2'
			path = pathlib.Path().absolute()
			size = 72
			#add image rma
			sh.pictures.add(os.path.join(path,folder_name,f'{rma}.png'),name = rma,
							top = sh.range(rma_img).offset(7*i,0).top,left = sh.range(rma_img).offset(7*i,0).left)
			sh.pictures(rma).width = size
			sh.pictures(rma).height = size

			#add image sn
			sh.pictures.add(os.path.join(path,folder_name,f'{sn}.png'),name = sn,
							top = sh.range(sn_img).offset(7*i,0).top,left = sh.range(sn_img).offset(7*i,0).left)
			sh.pictures(sn).width = size
			sh.pictures(sn).height = size

			#add image link
			sh.pictures.add(os.path.join(path,folder_name,f'{rma}_{sn}.png'),name = f'{rma}_{sn}',
							top = sh.range(link_img).offset(7*i,0).top,left = sh.range(link_img).offset(7*i,0).left)
			size = 87
			sh.pictures(f'{rma}_{sn}').width = size
			sh.pictures(f'{rma}_{sn}').height = size

			#word
			sh.range('B7').offset(7*i,0).value = rma
			sh.range('E7').offset(7*i,0).value = sn
			sh.range('G2').offset(7*i,0).value = sn
			sh.range('I3').offset(7*i,0).value = 'Scan'
			sh.range('I4').offset(7*i,0).value = 'Me'

			scan_img ='I4'
			sh.pictures.add(os.path.join(path,'templates','scan_me.png'),name = f'scan{i}',
						   top = sh.range(scan_img).offset(7*i,0).top + 8,
							left =sh.range(scan_img).offset(7*i,0).left + 5)

		print('Done')
		lrowB = sh.range('B' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
		sh.range(f'{lrowB+1}:1000').api.Delete()
