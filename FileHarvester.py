# DKats 2020 
# This script will take a target folder and collect all files the user chose for harvest. Copied files are
# being renamed if needed and a CSV report file is being created with the original file metadata. Copied files and
# CSV report are being ziped and md5 hashed. A log file is being created after the completion of the script. 
# A calculate action will provide the user with the total size of the files chosen to collect.
# IMPORTANT: 1. ALWAYS run as admin, 2. Do NOT choose an output Folder in the Input folder chosen to be searched...
# More info on the documentation at https://github.com/D-Kats/FileHarvester
#---Thanks to---
# textract module from Dean Malmgren at https://pypi.org/project/textract/
# PySimpleGUI module from  MikeTheWatchGuy at https://pypi.org/project/PySimpleGUI/
# executable's icon downloaded from www.freeiconspng.com
import PySimpleGUI as sg
import textract
import os
import webbrowser
import shutil
import hashlib
from datetime import datetime


#---functions definition
def finalize(outputFolder, destinationFolderName):
	os.chdir(outputFolder)
	shutil.make_archive(f'{destinationFolderName}', 'zip', f'{destinationFolderName}')
	print(f'Archive file created with name: {destinationFolderName}.zip')
	md5 = hashlib.md5(open(f'{destinationFolderName}.zip', 'rb').read()).hexdigest()
	print(f'Md5 checksum of the archive file is {md5}')
	try:
		with open(f'{destinationFolderName}\\FileHarvester.log', 'w', encoding='utf-8') as fout:						
			fout.write(window['-OUT-'].Get())
			print('Log file was created successfully!')
	except:
		print('Error writing log')
	sg.PopupOK(f'Harvesting Completed successfully!', title=':)', background_color='#2a363b')

def calculate(CalculateInputFolder, CalculateFileCategory):
	totalSize = 0
	errorFlag = False
	if CalculateFileCategory == 'documents':
		os.chdir(CalculateInputFolder)
		for root, folders, files in os.walk(CalculateInputFolder):
			os.chdir(root)
			for file in files:
				if file.endswith(('doc', 'DOC', 'docx', 'DOCX', 'xls', 'XLS', 'xlsx', 'XLSX', 'PDF', 'pdf', 'pptx', 'PPTX', 'ppt', 'PPT', 'ODT', 'odt', 'ODF', 'odf')):
					try:
						totalSize+=os.stat(file).st_size
					except:
						errorFlag = True
						continue				
	else:
		os.chdir(CalculateInputFolder)
		for root, folders, files in os.walk(CalculateInputFolder):
			os.chdir(root)
			for file in files:
				if file.endswith(('jpg', 'JPG', 'jpeg', 'JPEG', 'PNG', 'png', 'AVI', 'avi', 'MOV', 'mov', 'MP4', 'mp4', '3GP', '3gp')):
					try:
						totalSize+=os.stat(file).st_size
					except:
						errorFlag = True
						continue
	if errorFlag:
		print('Warning! Some files could not be read. Total file size may be bigger than reported')
	return totalSize

def harvester(inputFolder, destinationFolder, FileCategory, OptionalKeywords=''):
	TotalFilesCopied = 0
	count = 0
	csvContent = 'No,File Original Path,Original Creation Time, Original Modified Time,Filename in Output Folder,Flag\n'
	print('Initializing...')

	if FileCategory == 'media': # media files harvesting
		for root, folders, files in os.walk(inputFolder):
			for file in files:
				os.chdir(root)
				if file.endswith(('jpg', 'JPG', 'jpeg', 'JPEG', 'PNG', 'png', 'AVI', 'avi', 'MOV', 'mov', 'MP4', 'mp4', '3GP', '3gp')):
					count+=1
					if file in os.listdir(destinationFolder): # an yparxei arxeio me idio onoma hdh ston output fakelo, to metonomazw gia na mh diagrafei
						print(f'{file} has the same name with another file in the output folder...')
						print('Renaming...')
						destname = file
						while destname in os.listdir(destinationFolder): # oso to arxeio exei idio onoma me kapoio sto output toso synexizw na tou pros8etw to _RN sto onoma
							destfilename = f'{os.path.splitext(destname)[0]}'+'_RN'
							destname = f'{destfilename}{os.path.splitext(destname)[1]}'
						print(f'Renamed Successfully! New name: {destname}')
						try:
							shutil.copy2(file, f'{destinationFolder}\\{destname}')							
							print(f'Copying initial file: {file} to output folder as: {destname}')
							TotalFilesCopied += 1
							csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{destname},-\n'
						except:
							print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
							csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
					else: # den yparxei arxeio me to idio onoma ston output folder
						try:
							shutil.copy2(file,destinationFolder)
							print(f'Copying {file} to output folder')
							TotalFilesCopied += 1
							csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{file},-\n'
						except:
							print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
							csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
	else: # Document files harvesting
		if len(OptionalKeywords) == 0: #den exei valei keywords na psaksei
			for root, folders, files in os.walk(inputFolder):
				for file in files:
					os.chdir(root)
					if file.endswith(('doc', 'DOC', 'docx', 'DOCX', 'xls', 'XLS', 'xlsx', 'XLSX', 'PDF', 'pdf', 'pptx', 'PPTX', 'ppt', 'PPT', 'ODT', 'odt', 'ODF', 'odf')):
						count+=1
						if file in os.listdir(destinationFolder): # an yparxei arxeio me idio onoma hdh ston output fakelo, to metonomazw gia na mh diagrafei
							print(f'{file} has the same name with another file in the output folder...')
							print('Renaming...')
							destname = file
							while destname in os.listdir(destinationFolder): # oso to arxeio exei idio onoma me kapoio sto output toso synexizw na tou pros8etw to _RN sto onoma
								destfilename = f'{os.path.splitext(destname)[0]}'+'_RN'
								destname = f'{destfilename}{os.path.splitext(destname)[1]}'
							print(f'Renamed Successfully! New name: {destname}')
							try:
								shutil.copy2(file, f'{destinationFolder}\\{destname}')								
								print(f'Copying initial file: {file} to output folder as: {destname}')
								TotalFilesCopied += 1
								csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{destname},-\n'
							except:
								print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
								csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
						else: # den yparxei arxeio me to idio onoma ston output folder
							try:
								shutil.copy2(file,destinationFolder)
								print(f'Copying {file} to output folder')
								TotalFilesCopied += 1 
								csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{file},-\n'
							except:
								print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
								csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
		else: # exei valei keywords na psaksei
			if ',' in OptionalKeywords: # exei dwsei parapanw apo ena keywords
				KeywordsList = OptionalKeywords.split(',')
				for root, folders, files in os.walk(inputFolder):
					for file in files:
						os.chdir(root)
						if file.endswith(('doc', 'docx', 'xls', 'xlsx', 'pdf', 'pptx', 'odt')):
							count+=1
							try: #prospa8w na kanw read to arxeio
								text = textract.process(file)
								decoded_text = text.decode('utf-8')
								print(f'{file} was read successfully')							
								if any(keyword in decoded_text for keyword in KeywordsList): # checkarw to arxeio me ola ta keywords kai parsarw mono an yparxei to keyword entos tou arxeiou
									print(f'keyword found inside file: {file}!')
									if file in os.listdir(destinationFolder): # an yparxei arxeio me idio onoma hdh ston output fakelo, to metonomazw gia na mh diagrafei
										print(f'{file} has the same name with another file in the output folder...')
										print('Renaming...')
										destname = file
										while destname in os.listdir(destinationFolder): # oso to arxeio exei idio onoma me kapoio sto output toso synexizw na tou pros8etw to _RN sto onoma
											destfilename = f'{os.path.splitext(destname)[0]}'+'_RN'
											destname = f'{destfilename}{os.path.splitext(destname)[1]}'
										print(f'Renamed Successfully! New name: {destname}')
										try:
											shutil.copy2(file, f'{destinationFolder}\\{destname}')								
											print(f'Copying initial file: {file} to output folder as: {destname}')
											TotalFilesCopied += 1
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{destname},KEYWORD FOUND\n'
										except:
											print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
									else: # den yparxei arxeio me to idio onoma ston output folder
										try:
											shutil.copy2(file,destinationFolder)
											print(f'Copying {file} to output folder')
											TotalFilesCopied += 1 
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{file},KEYWORD FOUND\n'
										except:
											print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
								else: #to arxeio egine read alla to keywords den einai entos tou periexomenou tou
									csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,NO KEYWORD FOUND\n'
									continue # synexizw sto epomeno file									
							except Exception as e: # to arxeio den egine read successfully opote apla to antigrafw xwris na kserw an yparxei to keyword entos tou periexomenou tou
								print(e) 
								print(f'{file} could not be read successfully. Cannot check for keyword in file. Proceeding with copy action')
								if file in os.listdir(destinationFolder): # an yparxei arxeio me idio onoma hdh ston output fakelo, to metonomazw gia na mh diagrafei
									print(f'{file} has the same name with another file in the output folder...')
									print('Renaming...')
									destname = file
									while destname in os.listdir(destinationFolder): # oso to arxeio exei idio onoma me kapoio sto output toso synexizw na tou pros8etw to _RN sto onoma
										destfilename = f'{os.path.splitext(destname)[0]}'+'_RN'
										destname = f'{destfilename}{os.path.splitext(destname)[1]}'
									print(f'Renamed Successfully! New name: {destname}')
									try:
										shutil.copy2(file, f'{destinationFolder}\\{destname}')								
										print(f'Copying initial file: {file} to output folder as: {destname}')
										TotalFilesCopied += 1
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{destname},KEYWORD NOT SEARCHED\n'
									except:
										print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
								else: # den yparxei arxeio me to idio onoma ston output folder
									try:
										shutil.copy2(file,destinationFolder)
										print(f'Copying {file} to output folder')
										TotalFilesCopied += 1 
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{file},KEYWORD NOT SEARCHED\n'
									except:
										print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
			else: # exei dwsei akribws ena keyword
				for root, folders, files in os.walk(inputFolder):
					for file in files:
						os.chdir(root)
						if file.endswith(('doc', 'docx', 'xls', 'xlsx', 'pdf', 'pptx', 'odt')):
							count+=1
							try: #prospa8w na kanw read to arxeio
								text = textract.process(file)
								decoded_text = text.decode('utf-8')
								print(f'{file} was read successfully')							
								if OptionalKeywords in decoded_text: # checkarw to arxeio me to ena keyword pou exei dwsei o xrhsths kai parsarw mono an yparxei to keyword entos tou arxeiou
									print(f'keyword found inside file: {file}!')
									if file in os.listdir(destinationFolder): # an yparxei arxeio me idio onoma hdh ston output fakelo, to metonomazw gia na mh diagrafei
										print(f'{file} has the same name with another file in the output folder...')
										print('Renaming...')
										destname = file
										while destname in os.listdir(destinationFolder): # oso to arxeio exei idio onoma me kapoio sto output toso synexizw na tou pros8etw to _RN sto onoma
											destfilename = f'{os.path.splitext(destname)[0]}'+'_RN'
											destname = f'{destfilename}{os.path.splitext(destname)[1]}'
										print(f'Renamed Successfully! New name: {destname}')
										try:
											shutil.copy2(file, f'{destinationFolder}\\{destname}')								
											print(f'Copying initial file: {file} to output folder as: {destname}')
											TotalFilesCopied += 1
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{destname},KEYWORD FOUND\n'
										except:
											print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
									else: # den yparxei arxeio me to idio onoma ston output folder
										try:
											shutil.copy2(file,destinationFolder)
											print(f'Copying {file} to output folder')
											TotalFilesCopied += 1 
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{file},KEYWORD FOUND\n'
										except:
											print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
											csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
								else: #to arxeio egine read alla to keyword den einai entos tou periexomenou tou
									csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,NO KEYWORD FOUND\n'
									continue # synexizw sto epomeno file									
							except Exception as e: # to arxeio den egine read successfully opote apla to antigrafw xwris na kserw an yparxei to keyword entos tou periexomenou tou
								print(e) 
								print(f'{file} could not be read successfully. Cannot check for keyword in file. Proceeding with copy action')
								if file in os.listdir(destinationFolder): # an yparxei arxeio me idio onoma hdh ston output fakelo, to metonomazw gia na mh diagrafei
									print(f'{file} has the same name with another file in the output folder...')
									print('Renaming...')
									destname = file
									while destname in os.listdir(destinationFolder): # oso to arxeio exei idio onoma me kapoio sto output toso synexizw na tou pros8etw to _RN sto onoma
										destfilename = f'{os.path.splitext(destname)[0]}'+'_RN'
										destname = f'{destfilename}{os.path.splitext(destname)[1]}'
									print(f'Renamed Successfully! New name: {destname}')
									try:
										shutil.copy2(file, f'{destinationFolder}\\{destname}')								
										print(f'Copying initial file: {file} to output folder as: {destname}')
										TotalFilesCopied += 1
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{destname},KEYWORD NOT SEARCHED\n'
									except:
										print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
								else: # den yparxei arxeio me to idio onoma ston output folder
									try:
										shutil.copy2(file,destinationFolder)
										print(f'Copying {file} to output folder')
										TotalFilesCopied += 1 
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},{file},KEYWORD NOT SEARCHED\n'
									except:
										print(f'Error copying {file} to output folder. Failed action logged in csv report file!')
										csvContent += f'{count},{os.path.abspath(file)},{datetime.fromtimestamp(os.stat(file).st_ctime)},{datetime.fromtimestamp(os.stat(file).st_mtime)},NOT COPIED,ERROR\n'
	print('----------')
	print(f'{count} total files were found in folder')
	if TotalFilesCopied>0: # an exw vrei arxeia na antigrapsw tote mono ftiaxnw to csv
		if FileCategory == 'media':
			try:
				with open(f'{destinationFolder}\\MediaReport.csv', 'w', encoding='utf-8') as fout:
					fout.write(csvContent)
				print('CSV media report file successfully created')
			except:
				print('CSV media report file could not be written to the output folder!')
		else:
			try:
				with open(f'{destinationFolder}\\DocumentReport.csv', 'w', encoding='utf-8') as fout:
					fout.write(csvContent)
				print('CSV document report file successfully created')
			except:
				print('CSV document report file could not be written to the output folder!')

	return TotalFilesCopied
			
	
#---menu definition
menu_def = [['File', ['Exit']],
			['Help', ['Documentation', 'About']],] 
#---layout definition
CheckboxFrameLayout = [[sg.Checkbox('Documents (doc, docx, xls, xlsx, pdf, pptx, odt, odf)', background_color='#2a363b', enable_events=True, key='-DOCCHK-')],
						[sg.Checkbox('Media (jpg, png, avi, mov, mp4, 3gp)', background_color='#2a363b', enable_events=True, key='-MEDIACHK-')]
						]
FolderFrameLayout = [[sg.Text('Choose the folder to search in for files', background_color='#2a363b')],
					[sg.In(key='-FOLDER-', readonly=True, background_color='#334147'), sg.FolderBrowse()]]

KeywordsFrameLayout = [[sg.Text('Give the keywords to search for in files (separate by coma)', background_color='#2a363b')],					
					[sg.In(key='-KEYWORDS-', disabled=True, disabled_readonly_background_color='#a8a7a7'), sg.Checkbox('Enable', background_color='#2a363b', enable_events=True, key='-KEYWORDCHK-')]]

OutputSaveFrameLayout = [[sg.Text('Choose folder to save the report and files', background_color='#2a363b')],
					[sg.In(key='-OUTPUT-', readonly=True, background_color='#334147'), sg.FolderBrowse()]] #key='-SAVEBTN-', disabled=True, enable_events=True

col_layout = [[sg.Frame('Choose File Category to Harvest', CheckboxFrameLayout, background_color='#2a363b')],
				[sg.Frame('Input Folder', FolderFrameLayout, background_color='#2a363b')],
				[sg.Frame('Keywords (Optional - only for Documents)', KeywordsFrameLayout, background_color='#2a363b', pad=((0,0),(0,65))) ],				
				[sg.Frame('Output Folder', OutputSaveFrameLayout, background_color='#2a363b')],				
				[sg.Button('Exit', size=(7,1)), sg.Button('Harvest', size=(7,1)), sg.Button('Calculate', size=(7,1))]]

#---GUI Definition
layout = [[sg.Menu(menu_def, key='-MENUBAR-')],
			[sg.Column(col_layout, element_justification='c',background_color='#2a363b'), sg.Frame('Verbose Analysis',
			[[sg.Output(size=(70,25), key='-OUT-', background_color='#334147', text_color='#fefbd8')]], background_color='#2a363b')],
			[sg.Text('FileHarvester Ver. 1.0.0', background_color='#2a363b', text_color='#b2c2bf')]]

window = sg.Window('FileHarvester', layout, background_color='#2a363b') 

#---run
while True:
	event, values = window.read()
	# print(event, values)
	if event in (sg.WIN_CLOSED, 'Exit'):
		break
#---menu events
	if event == 'Documentation':
		try:
			webbrowser.open_new('https://github.com/D-Kats/FileHarvester/blob/main/README.md')
		except:
			sg.PopupOK('Visit https://github.com/D-Kats for documentation', title='Documentation', background_color='#2a363b')
	if event == 'About':
		sg.PopupOK('FileHarvester Ver. 1.0.0 \n\n --DKats 2020', title='-About-', background_color='#2a363b')
#---checkbox events
	if event == '-KEYWORDCHK-': # an energopoihsei ta keywords
		if values['-KEYWORDCHK-']:
			window['-KEYWORDS-'].update(disabled=False)			
		else:
			window['-KEYWORDS-'].update(value='', disabled=True)
	
#---buttons events
	if event == "Calculate":
		if values['-FOLDER-'] == '': # den exei epileksei fakelo gia analysh
			sg.PopupOK('Please choose a folder to calculate its files total size!', title='!', background_color='#2a363b')
		elif values['-DOCCHK-'] == False and values['-MEDIACHK-'] == False: #epelekse fakelous gia analysh kai output alla den tikare ti arxeia 8elei na psaksei			
			sg.PopupOK('Please choose a file category to calculate total size in the folder!', title='!', background_color='#2a363b')
		else:
			CalculateInputFolder = values['-FOLDER-']
			if values['-DOCCHK-'] and values['-MEDIACHK-']: # exei epileksei kai ta dyo file category gia na ginoun calculate
				print('Initializing calculation for Document and Media files...')
				print('Calculating Document files total size...')
				CalculateFileCategory = 'documents'
				CalculateDocumentTotalSize = calculate(CalculateInputFolder, CalculateFileCategory)
				print(f'Document files total size is {CalculateDocumentTotalSize} bytes')
				#---neo call gia media
				print('Calculating Media files total size...')
				CalculateFileCategory = 'media'
				CalculateMediaTotalSize = calculate(CalculateInputFolder, CalculateFileCategory)
				print(f'Media files total size is {CalculateMediaTotalSize} bytes')
				total = CalculateDocumentTotalSize + CalculateMediaTotalSize
				print(f'All files have size of {total} bytes')
			elif values['-DOCCHK-']: # exei epileksei mono documents gia calculate
				print('Calculating Document files total size...')
				CalculateFileCategory = 'documents'
				CalculateDocumentTotalSize = calculate(CalculateInputFolder, CalculateFileCategory)
				print(f'Document files total size is {CalculateDocumentTotalSize} bytes')
			else: # exei epileksei mono media gia calculate
				print('Calculating Media files total size...')
				CalculateFileCategory = 'media'
				CalculateMediaTotalSize = calculate(CalculateInputFolder, CalculateFileCategory)
				print(f'Media files total size is {CalculateMediaTotalSize} bytes')

	if event == "Harvest":
		if values['-FOLDER-'] == '': # den exei epileksei fakelo gia analysh
			sg.PopupOK('Please choose a folder to search in for files!', title='!', background_color='#2a363b')
		elif values['-OUTPUT-'] == '': # den exei epileksei output fakelo gia na grapsei ta arxeia
			sg.PopupOK('Please choose an output folder to save the files!', title='!', background_color='#2a363b')
		elif values['-DOCCHK-'] == False and values['-MEDIACHK-'] == False: #epelekse fakelous gia analysh kai output alla den tikare ti arxeia 8elei na psaksei			
			sg.PopupOK('Please choose a file category to search for in the folder!', title='!', background_color='#2a363b')
		else:
			try: # elegxw an mporw ontws na grapsw sto output folder
				inputFolder = values['-FOLDER-']
				outputFolder = values['-OUTPUT-']
				destinationFolderName = str(datetime.now()).replace(':', '_')
				destinationFolder = f'{outputFolder}\\{destinationFolderName}' # dhmiourgw neo fakelo sto output folder me thn hmeroxronologia ths ereynas 
				os.mkdir(destinationFolder) # dhmiourgw fakelo gia na grapsw ta arxeia
				#--- actual parsarisma tou input fakelou---
				if values['-DOCCHK-'] and values['-MEDIACHK-']: # exei epileksei kai ta dyo file category gia na psaksei
					totalMediaFilesCopied = harvester(inputFolder, destinationFolder, 'media')
					if values['-KEYWORDCHK-']: # exei tickarei ta keywords
						totalDocumentFilesCopied = harvester(inputFolder, destinationFolder, 'documents', values['-KEYWORDS-'])
						print('----------')
						print('Harvesting Complete')
						print(f'{totalMediaFilesCopied} total media files were copied to output folder and')
						print(f'{totalDocumentFilesCopied} total document files were copied to output folder')
						if totalMediaFilesCopied == 0 and totalDocumentFilesCopied == 0: # den exw kanei copy kanena arxeio 
							os.chdir(outputFolder)
							os.rmdir(destinationFolderName)
							sg.PopupOK('File Harvesting finished!\nNo files met your criteria in the selected folder', title='!', background_color='#2a363b')
						else:
							finalize(outputFolder, destinationFolderName)
					else: # den exei tickarei ta keywords
						totalDocumentFilesCopied = harvester(inputFolder, destinationFolder, 'documents')
						print('----------')
						print('Harvesting Complete')
						print(f'{totalMediaFilesCopied} total media files were copied to output folder and')
						print(f'{totalDocumentFilesCopied} total document files were copied to output folder')
						if totalMediaFilesCopied == 0 and totalDocumentFilesCopied == 0: # den exw kanei copy kanena arxeio
							os.chdir(outputFolder)
							os.rmdir(destinationFolderName)
							sg.PopupOK('File Harvesting finished!\nNo files met your criteria in the selected folder', title='!', background_color='#2a363b')
						else:
							finalize(outputFolder, destinationFolderName)
				elif values['-DOCCHK-']: # exei epileksei mono documents gia na psaksei
					if values['-KEYWORDCHK-']: # exei tickarei ta keywords
						totalDocumentFilesCopied = harvester(inputFolder, destinationFolder, 'documents', values['-KEYWORDS-'])
						print('----------')
						print('Harvesting Complete')
						print(f'{totalDocumentFilesCopied} total document files were copied to output folder')
						if totalDocumentFilesCopied == 0: # den exw kanei copy kanena arxeio
							os.chdir(outputFolder)
							os.rmdir(destinationFolderName) 
							sg.PopupOK('File Harvesting finished!\nNo files met your criteria in the selected folder', title='!', background_color='#2a363b')
						else:
							finalize(outputFolder, destinationFolderName)
					else:
						totalDocumentFilesCopied = harvester(inputFolder, destinationFolder, 'documents')
						print('----------')
						print('Harvesting Complete')
						print(f'{totalDocumentFilesCopied} total document files were copied to output folder')
						if totalDocumentFilesCopied == 0: # den exw kanei copy kanena arxeio
							os.chdir(outputFolder)
							os.rmdir(destinationFolderName) 
							sg.PopupOK('File Harvesting finished!\nNo files met your criteria in the selected folder', title='!', background_color='#2a363b')
						else:
							finalize(outputFolder, destinationFolderName)
				else: # exei epileksei mono media gia na psaksei
					totalMediaFilesCopied = harvester(inputFolder, destinationFolder, 'media')
					print('----------')
					print('Harvesting Complete')
					print(f'{totalMediaFilesCopied} total media files were copied to output folder')
					if totalMediaFilesCopied == 0: # den exw kanei copy kanena arxeio
						os.chdir(outputFolder)
						os.rmdir(destinationFolderName) 
						sg.PopupOK('File Harvesting finished!\nNo files met your criteria in the selected folder', title='!', background_color='#2a363b')
					else:
						finalize(outputFolder, destinationFolderName)				
			except:
				sg.PopupOK('Cannot write to output folder!\nCheck permissions or choose another output folder', title='!', background_color='#2a363b')
	
window.close()
