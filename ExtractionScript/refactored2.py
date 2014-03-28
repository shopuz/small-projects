#!/usr/bin/python
# create html file from csv
# the filenames of csv and html are given during command execution
# python summary.py file-name.csv html-file.html [-p]
# -p is to denote that the summary should be generated by person, otherwise it is generated by questions

import sys
import csv
import shutil
import os
import stat
from os.path import basename, splitext
import time
import Tkinter, tkFileDialog
from Tkinter import *
from tkFileDialog import askopenfilename
from tkCommonDialog import Dialog
from time import localtime, strftime

global csv_file
global search_type
search_typ =""
global copied_flag
from sys import exit



class Chooser(Dialog):

    command = "tk_chooseDirectory"

    def _fixresult(self, widget, result):
        if result:
            # keep directory until next time
            self.options["initialdir"] = result
        self.directory = result # compatibility
        return result

#
# convenience stuff

def askdirectory(**options):
    "Ask for a directory name"

    return apply(Chooser, (), options).show()








# Open the CSV file for reading
#reader = csv.reader(open(sys.argv[1]))
#reader = csv.reader(open(csv_file))

# Create the HTML file for output
#htmlfile = open(sys.argv[2],"w")

file1 = open('../error-file.txt', 'w')

log_file = ''


included_columns = [59] # column numbers of questions which should be included (excluding timestamp and email)
questions = {} 	# to store all the questions on the top row
name_col = 0 # to store the column number which contains name
flag= 0
# Extracts all the column numbers to be included in html.. Also extracts the column number of name (name_col)

#image_path = 'H:\CATALOGUE 2014\Product Images high res'
#colour_chart_path = 'H:\CATALOGUE 2014\Colour Charts high res'
#logo_path = 'H:\CATALOGUE 2014\Logos High res'

global main_path
global source_path

errorlist = []

def get_logos():
	global flag
	global main_path
	global log_file
	search_type = ''
	
	file_list = []

	if (flag == 0):
		search_type = 'ColourChart'
		included_columns = [61]
		print ("looking for colourcharts...")

	elif (flag ==1 ):
		search_type = 'LOGO'
		included_columns = [64]
		print ("looking for logos...")
	elif (flag ==2):
		search_type = 'IMAGE'
		included_columns = [59,60]
		print ("looking for images ... ")
	
	file_name ='../' 'logfile-' + search_type + '-' + strftime("%Y%m%d-%H%M%S", localtime()) + '.txt'
	log_file = open(file_name, 'w')
	
	
	print ("Source Directory : " + source_path)

	if (flag==0):
		#save_path = main_path = '.\ColourChart'
		path = colour_chart_path
		#save_path += '\\' + supplier_code + ' - ColourChart'

	elif (flag==1):
		#save_path = main_path = '.\Logos'
		path = logo_path
		#save_path += '\\' + supplier_code + ' - LOGO'
	elif (flag == 2):
		#save_path = main_path = '.\Images'
		path = image_path
		#save_path += '\\' + supplier_code + ' - Images'


	fileList_without_extension = []
	arr = []
	subdir = []
	#dir_arr = []
	# Get the list of all the file names in the path to be searched and remove extension from it
	for dirName, subdirList, fileList in os.walk(source_path):
		#dir_arr.append({'name'})
		subdir += subdirList
		for fname in fileList:
			#print ("image_name: " + image_name)
			fname_without_extension = splitext(fname)[0]
			fileList_without_extension.append(fname_without_extension)
			arr.append({'type':'file', 'file_name' : fname_without_extension , 'path' : os.path.abspath(os.path.join(dirName, fname))})

	#print subdir
	#exit()

	#print (fileList_without_extension)

	#exit()

	for included_column in included_columns:
		#print 'included_column : ' + str(included_column)
		
		with open(csv_file) as csvFile:
			reader = csv.reader(csvFile)

			#print ('after reading')
			
			i =1 
			for row in reader:
				#print i
				supplier_code = row[15]

			
				if (flag==0):
					save_path = main_path = '..\ColourChart'
					path = colour_chart_path
					save_path += '\\' + supplier_code + ' - ColourChart'

				elif (flag==1):
					save_path = main_path = '..\Logos'
					path = logo_path
					save_path += '\\' + supplier_code + ' - LOGO'
				elif (flag == 2):
					save_path = main_path = '..\Images'
					path = image_path
					save_path += '\\' + supplier_code + ' - Images'
				
				#save_path += '\\' + supplier_code + ' - ' + search_type
				# Skip the top row which contains the column headings
				if (i==1):
					i = i +1
					continue
				
				# Get the names of the Persons
				#print row
				
				image_name = row[included_column]

				if image_name == "" :
					continue
				#print (row[x])
				#print ("supplier: " + row[15])
				print "------------------------------"
				
				

				
				if not os.path.exists(save_path):
						os.makedirs(save_path)


				images_list = image_name.split(',')

				for item in images_list:
					try:
						print "------------------------------"
						copied_flag = 0
						item = splitext(item)[0]
						

						if (item in file_list or item =='' or item=='COLOUR CHART WITH THE DESIGNERS' ):
							continue
						elif ( item not in fileList_without_extension):
								file_list.append(item)
								print "Couldnot find file: " + item
								log_file.write('File NOT FOUND: ' + item + '\n')

								continue
						else:
							file_list.append(item)
							print ('searching for: ' + item)



						for myFile in arr:
							if item == myFile['file_name']:
								print 'copying file: ' + item
								shutil.copy2(myFile['path'], save_path)
								log_file.write('File copied: ' + item + '\n')
								copied_flag = 1


						if  item  in subdir :
							for dirName, subdirList, fileList in os.walk(source_path):
								if item in subdirList:
									
									save_dir =save_path+ '\\' + item

									print ('found directory: ' + os.path.join(dirName, item) )

									os.chmod(dirName+'\\'+item, stat.S_IWRITE)

									print ('copying directory: ' + os.path.join(dirName, item))

									shutil.copytree(os.path.join(dirName,item), save_dir)

									log_file.write('Directory Copied: ' + save_dir + '\n')
									copied_flag = 1
									
									continue
								
								#time.sleep(1)
							'''
								for fname in fileList:
									#print ("image_name: " + image_name)
									fname_without_extension = splitext(fname)[0]
									
									#print ("found file: " + fname)
									
									#print ('jpt')
									#print (fname_without_extension)
									#print('image after: ' + image)
									
									if fname_without_extension == item:
										#print('inside')
										#print (os.path.join(dirName,fname))
										os.chmod(os.path.join(dirName, fname), stat.S_IWRITE)
										print('copying file : ' + item )
										#print ('file list : ')
										#print (file_list)
										shutil.copy2(os.path.join(dirName, fname), save_path)
										copied_flag = 1

										log_file.write('File copied: ' + item + '\n')
				 			'''
			 			
					except Exception:
						global errorlist
						errorlist.append(item)
						log_file.write('Error copying : ' + item + '\n')
						print ('error copying: ' + item)





def removeEmptyFolders(path):

  if not os.path.isdir(path):
    
    return

  # remove empty subfolders
  files = os.listdir(path)
  if len(files):
    for f in files:
      fullpath = os.path.join(path, f)
      if os.path.isdir(fullpath):
        removeEmptyFolders(fullpath)

  # if folder empty, delete it
  files = os.listdir(path)
  if len(files) == 0:
  	print ''
  	print ("Removing empty folder:", path)
  	os.rmdir(path)

def exiting():
	sys.stdout.write("Exiting ")

	for i in range(1,10):
		sys.stdout.write('.')
		time.sleep(1)
	print ' '
	exit(0)


root = Tk()
#search_type = tkSimpleDialog.askstring('Extract')

csv_file =  tkFileDialog.askopenfilename(title="Open CSV file")
root.withdraw()

root = Tk()

# use width x height + x_offset + y_offset (no spaces!)
root.geometry("%dx%d+%d+%d" % (330, 80, 200, 150))
root.title("tk.Optionmenu as combobox")
var = StringVar(root)
# initial value
var.set('Select Extraction Type')
choices = ['Image', 'Colour Chart', 'Logos']
option = OptionMenu(root, var, *choices)
option.pack(side='left', padx=10, pady=10)
button = Button(root, text="OK", command=root.quit)
button.pack(side='left', padx=20, pady=10)
root.mainloop()


search_type = var.get()
root.withdraw()


root = Tk()
if search_type == "Colour Chart":
	flag = 0
	source_path = colour_chart_path = askdirectory(title="Choose the Directory where Colour Charts are present")
	
elif (search_type == "Logos"):
	flag = 1
	source_path = logo_path  = askdirectory(title="Choose the Directory where LOGOS are present")
elif (search_type ==  "Image"):
	flag = 2
	source_path = image_path  = askdirectory(title="Choose the Directory where Images are present")
else:
	print ("Please select a valid Extraction Type")
	exiting()
	#exit(1)

root.withdraw()





get_logos()

removeEmptyFolders(main_path)

str1 = ''.join(errorlist)
file1.write(str1)
print ''
print '##################################################'
print 'Please refer to logfile for a list of errors'
print 'Extraction process successfully completed....'

print ' '
exiting()

#print (main_path)


