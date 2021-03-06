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
# Open the CSV file for reading
reader = csv.reader(open(sys.argv[1]))

# Create the HTML file for output
#htmlfile = open(sys.argv[2],"w")

file1 = open('error-file.txt', 'w')

log_file = open('log.txt', 'w')


included_columns = [59] # column numbers of questions which should be included (excluding timestamp and email)
questions = {} 	# to store all the questions on the top row
name_col = 0 # to store the column number which contains name
flag= 0
# Extracts all the column numbers to be included in html.. Also extracts the column number of name (name_col)

catalogue_image_path = 'H:\CATALOGUE 2013\Product Images high res'
colour_chart_path = 'H:\CATALOGUE 2013\Colour Charts high res'
logo_path = 'H:\CATALOGUE 2013\Logos High res'

global main_path
def get_images():
	global main_path
	for x in included_columns:
		
		reader = csv.reader(open(sys.argv[1]))

		
		i =1 
		for row in reader:
			save_path = main_path = '.\Images'
			# Skip the top row which contains the column headings
			if (i==1):
				i = i +1
				continue
			
			# Get the names of the Persons
			#print row
			
			
			image_name = row[x]
			#print (row[x])
			#print ("supplier: " + row[15])
			supplier_code = row[15]
			save_path += '\\' + supplier_code + ' - IMAGES'

			if not os.path.exists(save_path):
					os.makedirs(save_path)


			images_list = image_name.split(',')

			for image in images_list:

				for dirName, subdirList, fileList in os.walk(catalogue_image_path):
					for fname in fileList:
						#print ("image_name: " + image_name)
						
						fname_without_extension = splitext(fname)[0]
						#print ("found file: " + fname)

						if fname_without_extension == image:
	 						print('found : ' + image_name )
	 						shutil.copy2(os.path.join(dirName, fname), save_path)



				
errorlist = []

def get_logos():
	global flag
	global main_path
	
	file_list = []
	if (flag == 0):
		
		included_columns = [61]
		print ("looking for colourcharts...")

	elif (flag ==1 ):
		included_columns = [64]
		print ("looking for logos...")
		
	
	for x in included_columns:
		
		reader = csv.reader(open(sys.argv[1]))

		
		i =1 
		for row in reader:
			
			supplier_code = row[15]
			if (flag==0):
				save_path = main_path = '.\ColourChart'
				path = colour_chart_path
				save_path += '\\' + supplier_code + ' - ColourChart'

			else:
				save_path = main_path = '.\Logos'
				path = logo_path
				save_path += '\\' + supplier_code + ' - LOGO'

			# Skip the top row which contains the column headings
			if (i==1):
				i = i +1
				continue
			
			# Get the names of the Persons
			#print row
			
			image_name = row[x]
			#print (row[x])
			#print ("supplier: " + row[15])
			
			

			
			if not os.path.exists(save_path):
					os.makedirs(save_path)


			images_list = image_name.split(',')

			for image in images_list:
				try:
					image = splitext(image)[0]
					if (image in file_list):
						#print ('first')
						#print ('image before: ' + image)
						#time.sleep(2)
						continue
					else:
						#print (fname_without_extension)
						#print ('second')
						#print(file_list)
						file_list.append(image)



					for dirName, subdirList, fileList in os.walk(path):
						if image in subdirList:
							
							save_path1 =save_path+ '\\' + image

							print ('found directory: ' + os.path.join(dirName, image) )

							os.chmod(dirName+'\\'+image, stat.S_IWRITE)
							shutil.copytree(os.path.join(dirName,image), save_path1)

							log_file.write('Directory Copied: ' + image + '\n')
							continue
						
						#time.sleep(1)

						for fname in fileList:
							#print ("image_name: " + image_name)
							if (flag==1):
								fname_without_extension = splitext(fname)[0]
								
							else:
								fname_without_extension = fname

							#print ("found file: " + fname)
							
							#print ('jpt')
							#print (fname_without_extension)
							#print('image after: ' + image)
							
							if fname_without_extension == image:
								#print('inside')
								#print (os.path.join(dirName,fname))
								os.chmod(os.path.join(dirName, fname), stat.S_IWRITE)
								print('copying file : ' + image )
								#print ('file list : ')
								#print (file_list)
								shutil.copy2(os.path.join(dirName, fname), save_path)
								log_file.write('File copied: ' + image + '\n')
		 				

				except Exception:
					global errorlist
					errorlist.append(image)
					log_file.write('Error copying : ' + image + '\n')
					print ('error copying: ' + image)
	
	





def get_colourchart():
	
	global main_path
	
	file_list = []
	included_columns = [61]
	print ("looking for colourcharts...")

		
	for x in included_columns:
		
		reader = csv.reader(open(sys.argv[1]))

		
		i =1 
		for row in reader:
			
			supplier_code = row[15]
			save_path = main_path = '.\ColourChart'
			path = colour_chart_path
			save_path += '\\' + supplier_code + ' - ColourChart'

			# Skip the top row which contains the column headings
			if (i==1):
				i = i +1
				continue
			
			# Get the names of the Persons
			#print row
			
			image_name = row[x]
			#print (row[x])
			#print ("supplier: " + row[15])
			
			

			
			if not os.path.exists(save_path):
					os.makedirs(save_path)


			images_list = image_name.split(',')

			for image in images_list:
				try:
					image = splitext(image)[0]
					if (image in file_list):
						#print ('first')
						#print ('image before: ' + image)
						#time.sleep(2)
						continue
					else:
						#print (fname_without_extension)
						#print ('second')
						#print(file_list)
						file_list.append(image)


					if image == '' or image == 'COLOUR CHART WITH THE DESIGNERS':
						continue
					else:
						print ("Searching for: " + image)


					for dirName, subdirList, fileList in os.walk(path):
						if image in subdirList:
							
							save_path1 =save_path+ '\\' + image

							print ('found directory: ' + os.path.join(dirName, image) )

							os.chmod(dirName+'\\'+image, stat.S_IWRITE)
							shutil.copytree(os.path.join(dirName,image), save_path1)
							continue
						
						#time.sleep(1)

						for fname in fileList:
							#print ("image_name: " + image_name)


							#print ("found file: " + fname)
							fname_without_extension = splitext(fname)[0]
							#print ('jpt')
							#print (fname_without_extension)
							#print('image after: ' + image)
							#print ('fname: ' + fname)
							#print ('image: ' + image)
							if fname_without_extension == image:
								#print('inside')
								#print (os.path.join(dirName,fname))
								os.chmod(os.path.join(dirName, fname), stat.S_IWRITE)
								print('copying file : ' + image )
								#print ('file list : ')
								#print (file_list)
								shutil.copy2(os.path.join(dirName, fname), save_path)
		 				

				except Exception:
					global errorlist
					errorlist.append(image)
					print ('error copying: ' + image)
	
	






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
    print ("Removing empty folder:", path)
    os.rmdir(path)



			
if (len(sys.argv) == 3):
	if (sys.argv[2].lower() == '-c'):
		flag = 0
		get_colourchart()
	elif (sys.argv[2].lower() == '-l'):
		flag = 1
		get_logos()

	
else:
	get_images()

print (main_path)
removeEmptyFolders(main_path)

str1 = ''.join(errorlist)
file1.write(str1)



exit(0)

		


    	

