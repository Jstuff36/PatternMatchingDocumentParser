#-------------#
#Extract Specs from masterspec expert team and add them an excel sheet
#Programmer: Justin White
#Created on 6/1/2016
#Updated on 6/6/2016 to allow for input to include multiple divisions
#All word documents must be in the .docx extension. Program will ignore if they have the .doc extension
#There are various bulk converters online that will convert all files in a folder to .docx from .doc format
#-------------#

import docx2txt
import os
import re
import xlwt
import shutil

#Must add path to folder that contains the word documents parse with the variable folder_ path 
#Must add path to directory where files will be sorted and new folders will be created by division with the variable directory_path
#Keep r before string containing path name. Ex: r'C:\Users\justin.white\Desktop\Div_01' or r'C:\Users\justin.white\Desktop\multiDownloadArchive'
folder_path = r'C:\Users\justin.white\Desktop\Master_specs_Master_File'
directory_path = r'C:\Users\justin.white\Desktop\Master_specs_With_Divisions'

files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

files_good = [f for f in files if f.endswith('.docx')] #find all files that end with .docx in the folder

#Sort document by division name then add them to a folder of thier respective division. If the folder doesn't exist create it 
for filename in files_good:
    folder = 'Div' + filename[:2]
    new_path = directory_path + '\\' + folder
    if not os.path.exists(new_path):
        os.makedirs(new_path)
    original_path = folder_path + '\\' + filename
    if os.path.isfile(original_path):
        shutil.copy2(original_path, new_path)

path_with_divisions = []
docx_files = []

a = 0
for dirpath, dirnames, filenames in os.walk(directory_path):
    path_with_divisions.append([])
    docx_files.append([])
    for i in filenames:
        path_with_divisions[a].append(os.path.join(dirpath, i))
        docx_files[a] = filenames
    a += 1

path_with_divisions.pop(0)
docx_files.pop(0)
    
text_good = []
useful_text = []
useful_file_names = []
useful_paths = []

for i in range(0, len(path_with_divisions)):
        useful_text.append([])
        useful_file_names.append([])
        useful_paths.append([])
            
for a in range(0, len(path_with_divisions)):
        
    for j in range(0, len(path_with_divisions[a])):
        
        if '~$' in path_with_divisions[a][j]: #file is a zipfile do not open it
            continue
        text = docx2txt.process(path_with_divisions[a][j])
        
        if 'PART ' not in text: #file is not formatted correctly skip it
            continue
        else:
            useful_text[a].append(text) #add files with correctly formatted
            useful_file_names[a].append(docx_files[a][j]) #add names of files with correctly formatted files
            useful_paths[a].append(path_with_divisions[a][j]) #add paths of files with correctly formated files
        j += 1
    a += 1

good_format_text = []
good_format_filenames = []
good_format_paths = []
i = 0
for a in useful_text: #If a list is empty because all files were formatted incorrectly remove
    if a:
        good_format_text.append(a)
        good_format_filenames.append(useful_file_names[i])
        good_format_paths.append(useful_paths[i])
    i += 1
    
wb = xlwt.Workbook()
sheet = wb.add_sheet('Test')

sheet.write(0, 0, 'Document')
sheet.write(0, 1, 'Section')
sheet.write(0, 2, 'Title')
sheet.write(0, 3, 'Page Number')
sheet.write(0, 4, 'Division')      
sheet.write(0, 6, 'Section')
sheet.write(0, 7, 'Title')
sheet.write(0, 8, 'Division')
sheet.write(0, 9, 'Document')
sheet.write(0, 10, 'Page Number')
    
index = 0
counter = 1
counter2 = 1

for div_text in good_format_text: 
    Section = [] #Create list to store data in
    Title = []
    Page = [] 
    
    for a in range(0, len(div_text)):    
        Section.append([])
        Title.append([])
        Page.append([])
        
    for a in range(0, len(div_text)):
        
        text_between_parts = div_text[a].split('PART ')[1:4] #Split text into strings between the ocurences of the word PART
        
        for i in range(0, len(text_between_parts)):
            text_good.append([])
            
        p = re.compile(r'^\s*(\b\d+(?:[.]\d+)?)(?:\W+|^\.)([^0-9].*?)\s*(\b\d+\b)$', re.MULTILINE) #With Regex find all ocurrences of Section, Content, Page number
        
        for i in range(0, len(text_between_parts)):
            text_good[i] = re.findall(p, text_between_parts[i]) 
            
        Section[a], Title[a], Page[a] = zip(*[t for l in text_good for t in l]) 

    files_good_list = [] 
    divisions = []
    
    p = re.compile(r'\\(Div[a-z]*?\d+)\\', re.IGNORECASE) #With Regex find the Division of each file by their filepath
    x = 0
    for i in Section:
        for a in range(0, len(i)):
            files_good_list.append(good_format_filenames[index][x])
            divisions.append(re.findall(p, good_format_paths[index][x])) 
        x += 1
    
    l1 = [item for sublist in Section for item in sublist] #Sepeate to sort by section number
    l2 = [item for sublist in Page for item in sublist]
    l3 = [item for sublist in Title for item in sublist]
    
    def section(s): #Clever way to sort text with decimals greater than .9
        return[int(_) for _ in s.split(".")]
        
    result = list(zip(*sorted(zip(l1, l2, l3, files_good_list, divisions), key = lambda x: section(x[0])))) #Sort all files by section number
    
    #Write data to excel sorted by filename
    j = 0
    
    for x in good_format_filenames[index]:       
        jj = 0
        for i in Section[j]:
            sheet.write(counter + jj, 0, x)
            sheet.write(counter + jj, 1, i)
            sheet.write(counter + jj, 4, divisions[jj][0])            
            jj += 1
        jj = 0
        for i in Title[j]:
            sheet.write(counter + jj, 2, i)
            jj += 1
        jj = 0
        for i in Page[j]:
            sheet.write(counter + jj, 3, i)
            jj += 1
        jj = 0
        counter = counter + len(Title[j])
        j += 1
    
   
    #Write data to excel sorted by section number
    length = len(result[0])
    for i in range(0, length):
        sheet.write(counter2, 6, result[0][i])
        sheet.write(counter2, 7, result[2][i])
        sheet.write(counter2, 8, result[4][i])
        sheet.write(counter2, 9, result[3][i])
        sheet.write(counter2, 10, result[1][i])
        counter2 += 1
    
    
    for i in range(11):
        sheet.col(i).width = 500*20
    
    index += 1
    
wb.save('Final_Excel_Doc.xls') #Change name of excel file by changing the name of the string inside save(). Ex: wb.save('File_name.xls')
#The Excel sheet that is written to must be closed in order to save it. 

    