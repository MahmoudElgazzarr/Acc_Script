#This script is being used to Rename Implementation data type to Application data Type

#import libraries being used
import os
import openpyxl

#add folder that contain .arxml file 
Acc_arxml =  "D:/CtAp_ACC.arxml"

#add path that contain .xlsx file which have data type
excel_DataTypes =  "D:/table_info_Data_Stage_B_PATH_3.xlsx"

# Create temp file to copy the arxmls, and edit in the new file 
data_index = Acc_arxml.find("\\")
new_Acc_arxml = Acc_arxml[:(data_index)] +"/new_"+Acc_arxml[(data_index+1):]

# workbook object is created 
wb_obj = openpyxl.load_workbook(excel_DataTypes)

#create Sheet object
sheet_obj = wb_obj.get_sheet_by_name('1D_Tables')

#get max numbers of rows
maxmium_row = sheet_obj.max_row

#Line Start
Line_Start = '<TYPE-TREF DEST="IMPLEMENTATION-DATA-TYPE">'
#Line End
Line_End = '</TYPE-TREF>'

for j in range(1,3):
    #for loop for the max number of row in the sheet , serach one by one
    for i in range(1, maxmium_row + 1 ):
        #Copy ACC File Line by Line
        new_file = open(new_Acc_arxml,'w')
        with open(Acc_arxml,'r') as inFile:
            #Get Cell object Data
            cell_obj = sheet_obj.cell(row = i, column = j)
            #search in the whole file , Sequencal search Slow $need to be updated to binary search for example
            for num_line, line_content in enumerate(inFile, 1):
                #get calibration Parameters from excel sheet
                DataType = str(cell_obj.value)
                #search for Implementation line
                if line_content.find(Line_Start) != -1 and line_content.find('Idt_') and line_content.find(DataType) and line_content.find(Line_End):
                    #get How many spaces before the line
                    X_Spaces = line_content.find("<TYPE")
                    #variable to hold the number of spaces
                    Spaces = line_content[:(X_Spaces)]
                    Application_Type = str(sheet_obj.cell(row = i, column = j + 4))
                    line_content = line_content.replace(Spaces + '<TYPE-TREF DEST="APPLICATION-PRIMITIVE-DATA-TYPE">/Package_Autocode/Data_Type/Application_Types/'+ Application_Type + '</TYPE-TREF>')
                # Write the line after edits "if needed" in new file     
                new_file.write(line_content)
            #close and save temp file         
            new_file.close()
print("Renaming Completed Successfully")
