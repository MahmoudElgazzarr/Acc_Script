#This script is being used to Rename Implementation data type to Application data Type

#import libraries being used
import os
import openpyxl

#add folder that contain .arxml file 
Acc_arxml =  "D:\\CtAp_ACC.arxml"

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
#maxmium_row = sheet_obj.max_row
maxmium_row = 80

#Line Start
Line_Start = '<TYPE-TREF DEST="IMPLEMENTATION-DATA-TYPE">'
#Middle Line
Middle_Line = 'Cal_Datatype'
#Line End
Line_End = '</TYPE-TREF>'

#Changed Lines Numbers
Changed_list_Lines = [0] * maxmium_row

flag = 1

for j in range(1,2):
    #for loop for the max number of row in the sheet , serach one by one
    for i in range(1, maxmium_row):
        #Copy ACC File Line by Line
        new_file = open(new_Acc_arxml,'w')
        with open(Acc_arxml,'r') as inFile:
            #Get Cell object Data
            cell_obj = sheet_obj.cell(row = i, column = j)
            #search in the whole file , Sequencal search Slow $need to be updated to binary search for example
            for num_line, line_content in enumerate(inFile, 1):
                #get calibration Parameters from excel sheet
                DataType = str(cell_obj.value)
                #search for Implementation cal line
                if line_content.find(Line_Start) != -1 and line_content.find(Middle_Line) and line_content.find('Idt_') !=-1 and line_content.find(DataType) != -1 and line_content.find(Line_End) != -1:
                    #get How many spaces before the line
                    X_Spaces = line_content.find("<TYPE")
                    #variable to hold the number of spaces
                    Spaces = line_content[:(X_Spaces)]
                    #get application Type
                    Application_Type = str(sheet_obj.cell(row = i, column = j + 4).value)
                    #replace line contenet with the new line must be not none
                    if (Application_Type != 'None'):
                        #get place of cal implementation type
                        X = line_content.find('Idt_')
                        Y = line_content.find('_T')
                        #rename implementation to Application
                        line_content = line_content.replace(line_content[X:(Y+2)], Application_Type)
                        #Change other headers
                        line_content = line_content.replace('"IMPLEMENTATION-DATA-TYPE">' , '"APPLICATION-PRIMITIVE-DATA-TYPE">')
                        line_content = line_content.replace('ComponentType/CtAp_ACC/Cal_Datatype/','Data_Type/Application_Types/')
                        #print Final line
                        print(line_content)
                        #Save Changed Line Number
                        Changed_list_Lines.insert(i,num_line)
                        # Write the line after edits "if needed" in new file
                        new_file.write(line_content)
                        new_file.close()
                #else:
                    # Write the line after edits "if needed" in new file
                    #new_file.write(line_content)
        #close and save temp file         
        new_file.close()
print("Renaming Completed Successfully")