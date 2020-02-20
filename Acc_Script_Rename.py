#This script is being used to Rename Implementation data type to Application data Type

#import libraries being used
import os
import openpyxl


#add folder that contain .arxml file 
Input_arxml =  "D:/Workspaces/RTE_Workspace/aptiv_sw/autosar_cfg/davinci/Config/Developer/ComponentTypes\\CtAp_TSR.arxml"

#add path that contain .xlsx file which have data type
excel_DataTypes =  "D:\Acc_Path1/table_info_Data_Stage_L_PATH_1.xlsx"

# Create temp file to copy the arxmls, and edit in the new file 
data_index = Input_arxml.find("\\")
new_Input_arxml = Input_arxml[:(data_index)] +"/new_"+Input_arxml[(data_index+1):]

# workbook object is created 
wb_obj = openpyxl.load_workbook(excel_DataTypes)

#create Sheet object
sheet_obj = wb_obj.get_sheet_by_name('1D_Tables')

#get max numbers of rows
#maxmium_row = sheet_obj.max_row

#Start Row
Start_Row = 345

#Todo This is the max number for Acc SWC , need to be changed according to each SWC or create an excel sheet for each swc
maxmium_row = 382

#Line Start
Line_Start = '<TYPE-TREF DEST="IMPLEMENTATION-DATA-TYPE">'
#Middle Line
Middle_Line = 'Cal_Datatype'
#Line End
Line_End = '</TYPE-TREF>'

#Changed Lines Numbers
Changed_list_Lines = [0] * maxmium_row

flag = 1

new_file = open(new_Input_arxml,'w')
with open(Input_arxml,'r') as inFile:
    #search in the whole file , Sequencal search Slow $need to be updated to binary search for example
    for num_line, line_content in enumerate(inFile, 1):
        #Copy ACC File Line by Line
        for j in range(1,2):
            #for loop for the max number of row in the sheet , serach one by one
            for i in range(Start_Row , maxmium_row):
                #Get Cell object Data
                cell_obj = sheet_obj.cell(row = i, column = j + 6 )
                #get Implementation Parameters from excel sheet
                Imp_DataType = str(cell_obj.value)
                #get Cal Parameters from excel sheet
                Cal_obj = sheet_obj.cell(row = i, column = j )
                #get Cal Name
                Cal_Name = str(Cal_obj.value)
                #search for Implementation cal line
                if line_content.find('<SHORT-NAME>'+Cal_Name+'</SHORT-NAME>') != -1:
                    Cal_Name_Found_Flag  = 1
                    #print (line_content)
                #replace
                if line_content.find(Line_Start) != -1 and line_content.find(Middle_Line) and line_content.find('Idt_') !=-1 and line_content.find(Imp_DataType) != -1 and line_content.find(Line_End) != -1 and Cal_Name_Found_Flag == 1 :
                    #Set Flag back to zero
                    Cal_Name_Found_Flag = 0
                    #get application Type
                    Application_Type = str(sheet_obj.cell(row = i, column = j + 4).value)
                    #replace line contenet with the new line must be not none
                    if (Application_Type != 'None'):
                        #rename implementation to Application
                        line_content = line_content.replace(Imp_DataType, Application_Type)
                        #Change other headers
                        line_content = line_content.replace('"IMPLEMENTATION-DATA-TYPE">' , '"APPLICATION-PRIMITIVE-DATA-TYPE">')
                        line_content = line_content.replace('ComponentType/CtAp_ACC/Cal_Datatype/','Data_Type/Application_Types/')
                        #print Final line
                        print(line_content)
                if line_content.find('</PARAMETER-DATA-PROTOTYPE>') != -1:
                    Cal_Name_Found_Flag = 0
        # Write the line after edits "if needed" in new file
        new_file.write(line_content)
    #close and save temp file         
    new_file.close()
print("Renaming Completed Successfully")
