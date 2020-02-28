#This script is being used to Rename Implementation data type to Application data Type

#Testing Stach

#import libraries being used
import os
import openpyxl


#add folder that contain .arxml file 
Input_arxml =  "D:/Ford_Dat2.1/aptiv_sw/autosar_cfg/davinci/Config/Developer/ComponentTypes\\CtAp_ACC.arxml"

#add path that contain .xlsx file which have data type
excel_DataTypes =  "D:/table_info_Data_Stage_B_PATH_3.xlsx"

# Create temp file to copy the arxmls, and edit in the new file 
data_index = Input_arxml.find("\\")
new_Input_arxml = Input_arxml[:(data_index)] +"/new_"+Input_arxml[(data_index+1):]

# workbook object is created 
wb_obj = openpyxl.load_workbook(excel_DataTypes)

#create Sheet object
sheet_obj = wb_obj.get_sheet_by_name('2D_Tables')

#get max numbers of rows
#maxmium_row = sheet_obj.max_row

#Start Row
Start_Row = 2

#Maxmium Row + 1
maxmium_row = 25

#Line Start
Line_Start = '<TYPE-TREF DEST="IMPLEMENTATION-DATA-TYPE">'
#Middle Line
Middle_Line = 'Cal_Datatype'
#Line End
Line_End = '</TYPE-TREF>'

#Check for number of changes
Number_Of_Changes = 0
Non_Null_Cells = 0
Number_Of_Found_Calibration = 0
Number_Of_None = 0
Idt_Found = 0
i = 0
j = 0

#Changed Lines Numbers
Changed_list_Lines = [0] * maxmium_row

flag = 1



#Search through the Excel Sheet
#Two Columns
for Coulumns in range(1,4):
    #Search in the Rows
    for rows in range(Start_Row , maxmium_row):
        #Get value of cells
        Application_Cell = sheet_obj.cell(row = rows, column = Coulumns + 4).value
        #If we find none increment the variable
        if Application_Cell is None :
            #increment the number of None Cell
            Number_Of_None = Number_Of_None + 1

new_file = open(new_Input_arxml,'w')
with open(Input_arxml,'r') as inFile:
    #search in the whole file , Sequencal search Slow $need to be updated to binary search for example
    for num_line, line_content in enumerate(inFile, 1):
        #Copy ACC File Line by Line
        for j in range(1,4):
            #for loop for the max number of row in the sheet , serach one by one
            for i in range(Start_Row , maxmium_row):
                #Get Cell object Data
                cell_obj = sheet_obj.cell(row = i, column = j + 7 )
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
                    #Number of Found Calibration
                    Number_Of_Found_Calibration = Number_Of_Found_Calibration + 1
                #replace
                if  (line_content.find(Imp_DataType) != -1) and (Cal_Name_Found_Flag == 1) :
                    #Found Cal and Idt In lines
                    Idt_Found = Idt_Found + 1
                    #get application Type
                    Application_Type = str(sheet_obj.cell(row = i, column = j + 4).value)
                    #replace line contenet with the new line must be not none
                    if (Application_Type != 'None'):
                        print(i,j)
                        #number of Non NULL Cells
                        Non_Null_Cells = Non_Null_Cells + 1
                        #rename implementation to Application
                        line_content = line_content.replace(Imp_DataType, Application_Type)
                        #Change other headers
                        line_content = line_content.replace('"IMPLEMENTATION-DATA-TYPE">' , '"APPLICATION-PRIMITIVE-DATA-TYPE">')
                        Component_Start_in_line = line_content.find('ComponentType/')
                        Component_End_in_line = line_content.find('Cal_Datatype/')
                        line_content = line_content.replace(line_content[Component_Start_in_line : Component_End_in_line + 13 ],'Data_Type/Application_Types/')
                        #print Final line
                        #print(line_content)
                        Number_Of_Changes = Number_Of_Changes + 1
                if line_content.find('</PARAMETER-DATA-PROTOTYPE>') != -1:
                    Cal_Name_Found_Flag = 0
        # Write the line after edits "if needed" in new file
        new_file.write(line_content)
    #print Number of cells in the excel sheet
    Cells_Count = maxmium_row - Start_Row
    #Two Columns
    Cells_Count = Cells_Count * 2
    #Print Number of Cells
    print('Number of Cells = ', Cells_Count)
    #print Number of found Cal
    print('Number_Of_Found_Calibration',Number_Of_Found_Calibration)
    #Print number of None
    print('Print Number of None : ',Number_Of_None)
    #print number of IDT found
    print('Idt_Found = ', Idt_Found)
    #Print Number on Non Null Cells
    print('Non Null Cells = ' , Non_Null_Cells)
    #print Number of changes in the file
    print('Number Of Changes = ' , Number_Of_Changes)
    #close and save temp file         
    new_file.close()

# remove the orignal file to replace the new one with it   
os.remove(Input_arxml)
# rename the new file to have the same name of the orignal one 
os.rename(new_Input_arxml,Input_arxml)

#Done
print("Renaming Completed Successfully")
