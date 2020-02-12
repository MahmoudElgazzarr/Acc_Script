import os
import openpyxl


#add folder that contain .arxml file 
arxml_DataTypes =  "D:\DataTypes.arxml"

#add path that contain .xlsx file which have data type
excel_DataTypes =  "D:/table_info_Data_Stage_B_PATH_3.xlsx"

# Create temp file to copy the arxmls, and edit in the new file 
data_index = arxml_DataTypes.find("\\")
new_arxml_DataTypes = arxml_DataTypes[:(data_index)] +"/new_"+arxml_DataTypes[(data_index+1):]

# workbook object is created 
wb_obj = openpyxl.load_workbook(excel_DataTypes)

#create Sheet object
sheet_obj = wb_obj.get_sheet_by_name('1D_Tables')

#get max numbers of rows
maxmium_row = sheet_obj.max_row

#The Implementation Datatype we are searching for start and end lines
Line_Start = '<IMPLEMENTATION-DATA-TYPE-REF'
Line_End = '</IMPLEMENTATION-DATA-TYPE-REF>'

#for loop for the max number of row in the sheet , serach one by one
for i in range(1, maxmium_row + 1 ):
    with open(arxml_DataTypes,'r') as inFile:
        #Get Cell object Data
        cell_obj = sheet_obj.cell(row = i, column = 1)
        
        #search in the whole file , Sequencal search Slow $need to be updated to binary search for example
        for num_line, line_content in enumerate(inFile, 1):
            
            #get calibration Parameters from excel sheet
            DataType = 'Idt_' + str(cell_obj.value) + '_struc_T'
            
            # check for required data mapping
            if line_content.find(Line_Start) != -1 and line_content.find(DataType) != -1 and line_content.find(Line_End) != -1:
                #found the Implementation data type But flag = 1
                Found_Datatype_Flag = 1
                #Get Application Data Type TODO
                break

            #couldn't found the calibartion implementation data set    
            else:
                Found_Datatype_Flag = 0
                
        if(Found_Datatype_Flag != 1):
            print("Couldn't find : "+str(cell_obj.value))
    

print("Hi It is Working")
