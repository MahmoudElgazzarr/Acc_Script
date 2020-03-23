#This Script is used to Remove Adt_& _T from datatypes Arxml File for calbibration parameters
#It removes from two places <DATA-TYPE-MAP> & <APPLICATION-PRIMITIVE-DATA-TYPE
#input : Excel sheet containing calibration names
#output : Application datatypes is same name as calibration names

import os
import openpyxl


#add folder that contain .arxml file 
arxml_DataTypes =  "D:\DataTypes.arxml"
#add path that contain .xlsx file which have data type
excel_DataTypes =  "D:/table_info_Data_Stage_B_PATH_3.xlsx"
# workbook object is created 
wb_obj = openpyxl.load_workbook(excel_DataTypes)
#create Sheet object
Sheet_object = wb_obj.get_sheet_by_name('1D_Tables')
#get max numbers of rows
maxmium_row = Sheet_object.max_row
# workbook object is created 
wb_obj = openpyxl.load_workbook(excel_DataTypes)
#Column_Max_Number = 3
Column_Max_Number = 3
#Start_Row
Start_Row = 1
#Flag for DATA_TYPE_MAP_Line_Start_found
DATA_TYPE_MAP_Line_Start_found = 0

new_Datatypes_arxml = 'D:/new_DataTypes.arxml'


#The APPLICATION Datatype we are searching for start and end lines
APPLICATION_Line_Start = '<APPLICATION-PRIMITIVE-DATA-TYPE'
APPLICATION_Line_End = '</APPLICATION-PRIMITIVE-DATA-TYPE>'
#Data type map tags
DATA_TYPE_MAP_Line_Start = '<DATA-TYPE-MAP>'
DATA_TYPE_MAP_Line_End = '</DATA-TYPE-MAP>'

def main():
    for Column in range(1 , Column_Max_Number ):
        for row in range(Start_Row , maxmium_row ) :
            #open the new file
            new_file = open(new_Datatypes_arxml,'w')
            #get Cal Parameters from excel sheet
            Cal_Name = Get_Value_Of_Cell(Sheet_object , row , Column )
            with open(arxml_DataTypes,'r') as inFile:
                #Get Line numbers - 1
                for num_line, line_content in enumerate(inFile, 1):
                    #Search for opening of Data-type-map opening tag
                    if line_content.find(DATA_TYPE_MAP_Line_Start) != -1 :
                        #if found set a flage 
                        DATA_TYPE_MAP_Line_Start_found = 1
                    #if cal found
                    if((line_content.find(Cal_Name) != -1) and (DATA_TYPE_MAP_Line_Start_found == 1) and (line_content.find('<APPLICATION-DATA-TYPE-REF') != -1 )):
                        #Remove Adt_ and also _T From the Line
                        line_content.replace('Adt_' , '')
                        line_content.replace('_T' , '')
                        print("Removing")
                    #Remove Flage for closing flag
                    if line_content.find(DATA_TYPE_MAP_Line_End) != -1:
                        DATA_TYPE_MAP_Line_Start_found = 0
                    #Write Line
                    new_file.write(line_content)
            #Close the File after Edits
            new_file.close()
            #Copt New File to the old file every Line Change
            Copy_File_Content(new_Datatypes_arxml , arxml_DataTypes )

#Function that takes sheet object , ROW , Coulumn
#Return value of a cell
def Get_Value_Of_Cell(sheet_obj , row , column):
    Cell_objj = sheet_obj.cell(row = row , column = column)
    Cell_Value = str(Cell_objj.value)
    return Cell_Value

#Function that copies All contenet of a file and copies to an other file
def Copy_File_Content(src , dest):
    with open(src,'r') as F1:
        with open(dest,'w') as F2:
            for Line in F1:
                F2.write(Line)
    F1.close()
    F2.close()

#Start App
if __name__ == '__main__':
    main()