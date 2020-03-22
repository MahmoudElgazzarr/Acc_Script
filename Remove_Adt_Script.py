#This Script is used to Remove Adt_& _T from datatypes Arxml File for calbibration parameters
#It removes from two places <DATA-TYPE-MAP> & <APPLICATION-PRIMITIVE-DATA-TYPE
#input : Excel sheet containing calibration names
#output : Application datatypes is same name as calibration names

import os
import openpyxl


#add folder that contain .arxml file 
arxml_DataTypes =  "C:\DataTypes.arxml"

#add path that contain .xlsx file which have data type
excel_DataTypes =  "C:/table_info_Data_Stage_B_PATH_3.xlsx"

# workbook object is created 
wb_obj = openpyxl.load_workbook(excel_DataTypes)

def main():
    

















if __name__ == '__main__':
    main()