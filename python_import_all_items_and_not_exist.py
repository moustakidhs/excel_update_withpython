
from xlrd import open_workbook
from dateutil.parser import parse
import pandas as pd
import xlsxwriter
import xlrd


lista_1 = []
lista_2 = []

am2 = []


count = -1

workbook = xlsxwriter.Workbook("master.xlsx")

worksheet1 = workbook.add_worksheet()
wb_1 = xlrd.open_workbook('test_1.xlsx')
sh_1 = wb_1.sheet_by_index(0)

wb_2 = xlrd.open_workbook('test_2.xlsx')
sh_2 = wb_2.sheet_by_index(0)



for sheet_1 in wb_1.sheets():
    number_of_rows_1 = sheet_1.nrows
    
    for i in range(0,number_of_rows_1):
        am_1 = sh_1.cell_value(rowx=i, colx=0)
        x_1  = sh_1.cell_value(rowx=i, colx=1)
        y_1  = sh_1.cell_value(rowx=i, colx=2)
        lista_1.append([am_1,x_1,y_1])

        
for sheet_2 in wb_2.sheets():
    number_of_rows_2 = sheet_2.nrows
    
    for row in range(0, number_of_rows_2):
        am_2 = sh_2.cell_value(rowx=row, colx=0)
        x_2  = sh_2.cell_value(rowx=row, colx=1)
        y_2  = sh_2.cell_value(rowx=row, colx=2)
        lista_2.append([am_2,x_2,y_2])
        am2.append(am_2)

def fun1(c,lst1,lst2,a2):
    
    try:
        for data_1 in lst1:

            am1 = data_1[0]
            x1  = data_1[1]     
            y1  = data_1[2]
            if (am1 not  in a2)==True:
                
                c += 1
                worksheet1.write(c, 0,am1)
                worksheet1.write(c, 1,x1)            
                worksheet1.write(c, 2,y1)
            else:
                for data_2 in lst2:
                        am2 = data_2[0]
                        x2  = data_2[1]      
                        y2  = data_2[2]
                        if am1==am2:
                            c += 1
                            worksheet1.write(c, 0,am2)
                            worksheet1.write(c, 1,x2)            
                            worksheet1.write(c, 2,y2)
                        else:
                            pass
    except:
        pass

fun1(count,lista_1,lista_2,am2)
workbook.close()
print("Done")
























