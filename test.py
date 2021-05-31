import random
import xlsxwriter
import math
outWorkbook = xlsxwriter.Workbook("out.xlsx")
outsheet = outWorkbook.add_worksheet()


names = ["kyle", "bob", "mary"]
values = [70,80,90]


outsheet.write("A1", "name")
outsheet.write("B1", "score")

cell_format1 = outWorkbook.add_format()   
cell_format1.set_fg_color('red')
cell_format2 = outWorkbook.add_format()   
cell_format2.set_fg_color('green')


outsheet.write("A2", names[0],cell_format1)
outsheet.write("A3", names[1], cell_format2)
outsheet.write("A4", names[2])
for i in range (3) :
    if values[i]>80 :
        outsheet.write(i+1,1, values[i],cell_format1)
    else :
        outsheet.write(i+1,1, values[i],cell_format2)

outWorkbook.close()
moy2 = 0
for i in range (50) :
    moy = 0
    for i in range (50) : 
        a = 250 +random.random()*50
        moy += a
    print(moy/50)
    moy2 += moy/50
print(moy2/50)