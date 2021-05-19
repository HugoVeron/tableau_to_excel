import download as dl
import csv_to_excel as cte

list_workbooks = []
list_sheets = []
with open("C:/Users/Hugo veron/OneDrive - MADIACOM/Bureau/python/tableau_to_excel/list_dl.csv", 'r') as list_dl :
    for lines in list_dl :
        line = lines.split(",")
        list_workbooks += [line[0].strip("\n")]
        list_sheets += [line[1].strip("\n")]
print(list_sheets)


for i in range (len(list_workbooks)) :
    if (i!=0) :
        dl.download(list_workbooks[i],list_sheets[i])
        cte.csv_excel(list_sheets[i])
