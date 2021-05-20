import download as dl
import csv_to_excel as cte
import csv_to_excel_crosstab as ctec

list_workbooks = []
list_sheets = []

#get the list of sheets we want to convert
with open("list_dl.csv", 'r') as list_dl :
    for lines in list_dl :
        line = lines.split(",")
        list_workbooks += [line[0].strip("\n")]
        list_sheets += [line[1].strip("\n")]
print(list_sheets)

#download and convert al the sheets needed
for i in range (len(list_workbooks)) :
    if (i!=0) :
        dl.download(list_workbooks[i],list_sheets[i])
        cte.csv_excel(list_sheets[i])
        ctec.csv_excel_crosstab(list_sheets[i])
