# -*- coding: utf-8 -*-
import xlsxwriter
import char_replace

# create the workbook in which we write 
def create_excel (path_excel) :
    workbook = xlsxwriter.Workbook(path_excel)
    workbook_sheet = workbook.add_worksheet()
    return(workbook_sheet,workbook_sheet)
#traitement du fichier config
def config_traitement(sheet) : 
    config = []
    with open("config.csv", 'r') as targets :
        for lines in targets :
            line = lines.split(",")
            if (line[0] == sheet) :
                config = [line[1], line[2],line[3],line[4],line[5],line[6].strip('\n').strip(' ')] 
    return config

#create data and max_length_data
def recup_data(path_csv) :
    with open(path_csv, "r") as file : 
        for lines in file :
            lines = lines.replace(",",".")
            lines = lines.strip('\n')
            line = lines.split('|')
            #data stores all the information contained in the csv file
            data = []
            # length_max_data is for finding the max length of a column for formating the excel
            length_max_data = []
            for i in range (len(line)) :
                line[i].strip('\n')
                data += [[]]
                length_max_data += [0]

    #read the information and put it in a data [[],[],[],...,[]]
    with open(path_csv, "r") as file : 
        for lines in file :
            lines = lines.replace(",",".")
            lines = lines.strip('\n')
            line = lines.split('|')
            for i in range (len(data)) :
                if (line[i].strip("\n").strip("\"")  != ''):
                    data[i] += [line[i]]
                else :
                    data[len(data)-1] += ["no value"]
            # else : 
            #      data[len(data)-1] += [char_replace.keep_int(line[len(data)-1])]
    return(data,length_max_data)

#find the place of str in data (in the first row so we know what "[]" to put to ge t the right data)
def  find_place(str,data) :
    place_valeur = 0
    for i in range (len(data)) :
        if (data[i][0] == str) :
            place_valeur = i 
    return(place_valeur)

#récupérer les colonnes afin de les afficher dans excel
def recup_column(config,data) :
    columns = config[3].split('|')
    place_column = []
    for i in range ( len(columns)) :
        place_column += [find_place(columns[i],data)]
    column_uni = []
    for i in range (len(columns)) :
        column_uni += [[]]
        for j in range (1,len(data[place_column[i]])) :
            if data[place_column[i]][j] != 'All' :
                if not (data[place_column[i]][j] in column_uni[i]) :
                    column_uni[i] += [data[place_column[i]][j]]
    return(column_uni)


#create the title of the columns of the excel sheet
def write_column_excel(workbook_sheet,column_uni) :
    prec = 1
    nb_elem_columns = []
    for i in range(len(column_uni)) :
        nb_elem_columns += [len(column_uni[i])]
    count_column = 0
    for i in range(len(nb_elem_columns)):
        for j in range(len(column_uni[i])) :
            workbook_sheet.write(0,1,"aa" )
            print(column_uni[i])
    print(workbook_sheet)

#the function to call to create the excel from the csv
def csv_excel_crosstab(sheet) :
    path_csv = sheet + ".csv"
    path_excel = sheet + ".xlsx"
    char_replace.replace_csv(path_csv)

    #creation of the excel file
    workbook_sheet,workbook = create_excel(path_excel)

    workbook = xlsxwriter.Workbook(path_excel)
    workbook_sheet = workbook.add_worksheet()
    config = config_traitement(sheet)
    data = recup_data(path_csv)[0]
    length_max_data = recup_data(path_csv)[1]
    column_uni = recup_column(config,data)
    write_column_excel(workbook_sheet,column_uni)