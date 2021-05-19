# -*- coding: utf-8 -*-
import xlsxwriter

import char_replace



def csv_excel(sheet):

    path_csv = sheet + ".csv"
    path_excel = sheet + ".xlsx"
    char_replace.replace_csv(path_csv)
    workbook = xlsxwriter.Workbook(path_excel)
    workbook_sheet = workbook.add_worksheet()

    #stores information from target
    target= []
    #store the place (in data) of the information we want to put in color
    place_valeur = 0

    #extracts the information from target 
    with open("tableau_to_excel/target.csv", 'r') as targets :
        for lines in targets :
            line = lines.split(",")
            if (line[0] == sheet) :
                target = [line[1], line[2].strip('\n').strip(' '),line[3].strip('\n').strip(' '),line[4].strip('\n').strip(' ')] 

    #create data and max_length_data
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

    # find the position of the desirable argument
    for i in range (len(data)) :
        if (data[i][0] == target[3]) :
            place_valeur = i 

    #define different color format
    cell_format_red = workbook.add_format()   
    cell_format_red.set_fg_color('red')
    cell_format_green = workbook.add_format()   
    cell_format_green.set_fg_color('green')
    cell_format_blue = workbook.add_format()   
    cell_format_blue.set_fg_color('blue')
    cell_format_orange = workbook.add_format()   
    cell_format_orange.set_fg_color('orange')
    cell_format_black = workbook.add_format()   
    cell_format_black.set_fg_color('black')

    #write the data in the excel
    for i in range (len(data[len(data)-1])) :
        for j in range (len(data)) :
            if (data[j][i] == "no value") :
                workbook_sheet.write(i,j, data[j][i],cell_format_blue)
            else :
                workbook_sheet.write(i,j, data[j][i])
        if (i!=0) :
            if ( target[2] == "up") :
                if (data[j][i] == "no value") :
                    workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_blue)
                else :
                    if ( float(data[place_valeur][i].strip('%')) > float(target[1])  ):
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_green)
                    elif ( float(data[place_valeur][i].strip('%')) < float(target[0])  ):
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_red)
                    elif ( float(target[1]) >= float(data[place_valeur][i].strip('%')) >= float(target[0])  ):
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_orange)
                    else :
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_black)
            elif ( target[2] == "down") :
                if (data[j][i] == "no value") :
                    workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_blue)
                else :
                    if ( float(data[place_valeur][i].strip('%')) < float(target[1])  ):
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_green)
                    elif ( float(data[place_valeur][i].strip('%')) > float(target[0])  ):
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_red)
                    elif ( float(target[1]) <= float(data[place_valeur][i].strip('%')) <= float(target[0])  ):
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_orange)
                    else :
                        workbook_sheet.write(i,place_valeur, data[place_valeur][i],cell_format_black)
    
    #calculate the max length of each columns
    for i in range (len(data[len(data)-1])) :
        for j in range (len(data)) :
            if len(data[j][i]) > length_max_data[j] :
                length_max_data[j] = len(data[j][i])

    #set the length of the columns            
    for j in range (len(data)):
        workbook_sheet.set_column(j,j,length_max_data[j])

    workbook.close()



