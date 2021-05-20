# -*- coding: utf-8 -*-
import xlsxwriter
import char_replace


#traitement du fichier config
def config_traitement(sheet) : 
    config = []
    with open("config.csv", 'r') as targets :
        for lines in targets :
            line = lines.split(",")
            if (line[0] == sheet) :
                config = [line[1], line[2],line[3],line[4],line[5],line[6].strip('\n').strip(' ')] 
    print(config)
    return config


def csv_excel_crosstab(sheet) :
    path_csv = sheet + ".csv"
    path_excel = sheet + ".xlsx"
    char_replace.replace_csv(path_csv)
    workbook = xlsxwriter.Workbook(path_excel)
    workbook_sheet = workbook.add_worksheet()

    config = config_traitement(sheet)