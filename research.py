import xlsxwriter
import char_replace
import time
import random

def create_excel (path_excel) :
    workbook = xlsxwriter.Workbook(path_excel)
    workbook_sheet = workbook.add_worksheet()
    return(workbook_sheet,workbook)


#find the place of str in data (in the first row so we know what "[]" to put to ge t the right data)
def  find_place(str,data) :
    place_valeur = 0
    for i in range (len(data)) :
        if (data[i][0] == str) :
            place_valeur = i 
    return(place_valeur)  


def recup_col_line_data(sheet) :
    nom_colonnes = []
    with open("colonnes.csv", 'r') as file :
            for lines in file:
                line = lines.split(",")
                if (line[0] == sheet) :
                    nom_colonnes = line[1].strip('\n').strip(' ').split("|")
    nom_lignes = []
    with open("lignes.csv", 'r') as file :
            for lines in file :
                line = lines.split(",")
                if (line[0] == sheet) :
                    nom_lignes = line[1].strip('\n').strip(' ').split("|")
    nom_data = []
    with open("donnees.csv", 'r') as file :
        for lines in file :
            line = lines.split(",")
            if (line[0] == sheet) :
                nom_data = line[1].strip('\n').strip(' ').split("|")
    return(nom_colonnes,nom_lignes,nom_data)


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
    return(data,length_max_data)

def place_col_lin_data (nom_colonnes, nom_lignes, nom_data, data) :
    place_colonnes = []
    place_lignes = []
    place_data = []
    for i in range (len(nom_colonnes)) :
        place_colonnes += [find_place(nom_colonnes[i], data)]
    for i in range (len(nom_lignes)) :
        place_lignes += [find_place(nom_lignes[i], data)]
    for i in range (len(nom_data)) :
        place_data += [find_place(nom_data[i], data)]
    return(place_colonnes, place_lignes, place_data)

def compare_list(list_1, list_2):
    s = sum(i == j for i, j in zip(list_1, list_2))
    return (s == len(list_1) == len(list_2) )


def find_in_dic(dic, list) :
    bool = False
    cle = -1
    for key,valeur in dic.items() :
        if (compare_list(list,valeur)) :
            bool = True
            cle = key
    return(bool,cle)
        


def write_excel(data, length_max_data, nom_colonnes, nom_lignes, nom_data,place_colonnes, place_lignes, place_data) :
    length_data = len(data[0])
    data_avec_col_lin = []
    id_colonnes = 0
    id_lignes = 0
    colonnes = {}
    lignes = {}
    #on parcourt toute les données en profondeur
    for i in range (length_data) :
        current_colonnes = []
        current_lignes = []
        current_data = []
        for j in place_colonnes :
            current_colonnes += [data[j][i]]
        for j in place_lignes :
            current_lignes += [data[j][i]]
        for j in place_data :
            current_data += [data[j][i]]
        data_avec_col_lin += [current_data]
        
        if(find_in_dic(lignes,current_lignes)[0]) :
            data_avec_col_lin[i] += [find_in_dic(lignes,current_lignes)[1]]

            
        else : 
            id_lignes += 1 
            lignes[id_lignes] = current_lignes
            data_avec_col_lin[i] += [id_lignes]

        if(find_in_dic(colonnes,current_colonnes)[0]) :
            data_avec_col_lin[i] += [find_in_dic(colonnes,current_colonnes)[1]]
        else : 
            id_colonnes += 1 
            data_avec_col_lin[i] += [id_colonnes]
            colonnes[id_colonnes] = current_colonnes
        
    return(data_avec_col_lin,colonnes,lignes)


def excel_write(data_avec_col_lin, colonnes, lignes,workbook_sheet) :

    for key,value in colonnes.items() :
            for i in range (len(value)) :
                if (key == 1) :
                    workbook_sheet.write(i, (len(colonnes.get(1)) -1 + key), value[i])
                else :
                    for j in range (len(data_avec_col_lin[0])-2) :
                        workbook_sheet.write(i, (len(colonnes.get(1)) -3 + key)*(len(data_avec_col_lin[0])-2)  + j , value[i])

    for key,value in lignes.items() :
        for i in range (len(value)) :
            workbook_sheet.write(len(lignes.get(1)) -1 + key, i, value[i])

    for i in range(1,len(data_avec_col_lin)):
        for j in range (len(data_avec_col_lin[i])-2) :
            workbook_sheet.write(data_avec_col_lin[i][-2] + 1, j + (len(data_avec_col_lin[0])-2)*(data_avec_col_lin[i][-1] - 1), data_avec_col_lin[i][j])
            print(data_avec_col_lin[i][j])
def lancer(sheet) :
    
    path_csv = sheet + ".csv"
    path_excel = sheet + ".xlsx"
    char_replace.replace_csv(path_csv)
    #creation of the excel file
    workbook_sheet,workbook = create_excel(path_excel)

    nom_colonnes, nom_lignes , nom_data= recup_col_line_data(sheet)
    data , length_max_data = recup_data(path_csv)
    place_colonnes , place_lignes , place_data = place_col_lin_data(nom_colonnes, nom_lignes, nom_data, data)
    data_avec_col_lin , colonnes , lignes = write_excel(data,length_max_data,nom_colonnes,nom_lignes,nom_data, place_colonnes, place_lignes, place_data)
    excel_write(data_avec_col_lin,colonnes,lignes,workbook_sheet)

    #workbook_sheet.write(0,1,"aa")        
    workbook.close()


lancer("WPS_ACPower")