import xlsxwriter
import char_replace
import time
import random

#créé le fichier excel pour la feuille dont on désire en faire un workbook 
def create_excel (path_excel) :
    workbook = xlsxwriter.Workbook(path_excel)
    workbook_sheet = workbook.add_worksheet()
    return(workbook_sheet,workbook)


#find the place of str in data (in the first row so we know what "[]" to put to get the right data)
def  find_place(str,data) :
    place_valeur = 0
    for i in range (len(data)) :
        if (data[i][0] == str) :
            place_valeur = i 
    return(place_valeur)  

#pour récupérer le nom des colonnes lignes et des données qu'on analyse
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

#récupère les indicateurs quipermettent de colorier les cellules sous forme de dictionnaire : {nom_KPI : [valeur_target_bad,valeur_target_good,validation_up_down] }
def recup_indicateurs(sheet,nom_data) :
    indic_dict = {} 
    key = False
    with open("target.csv", 'r') as file :
        for lines in file:
            line = lines.split(",")
            if (line[0] == sheet) :
                for item in line :
                    if item in nom_data :
                        key = item.strip('\n')
                        indic_dict[key] = []
                    elif(key) : 
                        indic_dict[key] += [item.strip('\n')]
    return(indic_dict)

#récupère la place des indicateurs dans data_avec_lin_col
def place_indicateur(indic_dict,data_avec_lin_col) :
    data_indic = indic_dict
    for i in range (len(data_avec_lin_col[0])) :
        item = data_avec_lin_col[0][i]
        if (item in indic_dict.keys()) :
            data_indic[item] +=[i]
    return(data_indic)

def prepare_indicator(data_indic) : 
    place_indic = {}
    for values in data_indic.values() :
        place_indic[values[-1]] = values[0:-1]
    return(place_indic)
#récupère les données du csv (toutes) et renvoie un [ [type 1] [type 2] [type n]] contentant toutes les datas ainsi qu'un tableau vide de la taille du nb de types différents
def recup_data(path_csv) :
    with open(path_csv, "r") as file : 
        for lines in file :
            lines = lines.replace(",",".")
            lines = lines.strip('\n')
            line = lines.split('|')
            #data stores all the information contained in the csv file
            data = []
            # tab_length_max_data is for finding the max length of a column for formating the excel
            tab_length_max_data = []
            for i in range (len(line)) :
                line[i].strip('\n')
                data += [[]]
                tab_length_max_data += [0]

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
    return(data,tab_length_max_data)

#renvoie la place des colonnes/lignes/données dans data
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

#compare deux listes et renvoie true si elles sont identiques
def compare_list(list_1, list_2):
    s = sum(i == j for i, j in zip(list_1, list_2))
    return (s == len(list_1) == len(list_2) )


#prend un dictionnaire et une liste en entrée. Regarde si la liste est parmis les items du dictionnaire. Renvoie un booléen et la clé de la liste du dictionnaire
def find_in_dic(dic, list) :
    bool = False
    cle = -1
    for key,valeur in dic.items() :
        if (compare_list(list,valeur)) :
            bool = True
            cle = key
    return(bool,cle)
        

#renvoie les datas avec leur place (colonnes, lignes) à la fin  ainsi que des dictionnaires de (col/lin) { n°lin/col : (tuples de col/lign)}
def prepare_data(data,place_colonnes, place_lignes, place_data) :
    length_data = len(data[0])
    data_avec_lin_col = []
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
        data_avec_lin_col += [current_data]
        
        if(find_in_dic(lignes,current_lignes)[0]) :
            data_avec_lin_col[i] += [find_in_dic(lignes,current_lignes)[1]]

            
        else : 
            id_lignes += 1 
            lignes[id_lignes] = current_lignes
            data_avec_lin_col[i] += [id_lignes]

        if(find_in_dic(colonnes,current_colonnes)[0]) :
            data_avec_lin_col[i] += [find_in_dic(colonnes,current_colonnes)[1]]
        else : 
            id_colonnes += 1 
            data_avec_lin_col[i] += [id_colonnes]
            colonnes[id_colonnes] = current_colonnes

    return(data_avec_lin_col,colonnes,lignes)



# écris dans l'excel les valeurs
def write_excel(data_avec_lin_col, colonnes, lignes,workbook_sheet,place_indic,workbook) :
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
    #ecris les noms des colonnes
    for key,value in colonnes.items() :
            for i in range (len(value)) :
                if (key == 1) :
                    workbook_sheet.write(i, (len(colonnes.get(1)) -1 + key), value[i])
                else :
                    for j in range (len(data_avec_lin_col[0])-2) :
                        workbook_sheet.write(i, (len(colonnes.get(1)) -3 + key)*(len(data_avec_lin_col[0])-2)  + j , value[i])

    #ecris les noms des lignes
    for key,value in lignes.items() :
        for i in range (len(value)) :
            workbook_sheet.write(len(lignes.get(1)) -1 + key, i, value[i])

    #ecris les "datas"
    for i in range(1,len(data_avec_lin_col)):
        for j in range (len(data_avec_lin_col[i])-2) :
            if(j in place_indic.keys()) :
                target = place_indic[j]
                if ( target[2] == "up") :
                    if (data_avec_lin_col[i][j] == "no value") :
                        workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_blue)
                    else :
                        if ( float(data_avec_lin_col[i][j].strip("%")) >= float(target[1]) ):
                            workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_green)
                        elif (float(target[0]) >= float(data_avec_lin_col[i][j].strip("%")) ):
                            workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_red)
                        elif (float(target[0]) < float(data_avec_lin_col[i][j].strip("%"))< float(target[1])) :
                            workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_orange)
                        else :
                            workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_black)
                elif ( target[2] == "down") :    
                    if ( float(data_avec_lin_col[i][j].strip("%")) >= float(target[1]) ):
                        workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_red)
                    elif (float(target[0]) >= float(data_avec_lin_col[i][j].strip("%")) ):
                        workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_green)
                    elif (float(target[0]) < float(data_avec_lin_col[i][j].strip("%")) < float(target[1])) :
                        workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_orange)
                    else :
                        workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j],cell_format_black)
            else :
                workbook_sheet.write(data_avec_lin_col[i][-2] + 1, j + (len(data_avec_lin_col[0])-2)*(data_avec_lin_col[i][-1] - 1), data_avec_lin_col[i][j])

#lance toutes les fonctions
def lancer(sheet) :
    
    path_csv = sheet + ".csv"
    path_excel = sheet + ".xlsx"
    char_replace.replace_csv(path_csv)
    #creation of the excel file
    workbook_sheet,workbook = create_excel(path_excel)

    nom_colonnes, nom_lignes , nom_data= recup_col_line_data(sheet)
    indic_dict = recup_indicateurs(sheet,nom_data)
    data , tab_length_max_data = recup_data(path_csv)
    place_colonnes , place_lignes , place_data = place_col_lin_data(nom_colonnes, nom_lignes, nom_data, data)
    data_avec_lin_col , colonnes , lignes = prepare_data(data, place_colonnes, place_lignes, place_data)
    data_indic = place_indicateur(indic_dict,data_avec_lin_col)
    place_indic = prepare_indicator(data_indic)
    write_excel(data_avec_lin_col,colonnes,lignes,workbook_sheet,place_indic,workbook)

    #workbook_sheet.write(0,1,"aa")        
    workbook.close()


lancer("WPS_ACPower")