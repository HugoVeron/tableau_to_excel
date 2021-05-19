# -*- coding: iso-8859-1 -*-
import re
def replace_csv(path) :
# open your csv and read as a text string
    with open (path,"r") as file :
        my_csv_text = file.read()

    

    # substitute
    find_str = 'é'
    replace_str = 'e'
    my_csv_text = re.sub(find_str, replace_str, my_csv_text)
    

    # substitute
    find_str = 'è'
    replace_str = 'e'
    my_csv_text = re.sub(find_str, replace_str, my_csv_text)
    

    # substitute
    find_str = 'ê'
    replace_str = 'e'
    my_csv_text = re.sub(find_str, replace_str, my_csv_text)
    
    
    # substitute
    find_str = 'à'
    replace_str = 'a'
    my_csv_text = re.sub(find_str, replace_str, my_csv_text)

    # substitute
    find_str = 'â'
    replace_str = 'a'
    my_csv_text = re.sub(find_str, replace_str, my_csv_text)

    # substitute
    find_str = 'ù'
    replace_str = 'u'
    my_csv_text = re.sub(find_str, replace_str, my_csv_text)

    # substitute
    find_str = 'ç'
    replace_str = 'c'
    my_csv_text = re.sub(find_str, replace_str, my_csv_text)

    # open new file and save
    with open(path, 'w') as file:
        file.write(my_csv_text)

def keep_int(string) :
    new_string = ""
    for char in string :
        try :
            int(char)
            new_string += char
        except : 
            pass
    return(new_string)
