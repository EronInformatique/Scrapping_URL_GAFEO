import unidecode
import re
import os
import pandas as pd
import pygsheets
from Packages.get_info_apprenant_session_loop import get_info_apprenant_session_loop
from Packages.update_googlesheet_data import update_workseet_suivi_eron
from Packages.update_layout_worksheet import update_layout_worksheet

gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/Scrapping_Url_GAFEO/Oauth_gg/code_secret_client_95743482524-gj2mnoav9naiqt454ggvt71r28r4n3dk.apps.googleusercontent.com.json')

# Suivi Eron 2022 (derniere version)
#sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')

# DUPLICATA Suivi Eron 2022 (derniere version)
sh = gc.open_by_key('1Ix4xc_kJPrIBXQL8JmGVz_XNnY4t7AMHdHvmyVlExiA')
header_col_num=6

format_sheet=False
#Format all sheet if different
if format_sheet:
    update_layout_worksheet(gc,header_col_num,4,22)


for sheetNumber in range(4,22):
    # print(sheetNumber)
    wksheet = sh[sheetNumber]

    list_header=wksheet.get_row(header_col_num)
    index_col_formation = [index+1 for index,elem in enumerate(list_header) if elem == "Thématiques" ][0]
    index_col_apprenant = [index+1 for index,elem in enumerate(list_header) if elem == "Nom Prénom" ][0]
    col_formation_raw = wksheet.get_col(index_col_formation)

    # print(col_name_apprenant)
    print(wksheet.title)
    print(sh.title)

    try:
        path_directory=os.path.join(os.path.dirname(__file__),"data",wksheet.title.replace("/","-").replace("(","").replace(")","").replace(" ","-"))
        print(path_directory)
        os.makedirs(path_directory)
    except FileExistsError:
        # directory already exists
        pass

    if len(os.listdir(path_directory) ) == 0 or  (len(os.listdir(path_directory) ) == 1 and ".DS_Store" in os.listdir(path_directory)):
        max_value=0
        col_name_apprenant = wksheet.get_col(index_col_apprenant)
        col_apprenant = [string for string in col_name_apprenant if string ]

        col_formation_name=[]
        col_formation_raw = wksheet.get_col(index_col_formation)
        col_formation = [ string.split("-")[0] for string in col_formation_raw if string]
        for string in col_formation_raw:
            if string and len(string.split("-"))==3:
                col_formation_name.append(string.split("-")[2])
            elif string!="":
                col_formation_name.append(string)
            else:
                continue
        # print(col_formation)
        # print(col_formation_name)


        datasheet = pd.DataFrame(list(zip(col_formation[2:],col_formation_name[2:],col_apprenant[2:])),columns=['Formation','Formation_name','Apprenant'])
        datasheet["Course_id"]=""
        datasheet["Gafeo_People_id"]=""
        datasheet['Dernier Accès'] = ''
        datasheet['Pourcentage']=''
        datasheet['Apprenant_Normalize']=''
        datasheet["Apprenant_Gafeo"]=""

        #Normalize name 
        for index,row in datasheet.iterrows():
            # print(unidecode.unidecode(row['Apprenant']))
            # print(row['Apprenant'].lower())
            # print(unidecode.unidecode(row['Apprenant'].lower()))
            # print(re.sub(r'[^a-zA-Z0-9]','',unidecode.unidecode(row['Apprenant'].lower().replace(' ', ''))))
            # row['Apprenant_Normalize']= re.sub(r'[^a-zA-Z0-9]','',unidecode.unidecode(row['Apprenant']).lower().replace(' ', ''))
            row['Apprenant_Normalize']= unidecode.unidecode(row['Apprenant']).lower()
        
        formation_unique= datasheet.Formation.unique().tolist()

    else:
        files_name =[f for f in os.listdir(path_directory) if not f.startswith('.')] 
        list_number = [ int(f.split("_")[-1]) for f in files_name if "DS_Store" not in f]
        max_value = max(list_number)
        max_index = list_number.index(max_value)
        datasheet = pd.read_pickle(os.path.join(path_directory,files_name[max_index]))  
        formation_unique= datasheet.Formation.unique().tolist()

        col_formation_raw = wksheet.get_col(index_col_formation)
        col_formation = [ string.split("-")[0] for string in col_formation_raw if string]
        # unique_formation_from_googlesheet=set(col_formation[2:])

    if len(formation_unique) != max_value:
        dic_dataframe={}
        for formation in formation_unique:
            dic_dataframe[formation]=datasheet[datasheet["Formation"]==formation]

        updated_datasheet = get_info_apprenant_session_loop(dic_dataframe,wksheet,datasheet,path_directory,start_formation=max_value)
        update_workseet_suivi_eron(wksheet,updated_datasheet)
    else:
        update_workseet_suivi_eron(wksheet,datasheet)