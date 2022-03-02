import time
import os
import sys
sys.path.append('C:\\Users\\Linda\\scrapping_gafeo\\Packages')
import re
import pandas as pd
import pygsheets
import unidecode
from Packages.get_info_apprenant_session_loop import get_info_apprenant_session_loop
from Packages.update_googlesheet_data import update_workseet_suivi_eron
from Packages.update_layout_worksheet import update_layout_worksheet

#Format all sheet if different
# if format_sheet:
#     update_layout_worksheet(gc,sh,header_col_num,4,22)

def pygsheet_suivi_eron(sh,header_col_num,nbsheet):
    """start connection to suivi eron 2022"""
    start_time_worksheet = time.perf_counter()

    # for sheet_number in range(4,22):
    # DUPLICATA Suivi Eron 2022 (derniere version)

    if sh[4].title != "(Départ 01/01/2022) INF":
        return

    for sheet_number in range(4,34):
        # print(sheet_number)
        wksheet = sh[sheet_number]
        
        list_header=wksheet.get_row(header_col_num)
        index_col_formation = [index+1 for index,elem in enumerate(list_header) if elem == "Thématiques" ][0]
        index_col_apprenant = [index+1 for index,elem in enumerate(list_header) if elem == "Nom Prénom" ][0]
        index_col_course_id = [index+1 for index,elem in enumerate(list_header) if elem == "Cours_ID_Static" ][0]
        index_col_statut_crm = [index+1 for index,elem in enumerate(list_header) if elem == "Statut CRM" ][0]
        # index_col_stag_id = [index+1 for index,elem in enumerate(list_header) if elem == "Apprenant_GAFEO_ID" ][0]
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

            col_name_course_id = wksheet.get_col(index_col_course_id)
            col_course_id = [string for string in col_name_course_id if string ]
            
            col_name_statut_crm = wksheet.get_col(index_col_statut_crm)
            col_statut_crm = [string for string in col_name_statut_crm if string ]

            # col_name_stag_id  = wksheet.get_col(index_col_stag_id)
            # col_stag_id = [string for string in col_name_stag_id if string ]

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

            l = [""] * len(col_formation[2:])
            datasheet = pd.DataFrame(list(zip(col_course_id[1:],l,col_formation[2:],col_formation_name[2:],col_apprenant[2:],col_statut_crm[2:])),columns=['Cours_ID_Static','Apprenant_GAFEO_ID','Formation','Formation_name','Apprenant','Statut CRM'])
            # datasheet["Course_id"]=""
            # datasheet["Gafeo_People_id"]=""
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
                row['Apprenant_Normalize']= re.sub(r'[^a-zA-Z0-9]','',unidecode.unidecode(row['Apprenant']).lower())
            
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

        nbsheet+=nbsheet+1


    end_time_wksheet=time.perf_counter()
    duree_total_update=end_time_wksheet-start_time_worksheet
    print("Duree total pour update {number_sheet} Sheet:".format(number_sheet=str(nbsheet))+str(duree_total_update)+'sec')



if __name__ == "__main__":
    gc = pygsheets.authorize(client_secret='C:\\Users\\Linda\\scrapping_gafeo\\sheets.googleapis.com-python.json')
    # sh = gc.open_by_key('1Ix4xc_kJPrIBXQL8JmGVz_XNnY4t7AMHdHvmyVlExiA')
    # Suivi Eron 2022 (derniere version)
    gc.set_batch_mode(False)
    sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')
    header_col_num=6
    nbsheet=0
    format_sheet=False
    pygsheet_suivi_eron(sh,header_col_num,nbsheet)