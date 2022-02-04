import pandas as pd

# gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/Scrapping_Url_GAFEO/Oauth_gg/code_secret_client_95743482524-gj2mnoav9naiqt454ggvt71r28r4n3dk.apps.googleusercontent.com.json')

# # Suivi Eron 2022 (derniere version)
# #sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')

# # DUPLICATA Suivi Eron 2022 (derniere version)
# sh = gc.open_by_key('1Ix4xc_kJPrIBXQL8JmGVz_XNnY4t7AMHdHvmyVlExiA')

# wksheet_depart_inf_01_01 = sh[4]

def update_workseet_suivi_eron(wksheet,datasheet):
    """Update workseet Suivi ERON 2022"""
# path=Path(os.path.dirname(__file__))

# path_directory=os.path.join(path.parent.absolute(),"data",wksheet_depart_inf_01_01.title.replace("/","-").replace("(","").replace(")","").replace(" ","-"))

# if len(os.listdir(path_directory) ) == 0:
#     print("no files")
# else:
#     files_name = os.listdir(path_directory)
#     list_number = [ int(f.split("_")[-1]) for f in files_name]
#     max_value = max(list_number)
#     max_index = list_number.index(max_value)
#     datasheet = pd.read_pickle(os.path.join(path_directory,files_name[max_index])) 
    df_to_paste = datasheet[['Dernier Accès', 'Pourcentage']].copy()
    wksheet.set_dataframe(df_to_paste,(7,5),copy_head=False)