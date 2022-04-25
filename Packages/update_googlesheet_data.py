import pandas as pd
from datetime import datetime

def update_workseet_suivi_eron(wksheet,datasheet):
    """Update workseet Suivi ERON 2022"""
    # current date and time
    now = datetime.now()
    time_stamp_str = now.strftime("%d-%b-%Y (%H:%M:%S)")
    df_to_paste = datasheet[['Apprenant_GAFEO_ID','Dernier Accès', 'Pourcentage']].copy()
    wksheet.set_dataframe(df_to_paste,(7,9),copy_head=False)
    wksheet.update_value('Q1', "Mise à jour des infos de progression/suivi apprenant effectué le:"+time_stamp_str)
    # wksheet.update_value('N1', time_stamp_str)