import pandas as pd
from datetime import datetime

def update_workseet_suivi_eron(wksheet,datasheet):
    """Update workseet Suivi ERON 2022"""
    # current date and time
    now = datetime.now()
    time_stamp_str = now.strftime("%d-%b-%Y (%H:%M:%S)")
    df_to_paste = datasheet[['Course_id','Gafeo_People_id','Dernier Accès', 'Pourcentage']].copy()
    wksheet.set_dataframe(df_to_paste,(7,6),copy_head=False)
    wksheet.update_value('M1', "Mise à jour des infos de progression/suivi apprenant effectué le:")
    wksheet.update_value('N1', time_stamp_str)