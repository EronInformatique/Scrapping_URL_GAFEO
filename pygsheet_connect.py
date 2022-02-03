import asyncio
from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

import pygsheets

gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/Scrapping_Url_GAFEO/Oauth_gg/code_secret_client_95743482524-gj2mnoav9naiqt454ggvt71r28r4n3dk.apps.googleusercontent.com.json')

# Suivi Eron 2022 (derniere version)
#sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')

# DUPLICATA Suivi Eron 2022 (derniere version)
#sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')
sh = gc.open_by_key('1Ix4xc_kJPrIBXQL8JmGVz_XNnY4t7AMHdHvmyVlExiA')

wksheet_depart_inf_01_01 = sh[4]
col_name_apprenant = wksheet_depart_inf_01_01.get_col(4)

print(col_name_apprenant)
print(wksheet_depart_inf_01_01.title)
print(sh.title)

cells = wksheet_depart_inf_01_01.get_all_values(include_empty_rows=False, include_tailing_empty=False, returnas='cells')
