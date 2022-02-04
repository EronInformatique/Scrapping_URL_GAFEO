import unidecode
import re
from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException,TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os
from fuzzywuzzy import process,fuzz



def get_info_apprenant_session_loop(dic_dataframe,datasheet,path_directory,start_formation):
    """
    Robot/ automatisation de récupération d'infos sur le suivi des apprenants dans leur formation
    """

    def initialize_browser_pagelogin(time_out):
        if os.path.exists("/Applications/Internet/Google Chrome.app/Contents/MacOS/Google Chrome"):
                path_google_chrome="/Applications/Internet/Google Chrome.app/Contents/MacOS/Google Chrome"
        else:
            path_google_chrome="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
        options = Options()
        options.binary_location = path_google_chrome
        options.add_experimental_option("detach", True)
        # options.add_argument("--headless")
        # options.add_argument("--disable-gpu")
        # options.add_experimental_option('prefs', {'download.default_directory':'/Users/acapai/Downloads/'})
        # options.add_experimental_option('prefs',{"profile.default_content_setting_values.automatic_downloads": 1})

        rel_path="/Applications/chromedriver"
        PATH = os.path.abspath(rel_path)
        id="andria.c@eronservice.fr"
        psswd="Andria-2021!"
        browser = webdriver.Chrome(options=options,executable_path=PATH)
        browser.get("https://www.gafeo.fr/my/")
        browser.find_element(By.ID,'username').send_keys(id)
        browser.find_element(By.ID,'password').send_keys(psswd)
        browser.find_element(By.CSS_SELECTOR,"button.btn-primary").click()
        return browser
    
    def page_recherche_formation(browser,time_out):
        # div_searchbox = WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.ID,"coursesearchbox")))
        inputs_field = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,"#coursesearch .coursesearchbox input")))
        inputs_field[0].send_keys(formation_table)
        inputs_field[1].click()
    
    def page_list_formation(browser,time_out):

        titre_datasheet = formation_df_copy["Formation_name"].iloc[0]
        element_title_formation_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CLASS_NAME,"coursetitle")))
        list_title_formation = [ elem.text for elem in element_title_formation_list]
        tuple_title_chosen = process.extractOne(titre_datasheet, list_title_formation, scorer=fuzz.token_sort_ratio)
        score_similarity_text=[]
        for i,elem in enumerate(element_title_formation_list):
            titre_html_gafeo = elem.text
            score_similarity_text.append(fuzz.token_set_ratio(titre_datasheet,titre_html_gafeo))
        
        max_score = max(score_similarity_text)
        max_index = score_similarity_text.index(max_score)
        
        for i,elem in enumerate(element_title_formation_list):
            titre_html_gafeo = elem.text
            if titre_html_gafeo.lower().find("old") ==0:
                continue
            # score_similarity_text= fuzz.token_set_ratio(titre_datasheet,titre_html_gafeo)
            # print(score_similarity_text)
            if len(element_title_formation_list)==1:
                formation_to_select =  elem
                break
            elif "v3" in titre_html_gafeo:
                formation_to_select =  elem
                break
            elif "v2" in titre_html_gafeo:
                formation_to_select =  elem
            elif tuple_title_chosen[0] == titre_html_gafeo:
                print(tuple_title_chosen[0])
                print(titre_html_gafeo)
                formation_to_select =  elem
                break
            elif max_score > 90:
                formation_to_select =  element_title_formation_list[max_index]
                break
            else:
                continue

        element_title_formation_list = WebDriverWait(formation_to_select,time_out).until(EC.presence_of_element_located((By.XPATH,"..")))
        parent_element_title_formation_list = WebDriverWait(element_title_formation_list,time_out).until(EC.presence_of_element_located((By.XPATH,"..")))
        image_link = WebDriverWait(parent_element_title_formation_list,time_out).until(EC.presence_of_element_located((By.CLASS_NAME,"hovercoursebox")))
        image_link.send_keys(Keys.RETURN)
        # action_link.click()

    def page_formation_detail(browser,time_out):
        div_button_config_course = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".teacherdash a")))
        div_button_config_course[0].click()
    
    def page_formation_detail_popup(browser,time_out):
        WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.LINK_TEXT,"Rapports de session"))).click()
    
    def page_details_rapport_prefill_list_apprenant(browser,time_out):
        dp_down_list_participant = Select(WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"select[name='userid']"))))

        list_apprenant= formation_df_copy['Apprenant_Normalize'].tolist()
        for i,element in enumerate(dp_down_list_participant.options):
            print(element.get_attribute("value"))
            print(element.get_attribute("text"))
            # element_normalize=re.sub(r'[^a-zA-Z0-9]','',unidecode.unidecode(element.get_attribute("text")).lower().replace(' ', ''))
            element_normalize=unidecode.unidecode(element.get_attribute("text")).lower()
            print("element_normalize:",element_normalize)
            print("list_apprenant:",list_apprenant)
            ratios = process.extract(element_normalize,list_apprenant,limit=len(list_apprenant))
            print(ratios)
            highest = process.extractOne(element_normalize,list_apprenant)
            print("Highest",highest)
            index= [x for x, y in enumerate(list_apprenant) if y == highest[0]]
            if highest[1]>75:
                print("Ratio",highest[1])
                print("index",index[0])
                formation_df_copy.iloc[index[0],formation_df_copy.columns.get_loc("Ratio_Fuzzy")]=highest[1]
                formation_df_copy.iloc[index[0],formation_df_copy.columns.get_loc("Name_Fuzzy")]=highest[0]
                formation_df_copy.iloc[index[0],formation_df_copy.columns.get_loc("option_value")]=element.get_attribute("value")
                formation_df_copy.iloc[index[0],formation_df_copy.columns.get_loc("Apprenant_Gafeo")]=element.get_attribute("text")
        return 

    def get_details_suivi_formations_apprenant(browser,time_out):
        for ind in formation_df_copy.index:
            value = formation_df_copy["option_value"][ind]
            if formation_df_copy["Apprenant_Gafeo"][ind] !="":
                lastname=formation_df_copy["Apprenant_Gafeo"][ind].split()[0]
                firstname=formation_df_copy["Apprenant_Gafeo"][ind].split()[1]
                print(firstname+" "+lastname)
                print("value option",value)
            if value != "":
                dp_down_list_participant = Select(WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"select[name='userid']"))))
                dp_down_list_participant.select_by_value(value)
                opt_name = WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"input[name='go_btn']")))
                opt_name.send_keys(Keys.RETURN)
                WebDriverWait(browser,time_out).until(EC.text_to_be_present_in_element((By.CSS_SELECTOR,".userinfobox h1"),lastname))
                info_pourcentage_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".mandatory-items div")))
                info_pourcentage_text = info_pourcentage_list[2].text
                pourcentage_only=re.findall(r"^.*?\([^\d]*(\d+)[^\d]*\).*$",info_pourcentage_text)
                print("%:",pourcentage_only[0])

                info_last_access = WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"#sample-lastcourseaccess .sample-value")))
                print("info_last_access",info_last_access.text)
                formation_df_copy.loc[ind,'Dernier Accès'] = info_last_access.text
                formation_df_copy.loc[ind,'Pourcentage']=pourcentage_only[0]

    def update_datasheet():
        # print(datasheet.loc[pd_copy.index.values,"Dernier Accès"])
        # print(datasheet.loc[pd_copy.index.values,"Apprenant"])
        datasheet.loc[formation_df_copy.index.values,"Dernier Accès"] = formation_df_copy.loc[formation_df_copy.index.values,"Dernier Accès"]
        datasheet.loc[formation_df_copy.index.values,"Pourcentage"] = formation_df_copy.loc[formation_df_copy.index.values,"Pourcentage"]

    def save_step_process(nb_time_loop,path_directory):
        datasheet.to_pickle(os.path.join(path_directory,"datasheet_loop_"+str(nb_time_loop)))


    formation_key_dic=list(dic_dataframe.keys())
    time_out=30
    nb_time_loop=start_formation
    formation_key_dic_start=formation_key_dic[start_formation:]
    print(formation_key_dic_start)
    error=False
    for formation_table in formation_key_dic_start:
        print(formation_table)
        print(dic_dataframe[formation_table]["Formation_name"])
        formation_df_copy = dic_dataframe[formation_table]# A changer
        formation_df_copy["Ratio_Fuzzy"]=""
        formation_df_copy["Name_Fuzzy"]=""
        formation_df_copy["option_value"]=""
        # print(formation_key_dic)
        error=False
        try:
            browser = initialize_browser_pagelogin(time_out)
            page_recherche_formation(browser,time_out)
            page_list_formation(browser,time_out)
            page_formation_detail(browser,time_out)
            page_formation_detail_popup(browser,time_out)
            page_details_rapport_prefill_list_apprenant(browser,time_out)
            get_details_suivi_formations_apprenant(browser,time_out)
        except NoSuchElementException:
            print("There was an error, no such element")
            error=True
            browser.quit()
        except TimeoutException:
            error=True
            browser.quit()
        else:
            print("There were no errors.")
        finally:
            print("Process completed.")
            if error:
                time.sleep(5)
                browser.quit()
            else:
                update_datasheet()
                nb_time_loop+=1
                save_step_process(nb_time_loop,path_directory)
                browser.quit()
                time.sleep(10)


    
    return datasheet

