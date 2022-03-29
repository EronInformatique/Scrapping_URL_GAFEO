from black import wrap_stream_for_windows
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
import update_googlesheet_data as up_wk


def get_info_apprenant_session_loop(dic_dataframe,wksheet,datasheet,path_directory,start_formation):
    """
    Robot/ automatisation de récupération d'infos sur le suivi des apprenants dans leur formation
    """

    def initialize_browser_pagelogin(time_out,url,params):
        if os.path.exists("/Applications/Internet/Google Chrome.app/Contents/MacOS/Google Chrome"):
                path_google_chrome="/Applications/Internet/Google Chrome.app/Contents/MacOS/Google Chrome"
        else:
            path_google_chrome="C:\\Program Files\\Google\\Chrome\Application"
        options = Options()
        # options.binary_location = path_google_chrome
        options.add_experimental_option("detach", True)
        # options.add_argument("--headless")
        # options.add_argument("--disable-gpu")
        # options.add_experimental_option('prefs', {'download.default_directory':'/Users/acapai/Downloads/'})
        # options.add_experimental_option('prefs',{"profile.default_content_setting_values.automatic_downloads": 1})

        rel_path="C:\\Users\\Linda\\Documents\\chromedriver.exe"
        PATH = os.path.abspath(rel_path)
        id="automate"
        psswd="AutomateEron2!!"
        browser = webdriver.Chrome(options=options,executable_path=PATH)
        # browser.get("https://www.gafeo.fr/my/")
        browser.get(url+params)
        browser.find_element(By.ID,'username').send_keys(id)
        browser.find_element(By.ID,'password').send_keys(psswd)
        browser.find_element(By.CSS_SELECTOR,"button.btn-primary").click()
        return browser
    
    # def page_recherche_formation(browser,time_out):
    #     # div_searchbox = WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.ID,"coursesearchbox")))
    #     inputs_field = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,"#coursesearch .coursesearchbox input")))
    #     inputs_field[0].send_keys(formation_table)
    #     inputs_field[1].click()
    
    # def page_list_formation(browser,time_out):

    #     # element_nb_module_formation_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".local-my-fp-summary.boxspecialwola")))
    #     element_nb_modules_formation_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".local-my-fp-summary div > span:nth-child(1)")))
    #     element_nb_heures_formation_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".local-my-fp-summary div > span:nth-child(2)")))
        
    #     nb_modules_list=[]
    #     nb_heures_list=[]
    #     for elem in element_nb_modules_formation_list:
    #         if elem.text !="":
    #             nb_modules_list.append(int(elem.text.split()[0])) #if s.isdigit()

    #     for elem in element_nb_heures_formation_list:
    #         if elem.text !="":
    #             nb_heures_list.append(int(elem.text.split()[0])) #if s.isdigit()
        
    #     # element_nb_modules_formation_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".local-my-fp-summary > div:first-child > span:nth-child(1)")))
    #     # element_nb_modules_formation_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.XPATH,"//div[@class='local-my-fp-summary boxspecialwola']/div/span")))
    #     titre_datasheet = formation_df_copy["Formation_name"].iloc[0]
    #     element_title_formation_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CLASS_NAME,"coursetitle")))
    #     list_title_formation = [ elem.text for elem in element_title_formation_list]
    #     tuple_title_chosen = process.extractOne(titre_datasheet, list_title_formation, scorer=fuzz.token_sort_ratio)
    #     score_similarity_text=[]
    #     for i,elem in enumerate(element_title_formation_list):
    #         titre_html_gafeo = elem.text
    #         if titre_html_gafeo.lower().find("old") ==0:
    #             titre_html_gafeo="old"
    #         score_similarity_text.append(fuzz.token_set_ratio(titre_datasheet,titre_html_gafeo))
        
    #     max_score = max(score_similarity_text)
    #     max_index = score_similarity_text.index(max_score)

    #     max_nb_modules = max(nb_modules_list)
    #     max_index_nb_modules = nb_modules_list.index(max_nb_modules)
        
    #     max_heures = max(nb_heures_list)
    #     max_index_heures = nb_heures_list.index(max_heures)
        
    #     for i,elem in enumerate(element_title_formation_list):
    #         titre_html_gafeo = elem.text
    #         if len(element_title_formation_list)==1:
    #             formation_to_select =  elem
    #             break
    #         elif titre_html_gafeo.lower().find("old") ==0  or nb_modules_list[i] == 0: #nb_heures_list[i] == 0
    #             continue
    #         elif "v3" in titre_html_gafeo and nb_modules_list[i] == max_nb_modules :
    #             formation_to_select =  elem
    #             break
    #         elif "v2" in titre_html_gafeo and nb_modules_list[i] == max_nb_modules:
    #             formation_to_select =  elem
    #             break
    #         # elif nb_heures_list[i] == max_heures:
    #         #     formation_to_select =  elem
    #         #     break
    #         elif nb_modules_list[i] == max_nb_modules:
    #             formation_to_select =  elem
    #             break
    #         # elif tuple_title_chosen[0] == titre_html_gafeo:
    #         #     print(tuple_title_chosen[0])
    #         #     print(titre_html_gafeo)
    #         #     formation_to_select =  elem
    #         #     break
    #         # elif max_score > 90:
    #         #     formation_to_select =  element_title_formation_list[max_index]
    #         #     break
    #         else:
    #             continue

    #     element_title_formation_list = WebDriverWait(formation_to_select,time_out).until(EC.presence_of_element_located((By.XPATH,"..")))
    #     parent_element_title_formation_list = WebDriverWait(element_title_formation_list,time_out).until(EC.presence_of_element_located((By.XPATH,"..")))
    #     image_link = WebDriverWait(parent_element_title_formation_list,time_out).until(EC.presence_of_element_located((By.CLASS_NAME,"hovercoursebox")))
    #     url=image_link.get_attribute("href")
    #     print('url',url)
    #     course_id = url.split("=")[1]
    #     print('course_id',course_id)
    #     course_id = url.split("=")[1]
    #     formation_df_copy["Course_id"]=course_id
    #     print(formation_df_copy["Course_id"])

    #     image_link.send_keys(Keys.RETURN)
    #     # action_link.click()

    # def page_formation_detail(browser,time_out):
    #     div_button_config_course = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".teacherdash a")))
    #     div_button_config_course[0].click()
    
    # def page_formation_detail_popup(browser,time_out):
    #     WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.LINK_TEXT,"Rapports de session"))).click()
    
    def check_url_still_connected():
        curent_url = browser.current_url
        if (curent_url == "https://www.gafeo.fr/login/index.php"):
            browser.find_element(By.ID,'username').send_keys(id)
            browser.find_element(By.ID,'password').send_keys(psswd)
            browser.find_element(By.CSS_SELECTOR,"button.btn-primary").click()
        else:
            pass


    def page_details_rapport_prefill_list_apprenant(browser,time_out):
        dp_down_list_participant = Select(WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"select[name='userid']"))))

        list_apprenant_dpdwn_gafeo=[]
        list_apprenant_dpdwn_gafeo_option=[]
        list_apprenant_dpdwn_gafeo_normalize=[]
        for i,element in enumerate(dp_down_list_participant.options):
            list_apprenant_dpdwn_gafeo.append(element.get_attribute("text"))
            list_apprenant_dpdwn_gafeo_option.append(element.get_attribute("value"))
            list_apprenant_dpdwn_gafeo_normalize.append(unidecode.unidecode(element.get_attribute("text")).lower())

        list_apprenant= formation_df_copy['Apprenant_Normalize'].tolist()
        for i,element in enumerate(list_apprenant):
            # print(element.get_attribute("value"))
            # print(element.get_attribute("text"))
            # element_normalize=re.sub(r'[^a-zA-Z0-9]','',unidecode.unidecode(element.get_attribute("text")).lower().replace(' ', ''))
            # element_normalize=unidecode.unidecode(element.get_attribute("text")).lower()
            print("element_google_sheet_apprenant:",element)
            print("list_apprenant_gafeo:",list_apprenant_dpdwn_gafeo)
            # w_ratio=[fuzz.WRatio(element,appr) for index,appr in enumerate(list_apprenant_dpdwn_gafeo_normalize)]
            # w_ratio_idx_appr=[(index,appr) for index,appr in enumerate(list_apprenant_dpdwn_gafeo_normalize)]
            # sup_80 =[ i for i in w_ratio if i >=80]

            ratios = process.extract(element,list_apprenant_dpdwn_gafeo_normalize,limit=len(list_apprenant),scorer=fuzz.ratio)
            print(ratios)
            ratios_bis = process.extract(element,list_apprenant_dpdwn_gafeo_normalize,limit=len(list_apprenant),scorer=fuzz.ratio)
            print(ratios_bis)
            highest = process.extractOne(element,list_apprenant_dpdwn_gafeo_normalize,scorer=fuzz.ratio)
            print("Highest",highest)
            index= [x for x, y in enumerate(list_apprenant_dpdwn_gafeo_normalize) if y == highest[0]]
            if highest[1]<75:
                print("WARNING Index fuzzy ratio < 75:"+"\nNom CRM Google Sheet :"+element+"\nNom Apprenant GAFEO :"+str(list_apprenant_dpdwn_gafeo[index[0]])+"\nRatio:"+str(highest[1]))
            print("Ratio",highest[1])
            print("index",index[0])
            formation_df_copy.iloc[i,formation_df_copy.columns.get_loc("Ratio_Fuzzy")]=highest[1]
            formation_df_copy.iloc[i,formation_df_copy.columns.get_loc("Name_Fuzzy")]=highest[0]
            formation_df_copy.iloc[i,formation_df_copy.columns.get_loc("Apprenant_GAFEO_ID")]=list_apprenant_dpdwn_gafeo_option[index[0]]
            formation_df_copy.iloc[i,formation_df_copy.columns.get_loc("Apprenant_Gafeo")]=list_apprenant_dpdwn_gafeo[index[0]]
        return 

    def get_details_suivi_formations_apprenant(browser,time_out):
        for ind in formation_df_copy.index:
            value = formation_df_copy["Apprenant_GAFEO_ID"][ind]
            statut = formation_df_copy["Statut CRM"][ind]
            if formation_df_copy["Apprenant_Gafeo"][ind] !="":
                lastname=formation_df_copy["Apprenant_Gafeo"][ind].split()[0]
                firstname=formation_df_copy["Apprenant_Gafeo"][ind].split()[1]
                print(firstname+" "+lastname)
                print("value option",value)
            if value != "" and statut !="Present" and statut !="Excusé" and statut!="Report" and statut!="Annulé" and statut!="Présent" and statut !="Excuse" and statut !="Annule" :
                dp_down_list_participant = Select(WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"select[name='userid']"))))
                dp_down_list_participant.select_by_value(value)
                opt_name = WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"input[name='go_btn']")))
                opt_name.send_keys(Keys.RETURN)
                check_url_still_connected()
                WebDriverWait(browser,time_out).until(EC.text_to_be_present_in_element((By.CSS_SELECTOR,".userinfobox h1"),lastname))
                info_pourcentage_list = WebDriverWait(browser,time_out).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,".mandatory-items div")))
                info_pourcentage_text = info_pourcentage_list[2].text
                pourcentage_only=re.findall(r"^.*?\([^\d]*(\d+)[^\d]*\).*$",info_pourcentage_text)
                print("%:",pourcentage_only[0])

                info_last_access = WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CSS_SELECTOR,"#sample-lastcourseaccess .sample-value")))
                print("info_last_access",info_last_access.text)
                formation_df_copy.loc[ind,'Dernier Accès'] = info_last_access.text
                formation_df_copy.loc[ind,'Pourcentage']=pourcentage_only[0]+" %"


    # def get_percentage_formations_apprenant(browser,time_out):
    #     for ind in formation_df_copy.index:
    #         id_user = formation_df_copy["Apprenant_GAFEO_ID"][ind]
    #         id_course = formation_df_copy["Cours_ID_Static"].iloc[0]
    #         course_ref = formation_df_copy["Formation"].iloc[0]
    #         if id_user !='':
    #             browser.get('https://www.gafeo.fr/report/learningtimecheck/index.php?itemid='+id_user+'&view=user&id='+id_course)
    #             #to refresh the browser
    #             # browser.refresh()
    #             WebDriverWait(browser, time_out).until(EC.presence_of_all_elements_located((By.TAG_NAME,'td')))
    #             # identifying the from row2 having <td> tag
    #             # rwdata = browser.find_elements_by_xpath("//table/tbody/tr[2]/td")
    #             try:
    #                 table_id =WebDriverWait(browser,time_out).until(EC.presence_of_element_located((By.CLASS_NAME, 'generaltable')))
    #                 #time.sleep(10)
    #                 rows = table_id.find_elements(By.TAG_NAME, "tr")
    #                 # rows = WebDriverWait(table_id,time_out).until(EC.presence_of_element_located((By.TAG_NAME, "tr")))
    #                 print("Rows length")
    #                 print(len(rows))
    #                 for row in rows[1:-1]:
    #                     # Get the columns
    #                     print("cols length")
    #                     print(len(row.find_elements(By.TAG_NAME, "td")))
    #                     # cols = row.find_elements(By.TAG_NAME, "td")
    #                     if row.find_elements(By.TAG_NAME, "td")[1].text == course_ref:
    #                         pourcentage_only=re.sub("[^0-9]", "",row.find_elements(By.TAG_NAME, "td")[3].text)
    #                         formation_df_copy.loc[ind,'Pourcentage']=pourcentage_only+" %"
    #                         break
    #                     else:
    #                         continue
    #                     # col_name = row.find_elements(By.TAG_NAME, "td")[1] # This is the Name column                print "col_name.text = "
    #             except NoSuchElementException as e:
    #                 return False


    def update_datasheet():
        # print(datasheet.loc[pd_copy.index.values,"Dernier Accès"])
        # print(datasheet.loc[pd_copy.index.values,"Apprenant"])
        # datasheet.loc[formation_df_copy.index.values,"Course_id"] = formation_df_copy.loc[formation_df_copy.index.values,"Course_id"]
        datasheet.loc[formation_df_copy.index.values,"Apprenant_GAFEO_ID"] = formation_df_copy.loc[formation_df_copy.index.values,"Apprenant_GAFEO_ID"]
        datasheet.loc[formation_df_copy.index.values,"Dernier Accès"] = formation_df_copy.loc[formation_df_copy.index.values,"Dernier Accès"]
        datasheet.loc[formation_df_copy.index.values,"Pourcentage"] = formation_df_copy.loc[formation_df_copy.index.values,"Pourcentage"]


    def save_step_process(nb_time_loop,path_directory):
        datasheet.to_pickle(os.path.join(path_directory,"datasheet_loop_"+str(nb_time_loop)))


    id="automate"
    psswd="AutomateEron2!!"

    formation_key_dic=list(dic_dataframe.keys())
    time_out=30
    nb_time_loop=start_formation
    formation_key_dic_start=formation_key_dic[start_formation:]
    print(formation_key_dic_start)
    error=False
    for formation_table in formation_key_dic_start:
        # print(formation_table)
        # print(dic_dataframe[formation_table]["Formation_name"])
        formation_df_copy = dic_dataframe[formation_table].copy()# A changer
        formation_df_copy["Ratio_Fuzzy"]=""
        formation_df_copy["Name_Fuzzy"]=""
        # print(formation_key_dic)
        error=False
        url = "https://www.gafeo.fr/report/trainingsessions/index.php?id="
        params = formation_df_copy["Cours_ID_Static"].iloc[0]
        try:
            browser = initialize_browser_pagelogin(time_out,url,params)
            # page_recherche_formation(browser,time_out)
            # page_list_formation(browser,time_out)
            # page_formation_detail(browser,time_out)
            # page_formation_detail_popup(browser,time_out)
            page_details_rapport_prefill_list_apprenant(browser,time_out)
            get_details_suivi_formations_apprenant(browser,time_out)
            # get_percentage_formations_apprenant(browser,time_out)
        except NoSuchElementException:
            print("There was an error, no such element")
            error=True
        except TimeoutException:
            print("Time Out - GAFEO is out")
            error=True
        except Exception as ex:
                print("Other Exception:",str(ex))
                error=True
        else:
            print("There were no errors.")
        finally:
            print("Process completed.")
            if error:
                time.sleep(5)
                if 'browser' in globals():
                    browser.quit()
            else:
                update_datasheet()
                up_wk.update_workseet_suivi_eron(wksheet,datasheet)
                nb_time_loop+=1
                # save_step_process(nb_time_loop,path_directory)
                browser.quit()
                time.sleep(5)
        


    
    return datasheet


