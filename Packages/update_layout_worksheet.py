import pygsheets


def update_layout_worksheet(gc,sh,HEADER_COL,start_sheet,end_sheet):
    """gc : googlesheet a modifier; HEADER_COL: numero de colonnes avec HEADERNAME, start_sheet: index col à modifier, end_sheet:fin colonne à modifier"""

    #Format all sheet if different
    for sheet_number in range(start_sheet,end_sheet):
        # print(sheetNumber)
        wksheet = sh[sheet_number]

        # A METTRE après création des 3 colonnes une fois créée Setting Format
        # first create a model cell with required properties
        gc.set_batch_mode(True)
        model_cell = pygsheets.Cell('D7')
        model_cell.color = (207/255, 226/255, 243/255, 1) #
        # model_cell.format = (pygsheets.FormatType.TEXT)
        model_cell.text_format['underline']=True
        model_cell.text_format['fontSize']=12
        # Setting format to multiple cells in one go
        # rng.apply_format(model_cell)  # will make all cell in this range rose color and percent format
        # Or if you just want to apply format, you can skip fetching data while creating datarange
        pygsheets.DataRange('E7','G1000', worksheet=wksheet).apply_format(model_cell)
        gc.run_batch() # All the above requests are executed here
        gc.set_batch_mode(False)
        wksheet.update_value('E7', '=ARRAYFORMULA((SI(ESTVIDE(F7:F);"";LIEN_HYPERTEXTE("https://www.gafeo.fr/report/trainingsessions/index.php?id=" & F7:F;"Lien Rapports session"))))')
  
        if wksheet.cols == 15:
           continue
        # A REMETTRE POUR NOUVELLE VERSION
        # wksheet.insert_cols(col=4,number=3)
        # wksheet.update_col(6,[['Course_ID'], ['Apprenant_GAFEO_ID']],row_offset=HEADER_COL-1)
        # wksheet.update_col(5,[['Lien Cours GAFEO']],row_offset=HEADER_COL-1)
        # pygsheets.Cell(f"E{HEADER_COL}").wrap_strategy="WRAP"
        # wksheet.update_value('B7', '=ARRAYFORMULA((SI(ESTVIDE(A7:A);"";LIEN_HYPERTEXTE("https://www.gafeo.fr/report/learningtimecheck/index.php?itemid=" & G7:G & "&view=user&id=" & F7:F;L7:L))))')
        # gc.set_batch_mode(True)
        # model_cell = pygsheets.Cell('D7')
        # model_cell.color = (207/255, 226/255, 243/255, 1) #
        # # model_cell.format = (pygsheets.FormatType.TEXT)
        # model_cell.text_format['underline']=True
        # model_cell.text_format['fontSize']=12
        # # Setting format to multiple cells in one go
        # # rng.apply_format(model_cell)  # will make all cell in this range rose color and percent format
        # # Or if you just want to apply format, you can skip fetching data while creating datarange
        # pygsheets.DataRange('E7','G1000', worksheet=wksheet).apply_format(model_cell)
        # gc.run_batch() # All the above requests are executed here
        # gc.set_batch_mode(False)
        # wksheet.update_value('E7', '=ARRAYFORMULA((SI(ESTVIDE(F7:F);"";LIEN_HYPERTEXTE("https://www.gafeo.fr/report/trainingsessions/index.php?id=" & F7:F;"Lien Rapports session"))))')

        wksheet.insert_cols(col=4,number=1)
        wksheet.update_col(5,[['Lien Cours GAFEO']],row_offset=HEADER_COL-1)
        # drange_lien_gafeo_id_appr = pygsheets.Datarange(start='E7', end='G1000', worksheet=wksheet)
        # drange_lien_gafeo_id_appr.protected = True
        # drange_lien_gafeo_id_appr.apply_format(model_cell_link_black)
        wksheet.update_value('E7', '=ARRAYFORMULA((SI(ESTVIDE(F7:F);"";LIEN_HYPERTEXTE("https://www.gafeo.fr/report/trainingsessions/index.php?id=" & F7:F;"Lien Rapports session"))))')



if __name__ == "__main__":

    gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/Scrapping_Url_GAFEO/Oauth_gg/code_secret_client_173621592213-0llal3ntmv316usvtboglpb52leq6jcu.apps.googleusercontent.com.json')
    sh = gc.open_by_key('1Ix4xc_kJPrIBXQL8JmGVz_XNnY4t7AMHdHvmyVlExiA')
    update_layout_worksheet(gc,sh,6,4,22)