import pygsheets


def update_layout_worksheet(gc,HEADER_COL,start_sheet,end_sheet):
    """gc : googlesheet a modifier; HEADER_COL: numero de colonnes avec HEADERNAME, start_sheet: index col à modifier, end_sheet:fin colonne à modifier"""

    # Suivi Eron 2022 (derniere version)
    #sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')

    # DUPLICATA Suivi Eron 2022 (derniere version)
    sh = gc.open_by_key('1Ix4xc_kJPrIBXQL8JmGVz_XNnY4t7AMHdHvmyVlExiA')

    #Format all sheet if different
    for sheet_number in range(start_sheet,end_sheet):
        # print(sheetNumber)
        wksheet = sh[sheet_number]
        if wksheet.cols == 14:
            return
        wksheet.insert_cols(col=4,number=2)
        wksheet.update_col(5,[['Course_ID'], ['Apprenant_GAFEO_ID']],row_offset=HEADER_COL-1)
        wksheet.update_value('B7', '=ARRAYFORMULA((SI(ESTVIDE(A7:A);"";LIEN_HYPERTEXTE("https://www.gafeo.fr/report/learningtimecheck/index.php?itemid=" & F7:F & "&view=user&id=" & E7:E;K7:K))))')




if __name__ == "__main__":

    gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/Scrapping_Url_GAFEO/Oauth_gg/code_secret_client_95743482524-gj2mnoav9naiqt454ggvt71r28r4n3dk.apps.googleusercontent.com.json')
    update_layout_worksheet(gc,6,4,22)