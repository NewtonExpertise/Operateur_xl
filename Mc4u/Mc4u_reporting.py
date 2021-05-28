import  xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta


try:
    ws = xw.sheets.active
    wb = ws.book
except:
    ws = False
    wb= False

def if_reporting():
    """controle sur 3 cellules pour identifier qu'il s'agisse bien de l'excel de reporting Mcdo"""
    if ws:
        if ws.range("A1").value == "JOINT VENTURE MONTHLY REPORT" and ws.range("B14").value == "REEL MOIS" and ws.range("B17").value == "REEL CUMUL":
            
            return True
        else:
            return False

def if_report_sheet(fin):
    """Vérifie cohérence onglet / prériode"""
    target = fin.strftime("%m%Y")
    exist = False
    print(target)
    for sheet in wb.sheets:
        print(sheet)
        if target ==  sheet.name:
            exist = True

    return exist


def dataReporting(dico, fin_periode, nb_resto):
    """Alimente le reporting avec les données extraites de la compta et les cumuls de l'onget du mois précédent."""
    print("dataReporting")
    vents_alim = dico["001"]["sold_mensuel"]
    vents_non_alim = dico["085"]["sold_mensuel"]
    marge_brut = dico["020"]["sold_mensuel"]
    pac = dico["060"]["sold_mensuel"]
    soi = dico["093"]["sold_mensuel"]
    g_a = dico["102"]["sold_mensuel"]
    resultat_pnl = dico["106"]["sold_mensuel"]
    # holding = 
    lastmonth = fin_periode + relativedelta(months=-1)
    target_sheet_periode = fin_periode.strftime("%m%Y")
    find_ws = False
    for sheet in wb.sheets:
        if target_sheet_periode ==  sheet.name:
            find_ws=True
    if find_ws:
        print(target_sheet_periode)
        target_sheet_result = wb.sheets[target_sheet_periode]
        target_sheet_result.range("F14").value = vents_alim
        target_sheet_result.range("G14").value = vents_non_alim
        target_sheet_result.range("J14").value = marge_brut
        target_sheet_result.range("L14").value = pac
        target_sheet_result.range("N14").value = soi
        target_sheet_result.range("P14").value = g_a
        target_sheet_result.range("Q14").value = resultat_pnl
        # target_sheet_result.range("R17").value = holding

        last_sheet_name = get_last_sheet(lastmonth)
        if last_sheet_name:    
            last_sheet = wb.sheets[last_sheet_name]
            if last_sheet.range("F17").value:
                vents_alim += last_sheet.range("F17").value
            if last_sheet.range("G17").value:
                vents_non_alim += last_sheet.range("G17").value
            if last_sheet.range("J17").value:
                marge_brut += last_sheet.range("J17").value
            if last_sheet.range("L17").value:
                pac += last_sheet.range("L17").value
            if last_sheet.range("N17").value:
                soi += last_sheet.range("N17").value
            if last_sheet.range("P17").value:
                g_a += last_sheet.range("P17").value
            if last_sheet.range("Q17").value:
                resultat_pnl += last_sheet.range("Q17").value
            # if last_sheet.range("R17").value:
                # holding += last_sheet.range("R17").value

        target_sheet_result = wb.sheets[target_sheet_periode]
        target_sheet_result.range("D14").value = nb_resto
        target_sheet_result.range("D17").value = nb_resto
        target_sheet_result.range("F17").value = vents_alim
        target_sheet_result.range("G17").value = vents_non_alim
        target_sheet_result.range("J17").value = marge_brut
        target_sheet_result.range("L17").value = pac
        target_sheet_result.range("N17").value = soi
        target_sheet_result.range("P17").value = g_a
        target_sheet_result.range("Q17").value = resultat_pnl
        # target_sheet_result.range("R17").value = holding


def get_last_sheet(date):
    """déduit le nom de l'onglet du mois précédent via la date de fin et vérifie son existance"""
    find=False
    target_name = date.strftime("%m%Y")
    for sheet in wb.sheets:
        if target_name ==  sheet.name:
            find= sheet.name
    return find

if __name__ == '__main__':
    from datetime import datetime
    fin_periode = datetime(year=2021, month = 4, day =1)
    dico = ""
    dataReporting("",fin_periode,"")