import os
from mdbagent import MdbConnect
import xlwings as xw
from datetime import timedelta,datetime
from Mc4u.mc4u import Gene_Mc4u
# from Mc4u.generateur_excel import generateur_excel
# from Mc4u.generateur_xl_ACD import generateur_excel

import re


def Mc4u_Minot(code_dossier, debut, fin):
    print("gene_Mc4u")
    Mc4u = Gene_Mc4u(code_dossier, debut, fin)
    Mc4u.gen_xl_qdra()
    print("getmc4u")
    Mc4u.get_Mc4u_qdra()

def Mc4u_Minot_acd(code_dossier, debut, fin):
    print("gene_Mc4u")
    Mc4u = Gene_Mc4u(code_dossier, debut, fin)
    Mc4u.gen_xl_acd()
    print("getmc4u")
    Mc4u.get_Mc4u_acd()



def ecritures_analytiques(mdbpath):
    """
    Renvoie vers le tableur la listes des écritures analytiques
    """

    sql = f"""
    SELECT
        E.CodeJournal AS Journal, E.Folio,
        DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture) AS DateEcr,
        E.NumeroCompte AS Compte, E.Libelle, E.MontantTenuDebit AS Debit, E.MontantTenuCredit AS Credit, 
        (E.MontantTenuDebit-E.MontantTenuCredit) AS Solde,
        E.NumeroPiece AS Piece, A.Centre, E.CodeOperateur AS Oper, E.DateSysSaisie
    FROM
        (
            SELECT 
                TypeLigne, NumUniq, NumeroCompte, CodeJournal,  Folio, LigneFolio, 
                PeriodeEcriture, JourEcriture, NumLigne, Libelle, MontantTenuDebit, MontantTenuCredit, 
                NumeroPiece, CodeOperateur, DateSysSaisie
            FROM Ecritures 
            WHERE TypeLigne='E' 
            AND (NumeroCompte LIKE '6%' OR NumeroCompte LIKE '7%')) E
    LEFT JOIN
        (
            SELECT 
                TypeLigne, CodeJournal, Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, Centre 
            FROM Ecritures WHERE TypeLigne='A') A
    ON E.CodeJournal=A.CodeJournal
    AND E.Folio=A.Folio
    AND E.LigneFolio=A.LigneFolio
    AND E.PeriodeEcriture=A.PeriodeEcriture
    """
    # Récupération data
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    # Création nouvelle onglet
    ws = xw.sheets.active
    wb = ws.book
    num_client = os.path.basename(os.path.dirname(mdbpath))
    base = os.path.basename(os.path.dirname(os.path.dirname(mdbpath)))
    Nom_feuille_excel = "EC_Analytique_"+num_client+"_"+base
    add_sheet_new_name(wb, Nom_feuille_excel)

    # Création des plages
    bnligne=len(data)
    if base == "DC":
        xw.Range('I1:I'+str(bnligne)).name = 'ColCentreN'
        xw.Range('A1:A'+str(bnligne)).name = 'ColJournauxN'
        xw.Range('G1:G'+str(bnligne)).name = 'ColSoldeN'
    else:
        xw.Range('I1:I'+str(bnligne)).name = 'ColCentreN1'
        xw.Range('A1:A'+str(bnligne)).name = 'ColJournauxN1'
        xw.Range('G1:G'+str(bnligne)).name = 'ColSoldeN1'

    # formatage
    xw.Range('H:H').number_format='@'
    xw.Range('C:C').number_format='@'
    xw.Range('E:G').number_format='# ##0,00'
    xw.Range('A1').value = data
    ws.autofit()
    set_AutoFilter(ws)

def ecritures(mdbpath):
    """
    Renvoie vers le tableur la listes des écritures 
    """
    sql = """
    SELECT
    DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture) as DateEcr,
    E.NumeroCompte, E.CodeJournal, E.Folio, E.Libelle, E.MontantTenuDebit,
    E.MontantTenuCredit, E.CodeLettrage, E.NumeroPiece, E.CodeOperateur, E.DateSysSaisie, A.Centre, T.DateEcheance
    FROM
    (
        (SELECT TypeLigne, NumUniq, NumeroCompte, CodeJournal,  Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, Libelle, MontantTenuDebit, MontantTenuCredit, NumeroPiece, CodeOperateur, DateSysSaisie, CodeLettrage FROM Ecritures WHERE TypeLigne='E') E
    LEFT JOIN
        (SELECT TypeLigne, CodeJournal, Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, Centre FROM Ecritures WHERE TypeLigne='A') A
    ON E.CodeJournal=A.CodeJournal
    AND E.Folio=A.Folio
    AND E.LigneFolio=A.LigneFolio
    AND E.PeriodeEcriture=A.PeriodeEcriture)
    LEFT JOIN
    (SELECT TypeLigne, CodeJournal, Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, DateEcheance FROM Ecritures WHERE TypeLigne='T') T
    ON E.CodeJournal=T.CodeJournal
    AND E.Folio=T.Folio
    AND E.LigneFolio=T.LigneFolio
    AND E.PeriodeEcriture=T.PeriodeEcriture
    """
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    ws = xw.sheets.active
    wb = ws.book
    num_client = os.path.basename(os.path.dirname(mdbpath))
    base = os.path.basename(os.path.dirname(os.path.dirname(mdbpath)))
    Nom_feuille_excel = "Ecritures_"+num_client+"_"+base
    ws = add_sheet_new_name(wb, Nom_feuille_excel)
    ws.range('G:G').number_format='@'
    ws.range('B:B').number_format='@'
    ws.range('E:F').number_format='# ##0,00'
    ws.range('A1').value = data
    ws.autofit()
    set_AutoFilter(ws)

def balance_generale_totale(mdbpath, fin_periode):
    """
    Renvoie vers le tableur la balance générale
    """    
    sql = f"""
    SELECT E.NumeroCompte, C.Intitule,
    SUM(E.MontantTenuDebit) AS Debit, SUM(E.MontantTenuCredit) AS Credit,
    (Debit - Credit) AS Solde
    FROM Ecritures E
    LEFT JOIN Comptes C
    ON E.NumeroCompte=C.Numero
    WHERE E.PeriodeEcriture < #{fin_periode}#
    AND E.TypeLigne = 'E'
    GROUP BY E.NumeroCompte, C.Intitule
    ORDER BY E.NumeroCompte
    """
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    ws = xw.sheets.active
    wb = ws.book
    num_client = os.path.basename(os.path.dirname(mdbpath))
    base = os.path.basename(os.path.dirname(os.path.dirname(mdbpath)))
    Nom_feuille_excel = "B_Generale_totale_"+num_client+"_"+base
    add_sheet_new_name(wb, Nom_feuille_excel)
    
    xw.Range('C:C').number_format='# ##0,00'
    xw.Range('D:D').number_format='# ##0,00'
    xw.Range('E:E').number_format='# ##0,00'
    xw.Range('A:A').number_format='@'
    xw.Range('A1').value = data
    ws.autofit()
    set_AutoFilter(ws)

def balance_generale(mdbpath, fin_periode):
    """
    Renvoie vers le tableur la balance simplifié
    """
    prefix_auxiliaire = get_auxiliaire_prefix(mdbpath) 
    # Contrôle présence auxiliaire : 
    # S'il n'y a pas de paramétrage des auxiliaires :
    if prefix_auxiliaire['client']== None:
        balance_generale_totale(mdbpath, fin_periode)
    # Sinon :
    else:
        sql = f"""
        SELECT E.NumeroCompte AS Numero_Compte, C.Intitule AS Libelle,
        SUM(E.MontantTenuDebit) AS Debit, SUM(E.MontantTenuCredit) AS Credit,
        (Debit - Credit) AS Solde
        FROM Ecritures E
        LEFT JOIN Comptes C
        ON E.NumeroCompte=C.Numero
        WHERE C.Type = 'G'
        AND E.TypeLigne = 'E'
        AND E.PeriodeEcriture < #{fin_periode}#
        GROUP BY E.NumeroCompte, C.Intitule
        ORDER BY E.NumeroCompte
        UNION
        SELECT '41100000' AS Numero_Compte, 'Clients' AS Libelle,
        SUM(E.MontantTenuDebit) AS Debit, SUM(E.MontantTenuCredit) AS Credit,
        (Debit - Credit) AS Solde
        FROM Ecritures E
        LEFT JOIN Comptes C
        ON E.NumeroCompte=C.Numero
        WHERE E.NumeroCompte LIKE '{prefix_auxiliaire['client']}%'
        AND E.TypeLigne = 'E'
        AND E.PeriodeEcriture < #{fin_periode}#
        UNION
        SELECT '40100000' AS Numero_Compte, 'Fournisseurs' AS Libelle,
        SUM(E.MontantTenuDebit) AS Debit, SUM(E.MontantTenuCredit) AS Credit,
        (Debit - Credit) AS Solde
        FROM Ecritures E
        LEFT JOIN Comptes C
        ON E.NumeroCompte=C.Numero
        WHERE E.NumeroCompte LIKE '{prefix_auxiliaire['fournisseur']}%'
        AND E.TypeLigne = 'E'
        AND E.PeriodeEcriture < #{fin_periode}#
        """
        with MdbConnect(mdbpath) as mdb:
            info, data = mdb.queryInfoData(sql)
        headers = [x[0] for x in info]
        data.insert(0, headers)

        ws = xw.sheets.active
        wb = ws.book
        num_client = os.path.basename(os.path.dirname(mdbpath))
        base = os.path.basename(os.path.dirname(os.path.dirname(mdbpath)))
        Nom_feuille_excel = "B_Generale_"+num_client+"_"+base
        add_sheet_new_name(wb, Nom_feuille_excel)
        
        xw.Range('C:C').number_format='# ##0,00'
        xw.Range('D:D').number_format='# ##0,00'
        xw.Range('E:E').number_format='# ##0,00'
        xw.Range('A:A').number_format='@'
        xw.Range('A1').value = data
        ws.autofit()
        set_AutoFilter(ws)

def balance_clients(mdbpath, fin_periode):
    """
    Renvoie vers le tableur la balance clients
    """
    prefix_auxiliaire = get_auxiliaire_prefix(mdbpath)
    sql = f"""
    SELECT E.NumeroCompte, C.Intitule,
    SUM(E.MontantTenuDebit) AS Debit, SUM(E.MontantTenuCredit) AS Credit,
    (Debit - Credit) AS Solde
    FROM Ecritures E
    LEFT JOIN Comptes C
    ON E.NumeroCompte=C.Numero
    WHERE C.Type = 'C'
    AND E.NumeroCompte LIKE '{prefix_auxiliaire['client']}%'
    AND E.TypeLigne = 'E'
    AND E.PeriodeEcriture < #{fin_periode}#
    GROUP BY E.NumeroCompte, C.Intitule
    ORDER BY E.NumeroCompte
    """
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    ws = xw.sheets.active
    wb = ws.book
    num_client = os.path.basename(os.path.dirname(mdbpath))
    base = os.path.basename(os.path.dirname(os.path.dirname(mdbpath)))
    Nom_feuille_excel = "B_clients_"+num_client+"_"+base
    add_sheet_new_name(wb, Nom_feuille_excel)

    xw.Range('C:C').number_format='# ##0,00'
    xw.Range('D:D').number_format='# ##0,00'
    xw.Range('E:E').number_format='# ##0,00'
    xw.Range('A:A').number_format='@'
    xw.Range('A1').value = data
    ws.autofit()
    set_AutoFilter(ws)

def balance_fournisseurs(mdbpath, fin_periode):
    """
    Renvoie vers le tableur la balance fournisseurs
    """
    prefix_auxiliaire = get_auxiliaire_prefix(mdbpath)
    sql = f"""
    SELECT E.NumeroCompte, C.Intitule,
    SUM(E.MontantTenuDebit) AS Debit, SUM(E.MontantTenuCredit) AS Credit,
    (Debit - Credit) AS Solde
    FROM Ecritures E
    LEFT JOIN Comptes C
    ON E.NumeroCompte=C.Numero
    WHERE C.Type = 'F'
    AND E.NumeroCompte LIKE '{prefix_auxiliaire['fournisseur']}%'
    AND E.TypeLigne = 'E'
    AND E.PeriodeEcriture < #{fin_periode}#
    GROUP BY E.NumeroCompte, C.Intitule
    ORDER BY E.NumeroCompte
    """
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    ws = xw.sheets.active
    wb = ws.book
    num_client = os.path.basename(os.path.dirname(mdbpath))
    base = os.path.basename(os.path.dirname(os.path.dirname(mdbpath)))
    Nom_feuille_excel = "B_fournisseurs_"+num_client+"_"+base
    add_sheet_new_name(wb, Nom_feuille_excel)

    xw.Range('C:C').number_format='# ##0,00'
    xw.Range('D:D').number_format='# ##0,00'
    xw.Range('E:E').number_format='# ##0,00'
    xw.Range('A:A').number_format='@'
    xw.Range('A1').value = data
    ws.autofit()
    set_AutoFilter(ws)

def codes_journaux(mdbpath):
    sql="""
    SELECT Code from Journaux ORDER BY Code;
    """
    with MdbConnect(mdbpath) as mdb:
        data = mdb.query(sql)
    set1 = {x[0] for x in data}

    # Requête sur la base des paramètres généraux QcomptaC
    drive, _ = os.path.splitdrive(mdbpath)
    QcomptaC = os.path.abspath(os.path.join(drive, "quadra/database/cpta/qcomptac.mdb"))
    with MdbConnect(QcomptaC) as mdb:
        data = mdb.query(sql)
    set2 = {x[0] for x in data}

    fullset = sorted(set1.union(set2))

    row, col = callSelectedCell()

    for i, value in enumerate(fullset, row):
        xw.Range((i, col)).value = value


def grand_livre(mdbpath, fin_periode):
    """
    retourne un grand livre 
    """

    sql = f"""
    SELECT DateEcr ,  NumeroCompte , CodeJournal, Folio, Libelle, MontantTenuDebit, MontantTenuCredit,  CodeLettrage,  NumeroPiece,  CodeOperateur , DateSysSaisie, Centre,  DateEcheance
    FROM ( SELECT '' as DateEcr , E.NumeroCompte as NumeroCompte ,'' as  CodeJournal, '' as Folio, 'Total compte' as Libelle, 
        SUM(E.MontantTenuDebit) as MontantTenuDebit, SUM(E.MontantTenuCredit) as MontantTenuCredit, '' as CodeLettrage, '' as NumeroPiece, '' as CodeOperateur , '' as DateSysSaisie, '' as Centre, '' as DateEcheance
        FROM Ecritures as E 
        WHERE TypeLigne='E'
        AND E.PeriodeEcriture < #{fin_periode}#
        GROUP BY E.NumeroCompte
        UNION 
        SELECT DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture) as DateEcr,
        E.NumeroCompte, E.CodeJournal, E.Folio, E.Libelle, E.MontantTenuDebit,
        E.MontantTenuCredit, E.CodeLettrage, E.NumeroPiece, E.CodeOperateur, E.DateSysSaisie, A.Centre, T.DateEcheance
        FROM ( (SELECT TypeLigne, NumUniq, NumeroCompte, CodeJournal,  Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, Libelle, MontantTenuDebit, MontantTenuCredit, NumeroPiece, CodeOperateur, DateSysSaisie, CodeLettrage 
                FROM Ecritures 
                WHERE TypeLigne='E') E
            LEFT JOIN
                (SELECT TypeLigne, CodeJournal, Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, Centre 
                FROM Ecritures 
                WHERE TypeLigne='A') A
                ON E.CodeJournal=A.CodeJournal
                AND E.Folio=A.Folio
                AND E.LigneFolio=A.LigneFolio
                AND E.PeriodeEcriture=A.PeriodeEcriture)
            LEFT JOIN
                (SELECT TypeLigne, CodeJournal, Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, DateEcheance 
                FROM Ecritures 
                WHERE TypeLigne='T') T
            ON E.CodeJournal=T.CodeJournal
            AND E.Folio=T.Folio
            AND E.LigneFolio=T.LigneFolio
            AND E.PeriodeEcriture=T.PeriodeEcriture
            WHERE E.PeriodeEcriture < #{fin_periode}# )
    ORDER BY NumeroCompte
    """
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    ws = xw.sheets.active
    wb = ws.book
    num_client = os.path.basename(os.path.dirname(mdbpath))
    base = os.path.basename(os.path.dirname(os.path.dirname(mdbpath)))
    Nom_feuille_excel = "G_L_"+num_client+"_"+base
    add_sheet_new_name(wb, Nom_feuille_excel)

    xw.Range('G:G').number_format='@'
    xw.Range('B:B').number_format='@'
    xw.Range('A:A').number_format='jj/mm/aaaa'
    xw.Range('E:F').number_format='# ##0,00'
    xw.Range('A1').value = data
    ws.autofit()
    set_AutoFilter(ws)


def set_AutoFilter(ws):
    used_range_rows = (ws.api.UsedRange.Row, ws.api.UsedRange.Row + ws.api.UsedRange.Rows.Count -1)
    used_range_cols = (ws.api.UsedRange.Column, ws.api.UsedRange.Column + ws.api.UsedRange.Columns.Count -1)
    ws.range(*zip(used_range_rows, used_range_cols)).api.AutoFilter(1)



def callSelectedCell():
    """
    Retourne les coordonnées de la cellule sélectionnée
    """
    wb = xw.books.active  
    selected_row = wb.app.selection.row
    selected_col = wb.app.selection.column
    return selected_row, selected_col

def get_mois_exercice(QcomptaC):
    """
    Renvoie la listes des mois de l'exercice.
    """
    sql = """
    SELECT DebutExercice, FinExercice, DateLimiteSaisie
    FROM Dossier1
    """
    with MdbConnect(QcomptaC) as mdb:
        periode = mdb.query(sql)
    
    debut = periode[0][0] # Début d'exercice
    fin = periode[0][1] # Fin d'exercice
    date_limite = periode[0][2] # Date limite en cas d'exercice multiple sur le DC
    no_date_limite = datetime(year = 1899, month = 12, day = 30) #Valeur renvoyé par Access lorsqu'un champ date est vide.
    test_liste = []
    # # Contrôle s'il y a une date limite de set alors on la prend en compte pour retourner la période total.
    if date_limite == no_date_limite:
        while debut <= fin :
            
            test_liste.append(end_of_month(debut))
            debut = end_of_month(debut)+ timedelta(1)
    # Sinon on utilise la date de fin d'exercice pour établir la période.
    else:
        date_limite= end_of_month(date_limite)
        while debut <= date_limite :
            test_liste.append(end_of_month(debut))
            debut = end_of_month(debut)+ timedelta(1)
 
    return test_liste

def get_auxiliaire_prefix(QcomptaC):
    """
    Récupère le préfixe des codes clients-fournisseurs :
    """
    sql = """
    SELECT
    CodifClasse0Seule
    FROM Dossier2
    """
    with MdbConnect(QcomptaC) as mdb:
        data = mdb.query(sql)
    rt = {}
    if data[0][0] == "F":
        rt['fournisseur']=0
        rt['client']=9
    elif data[0][0] == "C":
        rt['fournisseur']=9
        rt['client']=0
    else:
        rt['fournisseur']=None
        rt['client']=None

    return rt

def end_of_month(dt0):
    """
    Renvoi le dernier jour du mois de la date donnée
    """
    dt1 = dt0.replace(day=1)
    dt2 = dt1 + timedelta(days=32)
    dt3 = dt2.replace(day=1) - timedelta(days=1)
    return dt3

def add_sheet_new_name(wb, nom):
    """
    Génère une feuille excel avec un nom unique
    nb : une feuille excel ne peut contenir que 31 caractères
    """
    nom = nom[:29]
    increment = 0
    sheet = [sheet.name for sheet in wb.sheets]
    if nom in sheet:

        for sheet_name in sheet:
            try:
                name, compteur , _ =  re.split(nom+r'(\d+)',sheet_name) 
            except ValueError:
                compteur =0
            if increment< int(compteur):
                increment = int(compteur)

        new_name = nom + str(increment+1)
    else:
        new_name = nom 
    new_ws = wb.sheets.add(new_name)
    new_ws.book.activate(True)

    return new_ws


if __name__ == "__main__":


    # import pprint
    # pp=pprint.PrettyPrinter(indent=4)
    mdb=  r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DC\000948\qcompta.mdb'
    from datetime import datetime
    # pp.pprint(get_mois_exercice(mdb))
    # import xlwings as xw
    # ws = xw.sheets.active
    # wb=ws.book
    ecritures(mdb)

