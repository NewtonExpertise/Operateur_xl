from datetime import datetime
import logging
import pyodbc
import sys
import os
import pprint
try:
    from Mc4u.mdbagent import MdbConnect
except :
    from mdbagent import MdbConnect

pp = pprint.PrettyPrinter(indent=4)
try:
    sources = sys._MEIPASS
except:
    sources = ''

class reqBalanceAna(object) :
    """description of class"""
    def __init__(self, chem_base):
        self.chem_base = chem_base
        self.pnl_dico = None
        logging.info("Query db compta : {}".format(self.chem_base))
        
        # constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq='+ self.chem_base
    def get_data(self, debut, fin):
        sql = f"""SELECT
            CUM.Centre AS CodeAna, 
            MENS.Solde AS SoldeMens, 
            CUM.Solde AS SoldeCumul 
            FROM 
            (SELECT 
            Centre, 
            SUM(MontantTenuDebit) - SUM(MontantTenuCredit) AS Solde 
            FROM Ecritures 
            WHERE TypeLigne='A' 
            AND PeriodeEcriture=#{fin}# 
            AND (NumeroCompte LIKE '6%' 
            OR NumeroCompte LIKE '7%') 
            GROUP BY Centre) MENS RIGHT JOIN 
            (SELECT 
            Centre, 
            SUM(MontantTenuDebit) - SUM(MontantTenuCredit) AS Solde 
            FROM Ecritures 
            WHERE TypeLigne='A' 
            AND PeriodeEcriture>=#{debut}# 
            AND PeriodeEcriture<=#{fin}# 
            AND (NumeroCompte LIKE '6%' 
            OR NumeroCompte LIKE '7%') 
            GROUP BY Centre) CUM 
            ON MENS.Centre=CUM.Centre"""
        with MdbConnect(self.chem_base) as mdb:
            self.data = mdb.query(sql)
        return self.data

    def creaDic(self, codes):

        copy_codes = codes.copy()
        for item in codes.keys():
            copy_codes[item].setdefault("sold_mensuel", 0.0)
            copy_codes[item].setdefault("sold_cumule", 0.0)
        
        for centre, soldeMensuel, soldeCumule in self.data:

            for item in codes.keys():
                
                if centre == codes[item]['centre'] :

                    # Si le centre n'est pas alimenté dans Quadra
                    # faut remplacer Non par 0
                    if soldeMensuel == None:
                        soldeMensuel = 0.0
                    if soldeCumule == None:
                        soldeCumule = 0.0
                    # Les centres de produit doivent être 
                    # inversés
                    if item in ["001", "085", "087"] :
                        soldeMensuel = -soldeMensuel
                        soldeCumule = -soldeCumule

                    copy_codes[item].update({"sold_mensuel":soldeMensuel,
                                       "sold_cumule": soldeCumule})
                    break
        
        copy_codes["000"].update(
            {"sold_mensuel" : (copy_codes["001"]["sold_mensuel"] + 
                              copy_codes["085"]["sold_mensuel"]),
             "sold_cumule" : (copy_codes["001"]["sold_cumule"] + 
                              copy_codes["085"]["sold_cumule"]) })

        copy_codes["014"].update(
            {"sold_mensuel" : (copy_codes["010"]["sold_mensuel"] +
                               copy_codes["011"]["sold_mensuel"] +
                               copy_codes["012"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["010"]["sold_cumule"] +
                             copy_codes["011"]["sold_cumule"] +
                             copy_codes["012"]["sold_cumule"])})

        copy_codes["019"].update(
            {"sold_mensuel" : (copy_codes["013"]["sold_mensuel"] +
                               copy_codes["014"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["013"]["sold_cumule"] +
                             copy_codes["014"]["sold_cumule"])})

        copy_codes["020"].update(
            {"sold_mensuel" : (copy_codes["001"]["sold_mensuel"] -
                               copy_codes["019"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["001"]["sold_cumule"] -
                             copy_codes["019"]["sold_cumule"])})

        copy_codes["028"].update(
            {"sold_mensuel" : (copy_codes["023"]["sold_mensuel"] +
                               copy_codes["024"]["sold_mensuel"] +
                               copy_codes["026"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["023"]["sold_cumule"] +
                             copy_codes["024"]["sold_cumule"] +
                             copy_codes["026"]["sold_cumule"])})

        copy_codes["055"].update(
            {"sold_mensuel" : (copy_codes["028"]["sold_mensuel"] +
                               copy_codes["030"]["sold_mensuel"] +
                               copy_codes["032"]["sold_mensuel"] +
                               copy_codes["034"]["sold_mensuel"] +
                               copy_codes["036"]["sold_mensuel"] +
                               copy_codes["038"]["sold_mensuel"] +
                               copy_codes["040"]["sold_mensuel"] +
                               copy_codes["042"]["sold_mensuel"] +
                               copy_codes["044"]["sold_mensuel"] +
                               copy_codes["046"]["sold_mensuel"] +
                               copy_codes["048"]["sold_mensuel"] +
                               copy_codes["050"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["028"]["sold_cumule"] +
                             copy_codes["030"]["sold_cumule"] +
                             copy_codes["032"]["sold_cumule"] +
                             copy_codes["034"]["sold_cumule"] +
                             copy_codes["036"]["sold_cumule"] +
                             copy_codes["038"]["sold_cumule"] +
                             copy_codes["040"]["sold_cumule"] +
                             copy_codes["042"]["sold_cumule"] +
                             copy_codes["044"]["sold_cumule"] +
                             copy_codes["046"]["sold_cumule"] +
                             copy_codes["048"]["sold_cumule"] +
                             copy_codes["050"]["sold_cumule"])})

        copy_codes["060"].update(
            {"sold_mensuel" : (copy_codes["020"]["sold_mensuel"] -
                               copy_codes["055"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["020"]["sold_cumule"] -
                             copy_codes["055"]["sold_cumule"])})


        copy_codes["084"].update(
            {"sold_mensuel" : (copy_codes["062"]["sold_mensuel"] +
                               copy_codes["064"]["sold_mensuel"] +
                               copy_codes["065"]["sold_mensuel"] +
                               copy_codes["068"]["sold_mensuel"] +
                               copy_codes["071"]["sold_mensuel"] +
                               copy_codes["074"]["sold_mensuel"] +
                               copy_codes["076"]["sold_mensuel"] +
                               copy_codes["077"]["sold_mensuel"] +
                               copy_codes["078"]["sold_mensuel"] +
                               copy_codes["080"]["sold_mensuel"] +
                               copy_codes["082"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["062"]["sold_cumule"] +
                             copy_codes["064"]["sold_cumule"] +
                             copy_codes["065"]["sold_cumule"] +
                             copy_codes["068"]["sold_cumule"] +
                             copy_codes["071"]["sold_cumule"] +
                             copy_codes["074"]["sold_cumule"] +
                             copy_codes["076"]["sold_cumule"] +
                             copy_codes["077"]["sold_cumule"] +
                             copy_codes["078"]["sold_cumule"] +
                             copy_codes["080"]["sold_cumule"] +
                             copy_codes["082"]["sold_cumule"])})

        copy_codes["090"].update(
            {"sold_mensuel" : (copy_codes["085"]["sold_mensuel"] +
                               copy_codes["087"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["085"]["sold_cumule"] +
                             copy_codes["087"]["sold_cumule"])})

        copy_codes["093"].update(
            {"sold_mensuel" : (copy_codes["060"]["sold_mensuel"] -
                               copy_codes["084"]["sold_mensuel"] +
                               copy_codes["090"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["060"]["sold_cumule"] -
                             copy_codes["084"]["sold_cumule"] +
                             copy_codes["090"]["sold_cumule"])})

        copy_codes["104"].update(
            {"sold_mensuel" : (copy_codes["101"]["sold_mensuel"] +
                               copy_codes["102"]["sold_mensuel"] +
                               copy_codes["103"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["101"]["sold_cumule"] +
                             copy_codes["102"]["sold_cumule"] +
                             copy_codes["103"]["sold_cumule"])})

        copy_codes["106"].update(
            {"sold_mensuel" : (copy_codes["093"]["sold_mensuel"] -
                               copy_codes["104"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["093"]["sold_cumule"] -
                             copy_codes["104"]["sold_cumule"])})

        copy_codes["108"].update(
            {"sold_mensuel" : (copy_codes["106"]["sold_mensuel"] -
                               copy_codes["107"]["sold_mensuel"]),
            "sold_cumule" : (copy_codes["106"]["sold_cumule"] -
                             copy_codes["107"]["sold_cumule"])})

        logging.debug("Le dictionnaire est prêt")
        logging.debug("\n" + pp.pformat(copy_codes))

        return copy_codes


# if __name__ == '__main__':

#     from importCodes import importCodes

#     FORMAT = '%(asctime)s -- %(module)s -- %(levelname)s -- %(message)s'
#     logging.basicConfig(handlers=[logging.basicConfig(level=logging.DEBUG,
#                                                   format=FORMAT)])

#     xl = os.path.join(sources,"Mc4u\CodesAnalytiques.xlsx")
#     xl = "V:\Mathieu\PROJET\PROJET_COMPTA\operateur_xl\Operateurxl_sans_pandas\Mc4u\CodesAnalytiques.xlsx"


#     myobj = importCodes(xl)
#     codes = myobj.creaDic()

#     db_path = "//srvquadra/qappli/quadra/database/cpta/dc/000177/qcompta.mdb"
#     debut = datetime(year=2019, month=1, day=1)
#     fin = datetime(year=2019, month=6, day=1)

#     logging.debug("\nDossier : {}\nDebut : {}\nFin : {}".
#                   format(db_path, debut, fin))

#     myObj = reqBalanceAna(db_path, debut, fin)
#     #print ("constr p&l")
#     data = myObj.creaDic(codes)
#     print(data)
