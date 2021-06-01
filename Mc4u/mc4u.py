import logging
from sys import exit
from datetime import datetime
from Mc4u.generateur_excel import generateur_excel
from Mc4u.Mc4u_reporting import if_reporting, dataReporting, if_report_sheet
import xlrd
import ctypes


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s:%(funcName)s\t%(module)s -- %(message)s",
    datefmt="%m-%d %H:%M",
    filename="mc4u.log",
    filemode="w",)

class Gene_Mc4u():
    def __init__(self, code_dossier, debut, fin):
        
        self.code_dossier = code_dossier
        self.dt_debut = debut
        self.dt_fin = fin
        action=False

        if if_reporting():
            if if_report_sheet(self.dt_fin):
                 action = True
            else:
                action = self.Messagebox("Attention",f"Le fichier de reporting actuellement ouvert, pour les Mc4u consolidés, ne correspond pas à la période interrogé.\n\nSouhaitez vous tout de même générer le/les Mc4u sans le traitement dédier au reporting ?")

        else:
            action = self.Messagebox("Attention",f"Il semblerait que vous ne soyez pas sur l'onglet de reporting pour les Mc4u consolidés ou que vous n’ayez pas activé la modification de ce dernier.\n\nSouhaitez vous tout de même générer le/les Mc4u sans le traitement dédier au reporting ?")
        if action:
           self.generateur_excel = generateur_excel(self.code_dossier)
        else:
            exit()

    def get_Mc4u(self):

        debut = self.dt_debut
        fin =  self.dt_fin
        self.generateur_excel.invoke( debut, fin)
        if if_reporting():
            if self.generateur_excel.PNL_global:
                dataReporting(self.generateur_excel.PNL_global, fin, self.generateur_excel.nb_resto, self.generateur_excel.mdb_holding)
 
    def Messagebox(self,title, text):
        result = ctypes.windll.user32.MessageBoxW(0, text, title, 4)
        if result == 6:
            result = True
        elif result == 7:
            result = False
        return result


if __name__ == '__main__':
    import datetime
    debut = datetime.datetime(year = 2021, month= 1 , day = 1)
    fin =datetime.datetime(year = 2021, month= 4 , day = 30)
    mc4u = Gene_Mc4u("0966",debut,fin)
    mc4u.get_Mc4u()