# -*- coding: utf-8 -*-
import logging
import tempfile
import os
from tkinter.constants import FALSE
import xlrd
from xlwt import Workbook, easyxf
from datetime import datetime
from Mc4u.importCodes import importCodes
from Mc4u.ProfitAndLossRequest import reqBalanceAna
from Mc4u.quadraenv import QuadraSetEnv
import re
import sys

try:
    sources = sys._MEIPASS
except:
    sources = ''

ipl = r"\\srvquadra\qappli\quadra\database\client\quadra.ipl"

class generateur_excel(object):

    def __init__(self, code_dossier):
        self.qenv = QuadraSetEnv(ipl)
        self.code_dossier = code_dossier.zfill(6)
        self.liste_cli = self.qenv.gi_liste_clients()
        self.nom_dossier = self.liste_cli[self.code_dossier]["rs"]
        self.group, self.holding, self.list_groupe_Mc4u = self.get_groupe()
        self.PNL_global = False
        self.nb_resto = 0
        self.mdb_holding = self.qenv.make_db_path(self.holding, "DC")
        print(self.mdb_holding)


    def invoke(self, debut ,fin):
        """articule la génération d'un excel Mc4u."""
        book = Workbook()
        if self.group:
            for codeQuadra, codeMcdo in self.list_groupe_Mc4u:
                xl_codes_ana = os.path.join(sources,"Mc4u/CodesAnalytiques.xls")
                impcod = importCodes(xl_codes_ana)
                codes_ana = impcod.creaDic()
                mdb_path = self.qenv.make_db_path(codeQuadra, "DC")
                sa = reqBalanceAna(mdb_path, debut, fin)
                dico = sa.creaDic(codes_ana)
                if self.PNL_global:
                    self.create_global_PNL(dico)
                else:
                    self.PNL_global = dico
                    self.nb_resto+=1
                # import pprint
                # pp = pprint.PrettyPrinter(indent=4)
                # pp.pprint(dico)
                self.get_Mc4u(book, dico, fin, codeMcdo)
            
        else:
            xl_codes_ana = os.path.join(sources,"Mc4u/CodesAnalytiques.xls")
            impcod = importCodes(xl_codes_ana)
            codes_ana = impcod.creaDic()
            mdb_path = self.qenv.make_db_path(self.code_dossier, "DC")
            sa = reqBalanceAna(mdb_path, debut, fin)
            dico = sa.creaDic(codes_ana)
            self.get_Mc4u(book, dico, fin, self.nom_dossier)
        self.save_wb(book,fin)

    def get_Mc4u(self, book, dico, fin, codeMcdo):
        """
        Génère un Mc4u sur un nouvelle onglet.
        """

        sheet1 = book.add_sheet(f'{fin.strftime("%Y%m")}_{codeMcdo}')
        sheet1.col(1).width = 10000

        fmtHeader = easyxf(
            (
                "font: bold True; "
                "alignment: horizontal center; "
                "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ##0.00",
        )
        fmtNeutr = easyxf(
            (
                "font: bold False; "
                "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ##0.00",
        )
        fmtJaune = easyxf(
            (
                "font: bold True; "
                "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour light_orange;"
            ),
            num_format_str="# ### ##0.00",
        )
        fmtGris = easyxf(
            (
                "font: bold False; "
                "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour gray25;"
            ),
            num_format_str="# ### ##0.00",
        )

        sheet1.write(0, 0, "N° COMPTE", fmtHeader)
        sheet1.write(0, 1, "COMPTES", fmtHeader)
        sheet1.write(0, 2, "MENSUEL", fmtHeader)
        sheet1.write(0, 3, "CUMULE", fmtHeader)

        i = 1

        # Alimentation des cellules
        # Pour respecter l'inversion 013/014 on met la liste
        # en dur
        list_keys = [
            '000', '001', '010', '011', '012', '014', '013', '019', '020', '023', 
            '024', '026', '028', '030', '032', '034', '036', '038', '040', '042', 
            '044', '046', '048', '050', '055', '060', '062', '064', '065', '068', 
            '071', '074', '076', '077', '078', '080', '082', '084', '085', '087', 
            '090', '093', '101', '102', '103', '104', '106', '107', '108', '109', '110']
        for item in list_keys:
            if item in ["001", "020", "060", "093", "106", "108"]:
                format = fmtJaune
            elif item in ["014", "019", "028", "055", "084", "090", "104"]:
                format = fmtGris
            else:
                format = fmtNeutr
            sheet1.write(i, 0, item, format)
            sheet1.write(i, 1, dico[item]["libelle"], format)
            sheet1.write(i, 2, dico[item]["sold_mensuel"], format)
            sheet1.write(i, 3, dico[item]["sold_cumule"], format)
            i += 1

    def get_groupe(self):
        """retourn le nom du groupe et la correspondance entre les code dossier quadra et les code d'identification restaurant mcdo."""

        wb = xlrd.open_workbook(os.path.join(sources,"Mc4u/TableCorrespondance.xls"))
        ws = wb.sheet_by_name("Mc4u")
        list_group = []
        group = False
        holding = False
        for row in range(1, ws.nrows):
            if self.code_dossier == str(int(ws.cell(row, 3).value)).zfill(6):
                group = ws.cell(row, 0).value
        if group:
            for row in range(1, ws.nrows):
                if ws.cell(row, 0).value == group:
                    if ws.cell(row, 1).value == "Restaurant":
                        codeQuadra = str(int(ws.cell(row, 3).value)).zfill(6)
                        codeMcdo = str(int(ws.cell(row, 4).value))
                        list_group.append([codeQuadra, codeMcdo])
                    if ws.cell(row, 1).value == "Holding":
                        holding = str(int(ws.cell(row, 3).value)).zfill(6)
        else:
            pass
            ####
            ####
            #### showinfos messages erreur. ---> Génération Mc4u individuel.
            ####
            ####
        return group, holding, list_group

    def create_global_PNL(self, dico):
        """prend un P&L pour creer un p&l global"""
        self.nb_resto+=1
        for code_ana_PNL , detaille in dico.items():
            self.PNL_global[code_ana_PNL]['sold_mensuel']+= detaille['sold_mensuel']
            self.PNL_global[code_ana_PNL]['sold_cumule']+= detaille['sold_cumule']

    def save_wb(self, book, fin):
        """récupère un WB et l'enregistre dans un excel."""
        if self.group:
            tmp_dir = tempfile.gettempdir()
            self.nom_group = self.group.replace(" ", "_")
            filename = f"JV_{self.nom_group}_MC4U_{fin.strftime('%Y-%m')}.xls"
            filepath = os.path.join(tmp_dir, filename)
        else:
            tmp_dir = tempfile.gettempdir()
            self.nom_dossier = self.nom_dossier.replace(" ", "_")
            filename = self.nom_dossier + "_MC4U_" + fin.strftime("%Y-%m") + ".xls"
            filepath = os.path.join(tmp_dir, filename)
        try:
            filepath = self.get_unique_path(filepath)
            book.save(filepath)
            os.system(f"start excel.exe {filepath}")

        except IOError as e:
            logging.error(e)

    def get_unique_path(self, path):
        """
        Retourn un nom de fichier unique en fonction des fichiers déjà existant dans le dossier
        """

        # si le nom de fichier existe, on en cherche un autre
        while os.path.exists(path):
            # on vire l'extension
            base, ext = os.path.splitext(path)
            try:
                # on extrait le compteur si il existe
                base, counter, f = re.split(r"\((\d+)\)$", base)
            except ValueError:
                counter = 0
            # on reconstruit le path
            path = "%s(%s)%s" % (base, int(counter) + 1, ext)
        return path
