# -*- coding: utf-8 -*-
import logging
import tempfile
import os
from tkinter.constants import FALSE
import xlrd
from xlwt import Workbook, easyxf, Formula
from datetime import datetime, timedelta
from Mc4u.importCodes import importCodes
from Mc4u.ProfitAndLossRequest import reqBalanceAna
from Mc4u.ACD_env import get_group_mcdo, get_list_group_mcdo, get_balance_ana_exercice, get_holding_result
import re
from collections import defaultdict, OrderedDict
import sys
import copy
try:
    sources = sys._MEIPASS
except:
    sources = ''

class generateur_excel_ACD(object):

    def __init__(self, code_dossier):
        self.code_dossier = code_dossier
        self.group = get_group_mcdo(code_dossier)
        self.PNL_global = False
        self.nb_resto = 0
        self.l_mois_periode = None
        self.last_month_for_pnl = ""
        self.group, self.holding, self.list_groupe_Mc4u = self.get_groupe()
        self.rt_holding = None



    def invoke(self, debut ,fin):
        """articule la génération d'un excel Mc4u."""
        self.rt_holding = get_holding_result(self.holding, debut, fin)
        book = Workbook()
        if self.group:
            x =  0
            for codeQuadra, codeMcdo, _ in self.list_groupe_Mc4u:
                xl_codes_ana = os.path.join(sources,"Mc4u/CodesAnalytiques.xls")
                impcod = importCodes(xl_codes_ana)
                codes_ana = impcod.creaDic()
                data = get_balance_ana_exercice(codeQuadra, debut, fin)
                dico = self.creaDic(codes_ana, data)
     
                if self.PNL_global:
                    self.create_global_PNL(dico)
                else:
                    self.PNL_global = dico
                    self.nb_resto+=1
               
                self.get_Mc4u(book, dico, fin, codeMcdo)

            self.Invoke_pnl_minot(debut, fin)

            
        else:
            xl_codes_ana = os.path.join(sources,"Mc4u/CodesAnalytiques.xls")
            impcod = importCodes(xl_codes_ana)
            codes_ana = impcod.creaDic()
            mdb_path = self.qenv.make_db_path(self.code_dossier, "DC")
            sa = reqBalanceAna(mdb_path, debut, fin)
            dico = sa.creaDic(codes_ana)
            self.get_Mc4u(book, dico, fin, self.nom_dossier)
        filename = f"JV_{self.group}_MC4U_{fin.strftime('%Y-%m')}.xls"
        self.save_wb(book,filename)

    def get_Mc4u(self, book, dico, fin, codeMcdo):
        """
        Génère un Mc4u sur un nouvel onglet.
        """

        sheet1 = book.add_sheet(f'{fin.strftime("%Y%m")}_{codeMcdo}')
        sheet1.col(1).width = 10000
        sheet1.col(2).width = 5000
        sheet1.col(3).width = 5000

        fmtHeader = easyxf(
            (
                "font: bold True; "
                "alignment: horizontal center; "
                "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ### ##0.00",
        )
        fmtNeutr = easyxf(
            (
                "font: bold False; "
                "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ### ##0.00",
        )
        fmtJaune = easyxf(
            (
                "font: bold True; "
                "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour light_orange;"
            ),
            num_format_str="# ### ### ##0.00",
        )
        fmtGris = easyxf(
            (
                "font: bold False; "
                "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour gray25;"
            ),
            num_format_str="# ### ### ##0.00",
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
                        region = ws.cell(row, 5).value
                        list_group.append([codeQuadra, codeMcdo, region])
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

    def save_wb(self, book, filename):
        """récupère un WB et l'enregistre dans un excel."""
        if self.group:
            tmp_dir = tempfile.gettempdir()
            self.nom_group = self.group.replace(" ", "_")
            filepath = os.path.join(tmp_dir, filename)
        else:
            tmp_dir = tempfile.gettempdir()
            self.nom_dossier = self.nom_dossier.replace(" ", "_")
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


    def list_mois_periode(self, debut, fin):
        print(debut)
        print(fin)
        fin_annee = datetime(year= debut.year+1, month = debut.month,  day = debut.day)-timedelta(1)
        print(fin_annee)
        self.last_month_for_pnl =fin_annee.replace(day = 1)
 
        self.last_month_for_pnl
        list_periode_intervalle_mois = OrderedDict((datetime.strptime((debut + timedelta(_)).strftime(r"%m-%y"),"%m-%y"), None) for _ in range((fin - debut).days)).keys()
        list_periode_intervalle_mois = list(list_periode_intervalle_mois)
        if list_periode_intervalle_mois:
            list_periode_intervalle_mois.append(fin)
        else:
            list_periode_intervalle_mois.append(debut)
        print(list_periode_intervalle_mois)
        return list_periode_intervalle_mois


    def Invoke_pnl_minot(self,debut, fin):

        """articule la génération d'un excel PNL pour chaque resto de chaque mois pour une année."""
        self.rt_holding = get_holding_result(self.holding, debut, fin)
        book = Workbook()
        pnl_dict = defaultdict(dict)
        self.l_mois_periode = self.list_mois_periode(debut, fin)
        periodeannuel = self.list_mois_periode(debut, self.last_month_for_pnl)
        xl_codes_ana = os.path.join(sources,"Mc4u/CodesAnalytiques.xls")
        impcod = importCodes(xl_codes_ana)
        default_code_ana = impcod.creaDic()
        list_region = []
        if self.group:
            for codeQuadra, codeMcdo, region in self.list_groupe_Mc4u:
                
                
                for periode in self.l_mois_periode:
                    impcod = importCodes(xl_codes_ana)
                    codes_ana = impcod.creaDic()
                    data = get_balance_ana_exercice(codeQuadra, periode, periode)
                    dico = self.creaDic(codes_ana, data)

                    pnl_dict[codeMcdo].update({"region":region,
                                                periode:dico})
 
            
            pnl_dict = self.regroup_resto(pnl_dict, self.l_mois_periode)
            self.get_Pnl(book, pnl_dict, periodeannuel)

        else:
            xl_codes_ana = os.path.join(sources,"Mc4u/CodesAnalytiques.xls")
            impcod = importCodes(xl_codes_ana)
            codes_ana = impcod.creaDic()
            mdb_path = self.qenv.make_db_path(self.code_dossier, "DC")
            sa = reqBalanceAna(mdb_path, debut, fin)
            dico = sa.creaDic(codes_ana,data)
            self.get_Mc4u(book, dico, fin, self.nom_dossier)
        filename = f"PNL_Group_Minot_{fin.strftime('%Y%m')}.xls"
        self.save_wb(book,filename)


    def creaDic(self, codes, data):

        copy_codes = codes.copy()
        for item in codes.keys():
            copy_codes[item].setdefault("sold_mensuel", 0.0)
            copy_codes[item].setdefault("sold_cumule", 0.0)
        if data : 
            for centre, soldeMensuel, soldeCumule in data:

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



        return copy_codes


    def regroup_resto(self, dict_pnl, l_periode):
        """dict {code{dt[[ana]]}}"""
        dict_pnl2 = copy.deepcopy(dict_pnl)
        d_regroup = defaultdict(dict)

        for key, val in dict_pnl2.items():
            
            for dt_mois in l_periode:
                
                if val['region'] in d_regroup:
                    if dt_mois in d_regroup[val['region']]:
                        for centre, data, in val[dt_mois].items():
                            for key, montant in data.items():
                                if key in ['sold_cumule','sold_mensuel']:
                                    d_regroup[val["region"]][dt_mois][centre][key] += montant
                    else:
                        d_regroup[val["region"]][dt_mois] = val[dt_mois]
                else:
                    d_regroup[val["region"]]= {dt_mois:val[dt_mois]}

        dict_pnl.update(d_regroup)
        return dict_pnl



    def get_Pnl(self, book, dico, periode):
        """
        Génère des onglets PNL pour le groupe Minot.
        Un onglet par restaurant et par localité.
        Les données seront afficher par mois.
        un cumul sera repris à la fin.
        """
        fmtHeader = easyxf(
            (
                "font: bold True; "
                "alignment: horizontal center; "
                # "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ### ### ##0.00",
        )
        fmtNeutr = easyxf(
            (
                "font: bold False; "
                # "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ### ### ##0.00",
        )
        fmtJaune = easyxf(
            (
                "font: bold True; "
                # "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour light_orange;"
            ),
            num_format_str="# ### ### ### ##0.00",
        )
        fmtGris = easyxf(
            (
                "font: bold False; "
                # "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour gray25;"
            ),
            num_format_str="# ### ### ### ##0.00",
        )
        for codeMcdo, d_data in dico.items():
            
            sheet1 = book.add_sheet(f'{codeMcdo}')
            sheet1.col(1).width = 10000
            sheet1.col(3).width = 2051
            sheet1.col(5).width = 2051
            sheet1.col(7).width = 2051
            sheet1.col(9).width = 2051
            sheet1.col(11).width = 2051
            sheet1.col(13).width = 2051
            sheet1.col(15).width = 2051
            sheet1.col(17).width = 2051
            sheet1.col(19).width = 2051
            sheet1.col(21).width = 2051
            sheet1.col(23).width = 2051
            sheet1.col(25).width = 2051
            sheet1.col(27).width = 2051
            sheet1.col(2).width = 3407
            sheet1.col(4).width = 3883
            sheet1.col(6).width = 3883
            sheet1.col(8).width = 3883
            sheet1.col(10).width = 3883
            sheet1.col(12).width = 3883
            sheet1.col(14).width = 3883
            sheet1.col(16).width = 3883
            sheet1.col(18).width = 3883
            sheet1.col(20).width = 3883
            sheet1.col(22).width = 3883
            sheet1.col(24).width = 3883
            sheet1.col(26).width = 3883
            sheet1.write(0, 0, "N° COMPTE", fmtHeader)
            sheet1.write(0, 1, "COMPTES", fmtHeader)
            col = 2
            self.set_formula_pnl(sheet1)
            for mois in periode:
                if mois in d_data:
    
                    sheet1.write(0, col, mois.strftime("%B %Y"), fmtHeader)
            
                    sheet1.write(0, col+1, "%", fmtHeader)

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


                        ####controle datetime.

                        if item in ["001", "020", "060", "093", "106", "108"]:
                            format = fmtJaune
                        elif item in ["014", "019", "028", "055", "084", "090", "104"]:
                            format = fmtGris
                        else:
                            format = fmtNeutr
                        try:
                            sheet1.write(i, 0, item, format)
                            sheet1.write(i, 1, d_data[mois][item]["libelle"], format)
                        except:
                            pass
                        
                        sheet1.write(i, col, d_data[mois][item]["sold_mensuel"], format)
                        # sheet1.write(i, 3, dico[item]["sold_cumule"], format)
                        i += 1
                    col +=2
                else:
                    i=1
                    sheet1.write(0, col, mois.strftime("%B %Y"), fmtHeader)
                    sheet1.write(0, col+1, "%", fmtHeader)
                    for item in list_keys:
                        if item in ["001", "020", "060", "093", "106", "108"]:
                            format = fmtJaune
                        elif item in ["014", "019", "028", "055", "084", "090", "104"]:
                            format = fmtGris
                        else:
                            format = fmtNeutr
                        sheet1.write(i, col, 0, format)
                        i += 1
                    col += 2



    def set_formula_pnl(self, sheet):
        """
        Ajoute les formules de calcule pour obtenir la répartition en % des montants
        """
        fmtHeader = easyxf(
            (
                "font: bold True; "
                "alignment: horizontal center; "
                # "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ### ### ##0.00",
        )
        
        fmtN = easyxf(
            (
                "font: bold False; "
                # "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str='0.00%',
        )
        fmtJ = easyxf(
            (
                "font: bold True; "
                # "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour light_orange;"
            ),
            num_format_str='0.00%',
        )
        fmtG = easyxf(
            (
                "font: bold False; "
                # "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour gray25;"
            ),
            num_format_str='0.00%',
        )
        
        row = 2
        col =3
        col_formule={3:"C", 5:"E", 7:"G", 9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 19:"S", 21:"U", 23:"W", 25:"Y", 27:"AA"}
        for x, y in enumerate(range(13)):
            if x:
                col=col+2
            sheet.write(row, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+1}/{col_formule[col]}3)'),fmtJ)
            sheet.write(row+1, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+2}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+2, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+3}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+3, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+4}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+4, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+5}/{col_formule[col]}3)'),fmtG)
            sheet.write(row+5, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+6}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+6, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+7}/{col_formule[col]}3)'),fmtG)
            sheet.write(row+7, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+8}/{col_formule[col]}3)'),fmtJ)
            sheet.write(row+8, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+9}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+9, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+10}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+10, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+11}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+11, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+12}/{col_formule[col]}3)'),fmtG)

            sheet.write(row+12, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+13}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+13, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+14}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+14, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+15}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+15, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+16}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+16, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+17}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+17, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+18}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+18, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+19}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+19, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+20}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+20, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+21}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+21, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+22}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+22, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+23}/{col_formule[col]}3)'),fmtN)

            sheet.write(row+23, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+24}/{col_formule[col]}3)'),fmtG)
            sheet.write(row+24, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+25}/{col_formule[col]}3)'),fmtJ)

            sheet.write(row+25, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+26}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+26, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+27}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+27, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+28}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+28, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+29}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+29, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+30}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+30, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+31}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+31, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+32}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+32, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+33}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+33, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+34}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+34, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+35}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+35, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+36}/{col_formule[col]}3)'),fmtN)

            sheet.write(row+36, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+37}/{col_formule[col]}3)'),fmtG)

            sheet.write(row+37, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+38}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+38, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+39}/{col_formule[col]}3)'),fmtN)

            sheet.write(row+39, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+40}/{col_formule[col]}3)'),fmtG)
            sheet.write(row+40, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+41}/{col_formule[col]}3)'),fmtJ)

            sheet.write(row+41, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+42}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+42, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+43}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+43, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+44}/{col_formule[col]}3)'),fmtN)

            sheet.write(row+44, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+45}/{col_formule[col]}3)'),fmtG)
            sheet.write(row+45, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+46}/{col_formule[col]}3)'),fmtJ)

            sheet.write(row+46, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+47}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+47, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+48}/{col_formule[col]}3)'),fmtJ)
            sheet.write(row+48, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+49}/{col_formule[col]}3)'),fmtN)
            sheet.write(row+49, col, Formula(f'IF({col_formule[col]}3=0;0;{col_formule[col]}{row+50}/{col_formule[col]}3)'),fmtN)
         
        col-=1
        fmtN = easyxf(
            (
                "font: bold False; "
                # "borders: left thick, right thick, top thick, bottom thick;"
            ),
            num_format_str="# ### ### ##0.00",
        )
        fmtJ = easyxf(
            (
                "font: bold True; "
                # "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour light_orange;"
            ),
            num_format_str="# ### ### ##0.00",
        )
        fmtG = easyxf(
            (
                "font: bold False; "
                # "borders: left thick, right thick, top thick, bottom thick; "
                "pattern: pattern solid, fore_colour gray25;"
            ),
            num_format_str="# ### ### ##0.00",
        )
        
        sheet.write(0, col,    Formula(f'"Total "&right(Y1,4)'),fmtHeader)
        sheet.write(0, col+1, "%",fmtHeader)
        sheet.write(row-1, col,  Formula(f"C{row}+E{row}+G{row}+I{row}+K{row}+M{row}+O{row}+Q{row}+S{row}+U{row}+W{row}+Y{row}"),fmtN)
        sheet.write(row, col,    Formula(f"C{row+1}+E{row+1}+G{row+1}+I{row+1}+K{row+1}+M{row+1}+O{row+1}+Q{row+1}+S{row+1}+U{row+1}+W{row+1}+Y{row+1}"),fmtJ)
        sheet.write(row+1, col,  Formula(f"C{row+2}+E{row+2}+G{row+2}+I{row+2}+K{row+2}+M{row+2}+O{row+2}+Q{row+2}+S{row+2}+U{row+2}+W{row+2}+Y{row+2}"),fmtN)
        sheet.write(row+2, col,  Formula(f"C{row+3}+E{row+3}+G{row+3}+I{row+3}+K{row+3}+M{row+3}+O{row+3}+Q{row+3}+S{row+3}+U{row+3}+W{row+3}+Y{row+3}"),fmtN)
        sheet.write(row+3, col,  Formula(f"C{row+4}+E{row+4}+G{row+4}+I{row+4}+K{row+4}+M{row+4}+O{row+4}+Q{row+4}+S{row+4}+U{row+4}+W{row+4}+Y{row+4}"),fmtN)
        sheet.write(row+4, col,  Formula(f"C{row+5}+E{row+5}+G{row+5}+I{row+5}+K{row+5}+M{row+5}+O{row+5}+Q{row+5}+S{row+5}+U{row+5}+W{row+5}+Y{row+5}"),fmtG)
        sheet.write(row+5, col,  Formula(f"C{row+6}+E{row+6}+G{row+6}+I{row+6}+K{row+6}+M{row+6}+O{row+6}+Q{row+6}+S{row+6}+U{row+6}+W{row+6}+Y{row+6}"),fmtN)
        sheet.write(row+6, col,  Formula(f"C{row+7}+E{row+7}+G{row+7}+I{row+7}+K{row+7}+M{row+7}+O{row+7}+Q{row+7}+S{row+7}+U{row+7}+W{row+7}+Y{row+7}"),fmtG)
        sheet.write(row+7, col,  Formula(f"C{row+8}+E{row+8}+G{row+8}+I{row+8}+K{row+8}+M{row+8}+O{row+8}+Q{row+8}+S{row+8}+U{row+8}+W{row+8}+Y{row+8}"),fmtJ)
        sheet.write(row+8, col,  Formula(f"C{row+9}+E{row+9}+G{row+9}+I{row+9}+K{row+9}+M{row+9}+O{row+9}+Q{row+9}+S{row+9}+U{row+9}+W{row+9}+Y{row+9}"),fmtN)
        sheet.write(row+9, col,  Formula(f"C{row+10}+E{row+10}+G{row+10}+I{row+10}+K{row+10}+M{row+10}+O{row+10}+Q{row+10}+S{row+10}+U{row+10}+W{row+10}+Y{row+10}"),fmtN)
        sheet.write(row+10, col, Formula(f"C{row+11}+E{row+11}+G{row+11}+I{row+11}+K{row+11}+M{row+11}+O{row+11}+Q{row+11}+S{row+11}+U{row+11}+W{row+11}+Y{row+11}"),fmtN)
        sheet.write(row+11, col, Formula(f"C{row+12}+E{row+12}+G{row+12}+I{row+12}+K{row+12}+M{row+12}+O{row+12}+Q{row+12}+S{row+12}+U{row+12}+W{row+12}+Y{row+12}"),fmtG)
        sheet.write(row+12, col, Formula(f"C{row+13}+E{row+13}+G{row+13}+I{row+13}+K{row+13}+M{row+13}+O{row+13}+Q{row+13}+S{row+13}+U{row+13}+W{row+13}+Y{row+13}"),fmtN)
        sheet.write(row+13, col, Formula(f"C{row+14}+E{row+14}+G{row+14}+I{row+14}+K{row+14}+M{row+14}+O{row+14}+Q{row+14}+S{row+14}+U{row+14}+W{row+14}+Y{row+14}"),fmtN)
        sheet.write(row+14, col, Formula(f"C{row+15}+E{row+15}+G{row+15}+I{row+15}+K{row+15}+M{row+15}+O{row+15}+Q{row+15}+S{row+15}+U{row+15}+W{row+15}+Y{row+15}"),fmtN)
        sheet.write(row+15, col, Formula(f"C{row+16}+E{row+16}+G{row+16}+I{row+16}+K{row+16}+M{row+16}+O{row+16}+Q{row+16}+S{row+16}+U{row+16}+W{row+16}+Y{row+16}"),fmtN)
        sheet.write(row+16, col, Formula(f"C{row+17}+E{row+17}+G{row+17}+I{row+17}+K{row+17}+M{row+17}+O{row+17}+Q{row+17}+S{row+17}+U{row+17}+W{row+17}+Y{row+17}"),fmtN)
        sheet.write(row+17, col, Formula(f"C{row+18}+E{row+18}+G{row+18}+I{row+18}+K{row+18}+M{row+18}+O{row+18}+Q{row+18}+S{row+18}+U{row+18}+W{row+18}+Y{row+18}"),fmtN)
        sheet.write(row+18, col, Formula(f"C{row+19}+E{row+19}+G{row+19}+I{row+19}+K{row+19}+M{row+19}+O{row+19}+Q{row+19}+S{row+19}+U{row+19}+W{row+19}+Y{row+19}"),fmtN)
        sheet.write(row+19, col, Formula(f"C{row+20}+E{row+20}+G{row+20}+I{row+20}+K{row+20}+M{row+20}+O{row+20}+Q{row+20}+S{row+20}+U{row+20}+W{row+20}+Y{row+20}"),fmtN)
        sheet.write(row+20, col, Formula(f"C{row+21}+E{row+21}+G{row+21}+I{row+21}+K{row+21}+M{row+21}+O{row+21}+Q{row+21}+S{row+21}+U{row+21}+W{row+21}+Y{row+21}"),fmtN)
        sheet.write(row+21, col, Formula(f"C{row+22}+E{row+22}+G{row+22}+I{row+22}+K{row+22}+M{row+22}+O{row+22}+Q{row+22}+S{row+22}+U{row+22}+W{row+22}+Y{row+22}"),fmtN)
        sheet.write(row+22, col, Formula(f"C{row+23}+E{row+23}+G{row+23}+I{row+23}+K{row+23}+M{row+23}+O{row+23}+Q{row+23}+S{row+23}+U{row+23}+W{row+23}+Y{row+23}"),fmtN)
        sheet.write(row+23, col, Formula(f"C{row+24}+E{row+24}+G{row+24}+I{row+24}+K{row+24}+M{row+24}+O{row+24}+Q{row+24}+S{row+24}+U{row+24}+W{row+24}+Y{row+24}"),fmtG)
        sheet.write(row+24, col, Formula(f"C{row+25}+E{row+25}+G{row+25}+I{row+25}+K{row+25}+M{row+25}+O{row+25}+Q{row+25}+S{row+25}+U{row+25}+W{row+25}+Y{row+25}"),fmtJ)
        sheet.write(row+25, col, Formula(f"C{row+26}+E{row+26}+G{row+26}+I{row+26}+K{row+26}+M{row+26}+O{row+26}+Q{row+26}+S{row+26}+U{row+26}+W{row+26}+Y{row+26}"),fmtN)
        sheet.write(row+26, col, Formula(f"C{row+27}+E{row+27}+G{row+27}+I{row+27}+K{row+27}+M{row+27}+O{row+27}+Q{row+27}+S{row+27}+U{row+27}+W{row+27}+Y{row+27}"),fmtN)
        sheet.write(row+27, col, Formula(f"C{row+28}+E{row+28}+G{row+28}+I{row+28}+K{row+28}+M{row+28}+O{row+28}+Q{row+28}+S{row+28}+U{row+28}+W{row+28}+Y{row+28}"),fmtN)
        sheet.write(row+28, col, Formula(f"C{row+29}+E{row+29}+G{row+29}+I{row+29}+K{row+29}+M{row+29}+O{row+29}+Q{row+29}+S{row+29}+U{row+29}+W{row+29}+Y{row+29}"),fmtN)
        sheet.write(row+29, col, Formula(f"C{row+30}+E{row+30}+G{row+30}+I{row+30}+K{row+30}+M{row+30}+O{row+30}+Q{row+30}+S{row+30}+U{row+30}+W{row+30}+Y{row+30}"),fmtN)
        sheet.write(row+30, col, Formula(f"C{row+31}+E{row+31}+G{row+31}+I{row+31}+K{row+31}+M{row+31}+O{row+31}+Q{row+31}+S{row+31}+U{row+31}+W{row+31}+Y{row+31}"),fmtN)
        sheet.write(row+31, col, Formula(f"C{row+32}+E{row+32}+G{row+32}+I{row+32}+K{row+32}+M{row+32}+O{row+32}+Q{row+32}+S{row+32}+U{row+32}+W{row+32}+Y{row+32}"),fmtN)
        sheet.write(row+32, col, Formula(f"C{row+33}+E{row+33}+G{row+33}+I{row+33}+K{row+33}+M{row+33}+O{row+33}+Q{row+33}+S{row+33}+U{row+33}+W{row+33}+Y{row+33}"),fmtN)
        sheet.write(row+33, col, Formula(f"C{row+34}+E{row+34}+G{row+34}+I{row+34}+K{row+34}+M{row+34}+O{row+34}+Q{row+34}+S{row+34}+U{row+34}+W{row+34}+Y{row+34}"),fmtN)
        sheet.write(row+34, col, Formula(f"C{row+35}+E{row+35}+G{row+35}+I{row+35}+K{row+35}+M{row+35}+O{row+35}+Q{row+35}+S{row+35}+U{row+35}+W{row+35}+Y{row+35}"),fmtN)
        sheet.write(row+35, col, Formula(f"C{row+36}+E{row+36}+G{row+36}+I{row+36}+K{row+36}+M{row+36}+O{row+36}+Q{row+36}+S{row+36}+U{row+36}+W{row+36}+Y{row+36}"),fmtN)
        sheet.write(row+36, col, Formula(f"C{row+37}+E{row+37}+G{row+37}+I{row+37}+K{row+37}+M{row+37}+O{row+37}+Q{row+37}+S{row+37}+U{row+37}+W{row+37}+Y{row+37}"),fmtG)
        sheet.write(row+37, col, Formula(f"C{row+38}+E{row+38}+G{row+38}+I{row+38}+K{row+38}+M{row+38}+O{row+38}+Q{row+38}+S{row+38}+U{row+38}+W{row+38}+Y{row+38}"),fmtN)
        sheet.write(row+38, col, Formula(f"C{row+39}+E{row+39}+G{row+39}+I{row+39}+K{row+39}+M{row+39}+O{row+39}+Q{row+39}+S{row+39}+U{row+39}+W{row+39}+Y{row+39}"),fmtN)
        sheet.write(row+39, col, Formula(f"C{row+40}+E{row+40}+G{row+40}+I{row+40}+K{row+40}+M{row+40}+O{row+40}+Q{row+40}+S{row+40}+U{row+40}+W{row+40}+Y{row+40}"),fmtG)
        sheet.write(row+40, col, Formula(f"C{row+41}+E{row+41}+G{row+41}+I{row+41}+K{row+41}+M{row+41}+O{row+41}+Q{row+41}+S{row+41}+U{row+41}+W{row+41}+Y{row+41}"),fmtJ)
        sheet.write(row+41, col, Formula(f"C{row+42}+E{row+42}+G{row+42}+I{row+42}+K{row+42}+M{row+42}+O{row+42}+Q{row+42}+S{row+42}+U{row+42}+W{row+42}+Y{row+42}"),fmtN)
        sheet.write(row+42, col, Formula(f"C{row+43}+E{row+43}+G{row+43}+I{row+43}+K{row+43}+M{row+43}+O{row+43}+Q{row+43}+S{row+43}+U{row+43}+W{row+43}+Y{row+43}"),fmtN)
        sheet.write(row+43, col, Formula(f"C{row+44}+E{row+44}+G{row+44}+I{row+44}+K{row+44}+M{row+44}+O{row+44}+Q{row+44}+S{row+44}+U{row+44}+W{row+44}+Y{row+44}"),fmtN)
        sheet.write(row+44, col, Formula(f"C{row+45}+E{row+45}+G{row+45}+I{row+45}+K{row+45}+M{row+45}+O{row+45}+Q{row+45}+S{row+45}+U{row+45}+W{row+45}+Y{row+45}"),fmtG)
        sheet.write(row+45, col, Formula(f"C{row+46}+E{row+46}+G{row+46}+I{row+46}+K{row+46}+M{row+46}+O{row+46}+Q{row+46}+S{row+46}+U{row+46}+W{row+46}+Y{row+46}"),fmtJ)
        sheet.write(row+46, col, Formula(f"C{row+47}+E{row+47}+G{row+47}+I{row+47}+K{row+47}+M{row+47}+O{row+47}+Q{row+47}+S{row+47}+U{row+47}+W{row+47}+Y{row+47}"),fmtN)
        sheet.write(row+47, col, Formula(f"C{row+48}+E{row+48}+G{row+48}+I{row+48}+K{row+48}+M{row+48}+O{row+48}+Q{row+48}+S{row+48}+U{row+48}+W{row+48}+Y{row+48}"),fmtJ)
        sheet.write(row+48, col, Formula(f"C{row+49}+E{row+49}+G{row+49}+I{row+49}+K{row+49}+M{row+49}+O{row+49}+Q{row+49}+S{row+49}+U{row+49}+W{row+49}+Y{row+49}"),fmtN)
        sheet.write(row+49, col, Formula(f"C{row+50}+E{row+50}+G{row+50}+I{row+50}+K{row+50}+M{row+50}+O{row+50}+Q{row+50}+S{row+50}+U{row+50}+W{row+50}+Y{row+50}"),fmtN)
        sheet.write(row+50, col, Formula(f"C{row+51}+E{row+51}+G{row+51}+I{row+51}+K{row+51}+M{row+51}+O{row+51}+Q{row+51}+S{row+51}+U{row+51}+W{row+51}+Y{row+51}"),fmtN)





