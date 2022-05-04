from tkinter import *
from tkinter import messagebox, END
from tkinter.ttk import Combobox
import actions
from quadraenv import QuadraSetEnv
from espion import update_espion
from datetime import datetime
from time import strftime, strptime
from tkinter.messagebox import showwarning
from Mc4u.ACD_env import get_all_exercice, get_month_period
import locale
import os
import xlrd
import sys
try:
    sources = sys._MEIPASS
except:
    sources = ''
locale.setlocale(locale.LC_TIME,'')




class Application(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.couleur = "#E4AB5B"
        self.pack()
        self.create_widgets()
        self.index_list = 0
        self.reporting_minot = False


    def create_widgets(self):

        # Création des widgets
        self.var_dossiers = StringVar()
        self.var_dossiers.set("")
        self.lab_dossier = Label(self, text="Dossiers \ud83d\udd0e",font=('Helvetica', 12) , foreground='orange')
        self.lab_select_dossier = Label(self, text="\ud83d\udc49",font=('Helvetica', 14, 'bold'),foreground='orange')
        self.lab_base = Label(self, text="\ud83d\udc49",font=('Helvetica', 14, 'bold'),foreground='orange')
        self.lab_select_action = Label(self, text="\ud83d\udc49",font=('Helvetica', 14,'bold'),foreground='orange')
        self.lab_Fin_Periode = Label(self, text="\ud83d\udcc6",font=('Helvetica', 14, 'bold') , foreground='orange')
        self.saisie1 = Entry(self, width=25, textvariable=self.var_dossiers, cursor='question_arrow')
        self.saisie1.focus_set()
        self.liste_dossiers = Listbox(self, width=25, height=8, selectbackground=self.couleur, cursor="hand2")
        self.liste_bases = Listbox(self, width=25, height=5, selectbackground=self.couleur, cursor="hand2")
        self.liste_bases.config(height=0)
        self.liste_actions = Listbox(self,width=25, selectbackground=self.couleur, cursor="hand2")
        self.liste_actions.config(height=0)
        self.combobox_periode = Combobox(self, width=22, state="readonly", cursor="hand2")
        self.combobox_debut = Combobox(self, width=22, state="readonly", cursor="hand2")
        self.combobox_fin = Combobox(self, width=22, state="readonly", cursor="hand2")
        self.bt_reporting_minot = Button(self, width=22, text="Mc4u_Minot", cursor="hand2",background="#ff3300")

        # Positions
        self.lab_dossier.grid(row=0, column=0, padx=10, pady=3,sticky='e')
        self.saisie1.grid(row=0, column=1, padx=10, pady=3)
        self.liste_dossiers.grid(row=1, column=1, padx=10, pady=3)
        self.lab_select_dossier.grid(row=1, column=0, padx=10, pady=3,sticky='e')
        self.liste_bases.grid(row=4, column=1, padx=10, pady=3)
        self.liste_actions.grid(row=5, column=1, padx=10, pady=3)


        # Actions Binding
        self.liste_dossiers.bind("<ButtonRelease-1>", self.makeListeBase)
        self.liste_bases.bind("<ButtonRelease-1>", self.setMdbPath)
        self.liste_actions.bind("<ButtonRelease-1>", self.setAction)
        self.combobox_periode.bind("<<ComboboxSelected>>", self.setAction_periode)
        self.combobox_fin.bind("<<ComboboxSelected>>", self.setAction_debut_fin)
        self.bt_reporting_minot.bind("<ButtonRelease-1>", self.get_reporting)
        
        
        
        # Callback pour filtrage de la liste dossiers
        self.var_dossiers.trace("w", lambda name, index,
                                mode: self.filter_list_dossier())

        # Dictionnaires des actions
        self.dispatch = {
            actions.ecritures.__name__: actions.ecritures,
            actions.ecritures_analytiques.__name__: actions.ecritures_analytiques,
            actions.grand_livre.__name__: actions.grand_livre,
            actions.balance_generale_totale.__name__: actions.balance_generale_totale,
            actions.balance_generale.__name__: actions.balance_generale,
            actions.balance_clients.__name__: actions.balance_clients,
            actions.balance_fournisseurs.__name__: actions.balance_fournisseurs,
            actions.codes_journaux.__name__: actions.codes_journaux,
            actions.Mc4u_Minot.__name__: actions.Mc4u_Minot,
        }
        for i, action in enumerate(self.dispatch.keys()):
            self.liste_actions.insert(i, action)
        self.liste_actions.configure(state=DISABLED)
        self.makeListeDossier()
        self.filter_list_dossier()


    def get_reporting(self,e):

        if self.reporting_minot:
            for i, action in enumerate(self.dispatch.keys()):
                self.liste_actions.insert(i, action)
                
            self.liste_actions.grid(row=5, column=1, padx=10, pady=3)
            self.liste_actions.configure(state=DISABLED)
            self.makeListeBase('')
            self.liste_bases.unbind("<ButtonRelease-1>")
            self.liste_bases.bind("<ButtonRelease-1>", self.setAction)
            self.combobox_fin.unbind("<<ComboboxSelected>>")
            self.combobox_fin.bind("<<ComboboxSelected>>", self.setAction_debut_fin)
            self.reporting_minot = False
        else:
            self.exercice_ACD = get_all_exercice(self.code_dossier)
            self.liste_bases.delete(0,"end")
            self.liste_actions.grid_forget()

            for i, periode in enumerate(self.exercice_ACD):
                d, f = periode
                self.liste_bases.insert(i,f"{(datetime.strptime(f, '%Y%m%d')).strftime('%d/%m/%Y')}")
            self.liste_bases.unbind("<ButtonRelease-1>")
            self.liste_bases.bind("<ButtonRelease-1>", self.set_reporting_minot_view)
            self.combobox_fin.unbind("<<ComboboxSelected>>")
            self.combobox_fin.bind("<<ComboboxSelected>>", self.get_reporting_minot_debut_fin)
            self.reporting_minot = True

    # def action_acd(self, e):
    #     print(self.bt_ACD.cget("background"))
    #     self.bt_qdr.config(background= "SystemButtonFace")
    #     self.bt_ACD.config(background="#ff3300")
    #     self.dispatch = {
    #         actions.Mc4u_Minot.__name__: actions.Mc4u_Minot,
    #     }
    #     self.liste_actions.delete(0, 'end')
    #     for i, action in enumerate(self.dispatch.keys()):
    #         self.liste_actions.insert(i, action)



    # def action_qdra(self, e):
    #     self.bt_ACD.config(background= "SystemButtonFace")
    #     self.bt_qdr.config(background= "#ff3300")
       
    #     self.dispatch = {
    #         actions.ecritures.__name__: actions.ecritures,
    #         actions.ecritures_analytiques.__name__: actions.ecritures_analytiques,
    #         actions.grand_livre.__name__: actions.grand_livre,
    #         actions.balance_generale_totale.__name__: actions.balance_generale_totale,
    #         actions.balance_generale.__name__: actions.balance_generale,
    #         actions.balance_clients.__name__: actions.balance_clients,
    #         actions.balance_fournisseurs.__name__: actions.balance_fournisseurs,
    #         actions.codes_journaux.__name__: actions.codes_journaux,
    #         actions.Mc4u_Minot.__name__: actions.Mc4u_Minot,
    #     }
    #     self.liste_actions.delete(0, "end")
    #     for i, action in enumerate(self.dispatch.keys()):
    #         self.liste_actions.insert(i, action)


    def setMdbPath(self, e):
        """
        Définit le chemin complet vers la base qcompta (mdb)
        """
        print("bouyaa")
        self.lab_Fin_Periode.grid_forget()
        self.combobox_periode.grid_forget()
        self.lab_select_action.grid(row=4, column=0, padx=10, pady=3,sticky='e')
        self.liste_actions.configure(state=NORMAL)
        index, = self.liste_bases.curselection()
        self.base = self.liste_bases.get(index)
        self.mdb = False
        for base, chemin in self.dbList:
            if self.base == base:
                self.mdb = chemin
        if self.mdb:
            pass
        else:
            print("bouyaka")



    def makeListeDossier(self):
        """
        Prépare le liste des dossiers
        """
        self.dicDossier = {}
        ipl = r"\\srvquadra\qappli\quadra\database\client\quadra.ipl"
        self.qenv = QuadraSetEnv(ipl)
        for code, rs in self.qenv.gi_list_clients():
            label = f"{rs} ({code})"
            self.dicDossier.update({label: code})

    def makeListeBase(self, e):
        """
        Prépare la liste des bases (DC, DA, ...)
        """

        self.liste_actions.configure(state=DISABLED)
        self.lab_select_action.grid_forget()
        self.combobox_periode.grid_forget()
        self.lab_Fin_Periode.grid_forget()
        self.liste_bases.delete(0, END)
        index, = self.liste_dossiers.curselection()
        value = self.liste_dossiers.get(index)
        self.code_dossier = self.dicDossier[value]

        print(self.code_dossier)

        if self.code_dossier:

            wb = xlrd.open_workbook(os.path.join(sources,"Mc4u/TableCorrespondance.xls"))
            ws = wb.sheet_by_name("Mc4u")
            list_group = []
            group = False
            for row in range(1, ws.nrows):
                if self.code_dossier == str(int(ws.cell(row, 3).value)).zfill(6):
                    group = ws.cell(row, 0).value
            if group:
                for row in range(1, ws.nrows):
                    if ws.cell(row, 0).value == group:
                        list_group.append(str(int(ws.cell(row, 3).value)).zfill(6))


            if self.code_dossier in list_group:
                self.bt_reporting_minot.grid(row=2, column=1, padx=10, pady=3)
                self.lab_base.grid(row=4, column=0, padx=10, pady=3,sticky='e')
        
        self.dbList = self.qenv.recent_cpta(self.code_dossier, depth=3)
        for i, (base, _) in enumerate(self.dbList):
            self.liste_bases.insert(i, base)

    def filter_list_dossier(self):
        """
        Filtrage auto de la liste des dossiers
        """
        search_term = self.var_dossiers.get()
        lbox_list = [x for x in self.dicDossier.keys()]
        self.liste_dossiers.delete(0, END)
        for item in lbox_list:
            if search_term.lower() in item.lower():
                self.liste_dossiers.insert(END, item)

    def setAction(self, e):
        """
        Sélection du programmes qui sera lancé
        """
        print('set_action')
        # self.liste_actions.configure(state='disabled')
        index, = self.liste_actions.curselection()
        value = self.liste_actions.get(index)

        if "balance" in value or "livre" in value:
            self.combobox_periode.set('')
            self.combobox_periode.grid(row = 8, column = 1, padx = 10, pady = 3)
            self.lab_Fin_Periode.grid(row=9, column=0, rowspan= 4, padx=10, pady=3,sticky='e')
            index, = self.liste_actions.curselection()
            self.select_action = self.liste_actions.get(index)
            periode = [date.strftime("%Y-%B") for date in actions.get_mois_exercice(self.mdb)]
            print(periode)
            self.combobox_periode['values'] = periode
            update_espion(self.code_dossier, self.base, value)
        elif "Mc4u" in value:
            self.combobox_debut.set('')
            self.combobox_fin.set('')
            self.combobox_debut.grid(row = 8, column = 1, padx = 10, pady = 3)
            self.combobox_fin.grid(row = 9, column = 1, padx = 10, pady = 3)
            self.lab_Fin_Periode.grid(row=8, column=0, rowspan= 4, padx=10, pady=3,sticky='e')
            self.combobox_fin.bind("<<ComboboxSelected>>", self.setAction_debut_fin)
            index, = self.liste_actions.curselection()
            self.select_action = self.liste_actions.get(index)
            periode = [ date.strftime("%Y-%B") for date in actions.get_mois_exercice(self.mdb)]
            self.combobox_debut['values'] = periode
            self.combobox_fin['values'] = periode
        else:
            self.dispatch[value](self.mdb)
            messagebox.showinfo("Annonce", "Export terminé")
            update_espion(self.code_dossier, self.base, value)
            sys.exit()


    def set_reporting_minot_view(self, e):
        self.combobox_debut.set('')
        self.combobox_fin.set('')
        self.combobox_debut.grid(row = 8, column = 1, padx = 10, pady = 3)
        self.combobox_fin.grid(row = 9, column = 1, padx = 10, pady = 3)
        self.lab_Fin_Periode.grid(row=8, column=0, rowspan= 4, padx=10, pady=3,sticky='e')
        index, = self.liste_bases.curselection()
        self.cloture = self.liste_bases.get(index)

        for dates in  self.exercice_ACD:
            if self.cloture == f"{dates[1][-2:]}/{dates[1][4:6]}/{dates[1][:4]}":
                d = datetime.strptime(dates[0], "%Y%m%d")
                f = datetime.strptime(dates[1], "%Y%m%d")
                periode = get_month_period(d, f)
        self.combobox_debut['values'] = periode
        self.combobox_fin['values'] = periode

    def show_traitement(self):
        self.lab_dossier.configure(text = "Traitement en cours ⌛")
    
    def setAction_periode(self, e):
        """
        Sélection du programmes qui sera lancé avec une période choisie
        """
        # mois sélectionné :
        select_mois = self.combobox_periode.get()
        select_mois = actions.end_of_month(datetime.strptime(select_mois, "%Y-%B"))
        self.dispatch[self.select_action](self.mdb, select_mois)
        messagebox.showinfo("Annonce", "Export terminé")
        update_espion(self.code_dossier, self.base, self.select_action)
        sys.exit()

    def setAction_debut_fin(self, e):
        """
        Sélection du programmes qui sera lancé avec une période choisie
        """

        # mois sélectionné :
        self.show_traitement()
        select_debut = self.combobox_debut.get()
        select_debut = datetime.strptime(select_debut, "%Y-%B")
        select_fin = self.combobox_fin.get()
        select_fin = datetime.strptime(select_fin, "%Y-%B")
        self.dispatch[self.select_action](self.code_dossier, select_debut, select_fin)
        messagebox.showinfo("Annonce", "Export terminé")
        update_espion(self.code_dossier, self.base, self.select_action)
        sys.exit()

    def get_reporting_minot_debut_fin(self, e):
        """
        Sélection du programmes qui sera lancé avec une période choisie
        """

        print("get_reporting_minot_debut_fin")
        self.show_traitement()
        select_debut = self.combobox_debut.get()
        select_debut = datetime.strptime(select_debut, "%B %Y")
        select_fin = self.combobox_fin.get()
        select_fin = datetime.strptime(select_fin, "%B %Y")
        actions.Mc4u_Minot_acd(self.code_dossier, select_debut, select_fin)
        
        # self.dispatch[self.select_action](self.code_dossier, select_debut, select_fin)
        # messagebox.showinfo("Annonce", "Export terminé")
        # update_espion(self.code_dossier, self.base, self.select_action)
        sys.exit()



import traceback
try:
    root = Tk()
    root.title('Opérateur Excel v2')
    root.wm_attributes("-topmost", 1)
    ressources = os.path.dirname(sys.argv[0])
    root.iconbitmap(os.path.join(ressources,"IMG/favicon.png"))
    app = Application(master=root)
    app.mainloop()
except Exception as e:
    showwarning(title = f'{e}', message=f"{traceback.format_exc()}")