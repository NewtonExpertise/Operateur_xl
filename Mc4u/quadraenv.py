import logging
import os
import sys
import pyodbc


class QuadraSetEnv(object):
    def __init__(self, ipl_file):

        self.cpta = ""
        self.paie = ""
        self.gi = ""
        self.conn = ""
        self.cur = ""
        self.liste_cli = {}

        with open(ipl_file, "r") as f:
            lines = f.readlines()

        for line in lines:
            line = line.rstrip().replace("\\", "/")
            if "=" in line:
                key, item = line.split("=")[0:2]
                if key == "RACDATACPTA":
                    self.cpta = item
                elif key == "RACDATAPAIE":
                    self.paie = item
                elif key == "RACDATAGI":
                    self.gi = item

    def make_db_path(self, num_dossier, type_dossier):
        type_dossier = type_dossier.upper()
        num_dossier = num_dossier.upper()
        db_path = ""
        if (
            type_dossier == "DC"
            or type_dossier.startswith("DA")
            or type_dossier.startswith("DS")
        ):
            db_path = "{}{}/{}/qcompta.mdb".format(self.cpta, type_dossier, num_dossier)

        elif type_dossier == "PAIE":
            db_path = "{}{}/qpaie.mdb".format(self.paie, num_dossier)

        return db_path

    def gi_liste_clients(self):

        mdb_path = os.path.join(self.gi, "0000", "qgi.mdb")
        constr = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + mdb_path
        logging.info("openning qgi : {}".format(mdb_path))
        sql = """
            SELECT I.Code, I.Nom 
            FROM Intervenants I 
            INNER JOIN Clients C ON I.Code=C.Code 
            WHERE I.IsClient='1'
            """
        try:
            self.conn = pyodbc.connect(constr, autocommit=True)
            self.cur = self.conn.cursor()
            self.cur.execute(sql)
            data = list(self.cur)
            logging.info("connection OK")
        except pyodbc.Error:
            logging.error(
                ("erreur requete base {} \n {}".format(mdb_path, sys.exc_info()[1]))
            )
            return False
        
        for code, rs in data:
            self.liste_cli.update({code : {"rs" : rs}})


        return self.liste_cli


if __name__ == "__main__":
    import pprint

    pp = pprint.PrettyPrinter(indent=4)
    ipl = "C:/Users/nicolas/Documents/Pydio/mono.ipl"
    o = QuadraSetEnv(ipl)
    pp.pprint("---".join([o.cpta, o.paie, o.gi]))
    print(o.make_db_path("FORM05", "DC"))
    print(o.make_db_path("FORM05", "DA2017"))
    print(o.make_db_path("form05", "PAIE"))
    lst = o.gi_liste_clients()
    print(lst)
