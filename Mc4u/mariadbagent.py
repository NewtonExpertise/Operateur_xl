from email import header
import mariadb
import logging
from tkinter import messagebox

logging.basicConfig( level=logging.DEBUG)

class Mariadb():
    """
    Pilotage d'une connexion a une db postgresql
    les paramètres sont passés dans le fichier .ini
    """
    def __init__(self, db_locate):
        self.connection = ""
        self.cursor = ""
        if db_locate == "baku":
            self.user="root"
            self.password="acdpass"
            self.host="10.0.0.124"
            self.port=3306
        if db_locate == "rome":
            self.user="sysadmin"
            self.password="ShoQue8Xaph6Jew3"
            self.host="rome.axe.lan"
            self.port=3306



    def _connect(self):
        try:
            self.connection = mariadb.connect(user=self.user,
                                            password = self.password,
                                            host = self.host,
                                            port = self.port,
                                            )
            self.cursor = self.connection.cursor()
        except Exception as e:
            messagebox.showerror(title = "Erreur de connection", message=f"La tentative de connection à la base de donnée à échoué.\n\nVeuillez recommencer l'opération.\n\nCode erreur : \n\n{e}")
            print(self.database)

        
    def _close(self):
        if self.connection:
            try:
                self.connection.commit()
                self.connection.close()
            except :
                self.connection.close()

    def __enter__(self):
        """Pour context manager"""
        self._connect()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self._close()

    def execute(self, sql):
        try:
            self.cursor.execute(sql)
        except (Exception, mariadb.Error) as error:
            print(error)
            print(sql)
        data = self.cursor.fetchall()
        return data

    def execute_with_header(self, sql):

        try:
            self.cursor.execute(sql)
        except (Exception, mariadb.Error) as error:
            print(error)
            print(sql)

        description_colonnes = self.cursor.description
        header_columns = []
        if description_colonnes:
            for hc in description_colonnes:
                header_columns.append(hc[0])
        data = self.cursor.fetchall()
        data = [tuple(header_columns)]+data

        return data

    def insert(self, sql):
        try:
            self.cursor.execute(sql)
            return True
        except (Exception, mariadb.Error) as error:
            print(error)
            print(sql)
            return False




if __name__ == '__main__':

    sql = """SELECT * FROM compte"""
    with Mariadb() as Mdb:
        # print(Mdb.connection.user)
        # c = Mdb.list_table()

        print(db)
        # print(db[3][0])
        # print(Mdb.connection.database)
        # Mdb.connection.database = db[3][0]
        # print(Mdb.connection.database)
        # print(Mdb.query(sql))