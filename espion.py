import getpass
import logging
from collections import OrderedDict
from datetime import datetime
from postgreagent import PostgreAgent


def update_espion(dossier = "", base = "", operation = ""):

    conf = OrderedDict(
        [
            ('host', '10.0.0.17'), 
            ('user', 'admin'), 
            ('password', 'Zabayo@@'), 
            ('port', '5432'), 
            ('dbname', 'outils')
            ]
        )

    horodat = datetime.now()
    collab = getpass.getuser()
    table = "operateur_xl"
    values = [collab, horodat, dossier, base, operation]

    sql = """
    INSERT INTO operateur_xl (collab, horodat, code_client, base, operation) 
    VALUES (%s, %s, %s, %s, %s);
    """
    try:
        with PostgreAgent(conf) as db:
            if db.connection:
                if db.table_exists(table):
                    logging.debug(f"table {table} exists")        
                    db.cursor.execute(sql, values)
        return True
    except:
        return False

# if __name__ == "__main__":
#     import logging
#     update_espion("FORM05", "abc")

