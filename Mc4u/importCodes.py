# -*- coding: utf-8 -*-
import logging
import xlrd
import pprint
import sys

PP = pprint.PrettyPrinter(indent=4)


class importCodes(object):

    def __init__(self, chemXl):

        try:
            logging.info("Ouverture de {}".format(chemXl))
            wb = xlrd.open_workbook(chemXl, encoding_override="cp1252")
        except IOError as e:
            logging.error(e)
            sys.exit()

        ws = wb.sheet_by_index(0)
        self.dico = {}

        for row in range(1, ws.nrows):
            lib = ws.cell(row, 2).value
            self.dico.update(
                {
                    ws.cell(row, 0).value: {
                        "centre": ws.cell(row, 3).value,
                        "libelle": lib,
                    }
                }
            )


    def creaDic(self):
        """
        Return la config P&L
        """

        return self.dico



if __name__ == "__main__":

    xl = "./Codes Analytiques.xlsx"
    myobj = importCodes(xl)
    data = myobj.creaDic()
    PP.pprint(myobj.creaDic())
