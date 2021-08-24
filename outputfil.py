import openpyxl
import os
import shutil

class OutputFil:

    #outputfil = "Docs\\Berperdata.xlsx"
    #wb2 = openpyxl.load_workbook(outputfil, data_only=True)
    #ws_stat = wb2["Stat"]
    #ws_avvikelser = wb2["Avvikelser"]
    #ws_avvikelser2 = wb2["Avvikelser2"]
   # ws_kommentar = wb2["Alla Berper"]
   # ws_ejberper = wb2["Filer som ej bedöms vara berper"]
    #ws_trasiga = wb2['Excelfiler som ej kunde öppnas']
    #ws_forstora = wb2['För stora excelfiler']



    def nollstall(self):
        try:
            os.remove('Docs\\Berperdata.xlsx')
            shutil.copy('Docs\\Orginal\\Berperdata.xlsx', 'Docs\\Berperdata.xlsx')
        except PermissionError:
            return True
       # for sheet in self.wb2.worksheets:
            #sheet.delete_rows(2,2000)
            #try:
                #self.wb2.save(self.outputfil)
                #self.wb2.close()
            #except PermissionError:
                #return True

    def spara_stang(self):
        self.wb2.save(self.outputfil)
        self.wb2.close()

    def starta(self):
        os.startfile(self.outputfil)

