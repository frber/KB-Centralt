import pandas as pd
import openpyxl
import math
import datetime
import os
from openpyxl.styles import Font



class SlaIhop:

    def __init__(self, wb3):
        self.wb3 = wb3
        #self.filvag_slaihop = "Docs\\Agressodata.xlsx"
        #self.wb = openpyxl.load_workbook(self.filvag_slaihop, data_only=False)

    #Används inte just nu
    def clear(self):
        for sheet in self.wb.worksheets:
            sheet.delete_cols(10, 20)
            #self.wb.save("IHOPSLAG.xlsx")



    def kontroll_browserdata(self):

        idag = datetime.datetime.now()

        # Kontrollera rätt tecken.
        for sheets in self.wb3.worksheets:
            if sheets == self.wb3["Klistra in Agressodata här"]:
                continue
            else:
                for row in sheets['A1:A1000']:
                    for cell in row:
                        if cell.value != None and isinstance(cell.value, int):
                            konto = cell.value
                            belopp = cell.offset(column=5).value

                            #Kontroll att periodisering är bokat åt rätt håll (Rätt tecken).
                            if isinstance(belopp, int) or isinstance(belopp, float):
                                if konto > 1000 and konto < 2000 and belopp < 0:
                                    cell.offset(column=13).value = "Kontrollera"

                            # Kontrolelra att slutdatum inte har passerat.
                            slutdatum = cell.offset(column=7).value
                            if isinstance(slutdatum, datetime.datetime):
                                if slutdatum < idag:
                                    cell.offset(column=14).value = "Kontrollera"







    def con(self):
        pd.set_option('display.width', 320)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)
        df_output = pd.read_excel('Docs\\Berperdata.xlsx', sheet_name='Avvikelser2')


        output_projektnr = df_output['PROJEKT']
        output_konto = df_output['PERIODISERINGSKONTO']
        output_lank = df_output['LÄNK']
        output_per_k = df_output['KONTROLL PER']
        output_anl_k = df_output['KONTROLL ANL']
        output_kommentar = df_output['KONTROLL KOMMENTAR']
        output_ekonom = df_output['BERPER UPPRÄTTAD AV']
        output_belopp = df_output['BELOPP PERIODISERING']
        output_faktisk_kommentar = df_output['KOMMENTAR']
        output_int_kost = df_output['INT KOST']


        rakna = 0

        for sheets in self.wb3.worksheets:
            if sheets == self.wb3["Klistra in Agressodata här"]:
                continue
            else:
                for row in sheets['A1:A1000']:
                    for cell in row:
                        if cell.value != None:
                            browser_konto = cell.value
                            browser_projektnr = cell.offset(column=3).value
                            browser_slutdatum = cell.offset(column=13).value

                            #Korsar ihop Agressodata med Berperdata.
                            for x in df_output.index:
                                if output_projektnr[x] == browser_projektnr and output_konto[x] == browser_konto and output_belopp[x] != 0:
                                    #rakna += 1
                                    #print(rakna)
                                    if output_per_k[x] == "Kontrollera":
                                        cell.offset(column=10).value = "Kontrollera"
                                    if output_anl_k[x] == "Kontrollera":
                                        cell.offset(column=11).value = "Kontrollera"
                                    if output_kommentar[x] == "Kontrollera":
                                        cell.offset(column=12).value = "Kontrollera"
                                    if output_int_kost[x] == "Kontrollera":
                                        cell.offset(column=15).value = "Kontrollera"
                                    cell.offset(column=19).value = output_ekonom[x]
                                    cell.offset(column=21).value = output_faktisk_kommentar[x]
                                    lank = cell.offset(column=20)
                                    lank.value = "Öppna berper"
                                    lank.hyperlink = output_lank[x]
                                    lank.style = "Hyperlink"

    def slutkontroll(self):
        #Om det inte finns något markerat med "Kontrollera" i dokumentet, skriv in "OK" i en kolumn.
        for sheets in self.wb3.worksheets:
            if sheets == self.wb3["Klistra in Agressodata här"]:
                continue
            else:
                for row in sheets['A1:A5000']:
                    for cell in row:
                        if isinstance(cell.value, int) and cell.offset(column=2) != None:
                            if cell.offset(column=10).value == None and cell.offset(column=11).value == None and cell.offset(column=12).value == None and cell.offset(column=13).value == None and cell.offset(column=14).value == None and cell.offset(column=15).value == None:
                                cell.offset(column=9).value = "OK"

    #Används inte just nu
    def lank_losning(self):
        #Om slutdatum har passerat baserat på agressodatan och motsvarande berper är trasig och har laddat 0 i slutdatum så länkas inte den berper.
        #Kolla därför om slutdatum har passerat och om det finns en berper i "Alla Berper"-fliken, i sådana fall länka till den berper.
        #Programmet har i slutet då kollat om det finns berper i avvikelser och alla berper, om inte finns skriv något i stil med "berper saknas", eftersom alla kontroller borde då vara gjorda.

        df_output_allaberper = pd.read_excel('Docs\\Berperdata.xlsx', sheet_name='Alla Berper')

        output_projektnr = df_output_allaberper['PROJEKT']
        output_kommentar = df_output_allaberper['KOMMENTAR']
        output_upp_av = df_output_allaberper['UPPRÄTTAD AV']
        output_lank = df_output_allaberper['LÄNK']

        for sheets in self.wb3.worksheets:
            if sheets == self.wb3["Klistra in Agressodata här"]:
                continue
            else:
                for row in sheets['O1:O5000']:
                    for cell in row:
                        if cell.value == "Kontrollera" and cell.offset(column=6).value == None:
                            upp_av = cell.offset(column=5)
                            lank = cell.offset(column=6)
                            kommentar = cell.offset(column=7)
                            projeknr = cell.offset(column=-11).value
                            for x in df_output_allaberper.index:
                                if output_projektnr[x] == projeknr:
                                    #print(output_projektnr[x])
                                    print(output_lank[x])
                                    lank.value = "Öppna berper"
                                    lank.hyperlink = output_lank[x]
                                    lank.style = "Hyperlink"
                                    kommentar.value = output_kommentar[x]
                                    upp_av.value = output_upp_av[x]
                        if cell.value == "Kontrollera" and cell.offset(column=6).value == None:
                            lank = cell.offset(column=6)
                            lank.value = "BERPER EJ KONTROLLERAD"

            for row in sheets['N1:N5000']:
                for cell in row:
                    if cell.value == "Kontrollera" and cell.offset(column=7).value == None:
                        upp_av = cell.offset(column=6)
                        lank = cell.offset(column=7)
                        kommentar = cell.offset(column=8)
                        projeknr = cell.offset(column=-10).value
                        for x in df_output_allaberper.index:
                            if output_projektnr[x] == projeknr:
                                #print(output_projektnr[x])
                                print(output_lank[x])
                                lank.value = "Öppna berper"
                                lank.hyperlink = output_lank[x]
                                lank.style = "Hyperlink"
                                kommentar.value = output_kommentar[x]
                                upp_av.value = output_upp_av[x]
                    if cell.value == "Kontrollera" and cell.offset(column=7).value == None:
                        lank = cell.offset(column=7)
                        lank.value = "BERPER EJ KONTROLLERAD"

    def lagg_in_lankar(self):

        df_output_allaberper = pd.read_excel('Docs\\Berperdata.xlsx', sheet_name='Alla Berper')

        output_projektnr = df_output_allaberper['PROJEKT']
        output_kommentar = df_output_allaberper['KOMMENTAR']
        output_upp_av = df_output_allaberper['UPPRÄTTAD AV']
        output_lank = df_output_allaberper['LÄNK']


        for sheets in self.wb3.worksheets:
            if sheets == self.wb3["Klistra in Agressodata här"]:
                continue
            else:
                for row in sheets['D1:D5000']:
                    for cell in row:
                        if cell.value != None:
                            projektnr = cell.value
                            upp_av = cell.offset(column=16)
                            lank = cell.offset(column=17)
                            kommentar = cell.offset(column=18)
                            konto = cell.offset(column=-3)
                            for x in df_output_allaberper.index:
                                if output_projektnr[x] == projektnr:
                                    # print(output_projektnr[x])
                                    #print(output_lank[x])
                                    lank.value = "Öppna berper"
                                    lank.hyperlink = output_lank[x]
                                    lank.style = "Hyperlink"
                                    kommentar.value = output_kommentar[x]
                                    upp_av.value = output_upp_av[x]
                                if isinstance(konto.value, int) and projektnr != None and lank.value == None:
                                    lank.value = "BERPER EJ KONTROLLERAD"




    def spara(self):
        try:
            self.wb.save(self.filvag_slaihop)
            self.wb.close()
        except PermissionError:
            return True



    def oppna(self):
        os.startfile(self.filvag_slaihop)









