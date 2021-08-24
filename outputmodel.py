class OutputModel:

    def __init__(self):

        self.filnamn_clean = 0
        self.filplats = 0
        self.kontroll_per = 0
        self.kontroll_anl = 0
        self.resultat = 0
        self.periodisering = 0
        self.kommentar = 0
        self.konto_per = 0
        self.belopp = 0
        self.kontroll_ladd = 0
        self.kst = 0
        self.projekt = 0
        self.anv_perkonto_belopp = 0
        self.upprattad_av = 0
        self.intakt = 1 #använder 1 på denna för att en kontroll utgörs av att se om intakt + kostand = 0. (int_kost_per_noll() i kontrollera)
        self.kostnad = 1 #använder 1 på denna för att en kontroll utgörs av att se om intakt + kostand = 0 (int_kost_per_noll() i kontrollera)
        self.int_kost_per_noll = 0
        self.slutdatum = 0







