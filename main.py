from voortgangsdoc import SourceXls
from rapportagedoc import TargetXls
import constant
from tkinter import *
import tkinter


def vraagpad():
    root = Tk()
    root.directory = tkinter.filedialog.askdirectory(title="Selecteer map voor de leerling documenten")
    return root.directory


mysrc = SourceXls()
# De lijst van leerlingen wordt uit de tab gehaald met de naam van de klas
mysrc.initleerlingen()
mysrc.actleerlingindex = -1


padleerlingdocs = vraagpad()

# Alle leerlingen aflopen
while mysrc.setnextleerling():
    firstvak = True
    print(mysrc.namen[mysrc.actleerlingindex])
    # per leerling een aparte sheet maken of van een bestaande openen
    # vervolgens een nieuwe tab maken of die actief maken
    # De rij overzetten met zijn voortgang in de desbetreffende tab

    # zet de mysrc.actsheetindex op de tab met als naam de klas
    mysrc.initsheetindex()

    # ga dan vervolgens alle sheets na de actsheetindex af om de voortgamg op te zoeken
    hrange = []
    while mysrc.setnextsheet():
        # leerlingvoortgangsdocument wordt geopend of aangemaakt
        if firstvak:
            voortgangLeerling = TargetXls(mysrc.klas, mysrc.namen[mysrc.actleerlingindex],
                                          mysrc.leerlingen[mysrc.actleerlingindex], padleerlingdocs)

        else:
            print(".")

        # maak een nieuwe tab na de laatste tab en maak het actief
        voortgangLeerling.maaktab(mysrc.sheet.title)

        # kopieer de header range naar de leerlingvoortgangs document
        hrange = mysrc.kopieerrange(constant.STARTCOLHEADER, constant.STARTROWHEADER, constant.AANTALCOLSVOORTGANG, constant.ENDROWHEADER)
        leerlingrij = mysrc.zoekleerlingrij()
        leerlingrange = mysrc.kopieerrange(constant.STARTCOLRESULTS, leerlingrij, constant.AANTALCOLSVOORTGANG, leerlingrij)
        voortgangLeerling.pasteRange(constant.STARTCOLHEADER, constant.STARTROWHEADER, constant.AANTALCOLSVOORTGANG, constant.ENDROWHEADER, hrange, mysrc.sheet.title, True)

        # plak de resultatenrij in de vrije rij met een volgend weeknummer
        vrijerij = voortgangLeerling.zoekvrijerij(mysrc.sheet.title)
        voortgangLeerling.pasteRange(constant.STARTCOLRESULTS, vrijerij, constant.AANTALCOLSVOORTGANG, vrijerij, leerlingrange,
                                     mysrc.sheet.title, False)

        hrange = []
        # sla de aanpassingen op
        voortgangLeerling.save()
    firstvak = False

    voortgangLeerling.save()
    voortgangLeerling.sluitdoc()
    del voortgangLeerling

