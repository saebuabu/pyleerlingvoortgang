from voortgangsdoc import SourceXls
from rapportagedoc import TargetXls

mysrc = SourceXls()
# De lijst van leerlingen wordt uit de tab gehaald met de naam van de klas
mysrc.initleerlingen()

mysrc.actleerlingindex = -1
# Alle leerlingen aflopen
while mysrc.setnextleerling():
    # actleerlingrij = mysrc.zoekleerlingrij()
    # print("rij " + str(actleerlingrij) + " " + mysrc.namen[mysrc.actleerlingindex])

    firstvak = True
    # per leerling een aparte sheet maken of van een bestaande openen
    # vervolgens een nieuwe tab maken of die actief maken
    # De rij overzetten met zijn voortgang in de desbetreffende tab

    # zet de mysrc.actsheetindex op de tab met als naam de klas
    mysrc.initsheetindex()

    # ga dan vervolgens alle sheets na de actsheetindex af om de voortgamg op te zoeken
    hrange = []
    while mysrc.setnextsheet():
        if firstvak:
            voortgangLeerling = TargetXls(mysrc.klas, mysrc.namen[mysrc.actleerlingindex],
                                          mysrc.leerlingen[mysrc.actleerlingindex])

        # maak een nieuwe tab na de laatste tab en maak het actief
        # print(mysrc.sheet)
        else:
            print(".")

        #print(mysrc.sheet.title)
        voortgangLeerling.maaktab(mysrc.sheet.title)

        # kopieer de header range naar de leerlingvoortgangs document
        hrange = mysrc.kopieerrange(1, 1, 40, 3)
        # plak het in de leer

        #voortgangLeerling.kopieerheader(mysrc.getheaderrange())

        voortgangLeerling.pasteRange(1, 1, 40, 3, hrange, mysrc.sheet.title)
        hrange = []
        #sla de aanpassingen op
        voortgangLeerling.save()
    firstvak = False

    voortgangLeerling.save()
    voortgangLeerling.sluitdoc()
    del voortgangLeerling
