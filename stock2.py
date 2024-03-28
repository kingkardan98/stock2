import sys, openpyxl, io, os
from dotenv import load_dotenv
from operator import itemgetter
from costUtil import *
from quantUtil import *
from codeUtil import *
from inventoryWriter import *
from applicationWriter import *
import msoffcrypto

def elabSuite(data):
    # Cost elaboration.
    costoTotBlu, costoTot, costoElBlu, costoEl = elaboraCosti(data)
    print(separator())
    print("\n\n")

    # Quantity elaboration
    quantEl, quantTot = elaboraQuantita(data)
    print(separator())
    print("\n\n")

    # Code elaboration
    codiciEl, codiciTot = elaboraCodici(data)
    print(separator())
    print("\n\n")

    # Stock generation.
    stockFile = "inv.txt"
    invWriter(stockFile, data)

    # Application file generation
    applFile = "appl.txt"
    applWriter(applFile, data)

    # Tactical return right now, so the functions below that are not implemented do not execute.
    return costoElBlu, costoEl, costoTotBlu, costoTot, quantEl, quantTot, codiciEl, codiciTot, stockFile, applFile

def separator():
    return "="*30

def main():
    filename = "InventarioPiaggio.xlsx"

    try:
        decrypted_workbook = io.BytesIO()

        with open(filename, 'rb') as file:
            load_dotenv()
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=os.getenv('XL_PASSWORD'))
            office_file.decrypt(decrypted_workbook)

    # `filename` can also be a file-like object.
        workbook = openpyxl.load_workbook(filename=decrypted_workbook)
        ws = workbook.active
    except FileNotFoundError:
        print("File non trovato.")
        sys.exit()

    print("File aperto. Inizio flusso in entrata.")
    print(separator())
    data = []

    # Let's start at two, since the first row is just the header.
    i = 1
    while True:
        if (not ws.cell(row=i, column=1).value):
            break

        # Don't include the header row
        if (ws.cell(row=i, column=1).value == 'code'):
            i = i + 1
        else:
            line = [ws.cell(row=i, column=j).value for j in range(1,7)]
            line[0] = str(line[0]).strip()
            data.append(line)
            i = i + 1

        if i % 100 == 0:
            print("Elaborati: {} elementi".format(i))

    print(separator())
    print("Fine flusso in entrata. Elementi elaborati: {}\n\n".format(i-2))  # Remove the start (since Excel is 1-indexed) and the header element that wasn't added.
    data = sorted(data, key=itemgetter(0)) # Sort the data code-wise.
    
    costoElBlu, costoEl, costoTotBlu, costoTot, quantEl, quantTot, codiciEl, codiciTot, stockFile, applFile = elabSuite(data)

    print("Fine elaborazione.")
    print(separator())
    print("\n\n")

    report = ''
    report += "COSTI\n"
    report += separator() + "\n"
    report += "Il valore dell'inventario ad eliminarsi (ivato) è: €{:.2f}\n".format(costoEl)
    report += "Il valore dell'inventario ad eliminarsi (netto) è: €{:.2f}\n".format(costoElBlu)
    report += "Il valore dell'intero inventario (ivato) è: €{:.2f}\n".format(costoTot)
    report += "Il valore dell'intero inventario (netto) è: €{:.2f}\n".format(costoTotBlu)
    report += "La percentuale di inventario (relativa al valore) ad eliminarsi è: {:.2f}%\n".format(costoEl / costoTot * 100)
    report += separator() + "\n"
    report += "\n\n\n"

    report += "QUANTITÀ\n"
    report += separator() + "\n"
    report += "Su {} pezzi fisici presenti in inventario, ne vanno eliminati {}\n".format(quantTot, quantEl)
    report += "La percentuale di inventario (relativa alle quantità) ad eliminarsi è: {:.2f}%\n".format(quantEl / quantTot * 100)
    report += separator() + "\n"
    report += "\n\n\n"

    report += "CODICI\n"
    report += separator() + "\n"
    report += "Su {} codici totali presenti in inventario, ne vanno eliminati {}\n".format(codiciTot, codiciEl)
    report += "La percentuale di inventario (relativa alla quantità di codici) ad eliminarsi è: {:.2f}%\n".format(codiciEl / codiciTot * 100)
    report += separator() + "\n"
    report += "\n\n\n"

    report += "ULTERIORI ELABORAZIONI\n"
    report += separator() + "\n"
    report += "Nel file {} si trova l'inventario Piaggio.\n".format(stockFile)
    report += "Nel file {} si trova una lista di ciascun codice, con la sua descrizione e la lista degli applicativi.\n".format(applFile)
    report += separator() + "\n"
    report += "Fine programma."

    with open('report.txt', 'w+') as reportFile:
        reportFile.write(report)
    print(report)
    sys.exit()

if __name__ == '__main__':
    main()