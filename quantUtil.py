import time
sleep_time = 0

# Structure of an Excel line:
# | code | quant | prz_blu | prz | desc | appl |
# ==============================================
# |  1   |   2   |    3    |  4  |  5   |  6   |

def elaboraQuantita(data):
    print("Inizio elaborazione quantità articoli.\n")
    quantEl = 0
    quantTot = 0
    for line in data:
        quantity = str(line[1]).strip().split(' ')
        if 'el' in quantity:
            quantEl += int(quantity[0])
        quantTot += int(quantity[0])

        i = data.index(line)
        if i % 100 == 0:
            print("Elementi elaborati: {}".format(i))
            time.sleep(sleep_time)
    print("\nFine elaborazione quantità articoli. Elementi elaborati: {}\n".format(i+1))
    return quantEl, quantTot