import time
sleep_time = 0

# Structure of an Excel line:
# | code | quant | prz_blu | prz | desc | appl |
# ==============================================
# |  1   |   2   |    3    |  4  |  5   |  6   |

def elaboraCodici(data):
    print("Inizio elaborazione codici articoli.\n")
    codEl = 0
    codTot = 0
    for line in data:
        quantity = str(line[1]).strip().split(' ')
        if 'el' in quantity:
            codEl += 1
        codTot += 1

        i = data.index(line)
        if i % 100 == 0:
            print("Elementi elaborati: {}".format(i))
            time.sleep(sleep_time)
    print("\nFine elaborazione codici articoli. Elementi elaborati: {}\n".format(i+1))
    return codEl, codTot