import time, sys
sleep_time = 0

# Structure of an Excel line:
# | code | quant | prz_blu | prz | desc | appl |
# ==============================================
# |  1   |   2   |    3    |  4  |  5   |  6   |

def elaboraCosti(data):
    print("Inizio elaborazione costi articoli.\n")
    totalCostBluEl = 0
    totalCostBlu = 0
    totalCostEl = 0
    totalCost = 0
    for line in data:
        # Converting the quantity to a string prevents a possible error.
        quantity = str(line[1]).strip().split(' ')

        if isInvalid(line[2]):
            costBlu = 0
        else:
            costBlu = line[2]

        if isInvalid(line[3]):
            cost = 0
        else:
            cost = line[3]

        if 'el' in quantity:
            # Since 'el' is in quantity, we need to take into account only the number and convert it.
            try:
                totalCostBluEl += int(quantity[0]) * costBlu
                totalCostEl += int(quantity[0]) * cost
            except ValueError:
                print(quantity[0], costBlu)
                sys.exit()
            except TypeError:
                print(line[0], costBlu)
                sys.exit()
        try:
            totalCostBlu += int(quantity[0]) * costBlu
            totalCost += int(quantity[0]) * cost
        except ValueError:
            pass
        except TypeError:
            print(line)
            print("Errore di tipo")
            sys.exit()
        
        i = data.index(line)
        if i % 100 == 0:
            print("Elementi elaborati: {}".format(i))
            time.sleep(sleep_time)

    print("\nFine elaborazione costi articoli in eliminazione. Elementi elaborati: {}\n".format(i+1))    
    return totalCostBlu, totalCost, totalCostBluEl, totalCostEl

def isInvalid(value):
    if value == '' or value == None or value == '//':
        return True
    return False