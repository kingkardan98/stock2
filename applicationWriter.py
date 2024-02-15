# Structure of an Excel line:
# | code | quant | prz_blu | prz | desc | appl |
# ==============================================
# |  1   |   2   |    3    |  4  |  5   |  6   |

# For writer functions, all indexes must be considered Pythonic, so subtract 1 from the structure of the Excel line.

def applWriter(applFile, data):
    with open(applFile, 'w+') as outfile:
        outfile.write('CODICE                  ||    DESCRIZIONE                                           ||    APPLICAZIONI                                                                            ||\n')
        outfile.write('='*180 + '\n')
        for line in data:
            # Check for None, so the program doesn't break.
            # When everything is filled out, it shouldn't matter.
            if line[4] == None:
                line[4] = ''
            if line[5] == None:
                line[5] = ''

            # Check for empty price and don't write
            if line[3] == '//':
                continue
        
            while len(line[0]) < 20:
                line[0] += ' '
            while len(line[4]) < 50:
                line[4] += ' '
            while len(line[5]) < 20:
                line[5] += ' '
            outfile.write(line[0].upper().replace("'", '') + '    ||    ' + line[4].upper() + '    ||    ' + line[5].upper() + '\n')
            outfile.write('_'*180 + '\n')