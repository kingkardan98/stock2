def invWriter(stockFile, data):
    with open(stockFile, 'w+') as outfile:
        outfile.write('CODICE                  ||    QUANT.\n')
        outfile.write('='*30 + '\n')
        for line in data:
            while len(line[0]) < 20:
                line[0] += ' '
            outfile.write(line[0].upper().replace("'", '') + '    ||    ' + str(line[1]) + '\n')
            outfile.write('_'*30 + '\n')