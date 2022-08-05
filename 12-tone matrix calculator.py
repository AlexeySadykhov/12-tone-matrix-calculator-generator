import random
import numpy as np
import xlsxwriter as xlsx

scale = {
    'c': 6000, 'cis': 6100, 'des': 6100,
    'd': 6200, 'dis': 6300, 'es': 6300,
    'e': 6400, 'f': 6500, 'fis': 6600,
    'ges': 6600, 'g': 6700, 'gis': 6800,
    'as': 6800, 'a': 6900, 'ais': 7000,
    'b': 7000, 'h': 7100
}


def get_key(dct, value):
    for k, v in dct.items():
        if v == value:
            return k


def calculate_prima(sr, sc):
    p = []
    for x in sr:
        if x in sc.keys():
            p.append(sc.get(x))
        else:
            print('Error')
            print('There is no', x, 'pitch')
            exit(1)

    if len(p) != len(set(p)):
        ans1 = input("""Your list has duplicates. 
        Do you want to continue? (y/n)""")
        if ans1 == 'y':
            print('OK')
        else:
            exit(0)
    return p


seria = input("""Enter seria splitting notes by space or keep this field empty. 
In the second case seria will be generated randomly:""").split()
if not seria:
    prima = list({x for x in scale.values()})
    random.shuffle(prima)
else:
    prima = calculate_prima(seria, scale)

intervals = np.diff(prima)

filename = input('Enter file name to save:')
workbook = xlsx.Workbook(filename + '.xlsx')
worksheet = workbook.add_worksheet('matrix')

for i, item in enumerate(prima):
    item = item
    worksheet.write(0, i, get_key(scale, item))
    for j, x in enumerate(intervals):
        item = item - x
        if item < 6000:
            while item < 6000:
                item = item + 1200
                if item >= 6000:
                    worksheet.write(j + 1, i, get_key(scale, item))
        else:
            if item > 7100:
                while item > 7100:
                    item = item - 1200
                    if item <= 7100:
                        worksheet.write(j + 1, i, get_key(scale, item))
            else:
                worksheet.write(j + 1, i, get_key(scale, item))
workbook.close()

ans2 = input("""Matrix has been calculated. 
Do you want to save P-form in mc? (y/n)""")
if ans2 == 'y':
    seria_s = ' '.join(str(x) for x in prima)
    txt_filename = input('Enter file name to save:')
    file = open(txt_filename + '.txt', 'w')
    file.write(seria_s)
    file.close()
print('Done')
