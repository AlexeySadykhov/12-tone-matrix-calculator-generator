from random import shuffle
from numpy import diff
import xlsxwriter as xlsx

prima = [ ]

scale = {'c': 6000, 'cis': 6100, 'd': 6200, 'dis': 6300,
         'e': 6400, 'f': 6500, 'fis': 6600, 'g': 6700,
         'gis': 6800, 'a': 6900, 'ais': 7000, 'h': 7100}

def get_key(scale, value):
    for k, v in scale.items():
        if v == value:
            return k

for x in scale:
    x = scale.get(x)
    prima.append(x)
shuffle(prima)

intervals = diff(prima)

workbook = xlsx.Workbook('Random 12-tone matrix.xlsx')
worksheet = workbook.add_worksheet()

for i, item in enumerate(prima):
    worksheet.write(0, i, get_key(scale, item))

for i, item in enumerate(prima):
    item = item
    for j, x in enumerate(intervals):
        item = item - x
        if item < 6000:
            while item < 6000:
                item = item + 1200
                if item >= 6000:
                    worksheet.write(j+1, i, get_key(scale, item))
        else:
            if item > 7100:
                while item > 7100:
                    item = item - 1200
                    if item <= 7100:
                        worksheet.write(j+1, i, get_key(scale, item))
            else:
                worksheet.write(j+1, i, get_key(scale, item))
    
workbook.close()

answer = input('Matrix has been generated. Do you want to save txt-file in mc? (y/n)')
if answer == 'y':
    seria = ' '.join(str(x) for x in prima)
    file = open('random_seria.txt', 'w')
    file.write(seria)
    file.close()
    print('Done')
else:
    print('Done')
