from numpy import diff
import xlsxwriter as xlsx

seria_s = input('Enter seria splitting notes by space:')
seria_l = seria_s.split()
prima = []

scale = {
    'c': 6000, 'cis': 6100, 'des': 6100,
    'd': 6200, 'dis': 6300, 'es': 6300,
    'e': 6400, 'f': 6500, 'fis': 6600,
    'ges': 6600, 'g': 6700, 'gis': 6800,
    'as': 6800, 'a': 6900, 'ais': 7000,
    'b': 7000, 'h': 7100
}
keys = scale.keys()


def get_key(dct, value):
    for k, v in dct.items():
        if v == value:
            return k


for x in seria_l:
    if x in keys:
        prima.append(scale.get(x))
    else:
        print('Error')
        print('There is no', x, 'pitch')
        exit(1)

for x in prima:
    if prima.count(x) > 1:
        ans1 = input('Your list has duplicates. '
                     'Do you want to continue? (y/n)')
        if ans1 == 'y':
            print('OK')
        else:
            exit(0)
        break
    
intervals = diff(prima)

workbook = xlsx.Workbook('matrix.xlsx')
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

ans2 = input('Matrix has been calculated. '
             'Do you want to save P-form in mc? (y/n)')
if ans2 == 'y':
    seria = ' '.join(str(x) for x in prima)
    file = open('my_seria.txt', 'w')
    file.write(seria)
    file.close()
    print('Done')
else:
    print('Done')
