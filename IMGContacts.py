import sys
import xlrd
import random

excelfile = input('Enter the path of your file relative to the dir in which this script resides(eg., excel/myexcelfile.xlsx)')

try:
    book = xlrd.open_workbook(excelfile)
except:
    print('The path you specified is incorrect')
    sys.exit(1)

print('This Excel File contains following{} worksheets:'.format(book.nsheets))

c=0
for sheet in book.sheet_names():
    print(c, sheet)
    c=c+1

sheetno = input('Enter sheet no. you want to import to vcf')
sheetno = int(sheetno)
try:
    sh = book.sheet_by_index(sheetno)
except:
    print('Sheet Doesn\'t exist')
    sys.exit(1)

print('Worksheet {} selected'.format(sh.name))

year = input('In which year are you(1/2/3/4/5)')
year = int(year)

exportedfile = 'ExportedContact_PG_'+str(random.randint(100,1001))+'.vcf'
file = open(exportedfile, 'w')
vcard = "hi"
Name = "PG"
Firstname = "PG"
counter=1
for rx in range(1,sh.nrows):
    yeardiff = int(sh.cell_value(rowx=rx, colx=3)) - year
    if(yeardiff > 0):
        yeardiff = '+'+str(yeardiff)
    elif(yeardiff==0):
        yeardiff = ''
    else:
        yeardiff = yeardiff

    Name=';{};{} {};;'.format(sh.cell_value(rowx=rx, colx=1), yeardiff, sh.cell_value(rowx=rx, colx=4))
    Firstname='{} {} {};;'.format(sh.cell_value(rowx=rx, colx=1), yeardiff, sh.cell_value(rowx=rx, colx=4))
    vcard="BEGIN:VCARD\nVERSION:3.0\nN:{}\nFN:{}\nTEL;TYPE=CELL:{}\nTEL;TYPE=CELL:{}\nEMAIL;TYPE=WORK:{}\nADR;TYPE=HOME:;{};;;;\nORG:IMG\nEND:VCARD\n".format(Name, Firstname, int(sh.cell_value(rowx=rx, colx=5)), int(sh.cell_value(rowx=rx, colx=5)), sh.cell_value(rowx=rx, colx=7), sh.cell_value(rowx=rx, colx=6))
    file.write(vcard)
    print("{}, Sucessfully Entered({})".format(sh.cell_value(rowx=rx, colx=1), counter))
    counter=counter+1
file.close()
print('File exported as {}'.format(exportedfile))
print('Have A Nice Day!!')
