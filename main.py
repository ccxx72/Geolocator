from geopy.geocoders import Photon
from openpyxl import load_workbook


def main():
    geolocator = Photon()

    wb = load_workbook(filename='File.xlsx')
    sheet = wb['Foglio1']
    row_count = sheet.max_row
    column_count = sheet.max_column
    x = 0
    for i in range(2, row_count):
        # if sheet.cell(row=i, column=25).value == '0':
        x = x + 1
        if x == 10:
            wb.save("File.xlsx")
            x = 0
        indirizzo = str(sheet.cell(row=i, column=5).value) + ' ' + str(sheet.cell(row=i, column=7).value) + ' ' + str(sheet.cell(row=i, column=6).value) + ' ITALIA'
        print(x, indirizzo)
        print(str(sheet.cell(row=i, column=23).value), str(sheet.cell(row=i, column=24).value))
        try:
            location = geolocator.geocode(indirizzo, timeout=10)
            sheet.cell(row=i, column=23).value = location.latitude
            sheet.cell(row=i, column=24).value = location.longitude
            sheet.cell(row=i, column=25).value = ''
            print((location.latitude, location.longitude))
        except:
            location = geolocator.geocode(str(sheet.cell(row=i, column=6).value) + ' ' + str(sheet.cell(row=i, column=7).value) + ' ITALIA', timeout=10)
            sheet.cell(row=i, column=23).value = location.latitude
            sheet.cell(row=i, column=24).value = location.longitude
            sheet.cell(row=i, column=25).value = '0'
            print('NP')
        else:
            sheet.cell(row=i, column=23).value = ''
            sheet.cell(row=i, column=24).value = ''
            sheet.cell(row=i, column=25).value = '-1'
    wb.save("File.xlsx")


if __name__ == "__main__":
    main()