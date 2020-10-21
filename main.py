from geopy.geocoders import Photon
from openpyxl import load_workbook

LAT_COLUMN = 23
LON_COLUMN = 24
PRECISION_COLUMN = 25


def main():
    geolocator = Photon()

    wb = load_workbook(filename='File.xlsx')
    sheets = wb.sheetnames
    sheet = wb[sheets[0]]  # selecting always the first sheet
    row_count = sheet.max_row
    column_count = sheet.max_column
    for i in range(2, row_count):
        if i % 10 == 0:
            wb.save("File.xlsx")
        indirizzo = str(sheet.cell(row=i, column=5).value) + ' ' + \
                    str(sheet.cell(row=i, column=7).value) + ' ' + str(sheet.cell(row=i, column=6).value) + ' ITALIA'
        print(i, indirizzo)
        print(str(sheet.cell(row=i, column=23).value), str(sheet.cell(row=i, column=24).value))
        try:
            location = geolocator.geocode(indirizzo, timeout=10)
            sheet.cell(row=i, column=LAT_COLUMN).value = location.latitude
            sheet.cell(row=i, column=LON_COLUMN).value = location.longitude
            sheet.cell(row=i, column=PRECISION_COLUMN).value = ''
            print((location.latitude, location.longitude))
        except:
            location = geolocator.geocode(str(sheet.cell(row=i, column=6).value) + ' ' +
                                          str(sheet.cell(row=i, column=7).value) + ' ITALIA', timeout=10)
            sheet.cell(row=i, column=LAT_COLUMN).value = location.latitude
            sheet.cell(row=i, column=LON_COLUMN).value = location.longitude
            sheet.cell(row=i, column=PRECISION_COLUMN).value = '0'
            print('NP')
        else:
            sheet.cell(row=i, column=LAT_COLUMN).value = ''
            sheet.cell(row=i, column=LON_COLUMN).value = ''
            sheet.cell(row=i, column=PRECISION_COLUMN).value = '-1'

    wb.save("File.xlsx")


if __name__ == "__main__":
    main()
