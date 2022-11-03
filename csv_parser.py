import openpyxl
from unicodedata import normalize

csv_path = 'C:\\Temp\\'

wb = openpyxl.load_workbook(csv_path + 'Parcel_CSV_Example.xlsx')
raw_sheet = wb.worksheets[0]
parsed_sheet = wb.worksheets[1]

values = []

for row in raw_sheet.values:
   for value in row:
    if value == 'NAME':
        continue
    else:
        values.append(normalize('NFKD',value))

parcel = {"ParcelID": '',"AltID": '',"Address": '',"Owner": '', "Acres": ''}
i = 2

for v in values:
    if "View: Report" in v:
        namecell = 'A' + str(i)
        addresscell = 'C' + str(i)
        numbercell = 'D' + str(i)
        streetcell = 'E' + str(i)

        owner = parcel["Owner"]
        address = parcel["Address"]
        add_list = parcel["Address"].split(' ',1)

        parcel = {"ParcelID": '',"AltID": '',"Address": '',"Owner": '', "Acres": ''}

        if add_list[0].isdigit() == False:
            continue

        parsed_sheet[namecell] = owner
        parsed_sheet[addresscell] = address
        parsed_sheet[numbercell] = add_list[0]
        parsed_sheet[streetcell] = add_list[1]
        i = i +1
        continue

    if "Parcel ID - " in v:
        parcel["ParcelID"] = v.replace('Parcel ID - ','')
        continue

    if "Alt Id - " in v:
        parcel["AltID"] = v.replace('Alt Id - ','')
        continue

    if "Address - " in v:
        parcel["Address"] = v.replace('Address - ','')
        continue

    if "Owner - " in v:
        parcel["Owner"] = v.replace('Owner - ','')
        continue

    if "Acres - " in v:
        parcel["Acres"] = v.replace('Acres - ','')
        continue

    parcel["Owner"] = parcel["Owner"] + ' ' + v

wb.save(csv_path + 'Parcel_CSV_Example_py.xlsx')