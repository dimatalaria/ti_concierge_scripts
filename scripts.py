import openpyxl

keynect = openpyxl.load_workbook("keystrom.xlsx")

sheet = keynect.active

value_key = sheet["A1"].value

cell_key = sheet.cell(row=1, column=1).value
sps = []
for row in sheet.iter_rows(min_row=1, max_row=8, values_only=True):
    json = {"CardNo":row[0],"UserID":row[1],"UserName":"1","VTOPosition":"","CardType":0,"CardStatus":0}
    sps.append(json)

file = open("key.txt", "w")
file.write(str(sps))
file.close()
print(sps)