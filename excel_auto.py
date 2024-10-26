import openpyxl

#Excelを開く
wb = openpyxl.load_workbook('./サンプル.xlsx')

#Excelのシートを選択する
sheet = wb['Sheet1']

#Excelシートのセルに文字を入力する
sheet.cell(row=1, column=1).value = "word"
#value = sheet['A1'].value
#value = sheet['A1'].valueでもOK

#Excelシートのセルの文字を出力する
ret = sheet.cell(row=2, column=1).value
print(ret)

#Excelシートのセルのデータを合計する
#合計をセルおよびプロンプトに出力する。
sum = 0
for row in range(3, 6):
    sum += sheet.cell(row=row, column=1).value
sheet['B2'].value = sum
print(sum)

#Excelを保存する
wb.save('./サンプル.xlsx')

#Excelを閉じる
wb.close()