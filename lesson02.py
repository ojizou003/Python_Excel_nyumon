import openpyxl

# 【演習問題】
# シート名が(≫raw, data)のExcelファイルを、5枚作成しましょう。
# また、シート≫rawについては、色を赤に変更しておきましょう。

for i in range(5):
    file_name = 'Excelfiles/ensyuu{}.xlsx'.format(i)
    wb = openpyxl.Workbook()
    
    wb.active
    ws = wb['Sheet']
    ws.title = '>>raw'
    ws.sheet_properties.tabColor = 'ff0000'
    wb.create_sheet(title='data')
    wb.save(file_name)