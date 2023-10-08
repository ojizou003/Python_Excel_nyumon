# 【Python x Excel 超入門】 はやたす youtube

2023/10/8 ~ 10/8

## 環境設定

anacondaインストール  
仮想環境作成 --> excel_lesson
openpyxl インストール

import openpyxl

## ファイルの作成・読み込み・保存

Excelファイルの作成  
wb = openpyxl.Workbook()  
wb.active  
wb.save('ファイル名.xlsx')  

Excelファイルの読み込み  
wb = openpyxl.load_workbook('ファイル名.xlsx')

シート名の確認  
wb.sheetnames  

for文を活用して、中身のシート名を確認する  
for sheet_name in wb.sheetnames:  
    print(sheet_name)  
または  
for sheet in wb:  
    print(sheet.title)

ファイルの保存  
wb.save('ファイル名.xlsx')

## シートの操作(作成・削除・コピー)

シートを選択する  
ws = wb['シート名']  
または  
ws = wb.worksheets[0]  

シート名の確認  
print(ws.title)  

シート名の変更  
ws.title = '変更後のシート名'  

シートを追加する  
wb.create_sheet(title='シート名')  

シートのタブ色を変更する  
ws.sheet_properties.tabColor='16進数の色'  

シートをコピーする
wb.copy_worksheet(wb['シート名'])  

シートを削除する  
wb.remove(wb['シート名'])  
または  
wb.remove(wb.worksheets[index])

## セルの操作

セルの選択  

1. セルのアドレスを指定する : ws['A1']  
2. セルの番号を指定する : ws.cell(row=1, column=1)  
3. セルの番号を指定する : ws.cell(1, 1)  

複数のセルを指定する  

1. range1 = ws['A1':'B3']
2. range2 = ws['A1:B3']

条件を指定したセルの操作

ヘッダー部分(=1行目)をスキップして読み込むには、以下のように書く  

```python
for row in ws.iter_rows(min_row=2):
    print(row)
```

ヘッダー(1行目)とインデックス(1列目)を飛ばして読み込む場合  

```python
for row in ws.iter_rows(min_row=2,min_col=2):
    format_l = []
    for cell in row:
        format_l.append(cell.coordinate)
    print(format_l)
```

行単位でセルを指定する

row1に1行目を格納する(0ではなく、1からスタート)  
row1 = ws[1]

セルの場所を確認する  
cell1の場所を確認する  
  アドレスを確認する  
  print(cell1.coordinate)  

  行番号を確認する  
  print(cell1.row)  

  列番号を確認する  
  print(cell1.column)  

  列のアルファベット部分を取得する  
  print(cell1.column_letter)  

セルの中身を確認する  
print(cell1.value)

cell1の中に、数字の1を格納する  
cell1.value = 1

A3セルに、数字の10を格納する  
cell3 = ws.cell(3,1,10)

文字列をセル内に書き込んであげることで、Excelの数式を使えるようになる  
cell4 = ws.cell(2,1,'=sum(a1:a3)')  

for文を使って、シート内のセル全体を表示する  

```python
for row in ws.values:
    print(row)
```

セルの書式設定  
現在の書式を確認するのも、書式を変更するのも、number_formatを使う  
B2セルを、小数第2ケタに変更  
```ws['B2']. number_format = '0.00'```  
B3セルを、2020年10月10日の形に変更  
```ws['B3'].number_format = 'yyyy月mm年dd日'```  

フォントの変更  
A3セルのフォントを、bold体 + イタリック体に変更する  
```ws['A3'].font = openpyxl.styles.Font(bold=True,italic=True)```  
