# 【演習問題】
# 選択したシートを見やすくするコードを関数化。
# ※関数化する目的は、いつでも分かりやすい形で閲覧できるようにするため

def show_formatted_cell(range_): #rangeが予約語なのでrange_
    for row in range_:
        format_l =[]
        for cell in row:
            format_l.append(cell.coordinate)
        print(format_l)

