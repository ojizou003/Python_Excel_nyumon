# 関数を使って、いつでも分かりやすい形で閲覧できるようにする
def show_formated_cell(ws,min_row=1,min_col=1):
    for row in ws.iter_rows(min_row=min_row,min_col=min_col):
        format_l = []
        for cell in row:
            format_l.append(cell.coordinate)
        print(format_l)