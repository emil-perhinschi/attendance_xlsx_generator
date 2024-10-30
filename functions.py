
def write_day_cells(worksheet, column, start_row, end_row, format): 
    for i in range(start_row, end_row):
        if i == end_row - 1:
            worksheet.write(i, column, '', format)
        else:
            worksheet.write(i, column, '', format)

def month_romanian(month_int):
    months = [
        'Ianuarie',
        'Februarie',
        'Martie',
        'Aprilie',
        'Mai',
        'Iunie',
        'Iulie',
        'August',
        'Septembrie',
        'Octombrie',
        'Noiembrie',
        'Decembrie'
    ]
    return months[month_int]

