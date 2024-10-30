#!/usr/bin/env python

import xlsxwriter
import calendar
import functions
from xlsxwriter.utility import xl_rowcol_to_cell
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--year', dest='year', help='Year', required=True)
parser.add_argument('--month', dest='month', help='Month', required=True)
args = parser.parse_args()
# year = 2024 # TODO take from arguments
# month = 10  # TODO take from arguments
year = int(args.year)
month = int(args.month)
month_name = functions.month_romanian(month)

workbook = xlsxwriter.Workbook("{}_{}.xlsx".format(year, month_name))
workbook.formats[0].set_font_size(10)
workbook.formats[0].set_font_name('Arial')
worksheet = workbook.add_worksheet()

cal = calendar.Calendar()
calendar.setfirstweekday(calendar.MONDAY)


days_start_column = 3
days_start_row = 7
max_size_table = 10
# set column widths
worksheet.set_column('A:A', 6)
worksheet.set_column('B:B', 25)
worksheet.set_column('C:C', 16)
worksheet.set_column(days_start_column, calendar.monthrange(year, month)[1] + days_start_column - 1, 4);

# set header height
worksheet.set_row(7, 50)

# print(calendar.monthrange(year, month)[1])

worksheet.write('B2', "Your Company Name Ltd")
worksheet.write('B3', "CUI: 123456789")


# title 
title_format = workbook.add_format({ 'align': 'center', 'bold': True, 'underline': True})
subtitle_format = workbook.add_format({ 'align': 'center', 'bold': True })
worksheet.merge_range('N3:Y3', 'FOAIE COLECTIVA PREZENTA', title_format)
# worksheet.write('N3', 'FOAIE COLECTIVA PREZENTA')
worksheet.merge_range('P5:W5', '{} {}'.format(month_name, year), subtitle_format)

# table headers
header_format = workbook.add_format({'top': 2, 'bottom':2, 'valign': 'vcenter', 'align': 'center'})
header_format_weekend = workbook.add_format({'top':2, 'bottom':2,'valign': 'vcenter', 'align': 'center', 'bg_color': 'gray'})

worksheet.merge_range('A8:A9', 'Nr crt', header_format)
plain_table_cell = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border': 1})
weekend_table_cell = workbook.add_format({'align': 'center', 'valign':'vcenter', 'border': 1, 'bg_color': 'gray'})
for i in range(1,max_size_table):
    worksheet.write(8+i, 0, i, plain_table_cell)

name_table_cell = workbook.add_format({'align': 'left', 'valign':'top', 'text_wrap': True, 'border': 1})
worksheet.merge_range('B8:B9', 'Nume Prenume Functia', header_format)
for i in range(1,max_size_table):
    worksheet.write(8+i, 1, '', name_table_cell)

worksheet.write('C8', 'Program lucru', header_format)
worksheet.write('C9', 'ore', header_format)
for i in range(1,max_size_table):
    worksheet.write(8+i, 2, '', plain_table_cell)
    
last_day_number = 0
days_letters = ["L", "M", "M", "J", "V", "S", "D"]
for day in cal.itermonthdays(year, month): 
    # print(day)
    if day == 0:
        continue
    day_string = day
    last_cell_in_days = days_start_column + day - 1
    day_of_week = calendar.weekday(year, month, day)
    if day_of_week in (5, 6):
        worksheet.write(days_start_row, last_cell_in_days, day_string , header_format_weekend)
        worksheet.write(days_start_row + 1, last_cell_in_days, days_letters[day_of_week], header_format_weekend)
        functions.write_day_cells(
            worksheet, last_cell_in_days, 
            days_start_row + 2, days_start_row + 1 + max_size_table, 
            weekend_table_cell
            )
    else:
        worksheet.write(days_start_row, last_cell_in_days, day_string, header_format)
        worksheet.write(days_start_row + 1, last_cell_in_days, days_letters[day_of_week], header_format)
        functions.write_day_cells(
            worksheet, last_cell_in_days, 
            days_start_row + 2, days_start_row + 1 + max_size_table, 
            plain_table_cell
            )
    
# totals
totals_format = workbook.add_format({
    'top':2, 'bottom':2,'valign': 'vcenter', 'align': 'center', 'bg_color': '#8eaadb',
    'left':2, 'right':2
})
totals_header_format = workbook.add_format({
    'top':2, 'bottom':2, 'left':2, 'right':2, 'valign': 'vcenter', 'align': 'center', 'bg_color': '#8eaadb', 'text_wrap': True
})
totals_column = last_cell_in_days + 1
totals_header_row = 7;
totals_headers_cell = xl_rowcol_to_cell(totals_header_row, totals_column)
totals_headers_cell_to_merge = xl_rowcol_to_cell(totals_header_row + 1, totals_column )
totals_cells = "{}:{}".format(totals_headers_cell, totals_headers_cell_to_merge)

worksheet.merge_range(totals_cells, "Total ore lucrate", totals_header_format)
worksheet.write(totals_header_row + 1, totals_column, "", totals_format)
first_worked_hours_column = 3
for i in range(1, max_size_table):
    start_cell = xl_rowcol_to_cell(totals_header_row + 1 + i, first_worked_hours_column);
    end_cell = xl_rowcol_to_cell(totals_header_row + 1 + i, last_cell_in_days)
    worksheet.write(totals_header_row + 1 + i, totals_column, '=sum({}:{})'.format(start_cell, end_cell), totals_format)

# hours worked per shift and other columns I don't care about
extra_headers = [
    "Ore lucrate schimb",
    "Ore suplimentare zi",
    "Ore supl noapte",
    "Ore supl S/D",
    "Ore de noapte zi lucratoare",
    "CO",
    "CM"
]
extra_header_column = totals_column + 1
extra_header_format = workbook.add_format({
    'top': 2, 'bottom':2, 'left': 1, 'right': 1,
    'valign': 'vcenter', 'align': 'center', 'rotation': 90,
    'text_wrap': True})
extra_cells_format = workbook.add_format({'border': 1})
for extra_header in extra_headers:
    start_cell_extra_headers = xl_rowcol_to_cell(totals_header_row, extra_header_column)
    end_cell_extra_headers   = xl_rowcol_to_cell(totals_header_row + 1, extra_header_column)
    cells_to_merge_extra_headers = '{}:{}'.format(start_cell_extra_headers, end_cell_extra_headers)
    worksheet.merge_range(cells_to_merge_extra_headers, extra_header, extra_header_format)
    for extra_row_count in range(1,max_size_table):
        worksheet.write(totals_header_row + 1 + extra_row_count, extra_header_column, '', extra_cells_format)

    extra_header_column = extra_header_column + 1


workbook.close()




