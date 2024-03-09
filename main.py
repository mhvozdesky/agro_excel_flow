import os

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, Alignment

from styles import Styles
from current_books import AgroBook


DATA_DIR = os.getenv('DATA_DIR', None)
if not DATA_DIR:
    raise Exception('The application must know DATA_DIR')
FILE_NAME = 'Aggregate_Yield.xlsx'


def get_wb(input_file):
    # TODO error checking
    workbook = load_workbook(input_file)
    return workbook


def get_ws(wb):
    # TODO need to check if there is such a sheet (work_book.sheetnames)
    ws = wb['Whole data']
    return ws


def open_book(input_file):
    wb = get_wb(input_file)
    ws = get_ws(wb)
    return ws


def get_columns(input_ws):
    columns = {}
    for row in input_ws:
        for col in row:
            # columns[col.value] = col.column_letter
            columns[col.column_letter] = col.value
        break
    return columns


def process_file(agro_book, input_ws):
    columns = get_columns(input_ws)
    for i, row in enumerate(input_ws.rows):
        if i == 0:
            continue
        agro_book.write_from_row(row, columns)


def fill_book(agro_book, file_list):
    for input_file in file_list:
        input_ws = open_book(input_file)
        process_file(agro_book, input_ws)


def main(file_list):
    agro_book = AgroBook()
    fill_book(agro_book, file_list)
    agro_book.save_book(FILE_NAME, DATA_DIR)


def test():
    # font = Font(name='Calibri', size=10)
    # font_bold = Font(name='Calibri', size=10, bold=True)
    # border = Border(
    #     left=Side(border_style='thin', color='00000000'),
    #     right=Side(border_style='thin', color='00000000'),
    #     top=Side(border_style='thin', color='00000000'),
    #     bottom=Side(border_style='thin', color='00000000'),
    # )
    # alignment = Alignment(horizontal='center', vertical='center')

    styles = Styles()

    wb = Workbook()
    ws = wb.active

    a1 = ws['A1']
    b1 = ws['B1']

    a2 = ws['A2']
    b2 = ws['B2']
    print

    a1.value = 'name'
    b1.value = 'price'

    a2.value = 'spam'
    b2.value = 132.15

    for cell in [a1, b1, a2, b2]:
        if cell in [a1, b1]:
            cell.font = styles.font_bold
        else:
            cell.font = styles.font
        cell.border = styles.border
        cell.alignment = styles.alignment

    ws['C1'].value = 'spam'
    ws['C1'].font = styles.font_bold

    wb.save(f'{DATA_DIR}/test1.xlsx')


if __name__ == '__main__':
    input_file_list = [f'{DATA_DIR}/input_data/231010_UASESF88142023.xlsx']
    main(input_file_list)

    # test()
