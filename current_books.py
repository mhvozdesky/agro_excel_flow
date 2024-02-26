from openpyxl import Workbook
from styles import Styles


class AgroBook:

    def __init__(self):
        self.wb = Workbook()
        self.ws = self.get_ws()
        self.num_rows = 1
        self.styles = Styles()
        self.columns = {
            'NN': {'type': 'handle', 'func': None, 'letter': 'A', 'width': 6.54},
            'Область': {'type': 'handle', 'func': self.state, 'letter': 'B', 'width': 19.35},
            'Район': {'type': 'simple', 'rel_col': 'City', 'letter': 'C', 'width': 17.36},
            'Господарство': {'type': 'handle', 'func': None, 'letter': 'D', 'width': 22.80},
            'GPS-координати поля': {'type': 'handle', 'func': None, 'letter': 'E', 'width': 26.59},
            'COMPANY': {'type': 'simple', 'rel_col': 'Hybrid Company Name', 'letter': 'F', 'width': 15.41},
            'HYBRIDS': {'type': 'simple', 'rel_col': 'Hybrid Name', 'letter': 'G', 'width': 20.24},
            'Обробіток грунту': {'type': 'simple', 'rel_col': 'PTL_C', 'letter': 'H', 'width': 15.77},
            'Густота на момент\nзбирання, тис/га': {'type': 'simple', 'rel_col': 'HAVPN', 'letter': 'I', 'width': 19.92},
            'Кіл-ть\nрядків': {'type': 'simple', 'rel_col': 'Number of rows', 'letter': 'J', 'width': 8.57},
            'Довжина\nділянки': {'type': 'simple', 'rel_col': 'Plot Length', 'letter': 'K', 'width': 8.57},
            'Площа, га': {'type': 'handle', 'func': None, 'letter': 'L', 'width': 10.0},
            'Вага з\nділянки, кг': {'type': 'simple', 'rel_col': 'GWTPN', 'letter': 'M', 'width': 10.65},
            'YIELD,\nBUNKER WEIGHT (q/ha)\nБункерна вага': {'type': 'handle', 'func': None, 'letter': 'N', 'width': 20.53},
            'Harvesting moisture,\n% Вологість': {'type': 'simple', 'rel_col': 'GMSTP', 'letter': 'O', 'width': 18.29},
            'Re-calculation\nof yield at basis\nmoisture (7 %) (UA)': {'type': 'handle', 'func': None, 'letter': 'P', 'width': 19.80},
            'Re-calculation\nof yield at basis\nmoisture (8 %) (F)': {'type': 'handle', 'func': None, 'letter': 'Q', 'width': 19.80},
            'Попередник': {'type': 'simple', 'rel_col': 'Previous Crop', 'letter': 'R', 'width': 16.71},
            'Дата посіву': {'type': 'simple', 'rel_col': 'Date of Planting', 'letter': 'S', 'width': 11.10},
            'Дата збирання': {'type': 'simple', 'rel_col': 'Date of Harvest', 'letter': 'T', 'width': 13.74},
            'ПІБ менеджера,\nщо створив протокол': {'type': 'simple', 'rel_col': 'Username', 'letter': 'U', 'width': 22.03},
            'Тип досліду\n(Demo/SBS/Strip)': {'type': 'simple', 'rel_col': 'Trial type', 'letter': 'V', 'width': 15.33},
            'Ширина\nміжряддя': {'type': 'simple', 'rel_col': 'Row spacing', 'letter': 'W', 'width': 16.63},
            'Коментарі': {'type': 'handle', 'func': None, 'letter': 'X', 'width': 35.20}
        }

        self.init_default_columns()

    def get_ws(self):
        ws = self.wb.active
        return ws

    def init_default_columns(self):
        for column_name, data in self.columns.items():
            cell = self.ws[f'{data["letter"]}1']
            cell.value = column_name
            self.apply_title_styles(cell)
            self.adjust_column_width(data["letter"], data['width'])
        self.ws.row_dimensions[1].height = 46.49

    def adjust_column_width(self, column_letter, width):
        self.ws.column_dimensions[column_letter].width = width

    def apply_title_styles(self, cell):
        self.apply_styles(cell)
        cell.font = self.styles.font_bold

    def apply_styles(self, cell):
        cell.font = self.styles.font
        cell.border = self.styles.border
        cell.alignment = self.styles.alignment

    def row_to_dict(self, row, column):
        row_dict = {}
        for cell in row:
            row_dict[column[cell.column_letter]] = cell.value
        return row_dict

    def write_data(self, row_dict):
        self.num_rows += 1
        for title, data in self.columns.items():
            cell = self.ws[f'{data["letter"]}{self.num_rows}']
            if data['type'] == 'handle':
                method = data['func']
                if method:
                    val = method(row_dict)
                    cell.value = val
            elif data['type'] == 'simple':
                cell.value = row_dict.get(data['rel_col'], None)
            self.apply_styles(cell)

    def write_from_row(self, file_row, file_columns):
        row_dict = self.row_to_dict(file_row, file_columns)
        self.write_data(row_dict)

    def save_book(self, file_name, file_dir):
        self.wb.save(f'{file_dir}/{file_name}')

    def state(self, row_dict):
        rel_col = 'State'
        value = row_dict.get(rel_col, '')
        return value.replace('область', '').strip()
