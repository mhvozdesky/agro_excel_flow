from openpyxl import Workbook
from styles import Styles


class AgroBook:

    def __init__(self):
        self.wb = Workbook()
        self.ws = self.get_ws()
        self.num_rows = 1
        self.styles = Styles()
        self.columns = {
            'NN': {'type': 'handle', 'func': self.line_number, 'letter': 'A', 'width': 6.54},
            'Область': {'type': 'handle', 'func': self.state, 'letter': 'B', 'width': 19.35},
            'Район': {'type': 'simple', 'rel_col': 'City', 'letter': 'C', 'width': 17.36},
            'Господарство': {'type': 'handle', 'func': self.household, 'letter': 'D', 'width': 22.80},
            'GPS-координати поля': {'type': 'handle', 'func': self.gps_coordinates, 'letter': 'E', 'width': 28.53},
            'COMPANY': {'type': 'simple', 'rel_col': 'Hybrid Company Name', 'letter': 'F', 'width': 15.41},
            'HYBRIDS': {'type': 'simple', 'rel_col': 'Hybrid Name', 'letter': 'G', 'width': 20.24},
            'Обробіток грунту': {'type': 'simple', 'rel_col': 'PTL_C', 'letter': 'H', 'width': 15.77},
            'Густота на момент\nзбирання, тис/га': {'type': 'simple', 'rel_col': 'HAVPN', 'letter': 'I', 'width': 19.92},
            'Кіл-ть\nрядків': {'type': 'simple', 'rel_col': 'Number of rows', 'letter': 'J', 'width': 8.57},
            'Довжина\nділянки': {'type': 'simple', 'rel_col': 'Plot Length', 'letter': 'K', 'width': 8.57},
            'Площа, га': {'type': 'handle', 'func': self.area_hectares, 'letter': 'L', 'width': 10.0, 'style': {'number_format': '0.0000'}},
            'Вага з\nділянки, кг': {'type': 'simple', 'rel_col': 'GWTPN', 'letter': 'M', 'width': 10.65},
            'YIELD,\nBUNKER WEIGHT (q/ha)\nБункерна вага': {'type': 'handle', 'func': self.yield_bunker_weight_field, 'letter': 'N', 'width': 20.53, 'style': {'number_format': '0.00'}},
            'Harvesting moisture,\n% Вологість': {'type': 'simple', 'rel_col': 'GMSTP', 'letter': 'O', 'width': 18.29},
            'Re-calculation\nof yield at basis\nmoisture (7 %) (UA)': {'type': 'handle', 'func': self.yield_recalculation_field_7, 'letter': 'P', 'width': 19.80, 'style': {'number_format': '0.00'}},
            'Re-calculation\nof yield at basis\nmoisture (8 %) (F)': {'type': 'handle', 'func': self.yield_recalculation_field_8, 'letter': 'Q', 'width': 19.80, 'style': {'number_format': '0.00'}},
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
            cell_coordinate = f'{data["letter"]}{self.num_rows}'
            cell = self.ws[cell_coordinate]

            if data['type'] == 'handle' and data['func']:
                method = data['func']

                try:
                    val = method(row_dict)
                except Exception as e:
                    print(f'Помилка при розрахунку {cell_coordinate}. {e}')
                    val = None

                cell.value = val

            elif data['type'] == 'simple':
                cell.value = row_dict.get(data['rel_col'], None)

            self.apply_styles(cell)

            if data.get('style'):
                style = data.get('style')
                for key, value in style.items():
                    setattr(cell, key, value)

    def write_from_row(self, file_row, file_columns):
        row_dict = self.row_to_dict(file_row, file_columns)
        self.write_data(row_dict)

    def save_book(self, file_name, file_dir):
        self.wb.save(f'{file_dir}/{file_name}')

    def state(self, row_dict):
        rel_col = 'State'
        value = row_dict.get(rel_col, '')
        return value.replace('область', '').strip()

    def line_number(self, row_dict):
        return self.num_rows - 1

    def household(self, row_dict):
        value_row = row_dict.get('Custom trial name', '')
        value = value_row.split(',')
        if len(value) >= 4:
            return value[3]
        return value_row

    def format_coordinates(self, coordinate):
        if coordinate is None:
            return ''

        coord = str(coordinate)
        coord.replace(',', '.')
        return coord

    def gps_coordinates(self, row_dict):
        longitude = row_dict.get('Longitude', None)
        longitude = self.format_coordinates(longitude)

        latitude = row_dict.get('Latitude', None)
        latitude = self.format_coordinates(latitude)
        return f'{longitude},{latitude}'

    def get_not_none_value(self, col_name, row_dict):
        value = row_dict.get(col_name, None)
        if value is None:
            raise ValueError(f'Не вдалось отримати дані колонки "{col_name}"')
        return value

    def get_float_value(self, value):
        try:
            value = float(value)
        except ValueError:
            raise ValueError(f'Дані {value} не можуть бути конвертовані в число')
        return value

    def calculate_the_area_from_hrvar(self, row_dict):
        hrvar = self.get_not_none_value('Hrvar', row_dict)
        hrvar = self.get_float_value(hrvar)
        return round((hrvar / 10000), 4)

    def area_hectares(self, row_dict):
        number_of_rows = row_dict.get('Number of rows', None)
        length_of_plot = row_dict.get('Plot Length', None)

        if number_of_rows is None or length_of_plot is None:
            area = self.calculate_the_area_from_hrvar(row_dict)
        else:
            area = f'=J{self.num_rows}*0.7*K{self.num_rows}/10000'

        return area

    def yield_bunker_weight_field(self, row_dict):
        return f'=M{self.num_rows}/(L{self.num_rows}*100)'

    def yield_recalculation_field_7(self, row_dict):
        return f'=N{self.num_rows}*((100-O{self.num_rows})/93)'

    def yield_recalculation_field_8(self, row_dict):
        return f'=N{self.num_rows}*((100-O{self.num_rows})/92)'


