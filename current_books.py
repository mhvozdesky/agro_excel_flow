import sys
from openpyxl import Workbook, load_workbook
from styles import Styles


class BaseBook:
    sheet_name = 'aggregate_yield'
    column_lib = {
        'Область': {'type': 'handle', 'width': 19.35},
        'Район': {'type': 'simple', 'width': 17.36},
        'Господарство': {'type': 'handle', 'width': 22.80},
        'GPS-координати поля': {'type': 'handle', 'width': 28.53},
        'COMPANY': {'type': 'simple', 'width': 15.41},
        'HYBRIDS': {'type': 'simple', 'width': 20.24},
        'Обробіток грунту': {'type': 'simple', 'width': 15.77},
        'Тип грунту': {'type': 'simple', 'width': 15.77},
        'Густота на момент\nзбирання, тис/га': {'type': 'simple', 'width': 19.92},
        'Кіл-ть\nрядків': {'type': 'simple', 'width': 8.57},
        'Довжина\nділянки': {'type': 'simple', 'width': 8.57},
        'Площа, га': {'type': 'handle', 'width': 10.0, 'style': {'number_format': '0.0000'}},
        'Вага з\nділянки, кг': {'type': 'simple', 'width': 10.65},
        'YIELD,\nBUNKER WEIGHT (q/ha)\nБункерна вага': {'type': 'handle', 'width': 20.53, 'style': {'number_format': '0.00'}},
        'Harvesting moisture,\n% Вологість': {'type': 'simple', 'width': 18.29},
        'Вологість на час\nзбирання, %': {'type': 'simple', 'width': 18.29},
        'Re-calculation\nof yield at basis\nmoisture (7 %) (UA)': {'type': 'handle', 'width': 19.80, 'style': {'number_format': '0.00'}},
        'Re-calculation\nof yield at basis\nmoisture (8 %) (F)': {'type': 'handle', 'width': 19.80, 'style': {'number_format': '0.00'}},
        'Урожайність у\nперерахунку на базисну\nвологість (9 %) (UA)': {'type': 'handle', 'width': 22.47, 'style': {'number_format': '0.00'}},
        'Попередник': {'type': 'simple', 'width': 16.71},
        'Дата посіву': {'type': 'simple', 'width': 11.10},
        'Дата збирання': {'type': 'simple', 'width': 13.74},
        'ПІБ менеджера,\nщо створив протокол': {'type': 'simple', 'width': 22.03},
        'Тип досліду\n(Demo/SBS/Strip)': {'type': 'simple', 'width': 15.33},
        'Ширина\nміжряддя': {'type': 'simple', 'width': 16.63},
        'Ширина\nділянки': {'type': 'simple', 'width': 16.63},
        'Коментарі': {'type': 'handle', 'width': 35.20}
    }

    def __init__(self, file_path=None):
        self.wb = None
        self.ws = None
        self.num_rows = 0
        self.styles = Styles()
        self.columns = self.init_columns()

        self.init_workbook(file_path)

    def get_column_from_lib(self, title, letter, func=None, rel_col=None):
        column_setting = self.column_lib.get(title)
        if column_setting is None:
            raise ValueError(f'В бібліотеці колонок не знайдено {title}')

        column_type = column_setting['type']
        if column_type == 'handle':
            column_setting['func'] = func
        elif column_type == 'simple':
            column_setting['rel_col'] = rel_col

        column_setting['letter'] = letter

        return column_setting

    def init_columns(self):
        return {'NN': {'type': 'handle', 'func': self.line_number, 'letter': 'A', 'width': 6.54}}

    def init_workbook(self, file_path):
        if file_path is None:
            self.wb = Workbook()
            self.ws = self.get_ws()
            self.num_rows = 1

            self.init_default_columns()
        else:
            try:
                self.wb = load_workbook(file_path)
            except Exception:
                raise RuntimeError(f'Не вдалось відкрити файл {file_path}')

            self.ws = self.get_ws_from_file(self.wb)
            self.num_rows = self.ws.max_row

            self.check_file_structure()

    def get_ws(self):
        ws = self.wb.active
        ws.title = self.sheet_name
        return ws

    def get_ws_from_file(self, wb):
        if self.sheet_name not in wb.sheetnames:
            raise ValueError(f'Неправильна структура файла. Відсутній лист {self.sheet_name}')

        ws = wb.get_sheet_by_name(self.sheet_name)
        return ws

    def check_file_structure(self):
        default_column_list = list(self.columns.keys())
        first_row = next(self.ws.iter_rows())
        file_column = [cell.value for cell in first_row]

        # the number of elements must be the same
        not_enough_items = max(len(default_column_list) - len(file_column), 0)
        file_column.extend([''] * not_enough_items)

        for i, col_name in enumerate(default_column_list):
            if col_name != file_column[i]:
                raise ValueError(f'Неправильна структура файла. Проблема в назві колонки {i+1}')

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
        cell.alignment = self.styles.wrap_text()

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
                    print(f'Помилка при розрахунку {cell_coordinate}. {e}', file=sys.stderr)
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

    def line_number(self, row_dict):
        return self.num_rows - 1

    def calculate_the_area_by_formula(self):
        pass


class OilSeedCropBook(BaseBook):
    def state(self, row_dict):
        rel_col = 'State'
        value = row_dict.get(rel_col, '')
        if value is None:
            return None
        return value.replace('область', '').strip()

    def household(self, row_dict):
        value_row = row_dict.get('Custom trial name', '')
        if value_row is None:
            return None
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

        if longitude == '' or latitude == '':
            return None

        return f'{latitude},{longitude}'

    def get_not_none_value(self, col_name, row_dict):
        value = row_dict.get(col_name, None)
        if value is None:
            raise ValueError(f'Не вдалось отримати дані колонки "{col_name}"')
        return value

    def get_float_value(self, value):
        try:
            value = float(value)
        except ValueError:
            raise ValueError(
                f'Дані {value} не можуть бути конвертовані в число')
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
            area = self.calculate_the_area_by_formula()

        return area

    def yield_bunker_weight_field(self, row_dict):
        return f'=M{self.num_rows}/(L{self.num_rows}*100)'

    def yield_recalculation_field_7(self, row_dict):
        return f'=N{self.num_rows}*((100-O{self.num_rows})/93)'

    def yield_recalculation_field_8(self, row_dict):
        return f'=N{self.num_rows}*((100-O{self.num_rows})/92)'

    def calculate_yield_bunker_weight(self, row_dict):
        return f'=Q{self.num_rows}/(P{self.num_rows}*100)'

    def yield_recalculation_field_9(self, row_dict):
        return f'=R{self.num_rows}*((100-S{self.num_rows})/91)'


class AgroBookSunflower(OilSeedCropBook):
    def init_columns(self):
        columns = super().init_columns()

        columns.update({
            'Область': self.get_column_from_lib('Область', func=self.state, letter='B'),
            'Район': self.get_column_from_lib('Район', rel_col='City', letter='C'),
            'Господарство': self.get_column_from_lib('Господарство', func=self.household, letter='D'),
            'GPS-координати поля': self.get_column_from_lib('GPS-координати поля', func=self.gps_coordinates, letter='E'),
            'COMPANY': self.get_column_from_lib('COMPANY', rel_col='Hybrid Company Name', letter='F'),
            'HYBRIDS': self.get_column_from_lib('HYBRIDS', rel_col='Hybrid Name', letter='G'),
            'Обробіток грунту': self.get_column_from_lib('Обробіток грунту', rel_col='PTL_C', letter='H'),
            'Густота на момент\nзбирання, тис/га': self.get_column_from_lib('Густота на момент\nзбирання, тис/га', rel_col='HAVPN', letter='I'),
            'Кіл-ть\nрядків': self.get_column_from_lib('Кіл-ть\nрядків', rel_col='Number of rows', letter='J'),
            'Довжина\nділянки': self.get_column_from_lib('Довжина\nділянки', rel_col='Plot Length', letter='K'),
            'Площа, га': self.get_column_from_lib('Площа, га', func=self.area_hectares, letter='L'),
            'Вага з\nділянки, кг': self.get_column_from_lib('Вага з\nділянки, кг', rel_col='GWTPN', letter='M'),
            'YIELD,\nBUNKER WEIGHT (q/ha)\nБункерна вага': self.get_column_from_lib('YIELD,\nBUNKER WEIGHT (q/ha)\nБункерна вага', func=self.yield_bunker_weight_field, letter='N'),
            'Harvesting moisture,\n% Вологість': self.get_column_from_lib('Harvesting moisture,\n% Вологість', rel_col='GMSTP', letter='O'),
            'Re-calculation\nof yield at basis\nmoisture (7 %) (UA)': self.get_column_from_lib('Re-calculation\nof yield at basis\nmoisture (7 %) (UA)', func=self.yield_recalculation_field_7, letter='P'),
            'Re-calculation\nof yield at basis\nmoisture (8 %) (F)': self.get_column_from_lib('Re-calculation\nof yield at basis\nmoisture (8 %) (F)', func=self.yield_recalculation_field_8, letter='Q'),
            'Попередник': self.get_column_from_lib('Попередник', rel_col='Previous Crop', letter='R'),
            'Дата посіву': self.get_column_from_lib('Дата посіву', rel_col='Date of Planting', letter='S'),
            'Дата збирання': self.get_column_from_lib('Дата збирання', rel_col='Date of Harvest', letter='T'),
            'ПІБ менеджера,\nщо створив протокол': self.get_column_from_lib('ПІБ менеджера,\nщо створив протокол', rel_col='Username', letter='U'),
            'Тип досліду\n(Demo/SBS/Strip)': self.get_column_from_lib('Тип досліду\n(Demo/SBS/Strip)', rel_col='Trial type', letter='V'),
            'Ширина\nміжряддя': self.get_column_from_lib('Ширина\nміжряддя', rel_col='Row spacing', letter='W'),
            'Коментарі': self.get_column_from_lib('Коментарі', func=None, letter='X')
        })

        return columns

    def calculate_the_area_by_formula(self):
        return f'=J{self.num_rows}*0.7*K{self.num_rows}/10000'


class AgroBookRapeSeed(OilSeedCropBook):
    def init_columns(self):
        columns = super().init_columns()

        sales_manager_full_name = self.get_column_from_lib('ПІБ менеджера,\nщо створив протокол', rel_col='Username', letter='U')
        sales_manager_full_name['width'] = 26.48

        soil_processing = self.get_column_from_lib('Обробіток грунту', rel_col='PTL_C', letter='I')
        soil_processing['width'] = 20.63

        columns.update({
            'Область': self.get_column_from_lib('Область', func=self.state, letter='B'),
            'Район': self.get_column_from_lib('Район', rel_col='City', letter='C'),
            'Господарство': self.get_column_from_lib('Господарство', func=self.household, letter='D'),
            'GPS-координати поля': self.get_column_from_lib('GPS-координати поля', func=self.gps_coordinates, letter='E'),
            'Компанія': self.get_column_from_lib('COMPANY', rel_col='Hybrid Company Name', letter='F'),
            'Гібрид': self.get_column_from_lib('HYBRIDS', rel_col='Hybrid Name', letter='G'),
            'Попередник': self.get_column_from_lib('Попередник', rel_col='Previous Crop', letter='H'),
            'Обробіток грунту': soil_processing,
            'Тип грунту': self.get_column_from_lib('Тип грунту', rel_col='USDAS', letter='J'),
            'Дата посіву': self.get_column_from_lib('Дата посіву', rel_col='Date of planting', letter='K'),
            'Дата збирання': self.get_column_from_lib('Дата збирання', rel_col='Date of Harvest', letter='L'),
            'Ширина\nміжряддя,\nсм': self.get_column_from_lib('Ширина\nміжряддя', func=None, letter='M'),
            'Довжина\nділянки,\nм': self.get_column_from_lib('Довжина\nділянки', rel_col='Plot Length', letter='N'),
            'Ширина\nділянки,\nм': self.get_column_from_lib('Ширина\nділянки', rel_col='Plot Width', letter='O'),
            'Площа, га': self.get_column_from_lib('Площа, га', func=self.area_hectares, letter='P'),
            'Вага з\nділянки, кг': self.get_column_from_lib('Вага з\nділянки, кг', rel_col='USY', letter='Q'),
            'YIELD,\nBUNKER WEIGHT (q/ha)\nБункерна вага': self.get_column_from_lib('YIELD,\nBUNKER WEIGHT (q/ha)\nБункерна вага', func=self.calculate_yield_bunker_weight, letter='R'),
            'Вологість на час\nзбирання, %': self.get_column_from_lib('Вологість на час\nзбирання, %', rel_col='HSH', letter='S'),
            'Урожайність у\nперерахунку на базисну\nвологість (9 %) (UA)': self.get_column_from_lib('Урожайність у\nперерахунку на базисну\nвологість (9 %) (UA)', func=self.yield_recalculation_field_9, letter='T'),
            'ПІБ Менеджер з продажів': sales_manager_full_name,
            'Тип досліду\n(Demo/SBS/Strip)': self.get_column_from_lib('Тип досліду\n(Demo/SBS/Strip)', rel_col='Trial type', letter='V'),
        })

        return columns

    def calculate_the_area_by_formula(self):
        return f'=(N{self.num_rows}*O{self.num_rows})/10000'

    def area_hectares(self, row_dict):
        width_of_plot = row_dict.get('Plot Width', None)
        length_of_plot = row_dict.get('Plot Length', None)

        if width_of_plot is None or length_of_plot is None:
            area = self.calculate_the_area_from_hrvar(row_dict)
        else:
            area = self.calculate_the_area_by_formula()

        return area

    def household(self, row_dict):
        value_row = row_dict.get('Custom trial name', None)
        return value_row
