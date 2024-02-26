from openpyxl.styles import Font, Border, Side, Alignment


class Styles:
    def __init__(self):
        self.font = self._get_simple_font()
        self.font_bold = self._get_bold_font()
        self.border = self._get_border()
        self.alignment = self._get_alignment()

    def _get_simple_font(self):
        return Font(name='Calibri', size=10)

    def _get_bold_font(self):
        font = self._get_simple_font()
        font.bold = True
        return font

    def _get_border(self):
        return Border(
            left=Side(border_style='thin', color='00000000'),
            right=Side(border_style='thin', color='00000000'),
            top=Side(border_style='thin', color='00000000'),
            bottom=Side(border_style='thin', color='00000000'),
        )

    def _get_alignment(self):
        return Alignment(horizontal='center', vertical='center')
