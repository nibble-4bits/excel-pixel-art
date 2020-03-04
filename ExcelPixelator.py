from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from PIL import Image


class ExcelPixelator:
    def __init__(self, image_path, scale):
        self.image = Image.open(image_path).convert('RGB')
        self.default_excel_font_size = 16
        self.cell_size = 10  # size of the cell in pixels
        self.scale = scale  # not used yet

    def create_pixel_art(self, output_path):
        wb = Workbook()
        ws = wb.create_sheet('Pixel-Art')
        wb.remove(wb['Sheet'])  # remove default worksheet

        width, height = self.image.size
        for _row in range(1, width):
            ws.row_dimensions[_row].height = self.cell_size * \
                10 / self.default_excel_font_size
            for _col in range(1, height):
                ws.column_dimensions[curr_col].width = self.cell_size / 9

                curr_col = get_column_letter(_col)
                rgbTuple = self.image.getpixel((_col, _row))
                fill_color = self.__rgbToHex(rgbTuple)
                ws[f'{curr_col}{_row}'].fill = PatternFill(
                    start_color=fill_color, end_color=fill_color, fill_type='solid')

        wb.save(f'{output_path}.xlsx')
        wb.close()

    def __rgbToHex(self, rgbTuple):
        return ('%02x%02x%02x' % rgbTuple).upper()
