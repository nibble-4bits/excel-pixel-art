from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from PIL import Image


class ExcelPixelator:
    def __init__(self, input_path, output_path, file_name, cell_size, scale=1):
        self.image = Image.open(input_path).convert('RGB')
        self.output_path = output_path
        self.file_name = file_name
        self.cell_size = cell_size  # size of the cell in pixels
        self.scale = scale  # not used yet

    def create_pixel_art(self):
        default_excel_font_size = 16
        wb = Workbook()
        ws = wb.create_sheet('Pixel-Art')
        wb.remove(wb['Sheet'])  # remove default worksheet

        width, height = self.image.size
        for _row in range(1, width):
            ws.row_dimensions[_row].height = self.cell_size * 10 / default_excel_font_size
            for _col in range(1, height):
                curr_col = get_column_letter(_col)
                ws.column_dimensions[curr_col].width = self.cell_size / 9

                rgbTuple = self.image.getpixel((_col, _row))
                fill_color = self.__rgbToHex(rgbTuple)
                ws[f'{curr_col}{_row}'].fill = PatternFill(
                    start_color=fill_color, end_color=fill_color, fill_type='solid')

        wb.save(f'{self.output_path}/{self.file_name}.xlsx')
        wb.close()

    def __rgbToHex(self, rgbTuple):
        return ('%02x%02x%02x' % rgbTuple).upper()
