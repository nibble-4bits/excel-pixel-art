from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from PIL import Image


class ExcelPixelator:
    def __init__(self, input_path, output_path, file_name, cell_size, scale):
        self.image = Image.open(input_path).convert('RGB')
        self.output_path = output_path
        self.file_name = file_name
        self.cell_size = cell_size  # size of the cell in pixels
        self.scale = scale

    def create_pixel_art(self):
        default_excel_font_size = 16
        wb = Workbook()
        ws = wb.create_sheet('Pixel-Art')
        wb.remove(wb['Sheet'])  # remove default worksheet

        width, height = self.image.size
        pixel_map = [[0 for x in range(width // self.scale)] for y in range(height // self.scale)]
        squares = width * height // self.scale ** 2
        i, j = 0, 0
        for square in range(squares):
            rAvg, gAvg, bAvg = 0, 0, 0
            row_start = (square * self.scale // width) * self.scale
            row_end = row_start + self.scale
            col_start = square * self.scale % height
            col_end = col_start + self.scale
            for coll in range(row_start, row_end):
                for roww in range(col_start, col_end):
                    r, g, b = self.image.getpixel((roww, coll))
                    rAvg += r
                    gAvg += g
                    bAvg += b
            rAvg //= self.scale ** 2
            gAvg //= self.scale ** 2
            bAvg //= self.scale ** 2
            pixel_map[i][j] = (rAvg, gAvg, bAvg)
            i = i + 1 if j >= (width // self.scale) - 1 else i
            j = (j + 1) % (height // self.scale)

        for _row in range(len(pixel_map)):
            ws.row_dimensions[_row + 1].height = self.cell_size * 10 / default_excel_font_size
            for _col in range(len(pixel_map[_row])):
                curr_col = get_column_letter(_col + 1)
                ws.column_dimensions[curr_col].width = self.cell_size / 9

                rgbTuple = pixel_map[_row][_col]
                fill_color = self.__rgbToHex(rgbTuple)
                ws[f'{curr_col}{_row + 1}'].fill = PatternFill(
                    start_color=fill_color, end_color=fill_color, fill_type='solid')

        wb.save(f'{self.output_path}/{self.file_name}.xlsx')
        wb.close()

    def __rgbToHex(self, rgbTuple):
        return ('%02x%02x%02x' % rgbTuple).upper()
