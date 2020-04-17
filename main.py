import argparse
import os
from ExcelPixelator import ExcelPixelator

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='A simple script that takes an image and outputs Excel pixel art')
    parser.add_argument('input', type=str,
                        help='The path to the input image')
    parser.add_argument('-o', '--output', type=str, default=f'{os.getcwd()}',
                        help='Path where you want to save your Excel pixel art, if not provided default value will be the same path of this script')
    parser.add_argument('-f', '--filename', type=str, default='Pixel Art',
                        help='Name of the resulting Excel workbook, if not provided default value will be "Pixel Art"')
    parser.add_argument('-c', '--cellsize', type=int, default=10,
                        help='The size (width and height) in pixels of the workbook cells where the image will be drawn, if not provided default value will be 10px')
    parser.add_argument('-p', '--pixelsize', type=int, default=1,
                        help='The size of the pixels of the input image. A bigger value means the resulting image will look more pixelated. This value must be a number divisible exactly by both the image width and height. If not provided default value will be 1, which means that one pixel in the input image will map to one cell in the Excel workbook')

    args = parser.parse_args()

    excelPixelator = ExcelPixelator(
        args.input, args.output, args.filename, args.cellsize, args.pixelsize)
    excelPixelator.create_pixel_art()
