# Excel Pixel Art
A simple Python program that takes an image as input and outputs an Excel workbook containing said image.

It works by getting the average of each pixel's RGB data and then applying the resulting color to each corresponding cell in Excel.

## Prerequisites
* Python 3.6 or above

## Install
```shell
git clone https://github.com/nibble-4bits/excel-pixel-art.git
cd excel-pixel-art
pip install -r requirements.txt
```

### Usage
```shell
python main.py [-h] [-o OUTPUT] [-f FILENAME] [-c CELLSIZE] [-p PIXELSIZE] input

positional arguments:
  input                 The path to the input image

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Path where you want to save your Excel pixel art, if
                        not provided default value will be the same path of
                        this script
  -f FILENAME, --filename FILENAME
                        Name of the resulting Excel workbook, if not provided
                        default value will be "Pixel Art"
  -c CELLSIZE, --cellsize CELLSIZE
                        The size (width and height) in pixels of the cells
                        where the image will be drawn, if not provided default
                        value will be 10px
  -p PIXELSIZE, --pixelsize PIXELSIZE
                        The size of the pixels of the input image. A bigger
                        value means the resulting image will look more
                        pixelated. This value must be a number divisible
                        exactly by both the image width and height. If not
                        provided default value will be 1, which means that one
                        pixel in the input image will map to one cell in the
                        Excel workbook
```
