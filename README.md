# Excel Pixel Art
A simple Python program that takes an image as input and outputs an Excel workbook containing said image.

It works by getting each pixel RGB data and then applying the resulting color to each cell in Excel.

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
python main.py [-h] [-o OUTPUT] [-f FILENAME] [-c CELLSIZE] input

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
```
