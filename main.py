import xlsxwriter
import os
from PIL import Image

# DISCLAIMER: For now works with small images
# file_name = 'selfie.jpg'
file_name = 'aperture-icon.png'

print('Image name: ', file_name)

# Create output folder
if not os.path.isdir('output'):
    os.makedirs('output')



def get_pixels_from_image(name):
    im = Image.open(name)
    return im.size, im.load()

def rgb_to_hex(r, g, b, a = 255):
    if a == 0:
        return '#FFFFFF'
    return '#%02x%02x%02x' % (r, g, b)

def write_cell(workbook, x, y, style):
    workbook.write(x, y, '', style)


im_size, im_pixels = get_pixels_from_image(f'images/{file_name}')
alpha = len(im_pixels[0, 0]) == 4 # check if alpha channel is included.

print('Number of cells: ', str(im_size[0] * im_size[1]))

workbook = xlsxwriter.Workbook(f'output/{file_name}.xlsx', {'constant_memory': True})
# workbook = xlsxwriter.Workbook(f'output/{file_name}.xlsx')
ws = workbook.add_worksheet('Image')
ws.set_default_row(10)

formats = {}

for row in range(0, im_size[0]):
    for col in range(0, im_size[1]):
        if not alpha:
            r, g, b = im_pixels[row, col]
            color = rgb_to_hex(r, g, b)
        else: 
            r, g, b, a = im_pixels[row, col]
            color = rgb_to_hex(r, g, b, a)
        
        if color not in formats.keys():
            formats[color] = workbook.add_format({'bg_color': color})

        write_cell(ws, row, col, formats[color])
        ws.set_column(col, col, 1)

print('Colours used: ', len(formats.keys()))

workbook.close()