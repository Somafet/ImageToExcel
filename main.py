import xlsxwriter
from PIL import Image

# Start
file_name = 'aperture-icon.png'

def get_pixels_from_image(name):
    im = Image.open(name)
    return im.size, im.load()

def rgb_to_hex(r, g, b):
    return '#%02x%02x%02x' % (r, g, b)

im_size, im_pixels = get_pixels_from_image(f'images/{file_name}')

workbook = xlsxwriter.Workbook(f'output/{file_name}.xlsx')
ws = workbook.add_worksheet('Image')

for i in range(0, im_size[0]):
    for j in range(0, im_size[1]):
        r, g, b, _ = im_pixels[i, j]
        ws.write(i, j, ' ', workbook.add_format({'bg_color': rgb_to_hex(r, g, b)}))
workbook.close()