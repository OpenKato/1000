# -*- coding: utf-8 -*-
import os
import glob
from PIL import Image
from openpyxl import Workbook
from openpyxl import load_workbook

files = glob.glob('C:\dir/*.JPG')

for f in files:
    img = Image.open(f)
    img_resize = img.resize((190, 190), Image.LANCZOS)
    
    ftitle, fext = os.path.splitext(f)
    img_resize.save(ftitle + '_half' + fext)

wb = load_workbook(filename = '1.xlsx')
ws = wb.active

from openpyxl.drawing.image import Image

img1 = Image('N_half.JPG')
img2 = Image('F_half.JPG')
img3 = Image('L_half.JPG')
img4 = Image('R_half.JPG')
img5 = Image('B_half.JPG')

ws.add_image(img1, 'D15')
ws.add_image(img2, 'L15')
ws.add_image(img3, 'D27')
ws.add_image(img4, 'L27')
ws.add_image(img5, 'D39')
wb.save('2.xlsx')
