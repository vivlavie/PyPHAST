#https://pypi.org/project/python-resize-image/
#Actually there was no need to use 'resize'

# https://stackoverflow.com/questions/16646183/crop-an-image-in-the-centre-using-pil
# img.crop((left,top,width,height)) works!

import locale
locale.setlocale(locale.LC_ALL, 'english')

from PIL import Image
# from resizeimage import resizeimage
from os.path import isfile
import sys

new_width = 1660
new_height = 1080
jobs = ['000000','100103','100104']
# jobs = ['000000']
for j in jobs:
    for jf in ['-20','-60']:
        if isfile(j+jf+'.png') == True:
            fd_img = open(j+jf+'.png', 'rb')
            img = Image.open(fd_img)
            width, height = img.size

            # Crop the center of the image
            # left = (width - new_width)/2
            # top = (height - new_height)/2
            # right = (width + new_width)/2
            # bottom = (height + new_height)/2
            # img = img.crop((left, top, right, bottom))


            img = img.crop((0, 0, new_width, new_height))

            #Alternative method to crop the imge from the center, which looks redundant
            # img = resizeimage.resize_width(img, 1920)
            # img = resizeimage.resize_cover(img, [1660,1080])
            img.save('.\\cropped\\'+j+jf+'.png', img.format)
            fd_img.close()
