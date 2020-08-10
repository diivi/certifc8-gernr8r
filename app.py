
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import random

y_coordinate = int(input('Y-cordinate , smun ke liye 260 ke aaspaas try'))
template_name = input("template name: ")
data_file = input('names wali file ka name: ')
template_name+='.jpg'
data_file +=".xlsx"
try:
    df = pd.read_excel('data/{}'.format(data_file))
    names = list(df['Name'].values)
except ValueError:
    print("There is a value error in the Dataframe please check the value")

try:
    for name in names:
        name = name.rstrip()

        #naam daal
        img = Image.open("Certificate_Template/{}".format(template_name))
        width, height = img.size
        draw = ImageDraw.Draw(img)

        # font
        font = ImageFont.truetype("Fonts/BirdsOfParadise.ttf", 32)
        x_coordinate = int(400 - font.getsize(name)[0] / 2) 
        print(x_coordinate)
        draw.text((x_coordinate, y_coordinate), name, (238, 33, 33), font=font)

        img.save("Certificates/" + str(name) + ".jpg")
        print("Certificate Created for {}".format(name))
except Exception:
    print('gadbad')
