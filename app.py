import itertools  
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import random
import ezgmail


template_name = input("template name: ")
data_file = input('names wali file ka name: ')
template_name+='.png'
data_file +=".xlsx"


x_name = 1637
y_name = 1350
#x_pos=2513
#y_pos=1533
# x_comp=1891
# y_comp=1465
# x_as=2841
# y_as=1465

try:
    df = pd.read_excel('data/{}'.format(data_file))
    names = list(df['Name'].values)
    emails = list(df['Email'].values)
except ValueError:
    print("There is a value error in the Dataframe please check the value")

try:
    for (name,email) in zip(names,emails):
        name = name.rstrip()

        #image khol
        img = Image.open("Certificate_Template/{}".format(template_name))
        width, height = img.size
        draw = ImageDraw.Draw(img)

        #fonts (agar multiple fonts)
        namefont = ImageFont.truetype("Fonts/Lato-Bold.ttf", 80)
        #posfont = ImageFont.truetype("Fonts/Lato-Bold.ttf", 100)

        # name daal 
        xhere_name = int(x_name - namefont.getsize(name)[0] / 2)    
        #xhere_pos = int(x_pos - namefont.getsize(pos)[0] / 2)    

        yhere_name = int(y_name - namefont.getsize(name)[1]) 
        #yhere_pos = int(y_pos - namefont.getsize(pos)[1]) 


        draw.text((xhere_name, yhere_name), name,fill ="black", font=namefont)
        #Position
        #draw.text((xhere_pos, yhere_pos),pos,fill ="black", font=namefont)

        img.save("Certificates/" +str(name) + ".png")
        print("Certificate Created for {}".format(name))
        ezgmail.send(email, 'Intellectus Model United Nations Certificate', 'Find your certificate attached below. If there is any discrepancy, please contact us at this email.',['./Certificates/'+name+'.png'])
        print("Email sent to {}".format(name))

except Exception:
    print('gadbad')
