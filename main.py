import datetime, xlrd, random
import Image, ImageDraw, ImageFont, ImageFilter
# https://www.geeksforgeeks.org/reading-excel-file-using-python/
# for profile pic - beta
import json, requests
from io import BytesIO

#loading in comon var
with open('picData.txt') as json_file:
  profileList = json.load(json_file)

# for circle image cropping
def crop_center(pil_img, crop_width, crop_height):
    img_width, img_height = pil_img.size
    return pil_img.crop(((img_width - crop_width) // 2,
                         (img_height - crop_height) // 2,
                         (img_width + crop_width) // 2,
                         (img_height + crop_height) // 2))
def crop_max_square(pil_img):
  return crop_center(pil_img, min(pil_img.size), min(pil_img.size))

def mask_circle_solid(pil_img, background_color, blur_radius, offset=0):
    background = Image.new(pil_img.mode, pil_img.size, background_color)

    offset = blur_radius * 2 + offset
    mask = Image.new("L", pil_img.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((offset, offset, pil_img.size[0] - offset, pil_img.size[1] - offset), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(blur_radius))

    return Image.composite(pil_img, background, mask)

def text_wrap(text, font, max_width):
    lines = []
    if font.getsize(text)[0] <= max_width:
        lines.append(text) 
    else:
        words = text.split(' ')  
        i = 0
        while i < len(words):
            line = ''         
            while i < len(words) and font.getsize(line + words[i])[0] <= max_width:                
                line = line + words[i] + " "
                i += 1
            if not line:
                line = words[i]
                i += 1
            lines.append(line)    
    return lines
loc = ("./jsSample.xlsx")
imgRaw = ("./template.png")

wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 

# generate unique random numbers
try:
 unqRandList = random.sample(range(0, 99), sheet.nrows-1)
except ValueError:
  print('Sample size exceeded population size.')

# dateTime,amt
fontCR = ImageFont.truetype('fonts/Calibri_Regular.ttf', size=48, encoding="unic")
# carType/Addr
fontCL = ImageFont.truetype('fonts/Calibri_Light.ttf', size=42, encoding="unic")
# rest
fontCLB = ImageFont.truetype('fonts/Calibri_Light.ttf', size=52, encoding="unic")
white = 'rgb(255, 255, 255)' # white color
black = 'rgb(0, 0, 0)' # white color
grey = 'rgb(169, 169, 169)' # white color
for row in range(1,sheet.nrows): 
    image = Image.open(imgRaw)
    image_size = image.size
    
    # profile pic adding
    # curThumbUrl = profileList[row]["picture"]["medium"]
    curThumbUrl = "https://randomuser.me/api/portraits/med/men/"+str(unqRandList[row-1])+".jpg"
    im_thumb = Image.open(BytesIO(requests.get(curThumbUrl).content))
    # im_square = crop_max_square(im_thumb).resize(im_thumb.size, Image.LANCZOS)
    im_circle = mask_circle_solid(im_thumb, (249, 249, 249), 0)
    image.paste(im_circle, (45,900))
    
    draw = ImageDraw.Draw(image)
    dateTime = str(sheet.cell_value(row,0))
    amount = '₹ '+str(sheet.cell_value(row,1))
    amount = 'Rs. '+str(sheet.cell_value(row,1))
    carName = str(sheet.cell_value(row,2))
    fromAdd = text_wrap(str(sheet.cell_value(row,3)),fontCL,image_size[0]-100)
    toAdd = text_wrap(str(sheet.cell_value(row,4)),fontCL,image_size[0]-100)
    driverName = str(sheet.cell_value(row,5))
    travelType = str(sheet.cell_value(row,6))
    fileName = str(sheet.cell_value(row,7))
    # for datetime
    draw.text((42, 312), dateTime, fill=black, font=fontCR)
    draw.text((850, 312), amount, fill=black, font=fontCR)
    draw.text((830, 1390), amount, fill=black, font=fontCLB)
    draw.text((830, 1490), amount, fill=black, font=fontCLB)
    draw.text((830, 1590), amount, fill=black, font=fontCLB)
    draw.text((830, 1690), amount, fill=black, font=fontCLB)
    draw.text((42, 395), carName, fill=grey, font=fontCL)
    draw.text((970, 395), "Cash", fill=grey, font=fontCL)
    x = 82
    y = 495
    for line in fromAdd:
        draw.text((x, y), line, fill=grey, font=fontCL)
        y = y + fontCL.getsize('hg')[1]
    y = 680
    for line in toAdd:
        draw.text((x, y), line, fill=grey, font=fontCL)
        y = y + fontCL.getsize('hg')[1]
    # draw.text((42, 530), fromAdd, fill=grey, font=fontCL)
    # draw.text((42, 700), toAdd, fill=grey, font=fontCL)
    draw.text((130, 915), "Your Ride With "+driverName, fill=black, font=fontCLB)
    draw.text((80, 1280), travelType+" Receipt", fill=black, font=fontCLB)
    draw.text((80, 1390), "Trip Fare", fill=black, font=fontCLB)
    draw.text((80, 1490), "Sub Total", fill=black, font=fontCLB)
    draw.text((80, 1590), "Total", fill=black, font=fontCLB)
    draw.text((130, 1690), "Cash", fill=black, font=fontCLB)

    
    image.save(fileName+'.png', optimize=True, quality=20)
    print("saved "+fileName+'.png')
print("completed...")


 