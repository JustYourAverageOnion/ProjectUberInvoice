import datetime, xlrd
import Image, ImageDraw, ImageFont
# https://www.geeksforgeeks.org/reading-excel-file-using-python/
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
    
    # save the edited image
    
    image.save(fileName+'.png')
    print("saved "+fileName+'.png')
print("completed...")


 