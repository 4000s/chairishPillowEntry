import pandas as pd
import os
import smtplib
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep, time
from selenium.webdriver.support.select import Select
import pyautogui as pg

# if any exception is happened, this function sends mail with short explanation
def send_mail(imagesNotFound=None, product=None, sku_num=None):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login("hof2018hof@gmail.com", "hof12345")
    if imagesNotFound is None and product is None and sku_num is None:
        msg = "Hata!"
    elif product is None and sku_num is None:
        msg = "Hata! \n Image List Number : " + str(imagesNotFound)
    elif sku_num is None:
        msg = "Hata! \nImage List Number : " + str(imagesNotFound) + " Product : " + str(product)
    else:
        msg = "Hata! \nImage List Number : " + str(imagesNotFound) + " Product : " + str(product) + " SKU : " + str(
            sku_num)

    server.sendmail("hof2018hof@gmail.com", "hasanemreari@gmail.com", msg)
    print(msg)
    server.quit()


def modifyCategory(category):
    # print(nonpillow)
    cat = category
    # print(cat)
    b = str(cat).split('(')
    inche = (b[0])
    # print(inche)

    char = []
    catList = []
    for i in inche:
        char.append(i)
    for t in cat:
        catList.append(t)

    if char.__contains__('\xa0'):
        char.remove('\xa0')
    if char.__contains__('\xa0'):
        char.remove('\xa0')
    if char.__contains__(' '):
        char.remove(' ')
    if char.__contains__(' '):
        char.remove(' ')

    if catList.__contains__('\xa0'):
        catList.remove('\xa0')
    if catList.__contains__('\xa0'):
        catList.remove('\xa0')
    if catList.__contains__(' '):
        catList.remove(' ')
    if catList.__contains__(' '):
        catList.remove(' ')

    result1 = ''
    result2 = ''
    for j in char:
        result1 += j

    for s in catList:
        result2 += s
    return str(result1), str(result2)


def getInches(input):
    return str(input[0:2]), str(input[4:6])

def getMaterial(input):
    materials = input.split(',')
    return materials

def colorMaptoChairish(colorsFromExcel):

    defaultChairishColorList = ["Sky Blue","Light Pink","Aqua","Beige","Black","Blue","Violet","Brown","Cinnamon",
                                "Royal Blue","Yellow","Brown", "Orange","Vanilla","Red","Navy","Teal","Chestnut",
                                "Dark Gray","Khaki","Magenta","Dark Green","Ruby Red","Salmon","Green","Turquoise",
                                "Bright Pink","Gray","Brick Red","White","Pink","Gold","Bright Green","Ivory","Lavender",
                                "Beige","Raspbery Pink","Scarlet","Kelly Green","Cerulean","Light Yellow","Neon Green",
                                "Kelly Green","Linen","Maroon","Purple","Light Green","Ink Blue","Cream","Rose","Mustard",
                                "Navy Blue","Bubble Gum","Tangerine","Brown","Raspberry Red","Chocolate", "Forest Green",
                                "Silver","Tan","Kelly Green","Black","Burgundy"]

    colorList = [x.strip() for x in colorsFromExcel.split(',')]

    resultColorList = []
    unFoundColrsList =[]
    for color in colorList:
        isFound = False
        for chairishColor in defaultChairishColorList:
            if color == chairishColor:
                resultColorList.append(chairishColor)
                isFound = True
                break
        if not isFound:
            unFoundColrsList.append(color)

    return  resultColorList,unFoundColrsList






####################################Test##################################################
print (colorMaptoChairish("Blue, Red,Silver,red"))

file = 'C:\\Users\\okkaa\\Desktop\\2573-2633.xlsx'

excel = pd.read_excel(file, index_col=None, header=None)
total_number_of_row = excel.count().iloc[0]


row=0
while row < total_number_of_row:
    colors = excel[12][row]
    print (colorMaptoChairish(colors))
    row= row +1
