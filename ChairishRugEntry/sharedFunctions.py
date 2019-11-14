import smtplib


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


def getMaterial(input):
    materials = input.split(',')
    result = []
    for mat in materials:
        if mat.__eq__("Wool"):
            result.append("Wool")
        elif mat.__eq__("Cotton"):
            result.append("Cotton")
        elif mat.__eq__("Cotton Fabric"):
            result.append("Fabric")
        elif mat.__eq__("Goat Hair"):
            result.append("Goat")
        elif mat.__eq__("Textile Recyle"):
            result.append("Recycle")

    return result


def getStyles(input):
    styles = input.split(',')
    result = []

    for style in styles:
        if style.__eq__("Mid-Century Modern"):
            result.append("Mid-Century Modern")
        elif style.__eq__("Bohemian"):
            result.append("Boho Chic")
        elif style.__eq__("Colorful"):
            result.append("Boho Chic")
            result.append("Traditional")
        elif style.__eq__("Floral"):
            result.append("Traditional")
        elif style.__eq__("Geometric"):
            result.append("Traditional")
            result.append("Mid-Century Modern")
        elif style.__eq__("Distressed"):
            result.append("Shabby Chic")
        elif style.__eq__("Abstract"):
            result.append("Abstract")
        else:
            result.append("Traditional")
            result.append("Contemporary")
            result.append("Mid-Century Modern")
    print(set(result[:3]))
    return set(result[:3])


def colorMaptoChairish(colorsFromExcel):
    defaultChairishColorList = ["Sky Blue", "Light Pink", "Aqua", "Beige", "Black", "Blue", "Violet", "Brown",
                                "Cinnamon",
                                "Royal Blue", "Yellow", "Brown", "Orange", "Vanilla", "Red", "Navy", "Teal", "Chestnut",
                                "Dark Gray", "Khaki", "Magenta", "Dark Green", "Ruby Red", "Salmon", "Green",
                                "Turquoise",
                                "Bright Pink", "Gray", "Brick Red", "White", "Pink", "Gold", "Bright Green", "Ivory",
                                "Lavender",
                                "Beige", "Raspbery Pink", "Scarlet", "Kelly Green", "Cerulean", "Light Yellow",
                                "Neon Green",
                                "Kelly Green", "Linen", "Maroon", "Purple", "Light Green", "Ink Blue", "Cream", "Rose",
                                "Mustard",
                                "Navy Blue", "Bubble Gum", "Tangerine", "Brown", "Raspberry Red", "Chocolate",
                                "Forest Green",
                                "Silver", "Tan", "Kelly Green", "Black", "Burgundy"]

    colorList = [x.strip() for x in colorsFromExcel.split(',')]

    resultColorList = []
    unFoundColrsList = []
    for color in colorList:
        isFound = False
        for chairishColor in defaultChairishColorList:
            if color == chairishColor:
                resultColorList.append(chairishColor)
                isFound = True
                break
        if not isFound:
            unFoundColrsList.append(color)
            if color == "Multicolor":
                resultColorList.append("red")
                resultColorList.append("green")
                resultColorList.append("blue")

    return resultColorList, unFoundColrsList

def cm2Inches(cm):
    return int(cm/2.54)

#
# ####################################Test##################################################
# print (colorMaptoChairish("Blue, Red,Silver,red"))
#
# file = 'C:\\Users\\okkaa\\Desktop\\2573-2633.xlsx'
#
# excel = pd.read_excel(file, index_col=None, header=None)
# total_number_of_row = excel.count().iloc[0]
#
#
# row=0
# while row < total_number_of_row:
#     colors = excel[12][row]
#     print (colorMaptoChairish(colors))
#     row= row +1
