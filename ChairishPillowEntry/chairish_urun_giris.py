import os
import smtplib

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep, time
from selenium.webdriver.support.select import Select
import pyautogui as pg
import pandas as pd

from ChairishPillowEntry import DominantColorRGBFinder, pillow_config


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


################################################################################################

file = pillow_config.file
dir_input = pillow_config.dir_input
imagesNotFound = []

start_time = time()
excel = pd.read_excel(file, index_col=None, header=None)


print("Upload Started!")
browser = webdriver.Chrome()
browser.get("https://www.chairish.com/account/login")
browser.maximize_window()
try:
    elem = browser.find_element_by_id("id_email")
    print("Test Pass : Email ID found")
except Exception as e:
    print("Exception found" + str(e))

elem.send_keys("hetyemez@yahoo.com")

try:
    elem = browser.find_element_by_id("id_password")

    print("Test Pass : Password ID found")
except Exception as e:
    print("Exception found" + str(e))
elem.send_keys("etyemez57")
elem.send_keys(Keys.ENTER)
sleep(2)

excel_row = pillow_config.startRow
number_of_row = pillow_config.endRow
if pillow_config.endRow == -1:
    number_of_row = excel.count().iloc[0]


try:
    while excel_row < number_of_row:
        sku_num = str(excel[0][excel_row])
        title = excel[8][excel_row]
        price = excel[4][excel_row]
        category = excel[10][excel_row]
        material = excel[2][excel_row]
        modifiedCategoryTuple = modifyCategory(category)
        title2 = modifiedCategoryTuple[0]
        inches = getInches(title2)

        print(str(excel_row))
        print(str(sku_num))
        browser.get("https://www.chairish.com/product/create")
        sleep(1)

        try:
            elem = browser.find_element_by_id("id_title")
            print("Test Pass : Title ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(title)

        try:
            elem = browser.find_element_by_id("id_categories")
            print("Test Pass : Categories ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('pillows')
        sleep(2.5)
        elem.send_keys(Keys.DOWN)
        sleep(0.5)
        elem.send_keys(Keys.ENTER)
        sleep(0.5)

        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-basic-fields"]/div[3]/div[2]/fieldset/div/div/div[2]/div[1]/label')
            print("Test Pass : Photo1 ID found")
        except Exception as e:
            print("Exception found" + str(e))
        elem.click()
        sleep(1)
        if os.path.isfile(dir_input + str(sku_num) + ".jpg"):
            pg.typewrite(dir_input + str(sku_num) + ".jpg")  # "C:\\Users\\hasan\\Desktop\\deneme\\a.jpg")
        else:
            imagesNotFound.append(sku_num)
            print("Bu ürünün resmi bulunamadı 1: " + str(sku_num) + " Satır numarası: " + str(excel_row))
            excel_row = excel_row + 1
            print("*******************************Sıradaki ürüne geçildi********************************************")
            continue
        sleep(1)
        pg.press("enter")
        sleep(1)
        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-basic-fields"]/div[3]/div[2]/fieldset/div/div/div[2]/div[2]/label')
            print("Test Pass : Photo2 ID found")
        except Exception as e:
            print("Exception found" + str(e))
        elem.click()
        sleep(1)
        if os.path.isfile(dir_input + str(sku_num) + "a.jpg"):
            pg.typewrite(dir_input + str(sku_num) + "a.jpg")
        else:
            imagesNotFound.append(sku_num)
            print("Bu ürünün resmi bulunamadı 2: " + str(sku_num) + " Satır numarası: " + str(excel_row))
            print("*******************************Sıradaki ürüne geçildi********************************************")
            excel_row = excel_row + 1
            continue
        sleep(1)
        pg.press("enter")
        sleep(1)
        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-basic-fields"]/div[3]/div[2]/fieldset/div/div/div[2]/div[4]/label')
            print("Test Pass : Photo3 ID found")
        except Exception as e:
            print("Exception found" + str(e))
        elem.click()
        sleep(1)
        if os.path.isfile(dir_input + str(sku_num) + "b.jpg"):
            pg.typewrite(dir_input + str(sku_num) + "b.jpg")
        else:
            imagesNotFound.append(sku_num)
            print("Bu ürünün resmi bulunamadı 2: " + str(sku_num) + " Satır numarası: " + str(excel_row))
            print("*******************************Sıradaki ürüne geçildi********************************************")
            excel_row = excel_row + 1
            continue
        sleep(1)
        pg.press("enter")
        sleep(1)

        try:
            elem = browser.find_element_by_id("id_description")
            print("Test Pass : Description ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(
            'This is a pillow cover made from a vintage kilim rug. The piece was properly washed and ready to use. '
            'Pillow inserts not included.')

        try:
            elem = browser.find_element_by_xpath('//*[@id="js-details-fields"]/div[1]/div[1]')
            print("Test Pass : Next button 1 ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()

        sleep(2)
        # =============================================================================
        #
        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/fieldset/div[1]/div/div[1]/label[2]/span[1]')
            print("Test Pass : Unknown checkbox ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()

        #
        try:
            elem = browser.find_element_by_id("id_styles")
            print("Test Pass : Styles ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('turkish')
        sleep(2)
        elem.send_keys(Keys.DOWN)
        sleep(0.5)
        elem.send_keys(Keys.ENTER)
        sleep(1)
        elem.send_keys('organic')
        sleep(1.5)
        elem.send_keys(Keys.DOWN)
        sleep(0.5)
        elem.send_keys(Keys.ENTER)

        try:
            elem = browser.find_element_by_id("id_materials")
            print("Test Pass : Materials ID found")
        except Exception as e:
            print("Exception found" + str(e))

        for mat in getMaterial(material):
            elem.send_keys(str(mat))
            sleep(2)
            elem.send_keys(Keys.DOWN)
            sleep(0.5)
            elem.send_keys(Keys.ENTER)

        # *********************************************************************
        try:
            elem = browser.find_element_by_id("id_colors")
            print("Test Pass : Colors ID found")
        except Exception as e:
            print("Exception found" + str(e))
        colors = DominantColorRGBFinder.dominantColorsSet(dir_input + str(sku_num) + ".jpg")
        print(colors, type(colors))
        if colors.__contains__('Dove Color'):
            colors.remove('Dove Color')
        for color in colors:
            print(color)
            elem.send_keys(color)
            sleep(2)
            elem.send_keys(Keys.DOWN)
            sleep(0.5)
            elem.send_keys(Keys.ENTER)

        # *********************************************************************

        dropDownId = 'id_origin_region'
        try:

            elem = browser.find_element_by_id(dropDownId)

            print("Test Pass : Origin Region ID found")
        except Exception as e:
            print("Exception found" + str(e))
        Select(elem).select_by_visible_text("Turkey")


        try:
            elem = browser.find_element_by_id("id_dimension_width")
            print("Test Pass : Dimension Width ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(inches[0])

        try:
            elem = browser.find_element_by_id("id_dimension_depth")
            print("Test Pass : Dimension Depth ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('0.2')

        try:
            elem = browser.find_element_by_id("id_dimension_height")
            print("Test Pass : Dimension Height ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(inches[1])

        dropDownId = 'id_period_made'
        try:
            elem = browser.find_element_by_id(dropDownId)
            print("Test Pass : Period Made ID found")
        except Exception as e:
            print("Exception found" + str(e))
        Select(elem).select_by_visible_text("2010s")

        # =============================================================================

        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/div[4]/fieldset/div[3]/ul/li[1]/label/span[1]')
            '//*[@id="id_used_code_0"]'
            print("Test Pass : Used Code SPAN ID found")
        except Exception as e:
            print("Exception found" + str(e))
        elem.click()
        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/div[4]/fieldset/div[4]/ul/li[1]/label/span[1]')
            print("Test Pass : Alterations Code SPAN ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()

        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/div[4]/fieldset/div[5]/ul/li[1]/label/span[1]')
            print("Test Pass : İmperfections Code SPAN ID found")
        except Exception as e:
            print("Exception found" + str(e))
        elem.click()

        # =============================================================================

        try:
            elem = browser.find_element_by_id("id_condition_notes")
            print("Test Pass : Condition Notes ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('In very good condition.')

        try:
            elem = browser.find_element_by_css_selector("div.form-actions:nth-child(8) > button:nth-child(1)")
            print("Test Pass : Next button 2 ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()
        sleep(2)

        try:
            elem = browser.find_element_by_id("id_sku")
            print("Test Pass : Sku ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(sku_num)

        try:
            elem = browser.find_element_by_id("id_price")
            print("Test Pass : Price ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(str(round(price * 1.2)))

        try:
            elem = browser.find_element_by_id("id_trade_discount_percent")
            print("Test Pass : Trade Discount Percent ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('10')

        try:
            elem = browser.find_element_by_id("id_reserve_price")
            print("Test Pass : Reserve Price ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(str(round(price * 1.2)))

        try:
            elem = browser.find_element_by_id("id_shipping_width")
            print("Test Pass : Shipping Width ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(inches[0])

        try:
            elem = browser.find_element_by_id("id_shipping_depth")
            print("Test Pass : Shipping Depth ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('0.2')

        try:
            elem = browser.find_element_by_id("id_shipping_height")
            print("Test Pass : Shipping Height ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(inches[1])

        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-logistics-fields"]/div[3]/div[3]/fieldset/div[1]/ul/li[2]/div/label/span[1]')
            print("Test Pass : Should Chairish Handle Shipping ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()
        sleep(1)

        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-seller-delegated-shipping-options"]/div/ul/li[2]/div/label/span[1]')
            print("Test Pass : Is Shipping Free ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()
        sleep(1)
        try:
            elem = browser.find_element_by_xpath('//*[@id="id_seller_delegated_shipping_charge"]')
            print("Test Pass : Seller Delegated Shipping Charge ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('8')

        try:
            elem = browser.find_element_by_xpath('/html/body/nav/div/div/div[2]/div[2]/button[2]')
            print("Test Pass : Submit ID found")
        except Exception as e:
            print("Exception found" + str(e))
        elem.click()  # submit butonuna tıklama"""
        sleep(6)

        excel_row = excel_row + 1
except Exception as e:
    print("Exception found" + str(e))
    send_mail(1, excel_row, sku_num)
send_mail(imagesNotFound, excel_row, sku_num)
print("--- %s seconds ---" % (time() - start_time))
