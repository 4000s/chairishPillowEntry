import os
import smtplib

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep, time
from selenium.webdriver.support.select import Select
import pyautogui as pg
import pandas as pd

from ChairishRugEntry import rug_config
from ChairishRugEntry import sharedFunctions

################################################################################################

file = rug_config.file
dir_input = rug_config.dir_input
imagesNotFound = []

start_time = time()
excel = pd.read_excel(file, index_col=None, header=None)

print("Upload Started!")
browser = webdriver.Chrome()
browser.get("https://www.chairish.com/account/login")
browser.maximize_window()
############################  EMAIL  ####################################
try:
    elem = browser.find_element_by_id("id_email")
    print("Test Pass : Email ID found")
except Exception as e:
    print("Exception found" + str(e))

elem.send_keys("wovenyrugs@gmail.com")
############################  PASSWORD  ####################################
try:
    elem = browser.find_element_by_id("id_password")

    print("Test Pass : Password ID found")
except Exception as e:
    print("Exception found" + str(e))
elem.send_keys("boyabat57")
elem.send_keys(Keys.ENTER)
sleep(2)

excel_row = rug_config.startRow
number_of_row = rug_config.endRow
if rug_config.endRow == -1:
    number_of_row = excel.count().iloc[0]
############################START OF PRODUCT ENTRY####################################
try:
    while excel_row < number_of_row:
        sku_num = str(excel[0][excel_row])
        title = excel[1][excel_row]
        category = excel[2][excel_row]
        width = excel[3][excel_row]
        length = excel[4][excel_row]
        width_i = excel[5][excel_row]
        length_i = excel[6][excel_row]
        materials = excel[7][excel_row]
        styles = excel[9][excel_row]
        age = excel[11][excel_row]
        colors = excel[12][excel_row]
        price = excel[13][excel_row]
        description = excel[14][excel_row]

        print(str(excel_row))
        print(str(sku_num))
        browser.get("https://www.chairish.com/product/create")
        sleep(1)
        ############################ TITLE ####################################
        try:
            elem = browser.find_element_by_id("id_title")
            print("Test Pass : Title ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(title)
        ############################CATEGORIES####################################
        try:
            elem = browser.find_element_by_id("id_categories")
            print("Test Pass : Categories ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('rug')
        sleep(2.5)
        elem.send_keys(Keys.DOWN)
        sleep(0.5)
        elem.send_keys(Keys.ENTER)
        sleep(0.5)

        elem.send_keys('rug')
        sleep(2.5)
        elem.send_keys(Keys.DOWN)
        sleep(0.5)
        elem.send_keys(Keys.ENTER)

        ############################PHOTOS####################################
        image_path = ""
        for image in range(1, 12):
            sleep(1)

            if image == 1:
                image_path = dir_input + str(sku_num) + ".jpg" #+ "-" + str(width) + "x" + str(length)
            elif image == 2:
                image_path = dir_input + str(sku_num) + "a.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 3:
                image_path = dir_input + str(sku_num) + "b.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 4:
                image_path = dir_input + str(sku_num) + "c.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 5:
                image_path = dir_input + str(sku_num) + "d.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 6:
                image_path = dir_input + str(sku_num) + "e.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 7:
                image_path = dir_input + str(sku_num) + "f.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 8:
                image_path = dir_input + str(sku_num) + "g.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 9:
                image_path = dir_input + str(sku_num) + "h.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 10:
                image_path = dir_input + str(sku_num) + "i.jpg"#+ "-" + str(width) + "x" + str(length)
            elif image == 11:
                image_path = dir_input + str(sku_num) + "j.jpg"#+ "-" + str(width) + "x" + str(length)

            if os.path.isfile(image_path):
                try:
                    elem = browser.find_element_by_xpath(
                        '//*[@id="js-basic-fields"]/div[3]/div[2]/fieldset/div/div/div[2]/div[' + str(
                            image) + ']/label')
                    print("Test Pass : Photo ID found")
                except Exception as e:
                    print("Exception found" + str(e))
                elem.click()
                sleep(0.5)
                pg.typewrite(image_path)
                sleep(0.5)
                pg.press("enter")
                sleep(0.5)
            # else:
            # break
            # imagesNotFound.append(sku_num)
            # # print("Bu ürünün resmi bulunamadı 1: " + str(sku_num) + " Satır numarası: " + str(excel_row))
            # excel_row = excel_row + 1
            # print(
            #     "*******************************Sıradaki ürüne geçildi*****************************************")
            # continue

        # try:
        #     elem = browser.find_element_by_xpath(
        #         '//*[@id="js-basic-fields"]/div[3]/div[2]/fieldset/div/div/div[2]/div[2]/label')
        #     print("Test Pass : Photo2 ID found")
        # except Exception as e:
        #     print("Exception found" + str(e))
        # elem.click()
        # sleep(0.5)
        # if os.path.isfile(dir_input + str(sku_num) + "a.jpg"):
        #     pg.typewrite(dir_input + str(sku_num) + "a.jpg")
        # else:
        #     imagesNotFound.append(sku_num)
        #     print("Bu ürünün resmi bulunamadı 2: " + str(sku_num) + " Satır numarası: " + str(excel_row))
        #     print("*******************************Sıradaki ürüne geçildi********************************************")
        #     excel_row = excel_row + 1
        #     continue
        # sleep(0.5)
        # pg.press("enter")
        # sleep(0.5)
        # try:
        #     elem = browser.find_element_by_xpath(
        #         '//*[@id="js-basic-fields"]/div[3]/div[2]/fieldset/div/div/div[2]/div[4]/label')
        #     print("Test Pass : Photo3 ID found")
        # except Exception as e:
        #     print("Exception found" + str(e))
        # elem.click()
        # sleep(0.5)
        # if os.path.isfile(dir_input + str(sku_num) + "b.jpg"):
        #     pg.typewrite(dir_input + str(sku_num) + "b.jpg")
        # else:
        #     imagesNotFound.append(sku_num)
        #     print("Bu ürünün resmi bulunamadı 2: " + str(sku_num) + " Satır numarası: " + str(excel_row))
        #     print("*******************************Sıradaki ürüne geçildi********************************************")
        #     excel_row = excel_row + 1
        #     continue
        # sleep(0.5)
        # pg.press("enter")
        # sleep(0.5)
        ############################DESCRIPTION####################################
        try:
            elem = browser.find_element_by_id("id_description")
            print("Test Pass : Description ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(str(description))
        ############################NEXT BUTTON 1####################################
        try:
            elem = browser.find_element_by_xpath('//*[@id="js-details-fields"]/div[1]/div[1]')
            print("Test Pass : Next button 1 ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()

        sleep(2)
        # =============================================================================
        #
        ############################ UNKNOWN CHECKBOX ####################################
        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/fieldset/div[1]/div/div[1]/label[2]/span[1]')
            print("Test Pass : Unknown checkbox ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()

        ############################ STYLE ####################################
        try:
            elem = browser.find_element_by_id("id_styles")
            print("Test Pass : Styles ID found")
        except Exception as e:
            print("Exception found" + str(e))
        styles_d = sharedFunctions.getStyles(styles)

        for style in styles_d:
            elem.send_keys(str(style))
            sleep(2)
            elem.send_keys(Keys.DOWN)
            sleep(0.5)
            elem.send_keys(Keys.ENTER)

        ############################ MATERIALS ####################################
        try:
            elem = browser.find_element_by_id("id_materials")
            print("Test Pass : Materials ID found")
        except Exception as e:
            print("Exception found" + str(e))

        for mat in sharedFunctions.getMaterial(materials):
            elem.send_keys(str(mat))
            sleep(2)
            elem.send_keys(Keys.DOWN)
            sleep(0.5)
            elem.send_keys(Keys.ENTER)

        ############################ COLORS ####################################
        try:
            elem = browser.find_element_by_id("id_colors")
            print("Test Pass : Colors ID found")
        except Exception as e:
            print("Exception found" + str(e))
        colors, declined_colors = sharedFunctions.colorMaptoChairish(colors)
        print(colors, type(colors))

        for color in colors:
            print(color)
            elem.send_keys(color)
            sleep(2)
            elem.send_keys(Keys.DOWN)
            sleep(0.5)
            elem.send_keys(Keys.ENTER)

        # *********************************************************************
        ############################ ORIGIN REGION ####################################
        dropDownId = 'id_origin_region'
        try:

            elem = browser.find_element_by_id(dropDownId)

            print("Test Pass : Origin Region ID found")
        except Exception as e:
            print("Exception found" + str(e))
        Select(elem).select_by_visible_text("Turkey")
        ############################ SKU ####################################
        try:
            elem = browser.find_element_by_id("id_sku")
            print("Test Pass : Sku ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(sku_num)
        ############################ WIDTH ####################################
        try:
            elem = browser.find_element_by_id("id_dimension_width")
            print("Test Pass : Dimension Width ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(str(width_i))
        ############################ DEPTH ####################################
        try:
            elem = browser.find_element_by_id("id_dimension_depth")
            print("Test Pass : Dimension Depth ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('0.2')
        ############################ HEIGHT ####################################
        try:
            elem = browser.find_element_by_id("id_dimension_height")
            print("Test Pass : Dimension Height ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(str(length_i))

        ############################ PERIOD MADE ####################################
        dropDownId = 'id_period_made'
        try:
            elem = browser.find_element_by_id(dropDownId)
            print("Test Pass : Period Made ID found")
        except Exception as e:
            print("Exception found" + str(e))
        if age == 'Vintage':
            Select(elem).select_by_visible_text("1960s")
        elif age == 'Antique':
            Select(elem).select_by_visible_text("1910s")

        # =============================================================================
        ############################ CONDITION ####################################
        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/div[3]/fieldset/div[3]/ul/li[1]/label/span[1]')
            print("Test Pass : Used Code SPAN ID found")
        except Exception as e:
            print("Exception found" + str(e))
        elem.click()
        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/div[3]/fieldset/div[4]/ul/li[1]/label/span[1]')
            print("Test Pass : Alterations Code SPAN ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()

        try:
            elem = browser.find_element_by_xpath(
                '//*[@id="js-details-fields"]/div[3]/div[3]/fieldset/div[5]/ul/li[1]/label/span[1]')
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
        ############################ NEXT BUTTON 2 ####################################
        try:
            elem = browser.find_element_by_css_selector("div.form-actions:nth-child(8) > button:nth-child(1)")
            print("Test Pass : Next button 2 ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()
        sleep(2)
        ############################ PRICE ####################################
        try:
            elem = browser.find_element_by_id("id_price")
            print("Test Pass : Price ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(str(round(price * 1.2)))

        # try:
        #     elem = browser.find_element_by_id("id_trade_discount_percent")
        #     print("Test Pass : Trade Discount Percent ID found")
        # except Exception as e:
        #     print("Exception found" + str(e))
        #
        # elem.send_keys('15')

        try:
            elem = browser.find_element_by_id("id_reserve_price")
            print("Test Pass : Reserve Price ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys(str(round(price * 1.2)))
        ############################ SHIPPING ####################################
        # try:
        #     elem = browser.find_element_by_id("id_shipping_width")
        #     print("Test Pass : Shipping Width ID found")
        # except Exception as e:
        #     print("Exception found" + str(e))
        #
        # elem.send_keys(inches[0])
        #
        # try:
        #     elem = browser.find_element_by_id("id_shipping_depth")
        #     print("Test Pass : Shipping Depth ID found")
        # except Exception as e:
        #     print("Exception found" + str(e))
        #
        # elem.send_keys('0.2')
        #
        # try:
        #     elem = browser.find_element_by_id("id_shipping_height")
        #     print("Test Pass : Shipping Height ID found")
        # except Exception as e:
        #     print("Exception found" + str(e))
        #
        # elem.send_keys(inches[1])

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
            print("Test Pass : Is Shipping FRee ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.click()
        sleep(1)
        try:
            elem = browser.find_element_by_xpath('//*[@id="id_seller_delegated_shipping_charge"]')
            print("Test Pass : Seller Delegated Shipping Charge ID found")
        except Exception as e:
            print("Exception found" + str(e))

        elem.send_keys('45')
        ############################ SUBMIT ####################################
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
    sharedFunctions.send_mail(1, excel_row, sku_num)
# send_mail(imagesNotFound, excel_row, sku_num)
print("--- %s seconds ---" % (time() - start_time))
