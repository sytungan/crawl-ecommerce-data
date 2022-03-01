import xlwt
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from xlwt.BIFFRecords import SaveRecalcRecord
from xlwt.Style import add_palette_colour
from dotenv import load_dotenv
import os

f = open("input.txt", 'r')

book = xlwt.Workbook(encoding="utf-8")

load_dotenv()
# chromeDriver = os.getenv("CHROME_DRIVER")

# Setting 
cService = Service(ChromeDriverManager().install())
options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")
options.add_argument("--log-level=3")
masterDriver = webdriver.Chrome(options=options, service=cService)
slaveDriver = webdriver.Chrome(options=options, service=cService)

numberUrl = int(f.readline())
for num in range (numberUrl) :
    f.readline()
    currentSheet = book.add_sheet("Sheet " + str(num))
    nameFile = str(f.readline())
    currentSheet.write(0, 0, nameFile)
    currentSheet.write(1, 0, "Name")
    currentSheet.write(1, 1, "Current_Price")
    currentSheet.write(1, 2, "Original_Price")
    currentSheet.write(1, 3, "Discount")
    currentSheet.write(1, 4, "Description_Table_Heading")
    currentSheet.write(1, 5, "Description_Table_Content")
    currentSheet.write(1, 6, "Description_Content")
    currentSheet.write(1, 7, "Rating_Point")
    currentSheet.write(1, 8, "Rating_Total")
    currentSheet.write(1, 9, "Image_URL")
    currentSheet.write(1, 10, "Product_URL")

    startingUrl = f.readline()
    masterDriver.get(startingUrl)
    time.sleep(1)

    page = 1; row = 2
    for x in range (200): #maximum page to crawl is 200
        allProduct = masterDriver.find_elements(By.CLASS_NAME, 'item_3x07')

        print("-----------------------------------------------")
        if not(len(allProduct)):
            print("THE END")
            break
        print("Page: ", page)
        print("Amount: ", len(allProduct))

        for product in allProduct:
            while True:
                try:
                    productUrlTemp = product.get_attribute('href')
                    productUrl = productUrlTemp.split('?')[0]
                    #productUrl = "https://www.sendo.vn/dong-ho-dinh-vi-tre-em-z6-q19-co-tieng-viet-lap-sim-nghe-goi-doc-lap-41826659.html"
                    #productUrl = "https://www.sendo.vn/dong-ho-cam-ung-dz09-va-z6-38681154.html?source_block_id=listing_products&source_info=desktop2_60___session_key___d9c4a5f0-5516-41c5-ad84-77e2cc899c93_5_algo13_0_14_7_-1&source_page_id=cate2_listing_v2_desc"
                    slaveDriver.get(productUrl)
                    time.sleep(1)
                    print(productUrl)
                    productName = slaveDriver.find_element(By.CLASS_NAME, 'd7e-ed528f.d7e-7dcda3.d7e-f56b44.d7e-fb1c84.undefined').text
                    # Price
                    productCurrentPrice = slaveDriver.find_element(By.CLASS_NAME, 'd7e-87b451.d7e-fb1c84.d7e-a4f628').text

                    #print(productUrl)
                    try:
                        productOriPrice = slaveDriver.find_element(By.CLASS_NAME, 'd7e-d87aa1.d7e-b61d5e.d7e-e3a0b4').text
                        productDiscountRate = slaveDriver.find_element(By.CLASS_NAME, '_314-fa4b74.d7e-d87aa1.d7e-b61d5e.d7e-a4f628').text
                    except NoSuchElementException as Exception:
                        productOriPrice = productCurrentPrice
                        productDiscountRate = "0%"
                    except StaleElementReferenceException as Exception:
                        productOriPrice = productCurrentPrice
                        productDiscountRate = "0%"
                    
                    # Description
                    try:
                        productDescription = slaveDriver.find_elements(By.CLASS_NAME, '_96e-5d268c')[0].text
                    except NoSuchElementException as Exception:
                        productDescription = ""
                    
                    print(productName)

                    a = slaveDriver.find_elements(By.CLASS_NAME, 'd7e-aa34b6.d7e-1b9468.d7e-13f811.d7e-f99ea6.d7e-dc4b7b')
                    a[2].click()

                    try:
                        table = slaveDriver.find_elements(By.CLASS_NAME, 'd7e-ed528f.d7e-fde242.d7e-d87aa1.d7e-b61d5e.d7e-a58302')
                        sHeading = ''
                        sContent = ''
                        for idx, val in enumerate(table):
                            if (idx % 2):
                                sContent += val.text + '\n'
                            else:
                                sHeading += val.text + '\n'
                        productDescriptionTableHeading = sHeading
                        productDescriptionTableContent = sContent
                    except NoSuchElementException as Exception:
                        productDescriptionTableHeading = ''
                        productDescriptionTableContent = ''
                        #productDescriptionContent = productDescription[0].text
                    print("table")


                    #Rating
                    try:
                        productRatingPoint = slaveDriver.find_element(By.CLASS_NAME, 'undefined.d7e-922765.d7e-fb1c84').text
                        productRatingTotal = slaveDriver.find_element(By.CLASS_NAME, '_39a-b49a28').text.split(' ')[8]
                    except NoSuchElementException as Exception:
                        productRatingPoint = "" 
                        productRatingTotal = ""

                    try:
                        content = slaveDriver.find_element(By.CLASS_NAME, 'd7e-f7453d.d7e-57f266.undefined.d7e-d87aa1.d7e-b61d5e').text
                    except NoSuchElementException as Exception:
                        content = ""

                    if (content != ""):
                        productDescription = content
                    print("content")


                    # Image
                    IMG_SIZE = '500x500'
                    sImage = ''
                    imageContainer = slaveDriver.find_elements(By.CLASS_NAME, 'swiper-wrapper')[1]
                    imageList = imageContainer.find_elements(By.TAG_NAME, 'img')
                    for img in imageList[:-1]:
                        imgUrl = img.get_attribute('src')
                        imgUrl = imgUrl.replace('100x100', IMG_SIZE)
                        sImage += imgUrl + '\n'

                    productImageUrl = sImage 
                    print("productImageUrl")

                    # Save Data
                    currentSheet.write(row, 0, productName)
                    currentSheet.write(row, 1, productCurrentPrice)
                    currentSheet.write(row, 2, productOriPrice)
                    currentSheet.write(row, 3, productDiscountRate)
                    currentSheet.write(row, 4, productDescriptionTableHeading)
                    currentSheet.write(row, 5, productDescriptionTableContent)
                    currentSheet.write(row, 6, productDescription)
                    currentSheet.write(row, 7, productRatingPoint)
                    currentSheet.write(row, 8, productRatingTotal)
                    currentSheet.write(row, 9, productImageUrl) 
                    currentSheet.write(row, 10, productUrl)
                    row = row + 1
                    # break
                    print("done")

                except StaleElementReferenceException as Exception:
                    print("Making except -- Try again")
                    continue
                except NoSuchElementException as Exception:
                    print("No element")
                except Exception:
                    pass
                break

        page = page + 1
        print("Next page")
        if startingUrl.find("hamburger_menu_fly_out_banner") != -1:
            masterDriver.get(startingUrl + "&page=" + str(page))
        else:
            masterDriver.get(startingUrl + "?page=" + str(page))
        time.sleep(1)
        # break

book.save("data.xls")
print("Completed")
masterDriver.quit()
slaveDriver.quit()