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

    page = 1; row = 2
    for x in range (200): #maximum page to crawl is 200
        time.sleep(2)
        allProduct = masterDriver.find_elements(By.CLASS_NAME, 'col-xs-2-4.shopee-search-item-result__item')
        print("-----------------------------------------------")
        if not(len(allProduct)):
            print("THE END")
            break
        print("Page: ", page)
        print("Amount: ", len(allProduct))

        for product in allProduct:
            try:
                productUrl = product.find_element(By.TAG_NAME, 'a').get_attribute('href')
                slaveDriver.get(productUrl)
                time.sleep(1)

                productName = slaveDriver.find_element(By.CLASS_NAME, 'attM6y').find_element(By.TAG_NAME, 'span').text
                # print(productName)
                

                # Price
                productCurrentPrice = slaveDriver.find_element(By.CLASS_NAME, 'Ybrg9j').text
                try:
                    productOriPrice = slaveDriver.find_element(By.CLASS_NAME, '_2MaBXe').text
                    productDiscountRate = slaveDriver.find_element(By.CLASS_NAME, '_3LRxdy').text.split(' ')[0]
                except NoSuchElementException as Exception:
                    productOriPrice = productCurrentPrice
                    productDiscountRate = "0%"
                except StaleElementReferenceException as Exception:
                    productOriPrice = productCurrentPrice
                    productDiscountRate = "0%"
                
                # print(productCurrentPrice, productOriPrice, productDiscountRate)

                # Description
                # productDescription = slaveDriver.find_elements(By.CLASS_NAME, '_3wdEZ5')
                productDescriptionContent = slaveDriver.find_element(By.CLASS_NAME, '_3yZnxJ').text
                try:
                    table = slaveDriver.find_elements(By.CLASS_NAME, 'aPKXeO')
                    sHeading = ''
                    sContent = ''
                    for idx, val in enumerate(table):
                        if idx == 0:
                            continue
                        sContent += val.find_element(By.CLASS_NAME, 'SFJkS3').text + '\n'
                        sHeading += val.find_element(By.TAG_NAME, 'div').text + '\n'
                    productDescriptionTableHeading = sHeading
                    productDescriptionTableContent = sContent

                except NoSuchElementException as Exception:
                    productDescriptionTableHeading = ''
                    productDescriptionTableContent = ''


                # print(productDescriptionContent, productDescriptionTableHeading, productDescriptionTableContent)
                
                # Rating
                try:
                    productRatingPoint = slaveDriver.find_element(By.CLASS_NAME, 'OitLRu._1mYa1t').text
                    if (len(slaveDriver.find_elements(By.CLASS_NAME, 'OitLRu')) > 1):
                        productRatingTotal = slaveDriver.find_elements(By.CLASS_NAME, 'OitLRu')[1].text
                    else:
                        productRatingTotal = ""
                except NoSuchElementException as Exception:
                    productRatingPoint = "" 
                    productRatingTotal = ""

                # print(productRatingPoint, productRatingTotal)

                # Image
                sImage = ''
                # try:
                #     imageEle = slaveDriver.find_element(By.CLASS_NAME, '_3Q7kBy._2GchKS')
                #     slaveDriver.execute_script("arguments[0].click();", imageEle)
                #     time.sleep(1)
                #     imageContainer = slaveDriver.find_element(By.CLASS_NAME, 'pWIaLy')
                #     print(imageContainer.get_attribute('innerHTML'))
                #     imageList = imageContainer.find_elements(By.TAG_NAME, 'img')
                #     print(len(imageList))
                #     for img in imageList[:-1]:
                #         imgUrl = img.find_element(By.TAG_NAME, 'img').get_attribute('src')
                #         imgUrl = imgUrl.replace('https://salt.tikicdn.com/cache/100x100', preImgAPI)
                #         sImage += imgUrl + '\n'
                # except NoSuchElementException as Exception:
                #     slaveDriver.find_element(By.CLASS_NAME, '_1T-M8Y._3VLS5X').click()


                productImageUrl = sImage 
                # print(productImageUrl)
                

                # Save Data
                currentSheet.write(row, 0, productName)
                currentSheet.write(row, 1, productCurrentPrice)
                currentSheet.write(row, 2, productOriPrice)
                currentSheet.write(row, 3, productDiscountRate)
                currentSheet.write(row, 4, productDescriptionTableHeading)
                currentSheet.write(row, 5, productDescriptionTableContent)
                currentSheet.write(row, 6, productDescriptionContent)
                currentSheet.write(row, 7, productRatingPoint)
                currentSheet.write(row, 8, productRatingTotal)
                currentSheet.write(row, 9, productImageUrl)
                currentSheet.write(row, 10, productUrl)
                row = row + 1
                # break

            except StaleElementReferenceException as Exception:
                print("Making except")
            except NoSuchElementException as Exception:
                print("No element")
            # break

        page = page + 1
        print("Next page")
        if startingUrl.find("hamburger_menu_fly_out_banner") != -1:
            masterDriver.get(startingUrl + "&page=" + str(page-1))
        else:
            masterDriver.get(startingUrl + "?page=" + str(page-1))
        time.sleep(1)
        # break

book.save("data.xls")
print("Completed")
masterDriver.quit()
slaveDriver.quit()