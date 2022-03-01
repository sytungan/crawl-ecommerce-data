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

    try:
        page = 1; row = 2
        for x in range (200): #maximum page to crawl is 200
            allProduct = masterDriver.find_elements(By.CLASS_NAME, "product-item")

            print("-----------------------------------------------")
            if not(len(allProduct)):
                print("THE END")
                break
            print("Page: ", page)
            print("Amount: ", len(allProduct))

            for product in allProduct:
                while True:
                    try:
                        productUrl = product.get_attribute('href')
                        slaveDriver.get(productUrl)
                        time.sleep(1)

                        print("Visit " + productUrl)
                        productName = slaveDriver.find_element(By.CLASS_NAME, 'title').text

                        # Price
                        productCurrentPrice = slaveDriver.find_element(By.CLASS_NAME, 'product-price__current-price').text
                        try:
                            productOriPrice = slaveDriver.find_element(By.CLASS_NAME, 'product-price__list-price').text
                            productDiscountRate = slaveDriver.find_element(By.CLASS_NAME, 'product-price__discount-rate').text
                        except NoSuchElementException as Exception:
                            productOriPrice = productCurrentPrice
                            productDiscountRate = "0%"
                        except StaleElementReferenceException as Exception:
                            productOriPrice = productCurrentPrice
                            productDiscountRate = "0%"
                        
                        # print(productCurrentPrice, productOriPrice, productDiscountRate)

                        # Description
                        productDescription = slaveDriver.find_elements(By.CLASS_NAME, 'content')
                        slaveDriver.find_element(By.CLASS_NAME, 'btn-more').click()
                        try:
                            table = productDescription[0].find_elements(By.TAG_NAME, 'td')
                            sHeading = ''
                            sContent = ''
                            for idx, val in enumerate(table):
                                if (idx % 2):
                                    sContent += val.text + '\n'
                                else:
                                    sHeading += val.text + '\n'
                            productDescriptionTableHeading = sHeading
                            productDescriptionTableContent = sContent
                            if (len(productDescription) > 1):
                                productDescriptionContent = productDescription[1].text
                            else:
                                productDescriptionContent = ""
                        except NoSuchElementException as Exception:
                            productDescriptionTableHeading = ''
                            productDescriptionTableContent = ''
                            productDescriptionContent = productDescription[0].text


                        # print(productDescriptionContent, productDescriptionTableHeading, productDescriptionTableContent)

                        # Rating
                        try:
                            productRatingPoint = slaveDriver.find_element(By.CLASS_NAME, 'review-rating__point').text
                            productRatingTotal = slaveDriver.find_element(By.CLASS_NAME, 'review-rating__total').text.split(' ')[0]
                        except NoSuchElementException as Exception:
                            productRatingPoint = "" 
                            productRatingTotal = ""

                        # print(productRatingPoint, productRatingTotal)

                        # Image
                        IMG_SIZE = '500x500'
                        preImgAPI = "https://salt.tikicdn.com/cache/" + IMG_SIZE
                        sImage = ''
                        imageContainer = slaveDriver.find_element(By.CLASS_NAME, 'review-images__list')
                        imageList = imageContainer.find_elements(By.TAG_NAME, 'a')
                        for img in imageList[:-1]:
                            imgUrl = img.find_element(By.TAG_NAME, 'img').get_attribute('src')
                            imgUrl = imgUrl.replace('https://salt.tikicdn.com/cache/100x100', preImgAPI)
                            sImage += imgUrl + '\n'

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
                        print("Making except -- Try again")
                        continue
                    except NoSuchElementException as Exception:
                        print("No element")
                    break

            page = page + 1
            print("Next page")
            if startingUrl.find("&") != -1:
                masterDriver.get(startingUrl + "&page=" + str(page))
            else:
                masterDriver.get(startingUrl + "?page=" + str(page))
            time.sleep(1)
            # break
    except:
        break

book.save("data.xls")
print("Completed")
masterDriver.quit()
slaveDriver.quit()