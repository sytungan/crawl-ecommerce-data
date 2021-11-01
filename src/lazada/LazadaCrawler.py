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
        allProduct = masterDriver.find_elements(By.CLASS_NAME, "_3VkVO")
        
        print("-----------------------------------------------")
        if not(len(allProduct)):
            print("THE END")
            break
        print("Page: ", page)
        print("Amount: ", len(allProduct))

        for product in allProduct:
            try:
                # productUrl = product.get_attribute('href')
                productUrl = product.find_elements(By.CSS_SELECTOR,"[href]")[1].get_attribute('href')
                slaveDriver.get(productUrl)
                time.sleep(1)

                productName = slaveDriver.find_element(By.XPATH, "//*[@class='pdp-mod-product-badge-wrapper']/h1").text

                # Price
                productCurrentPrice = slaveDriver.find_element(By.XPATH, "//span[@class=' pdp-price pdp-price_type_normal pdp-price_color_orange pdp-price_size_xl']").text
                try:
                    productOriPrice = slaveDriver.find_element(By.XPATH, "//span[@class=' pdp-price pdp-price_type_deleted pdp-price_color_lightgray pdp-price_size_xs']").text
                    productDiscountRate = slaveDriver.find_element(By.XPATH, "//span[@class='pdp-product-price__discount']").text
                except NoSuchElementException as Exception:
                    productOriPrice = productCurrentPrice
                    productDiscountRate = "0%"
                except StaleElementReferenceException as Exception:
                    productOriPrice = productCurrentPrice
                    productDiscountRate = "0%"
                
            
                # print(productCurrentPrice, productOriPrice, productDiscountRate)
                
                # Description
                btn = slaveDriver.find_element(By.CLASS_NAME,'expand-button').find_element(By.TAG_NAME, "button").send_keys("\n")
                time.sleep(1)
                productDescriptionContent = slaveDriver.find_element(By.CLASS_NAME, 'html-content.detail-content').text
                
                try:
                    table = slaveDriver.find_element(By.CLASS_NAME,'specification-keys').find_elements(By.TAG_NAME, 'li')
                    sHeading = ''
                    sContent = ''
                    for ele in table:
                        sHeading +=ele.find_element(By.CLASS_NAME,'key-title').text + '\n'
                        sContent +=ele.find_element(By.CLASS_NAME,'key-value').text + '\n'
                    
                    productDescriptionTableHeading = sHeading
                    productDescriptionTableContent = sContent
                
                except NoSuchElementException as Exception:
                    productDescriptionTableHeading = ''
                    productDescriptionTableContent = ''
                    


                # print(productDescriptionContent, productDescriptionTableHeading, productDescriptionTableContent)

                # break
                # Ratingcount
                try:
                    productRatingPoint = slaveDriver.find_element(By.CLASS_NAME, 'score-average').text
                    productRatingTotal = slaveDriver.find_element(By.CLASS_NAME, 'summary').find_element(By.CLASS_NAME, 'count').text.split(' ')[0]
                except NoSuchElementException as Exception:
                    productRatingPoint = "" 
                    productRatingTotal = ""

                # print(productRatingPoint, productRatingTotal)

                # Image
                IMG_SIZE = '500x500q80'
                imageContainer = slaveDriver.find_element(By.CLASS_NAME,'next-slick-track').find_elements(By.CLASS_NAME,"item-gallery__image-wrapper")
                sImage = ''
                for img in imageContainer[:-1]:
                    imgUrl = img.find_element(By.TAG_NAME,'img').get_attribute('src')
                    imgUrl = imgUrl.replace('120x120q80',IMG_SIZE)
                    sImage += imgUrl + '\n'

                productImageUrl = sImage 
                

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
        masterDriver.get(startingUrl + "&page=" + str(page))
        print(startingUrl + "&page=" + str(page))
        time.sleep(1)
        # break

book.save("data.xls")
print("Completed")
masterDriver.quit()
slaveDriver.quit()