# import requests
import shutil
from bs4 import BeautifulSoup
from xlrd import open_workbook
from xlutils.copy import copy
import os

ALIEXPRESS_PATH = "./ResultData"
ALIEXPRESS_PROFUCT_PATH = "/Users/clp/Documents/AliExpress/ProductImage"
ALIEXPRESS_IMAGE_PATH = ALIEXPRESS_PATH + "/ProductImage"
ALIEXPRESS_INFO_PATH = ALIEXPRESS_PATH + "/ProductInfo.xls"
ALIEXPRESS_IMAGE_TEMPLATE_PATH = ALIEXPRESS_PATH + "/ProductImageTemplate"
HTM_SAVE_FOLDER = "./1688"

PRODUCT_CLASS = {
    "smartwatch": "智能手表",
    "earphone": "耳机"
}


class ProductInfo:
    className = ""
    productName = ""
    costPrice = 0
    weight = 200
    withPackingWeight = 200
    withDiscountPrice = 0
    margin = 0.25
    PlatformPumping = 0.08
    sellPrice = 0
    exchangeRate = 6.5
    logisticsPrice = 0.13
    packingCost = 0.93
    address1688 = ""
    productImageFolder = ""


def getCarouselImagePath(number):
    return ALIEXPRESS_IMAGE_PATH + "/" + number + "/Carousel"


def getDetailImagePath(number):
    return ALIEXPRESS_IMAGE_PATH + "/" + number + "/Detail"


def getProductSheet(folder=ALIEXPRESS_INFO_PATH, sheetNumber=0):
    productExcel = open_workbook(folder)
    productSheet = productExcel.sheet_by_index(sheetNumber)
    return productSheet


def openProductNumber(folder=ALIEXPRESS_INFO_PATH, sheetNumber=0):
    productSheet = getProductSheet(folder, sheetNumber)
    rowSize = productSheet.nrows
    return {
        "rowSize": rowSize,
        "productNumber": str(1000 + rowSize - 1),
        "nextProductNumber": str(1000 + rowSize)
    }


def writeProductInfo(folder=ALIEXPRESS_INFO_PATH, sheetNumber=0, products=[]):
    productExcel = open_workbook(folder, )
    startRowIndex = productExcel.sheet_by_index(sheetNumber).nrows
    editingProductExcel = copy(productExcel)
    editingSheet = editingProductExcel.get_sheet(sheetNumber)
    for i in range(len(products)):
        products[i].productName = getProductName(1000 + startRowIndex + i, products[i].className)
        products[i].productImageFolder = ALIEXPRESS_PROFUCT_PATH + "/" + products[i].productName
        editingSheet.write(startRowIndex + i, 0, products[i].className)
        editingSheet.write(startRowIndex + i, 1, products[i].productName)
        editingSheet.write(startRowIndex + i, 2, products[i].costPrice)
        editingSheet.write(startRowIndex + i, 14, products[i].address1688)
        editingSheet.write(startRowIndex + i, 15, products[i].productImageFolder)
    editingProductExcel.save(folder)
    # productExcel.close()
    return products


def getHtm(file):
    print(file)
    return BeautifulSoup(open(file, "r", encoding='utf-8'), 'html.parser')


def getProductInfoFrom1688(file):
    productSoup = getHtm(file)
    product = ProductInfo()
    priceText = productSoup.select("tr[class='price'] .value")[0].text
    product.costPrice = priceText
    product.address1688 = productSoup.select("link[rel='canonical']")[0].attrs["href"]
    return product


def getProductName(rankNumber, className):
    return str(rankNumber) + "-" + className


def saveDetailImage(file, productNumber):
    mkdirFolderByProductNumber(productNumber=productNumber)
    soup = getHtm(file)
    images = soup.select("#desc-lazyload-container p img")
    for index in range(len(images)):
        image = images[index]
        imageSrc = image.attrs["src"]
        imageSrcSplitList = imageSrc.split(".")
        imageSrcSuffix = imageSrcSplitList[len(imageSrcSplitList) - 1]
        shutil.copy(HTM_SAVE_FOLDER + imageSrc[1:],
                    getDetailImagePath(productNumber) + "/" + str(index) + "." + imageSrcSuffix)


def getHtmFile(folder=HTM_SAVE_FOLDER):
    filesList = os.listdir(folder)
    htmFileslist = []
    for file in filesList:
        if (file.endswith(".htm")):
            htmFileslist.append(file)
    return htmFileslist


def checkImageFolderExist(fileName):
    return os.path.exists(ALIEXPRESS_IMAGE_PATH + "/" + fileName)


def mkdirFolderByProductNumber(source=ALIEXPRESS_IMAGE_TEMPLATE_PATH,
                               destination=ALIEXPRESS_IMAGE_PATH,
                               productNumber="1000"):
    print(source, destination, destination + "/" + productNumber)
    return shutil.copytree(source, destination + "/" + productNumber)

def main(className):
    fileList = getHtmFile()
    productList = []
    for file in fileList:
        product = getProductInfoFrom1688(HTM_SAVE_FOLDER + "/" + file)
        product.className = className
        productList.append(product)
    productListWithNumber = writeProductInfo(products=productList)
    for index in range(len(productListWithNumber)):
        filePath = HTM_SAVE_FOLDER + "/" + fileList[index]
        saveDetailImage(filePath, productListWithNumber[index].productName)

main("smartwatch")
