from bs4 import BeautifulSoup
import os
import xlwt as xt
import xlrd as xd
from nltk.util import ngrams
import re

productName = "Kichen Utensil"
# productName = "Silicone Food Bag"
productName = "Pepper Mill"
# productName = "French Press Coffee Maker"
# productName = "Coffee Grinder"
# productName = "Coffee Filter"

PRODUCTLIST_HTML_PATH = "./OriginalData/AmazonProductListHtml/" + productName
LIST_RESULT_DATA_PATH = "./ResultsData/" + productName
PRODUCT_LIST_INFO_PATH = LIST_RESULT_DATA_PATH + "/ProductListInfo.xls"

TITLES = ["Title", "Rating", "Reviews", "Price", "ASIN", "Sponsored", "Html Name", "Link"]

class Product:
    title = ""
    sponsored = ""
    rating = ""
    reviews = ""
    price = ""
    asin = ""
    htmlFile = ""
    link = ""


def getHtml(file):
    return BeautifulSoup(open(file, "r", encoding='utf-8'), 'html.parser')


def clearText(list):
    if (len(list) > 0):
        return list[0].text.strip().replace("  ", "").replace("\n", "")
    else:
        return ""


def getHtmFile(folder=PRODUCTLIST_HTML_PATH, format=".html"):
    filesList = os.listdir(folder)
    htmFileslist = []

    for file in filesList:
        if (file.endswith(format) or file.endswith(".htm")):
            htmFileslist.append(file)
    return htmFileslist


def getProductFromHtml(htmlSoup, htmlName):
    productHtmlWrapper = htmlSoup.select("div[class='s-result-list s-search-results sg-row']")
    productHtml = productHtmlWrapper[0]
    productHtmlList = productHtml.find_all("div", attrs={"data-asin": True})
    productList = []
    for productHtml in productHtmlList:
        product = Product()
        product.title = clearText(
            productHtml.select("h2[class='a-size-mini a-spacing-none a-color-base s-line-clamp-4']"))
        product.rating = clearText(productHtml.select("span[class='a-icon-alt']")).split("out of 5 stars")[0]
        product.reviews = clearText(productHtml.select("span[class='a-size-base']"))
        product.price = clearText(productHtml.select("div[class='a-row'] span[class='a-price']"))
        product.sponsored = clearText(
            productHtml.select("div[class='a-row a-spacing-micro'] span[class='a-size-base a-color-secondary']"))
        product.asin = productHtml.get("data-asin")
        product.htmlName = htmlName
        product.link = "https://www.amazon.com/dp/" + product.asin
        productList.append(product)
    return productList


def getProductListFromFiles():
    fileList = getHtmFile()
    productList = []
    for file in fileList:
        filePath = PRODUCTLIST_HTML_PATH + "/" + file
        htmlSoup = getHtml(filePath)
        productListFromHtml = getProductFromHtml(htmlSoup, file)
        productList = productList + productListFromHtml
    return productList


def setSheetHeader(sheet, titles=TITLES):
    for index in range(len(titles)):
        sheet.write(0, index, titles[index])


def saveProductTitle(productList):
    book = xt.Workbook(encoding='utf-8', style_compression=0)
    productSheet = book.add_sheet("Product", cell_overwrite_ok=True)
    setSheetHeader(productSheet)
    for index in range(len(productList)):
        productSheet.write(index + 1, 0, productList[index].title)
        productSheet.write(index + 1, 1, productList[index].rating)
        productSheet.write(index + 1, 2, productList[index].reviews)
        productSheet.write(index + 1, 3, productList[index].price)
        productSheet.write(index + 1, 4, productList[index].asin)
        productSheet.write(index + 1, 5, productList[index].sponsored)
        productSheet.write(index + 1, 6, productList[index].htmlName)
        productSheet.write(index + 1, 7, productList[index].link)
    book.save(PRODUCT_LIST_INFO_PATH)


def main():
    product = getProductListFromFiles()
    saveProductTitle(product)

main()
