from bs4 import BeautifulSoup
import os
import xlwt as xt
import xlrd as xd
from nltk.util import ngrams
import re

# productName = "Kichen Utensil"
# productName = "Silicone Food Bag"
# productName = "Pepper Mill"
# productName = "French Press Coffee Maker"
productName = "Coffee Grinder"
# productName = "Coffee Filter"

PRODUCTLIST_HTML_PATH = "./OriginalData/AmazonProductListHtml/" + productName
LISTRESULTDATA_PATH = "./ResultsData/ProductList/" + productName
DetailRESULTDATA_PATH = "./ResultsData/ProductDetail/" + productName
PRODUCTLISTINFO_PATH = LISTRESULTDATA_PATH + "/ProductListInfo.xls"
PRODUCTLIST_TITLE_NGRAMS = LISTRESULTDATA_PATH + "/ProductTitleNgrams.xls"

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
    book.save(PRODUCTLISTINFO_PATH)


def getTitlesFromExcel(fileName=PRODUCTLISTINFO_PATH):
    titlesExcel = xd.open_workbook(fileName)
    titlesSheet = titlesExcel.sheet_by_index(0)
    titleString = ""
    for rowIndex in range(titlesSheet.nrows):
        titleString = titleString + "  " + titlesSheet.row_values(rowIndex)[0]
    pat_letter = re.compile(r'[^a-zA-Z \']+')
    titleString = pat_letter.sub(' ', titleString).strip().lower()
    return titleString.split()

def writeDegree(degree, sheet):
    degree_c = {}
    for b in degree:
        if b not in degree_c:
            degree_c[b] = 1
        else:
            degree_c[b] += 1
    rowIndex2 = 0
    for key, val in sorted(degree_c.items(), key=lambda x: (x[1], x[0]), reverse=True):
        sheet.write(rowIndex2, 0, " ".join(key))
        sheet.write(rowIndex2, 1, val)
        rowIndex2 = rowIndex2 + 1

def twoWordNgrams(titleList=[], fileName=PRODUCTLIST_TITLE_NGRAMS):
    degree1 = ngrams(titleList, 1)
    degree2 = ngrams(titleList, 2)
    degree3 = ngrams(titleList, 3)
    titlesExcel = xt.Workbook(encoding='utf-8', style_compression=0)
    degree1Sheet = titlesExcel.add_sheet("Degree 1", cell_overwrite_ok=True)
    degree2Sheet = titlesExcel.add_sheet("Degree 2", cell_overwrite_ok=True)
    degree3Sheet = titlesExcel.add_sheet("Degree 3", cell_overwrite_ok=True)

    writeDegree(degree1, degree1Sheet)
    writeDegree(degree2, degree2Sheet)
    writeDegree(degree3, degree3Sheet)

    titlesExcel.save(fileName)


def main():
    product = getProductListFromFiles()
    saveProductTitle(product)


def main2():
    titleList = getTitlesFromExcel()
    twoWordNgrams(titleList)


# main()
main2()
