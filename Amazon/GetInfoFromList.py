from bs4 import BeautifulSoup
import os
import xlwt as xt
import xlrd as xd
from nltk.util import ngrams
import re

# productName = "Kichen Utensil"
# productName = "Silicone Food Bag"
# productName = "Peper Mill"
# productName = "French Press Coffee Maker"
# productName = "Coffee Grinder"
productName = "Coffee Filter"

PRODUCTLIST_HTML_PATH = "./OriginalData/AmazonProductListHtml/" + productName
LISTRESULTDATA_PATH = "./ResultsData/ProductList/" + productName
DetailRESULTDATA_PATH = "./ResultsData/ProductDetail/" + productName
PRODUCTLISTINFO_PATH = LISTRESULTDATA_PATH + "/ProductListInfo.xls"
PRODUCTLIST_TITLE_NGRAMS = LISTRESULTDATA_PATH + "/ProductTitleNgrams.xls"

def getHtml(file):
    return BeautifulSoup(open(file, "r", encoding='utf-8'), 'html.parser')

def getHtmFile(folder=PRODUCTLIST_HTML_PATH, format=".html"):
    filesList = os.listdir(folder)
    htmFileslist = []
    for file in filesList:
        if (file.endswith(format) or file.endswith(".htm")):
            htmFileslist.append(file)
    return htmFileslist

def getTitlesFromHtml(htmlSoup):
    titleHtmls = htmlSoup.select("div[class='sg-col-inner'] span[class='a-size-base-plus a-color-base a-text-normal']")
    titles = []
    for index in range(len(titleHtmls)):
        titles.append(titleHtmls[index].text)
    return titles

def getTitleListFromFiles():
    fileList = getHtmFile()
    htmlList = []
    titleList = []

    for index in range(len(fileList)):
        filePath = PRODUCTLIST_HTML_PATH + "/" + fileList[index]
        htmlList.append(getHtml(filePath))
    for index in range(len(htmlList)):
        titles = getTitlesFromHtml(htmlList[index])
        titleList = titleList + titles
    return titleList

def saveProductTitle(titles):
    book = xt.Workbook(encoding='utf-8', style_compression=0)
    titleSheet = book.add_sheet("Title", cell_overwrite_ok=True)
    wordSheet = book.add_sheet("Title Word", cell_overwrite_ok=True)
    for index in range(len(titles)):
        titleWords = titles[index].split()
        titleSheet.write(index, 0, titles[index])
        for wordIndex in range(len(titleWords)):
            wordSheet.write(index, wordIndex, titleWords[wordIndex])
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

def twoWordNgrams(titleList=[], fileName=PRODUCTLIST_TITLE_NGRAMS):
    degree2 = ngrams(titleList, 2)
    degree3 = ngrams(titleList, 3)
    titlesExcel = xt.Workbook(encoding='utf-8', style_compression=0)
    degree2Sheet = titlesExcel.add_sheet("Degree 2", cell_overwrite_ok=True)
    degree3Sheet = titlesExcel.add_sheet("Degree 3", cell_overwrite_ok=True)

    degree2_c = {}
    for b in degree2:
        if b not in degree2_c:
            degree2_c[b] = 1
        else:
            degree2_c[b] += 1
    rowIndex2 = 0
    for key, val in sorted(degree2_c.items(), key=lambda x: (x[1], x[0]), reverse=True):
        degree2Sheet.write(rowIndex2, 0, key[0] + " " + key[1])
        degree2Sheet.write(rowIndex2, 1, val)
        rowIndex2 = rowIndex2 + 1

    degree3_c = {}
    for b in degree3:
        if b not in degree3_c:
            degree3_c[b] = 1
        else:
            degree3_c[b] += 1
    rowIndex3 = 0
    for key, val in sorted(degree3_c.items(), key=lambda x: (x[1], x[0]), reverse=True):
        degree3Sheet.write(rowIndex3, 0, key[0] + " " + key[1]+ " " + key[2])
        degree3Sheet.write(rowIndex3, 1, val)
        rowIndex3 = rowIndex3 + 1

    titlesExcel.save(fileName)

def main():
    titleList = getTitleListFromFiles()
    saveProductTitle(titleList)

def main2():
    titleList = getTitlesFromExcel()
    twoWordNgrams(titleList)

# main()
main2()



