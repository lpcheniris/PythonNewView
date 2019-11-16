from bs4 import BeautifulSoup
import os
import xlwt as xt
import xlrd as xd
from nltk.util import ngrams
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.collocations import *
from nltk import FreqDist

ALIEXPRESS_PATH = "/Users/clp/Documents/AliExpress"
TITLE_HTML_SAVE_FOLDER = "./titleList"
def getHtml(file):
    return BeautifulSoup(open(file, "r", encoding='utf-8'), 'html.parser')

def getHtmFile(folder=TITLE_HTML_SAVE_FOLDER, format=".html"):
    filesList = os.listdir(folder)
    htmFileslist = []
    for file in filesList:
        if (file.endswith(format)):
            htmFileslist.append(file)
    return htmFileslist

def getTitleFromHtml(htmlSoup):
    titleHtmls = htmlSoup.select("a[class='item-title']")
    titles = []
    for index in range(len(titleHtmls)):
        titles.append(titleHtmls[index].attrs["title"])
    return titles

def getTitleList():
    fileList = getHtmFile()
    htmlList = []
    titleList = []

    for index in range(len(fileList)):
        filePath = TITLE_HTML_SAVE_FOLDER + "/" + fileList[index]
        htmlList.append(getHtml(filePath))
    for index in range(len(htmlList)):
        titles = getTitleFromHtml(htmlList[index])
        titleList = titleList + titles
    return titleList

def saveProductTitle(titles, className = "smartwatch"):
    book = xt.Workbook(encoding='utf-8', style_compression=0)
    titleSheet = book.add_sheet(className + " Title", cell_overwrite_ok=True)
    wordSheet = book.add_sheet(className+" Title Word", cell_overwrite_ok=True)
    for index in range(len(titles)):
        titleWords = titles[index].split()
        titleSheet.write(index, 0, titles[index])
        for wordIndex in range(len(titleWords)):
            wordSheet.write(index, wordIndex, titleWords[wordIndex])

    book.save(ALIEXPRESS_PATH + "/" + "ProductTitleList.xls")

def getTitlesFromExcel(fileName=ALIEXPRESS_PATH + "/" + "ProductTitleList.xls", className = "smartwatch" + " Title Word"):
        titleList = []
        titlesExcel = xd.open_workbook(fileName)
        titlesSheet = titlesExcel.sheet_by_name(className)
        for rowIndex in range(titlesSheet.nrows):
            titleRow = titlesSheet.row_values(rowIndex)
            for colindex in range(len(titleRow)):
                cellValue = titlesSheet.cell(rowIndex, colindex)
                if cellValue.value:
                   titleList.append(cellValue.value)
        return titleList

def oneWordCount(titleList, fileName=ALIEXPRESS_PATH + "/" + "ProductTitleWordFreq.xls", className = "smartwatch" + " Title Word Freq"):
    wordFreq = FreqDist(titleList)
    titlesExcel = xt.Workbook(encoding='utf-8', style_compression=0)
    wordFreqSheet = titlesExcel.add_sheet(className, cell_overwrite_ok=True)
    rowIndex = 0
    for key, val in sorted(wordFreq.items(), key=lambda x: (x[1], x[0]), reverse=True):
        wordFreqSheet.write(rowIndex, 0, key)
        wordFreqSheet.write(rowIndex, 1, val)
        rowIndex = rowIndex + 1
    titlesExcel.save(fileName)

def twoWordCount(titleList, fileName=ALIEXPRESS_PATH + "/" + "ProductTitleWordFreqThree.xls", className = "smartwatch" + " Title Word Freq Two"):
    bigrams = ngrams(titleList, 3)
    bigrams_c = {}
    for b in bigrams:
        if b not in bigrams_c:
            bigrams_c[b] = 1
        else:
            bigrams_c[b] += 1
    titlesExcel = xt.Workbook(encoding='utf-8', style_compression=0)
    wordFreqSheet = titlesExcel.add_sheet(className, cell_overwrite_ok=True)
    rowIndex = 0

    for key, val in sorted(bigrams_c.items(), key=lambda x: (x[1], x[0]), reverse=True):
        wordFreqSheet.write(rowIndex, 3, key[0] + " " + key[1] + " " + key[2])
        wordFreqSheet.write(rowIndex, 4, val)
        rowIndex = rowIndex + 1
    titlesExcel.save(fileName)



def main():
    # print(getTitlesFromExcel())
    titleList = getTitlesFromExcel()
    twoWordCount(titleList)
    # oneWordCount(titleList)
    # for i in range(num_words):
    #     freq_list.append([list(freq_dist.keys())[i], list(freq_dist.values())[i]])
    # freqArr = np.array(freq_list)

    # wordCount(titleList)
    # titleList = getTitleList()
    # saveProductTitle(titleList)

main()
# def saveTitleToText(filename = "Pro", docs):
#     fh = open(filename, 'w', encoding='utf-8')
#     for doc in docs:
#         fh.write(doc)
#         fh.write('\n')
#     fh.close()
#     save('/Users/Desktop/portrait/jour_paper_docs.txt', docs)


# getTitleList()

# def writeTitleToText(titleList):


