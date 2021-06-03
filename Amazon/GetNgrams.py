import os
import xlwt as xt
import xlrd as xd
from nltk.util import ngrams
import re

# productName = "Kichen Utensil"
# productName = "Silicone Food Bag"
# productName = "Pepper Mill"
# productName = "French Press Coffee Maker"
# productName = "Coffee Grinder"
# productName = "Coffee Filter"
# productName = "i11 pro max case"
# productName = "i11 pro case"
productName = "i11 pro max screen protector"


PRODUCT_RESULT_DATA_PATH = "./ResultsData/" + productName
PRODUCT_DETAIL_INFO_PATH = PRODUCT_RESULT_DATA_PATH + "/ProductDetail.xls"
PRODUCT_LIST_INFO_PATH = PRODUCT_RESULT_DATA_PATH + "/ProductListInfo.xls"
PRODUCT_LIST_TITLE_NGRAMS = PRODUCT_RESULT_DATA_PATH + "/ProductListTitleNgrams.xls"
PRODUCT_DETAIL_TITLE_NGRAMS = PRODUCT_RESULT_DATA_PATH + "/ProductDetailTitleNgrams.xls"

def getTitlesFromExcel(titleFile=PRODUCT_LIST_TITLE_NGRAMS, row = 0):
    titlesExcel = xd.open_workbook(titleFile)
    titlesSheet = titlesExcel.sheet_by_index(0)
    titleString = ""
    for rowIndex in range(titlesSheet.nrows):
        titleString = titleString + "  " + titlesSheet.row_values(rowIndex)[row]
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

def wordNgrams(titleList=[], fileName=PRODUCT_LIST_TITLE_NGRAMS):
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
    listTitles = getTitlesFromExcel(PRODUCT_LIST_INFO_PATH)
    wordNgrams(listTitles, PRODUCT_LIST_TITLE_NGRAMS)
    detailTitles = getTitlesFromExcel(PRODUCT_DETAIL_INFO_PATH, 1)
    wordNgrams(detailTitles, PRODUCT_DETAIL_TITLE_NGRAMS)

main()
