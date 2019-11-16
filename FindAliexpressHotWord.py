import os
import xlrd as xd
import xlwt as xt
import numpy as np

FOLDER = "./data/AliexpressData/"
INPUT_FOLDER = FOLDER+"Input/"
OUT_FOLDER = FOLDER+"Output/"
HEADER = ["行业","国家","商品关键词","成交指数","浏览-支付转化率排名","竞争指数"]

class HotSell:
    industry = ""
    country = ""
    commodity = ""
    transactionIndex = 0
    browseToPayConversionRateRanking = 0
    Competition = 0

def getHotWordList(fileName, HotWord):
    list  = []
    hotWordExcel = xd.open_workbook(fileName)
    hotWordSheet = hotWordExcel.sheet_by_index(0)
    rowSize = hotWordSheet.nrows
    titleRow = hotWordSheet.row_values(0)
    for i in range(1, rowSize):
        hotWord = HotWord()
        hotWord.industry = hotWordSheet.row_values(i)[0]
        hotWord.country = hotWordSheet.row_values(i)[1]
        hotWord.commodity = hotWordSheet.row_values(i)[2]
        hotWord.transactionIndex = hotWordSheet.row_values(i)[3]
        hotWord.browseToPayConversionRateRanking = hotWordSheet.row_values(i)[4]
        hotWord.Competition = hotWordSheet.row_values(i)[5]
        list.append(hotWord)
    return list

getHotWordList(INPUT_FOLDER+"hot_Sale_total_1day.xls", HotSell)