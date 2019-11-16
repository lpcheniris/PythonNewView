import os
import json
import xlwt as xt

def saveKeyWords(wordsList):
    wordsExcel = xt.Workbook(encoding='utf-8', style_compression=0)
    wordsSheet = wordsExcel.add_sheet("Key Words", cell_overwrite_ok=True)
    for index in range(len(wordsList)):
        value = wordsList[index]
        wordsSheet.write(index, 0, value["value"])
        wordsSheet.write(index, 2, value["resource"])
        wordsSheet.write(index, 3, value["prefix"])
    wordsExcel.save("./Key Words.xls")

def getWordsList():
    wordsFile = open("./keyWords.json", "r", encoding='utf-8')

    wordsJsonList = json.load(wordsFile)["key_word"]
    keyMap = {}
    keyWords = []
    for suggestions in wordsJsonList:
        prefix = suggestions["prefix"]
        suggestionList = suggestions["suggestions"]
        for item in suggestionList:
            # value = item["value"]
            if prefix != "" and item["value"] != "":
                valueWithPrefix = item["value"].strip()
                valueSplitList = valueWithPrefix.split(prefix)

                if (len(valueSplitList) > 0):
                    value = valueSplitList[1].strip()
                    if value not in keyMap.keys():
                        keyMap[value] = value
                        keyWord = {
                            "prefix": prefix,
                            "value": value,
                            "resource": "Amazon Search"
                        }
                        keyWords.append(keyWord)
    return keyWords

def main():
    wordsList = getWordsList()
    print(wordsList)
    saveKeyWords(wordsList)


main()