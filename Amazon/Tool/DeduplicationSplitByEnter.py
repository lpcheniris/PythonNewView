from bs4 import BeautifulSoup

def getWordListFromHtml(file="./ForToolKit.html"):
    soup = BeautifulSoup(open(file, "r", encoding='utf-8'), 'html.parser')
    wordList = soup.select("p")[0].text
    return list(set(wordList.split("\n")))

def main():
    wordsList = "\n".join(getWordListFromHtml())
    print(wordsList)

main()
