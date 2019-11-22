from bs4 import BeautifulSoup
import os
import xlwt as xt

# productName = "Kichen Utensil"
# productName = "Silicone Food Bag"
productName = "Pepper Mill"
# productName = "French Press Coffee Maker"
# productName = "Coffee Grinder"
# productName = "Coffee Filter"

TITLES = ["Class Name", "Title", "Rating", "Star", "Price", "ASIN", "Weight", "ShippingWeight", "Package Dimensions",
          "First List Date", "First Class Rank", "Second Class Rank", "Brand", "Store", "Bullet Point", "Description",
          "Reviews", "Html Name", "Link"]

PRODUCTDETAIL_HTML_PATH = "./OriginalData/AmazonProductDetailHtml/" + productName
DETAIL_RESULT_DATA_PATH = "./ResultsData/" + productName
PRODUCT_DETAIL_INFO_PATH = DETAIL_RESULT_DATA_PATH + "/ProductDetail.xls"

class Product:
    className = ""
    title = ""
    ratings = ""
    star = ""
    price = ""
    buyBoxPrice = ""
    fivePoint = []
    description = ""
    ASIN = ""
    brand = ""
    store = ""
    packageDimensions = ""
    weight = ""
    shippingWeight = ""
    firstListDate = ""
    firstClassRank = ""
    secondClassRank = ""
    reviews = []
    link = ""
    fromHtml = ""


def getHtml(file):
    return BeautifulSoup(open(file, "r", encoding='utf-8').read(), 'html.parser')


def getHtmFile(folder=PRODUCTDETAIL_HTML_PATH, format=".html"):
    filesList = os.listdir(folder)
    htmFileslist = []
    for file in filesList:
        if (file.endswith(format) or file.endswith(".htm")):
            htmFileslist.append(file)
    return htmFileslist


def clearText(list):
    if (len(list) > 0):
        return list[0].text.strip().replace("  ", "").replace("\n", "")
    else:
        return ""


def getProductDetailFromHtml(htmlSoup, htmlName):
    product = Product()
    product.title = clearText(htmlSoup.select("span[id='productTitle']"))
    product.className = clearText(htmlSoup.select("ul[class='a-unordered-list a-horizontal a-size-small']"))
    product.price = clearText(htmlSoup.select("span[id='priceblock_ourprice']")) + clearText(
        htmlSoup.select("span[id='priceblock_saleprice']"))
    product.buyBoxPrice = clearText(htmlSoup.select("span[id='price_inside_buybox']"))
    product.star = clearText(htmlSoup.select("span[id='acrPopover'] i > span")).split("out of")[0]
    product.ratings = clearText(htmlSoup.select("a[id='acrCustomerReviewLink'] span"))
    fivePointsHtml = htmlSoup.select("div[id='feature-bullets'] li span[class='a-list-item']")
    fivePointList = []
    for item in fivePointsHtml:
        fivePointList.append(item.text)
    product.fivePoint = fivePointList[1:6]
    product.description = clearText(htmlSoup.select("div[id='productDescription']"))
    detailHtml = htmlSoup.select(
        "div[id='productDetails_feature_div'] table[id=productDetails_detailBullets_sections1] tr")
    product.brand = clearText(htmlSoup.select("a[id='bylineInfo']"))
    product.store = clearText(htmlSoup.select("a[id='sellerProfileTriggerId']"))
    for item in detailHtml:
        key = clearText(item.select("th"))
        value = clearText(item.select("td"))
        if (key == "Package Dimensions"):
            product.packageDimensions = value
        elif (key == "Item Weight"):
            product.weight = value.split("pounds")[0]
        elif (key == "Shipping Weight"):
            product.shippingWeight = value.split("pounds")[0]
        elif (key == "ASIN"):
            product.ASIN = value
        elif (key == "Date first listed on Amazon"):
            product.firstListDate = value
        elif (key == "Best Sellers Rank"):
            rankStr = value.split("#")
            product.firstClassRank = rankStr[1].split("(See Top 100 in Kitchen & Dining)")[0]
            product.secondClassRank = rankStr[2]

    reviewsHtml = htmlSoup.find_all("div",
                                    class_="a-expander-content reviewText review-text-content a-expander-partial-collapse-content")
    reviewsList = []
    for item in reviewsHtml:
        reviewsList.append(item.text)
    product.reviews = reviewsList
    product.fromHtml = htmlName
    product.link = "https://www.amazon.com/dp/" + product.ASIN
    return product


def getProductDetailList():
    fileList = getHtmFile()
    productList = []
    for item in fileList:
        soup = getHtml(PRODUCTDETAIL_HTML_PATH + "/" + item)
        product = getProductDetailFromHtml(soup, item)
        productList.append(product)
    return productList

def setSheetHeader(sheet, titles=TITLES):
    for index in range(len(titles)):
        sheet.write(0, index, titles[index])


def saveProductDetail():
    productList = getProductDetailList()
    book = xt.Workbook(encoding='utf-8', style_compression=0)
    productSheet = book.add_sheet("Product Detail", cell_overwrite_ok=True)
    setSheetHeader(productSheet)
    for index in range(len(productList)):
        product = productList[index]
        productSheet.write(index + 1, 0, product.className)
        productSheet.write(index + 1, 1, product.title)
        productSheet.write(index + 1, 2, product.ratings)
        productSheet.write(index + 1, 3, product.star)
        productSheet.write(index + 1, 4, product.price)
        productSheet.write(index + 1, 5, product.ASIN)
        productSheet.write(index + 1, 6, product.weight)
        productSheet.write(index + 1, 7, product.shippingWeight)
        productSheet.write(index + 1, 8, product.packageDimensions)
        productSheet.write(index + 1, 9, product.firstListDate)
        productSheet.write(index + 1, 10, product.firstClassRank)
        productSheet.write(index + 1, 11, product.secondClassRank)
        productSheet.write(index + 1, 12, product.brand)
        productSheet.write(index + 1, 13, product.store)
        productSheet.write(index + 1, 14, "----#####----".join(product.fivePoint))
        productSheet.write(index + 1, 15, product.description)
        productSheet.write(index + 1, 16, "----#####----".join(product.reviews))
        productSheet.write(index + 1, 17, product.fromHtml)
        productSheet.write(index + 1, 18, product.link)
    book.save(PRODUCT_DETAIL_INFO_PATH)


def main():
    # os.mkdir(DETAIL_RESULT_DATA_PATH)
    # print(getProductDetailList())
    saveProductDetail()

main()
