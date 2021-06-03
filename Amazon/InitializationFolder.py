import os

# productName = "Kichen Utensil"
# productName = "Silicone Food Bag"
# productName = "Pepper Mill"
# productName = "French Press Coffee Maker"
# productName = "Coffee Grinder"
# productName = "Coffee Filter"
# productName = "i11 pro max case"
# productName = "i11 pro case"
# productName = "i11 pro max screen protector" 

# productName = "Smart Lock" 
# productName = "Camera" 
productName = "Robot Vacuum" 

DETAILORIGINALDATA_PATH = "./OriginalData/AmazonProductDetailHtml/" + productName
LISTORIGINALDATA_PATH = "./OriginalData/AmazonProductListHtml/" + productName
DETAILRESULTDATA_PATH = "./ResultsData/ProductDetail/" + productName
LISTRESULTDATA_PATH = "./ResultsData/ProductList/" + productName
RESULTDATA_PATH = "./ResultsData/" + productName


os.mkdir(DETAILORIGINALDATA_PATH)
os.mkdir(LISTORIGINALDATA_PATH)
os.mkdir(RESULTDATA_PATH)
# os.mkdir(DETAILRESULTDATA_PATH)
# os.mkdir(LISTRESULTDATA_PATH)
