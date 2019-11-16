import random
NameList = ["达", "视", "昇", "凯", "安", "视", "天", "旭" , "思", "奇", "驰", "环", "易"]
count = 0
while count < 50:
  firstName = random.randint(0, NameList.__len__()-1)
  lastName = random.randint(0, NameList.__len__()-1)
  threeName = random.randint(0, NameList.__len__()-1)
  if (firstName != lastName != threeName):
      print(NameList[firstName]+NameList[lastName])
      # +NameList[threeName]
      count = count + 1



print("Game Over!!")