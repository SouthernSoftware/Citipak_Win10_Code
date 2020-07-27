DEFINT A-Z
  '$INCLUDE: 'IC.BI'

   REDIM oItemRec(1) AS oItemRecType
   oICItemLen = LEN(oItemRec(1))

   REDIM ItemRec(1) AS ItemRecType
   ICItemLen = LEN(ItemRec(1))

   ICFile = FREEFILE
   OPEN "ICITEM.DAT" FOR RANDOM SHARED AS ICFile LEN = oICItemLen
   NumOfRecs& = LOF(ICFile) \ ICItemLen
   FOR Cnt = 1 TO NumOfRecs&
     GET ICFile, Cnt, oItemRec(1)
     ItemRec(1).ITEMNUMBER = oItemRec(1).ITEMNUMBER
     ItemRec(1).DESC1 = oItemRec(1).DESC1
     ItemRec(1).DESC2 = oItemRec(1).DESC2
     ItemRec(1).Vendor = oItemRec(1).Vendor
     ItemRec(1).VItem = oItemRec(1).VItem
     ItemRec(1).DeptNumb = oItemRec(1).DeptNumb
     ItemRec(1).COST = oItemRec(1).COST
     ItemRec(1).YTDCost = 0
     ItemRec(1).Unit = oItemRec(1).Unit
     ItemRec(1).QtyOnOrd = oItemRec(1).QtyOnOrd
     ItemRec(1).QtyOnOrd = oItemRec(1).QtyOnOrd
     ItemRec(1).UsedYTD = oItemRec(1).UsedYTD
     ItemRec(1).ReOrder = oItemRec(1).ReOrder
     ItemRec(1).QTD = 0
     ItemRec(1).Note1 = oItemRec(1).Note1
     ItemRec(1).Note2 = oItemRec(1).Note2
     ItemRec(1).Deleted = oItemRec(1).Deleted
     ItemRec(1).Future = oItemRec(1).Future
     PUT ICFile, Cnt, ItemRec(1)
   NEXT
   CLOSE

