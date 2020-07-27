DEFINT A-Z
  '$INCLUDE: 'IC.BI'
   
   REDIM ItemRec(1) AS ItemRecType
   ICItemLen = LEN(ItemRec(1))

   ICFile = FREEFILE
   OPEN "ICITEM.DAT" FOR RANDOM SHARED AS ICFile LEN = ICItemLen
   NumOfRecs& = LOF(ICFile) \ ICItemLen
   FOR Cnt = 1 TO NumOfRecs&
     GET ICFile, Cnt, ItemRec(1)
     ItemRec(1).COST = 0
     ItemRec(1).QtyOnHand = 0
     ItemRec(1).QtyOnOrd = 0
     ItemRec(1).UsedYTD = 0
     ItemRec(1).ReOrder = 0
     ItemRec(1).QTD = 0
     ItemRec(1).LastTran = 0
     PUT ICFile, Cnt, ItemRec(1)
   NEXT
   CLOSE

