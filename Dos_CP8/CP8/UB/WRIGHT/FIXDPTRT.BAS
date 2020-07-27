DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  GoodDate = Date2Num("05-03-1999")

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  
  CLS
  LOCATE 3, 1, 0
  PRINT "Unflaging Transactions. . ."
  WIDTH LPRINT 80
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  FOR CCnt& = 1 TO TNumOfRecs&
    GET UBTFile, CCnt&, UBTranRec(1)
    'LOCATE 6, 1
    'PRINT "Processing:"; CCnt&; " of"; TNumOfRecs&;
    'IF UBTranRec(1).TransDate >= GoodDate THEN

     IF UBTranRec(1).Posted2GL = "Y" THEN
       Posted = Posted + 1
     ELSE
       NotPosted = NotPosted + 1
     END IF
      'LPRINT Num2Date(UBTranRec(1).TransDate); "!",
      'Fixed = Fixed + 1
      'PUT UBTFile, CCnt&, UBTranRec(1)
    'END IF
  NEXT
  
  CLOSE
 ' LOCATE 9, 1
  PRINT "Posted:"; Posted; " Not Posted:"; NotPosted

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

