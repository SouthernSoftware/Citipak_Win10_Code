DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
    GET UBCFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      IF UBCustRec(1).Status = "B" THEN
        Fixed = Fixed + 1
      END IF
      UBCustRec(1).Status = "I"
      FOR FRCnt = 1 TO 15
        UBCustRec(1).CurrRevAmts(FRCnt) = 0
        UBCustRec(1).PrevRevAmts(FRCnt) = 0
      NEXT
      UBCustRec(1).CurrBalance = 0
      UBCustRec(1).PrevBalance = 0
      PUT UBCFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Fixed:"; Fixed

  PRINT
  KillFile "FIXPENGP.EXE"


