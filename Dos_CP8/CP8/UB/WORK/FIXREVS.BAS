DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transaction!"

  PRINT
  PRINT
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen

  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  
  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).RevAmt(15) <> 0 THEN
      SWAP UBTranRec(1).RevAmt(15), UBTranRec(1).RevAmt(14)
      DoIt = True
      IF DoIt THEN
        Fixed = Fixed + 1
        PUT UBTFile, Cnt&, UBTranRec(1)
        DoIt = False
      END IF
    END IF
  NEXT

  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Customer Information. . ."
  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen

  FOR CCnt& = 1 TO NumCust&
    LOCATE , 1
    PRINT "Processing:"; CCnt&; "of"; NumCust&;
    DoIt = False
    GET UBCFile, CCnt&, UBCustRec(1)
    FOR zz = 1 TO 4
      IF UBCustRec(1).flatrates(zz).revsrc = 15 THEN
        UBCustRec(1).flatrates(zz).revsrc = 14
        DoIt = True
      END IF
    NEXT
    IF DoIt = True THEN
      PUT UBCFile, CCnt&, UBCustRec(1)
      Changed = Changed + 1
    END IF
  NEXT

  CLOSE


SkiptoNext:
  CLOSE

  LOCATE 9, 1
  PRINT "Changed:"; Changed
  CALL KILLFILE("fixREVS.exe")

