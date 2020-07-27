DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'cmfiles.BI'
  CLS

  NDate = Date2Num%("03-11-1999")

  LOCATE 3, 1, 0
  PRINT "Processing Transactions . . ."

  REDIM CMTranRec(1) AS CMTransRecType
  REDIM UBTranRec(1) AS UBTransRecType

  UBTranRecLen = LEN(UBTranRec(1))
  CMTranRecLen = LEN(CMTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen
  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "UB Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransDate >= NDate THEN
      TTCnt& = TTCnt& + 1
      UBTranRec(1).Posted2GL = ""
      PUT UBFile, Cnt&, UBTranRec(1)
    END IF
  NEXT
  CLOSE

  CMFile = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMFile LEN = CMTranRecLen
  CNumOfRecs& = LOF(UBFile) / CMTranRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 8, 1
    PRINT "CM Processing:"; Cnt&; " of"; CNumOfRecs&;
    GET CMFile, Cnt&, CMTranRec(1)
    IF CMTranRec(1).TransDate >= NDate THEN
      TTCnt& = TTCnt& + 1
      CMTranRec(1).Trans2GL = ""
      PUT CMFile, Cnt&, CMTranRec(1)
    END IF
  NEXT
  CLOSE

LOCATE 12, 1
PRINT "  Found:"; TTCnt&;

