DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB FMakeSequenceIndex (IndexText$)

  REDIM SequenceIndex(1 TO 1)  AS UBMowSequenceIndexType
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Location Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs& = GetNumOfCust%
  
  IndexRecLen = LEN(SequenceIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs&
    GET CHandle, Cnt, UBCustRec(1)
    IF UBCustRec(1).Status = "F" THEN
      FCnt = FCnt + 1
      REDIM PRESERVE SequenceIndex(1 TO FCnt)  AS UBMowSequenceIndexType
      SequenceIndex(FCnt).SeqNumber = UBCustRec(1).SEQ
      SequenceIndex(FCnt).RecNum = Cnt
    END IF
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT SequenceIndex(1), FCnt, 0, 8, 0, -2
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  
  FCreate TempIndexName
  
  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  
  FOR Cnt = 1 TO FCnt
    Prec& = SequenceIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, FCnt               'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, SequenceIndex

END SUB

