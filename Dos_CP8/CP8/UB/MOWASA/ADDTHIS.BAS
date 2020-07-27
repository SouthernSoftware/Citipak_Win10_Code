
SUB MakeMOWSequenceIndex (Text$)

  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Location Records     ", 11, 25, -1

  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))

  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen

  NumCustRecs& = LOF(CHandle) / CustRecLen

  REDIM SequenceIndex(1 TO NumCustRecs&)  AS UBSequenceIndexTypeII
  IndexRecLen = LEN(SequenceIndex(1))

  FOR Cnt = 1 TO NumCustRecs&
    GET CHandle, Cnt, UBCustRec(1)
    SequenceIndex(Cnt).SeqNumber = UBCustRec(1).SEQ
    SequenceIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  NEXT

  CLOSE CHandle

  QPrintRC "         Sorting Index.        ", 11, 25, -1

  SortT SequenceIndex(1), CINT(NumCustRecs&), 0, 8, 0, -2

  QPrintRC "      Writing Index Records      ", 11, 25, -1

  FCreate TempIndexName

  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4

  FOR Cnt = 1 TO NumCustRecs&
    Prec& = SequenceIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  NEXT
  CLOSE IHandle

  ERASE UBCustRec, SequenceIndex

END SUB

