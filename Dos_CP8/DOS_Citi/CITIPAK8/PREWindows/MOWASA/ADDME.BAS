
SUB MakeZipCodeIndex (IndexText$)
  
  ShowProcessingScrn "Creating " + IndexText$ + " Index "
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  'REDIM ZipIndex(1 TO 1)  AS PSAZipIndexType
  REDIM UBCustRec(1)      AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  
  REDIM ZipIndex(1 TO NumOfBillRec)   AS PSAZipIndexType
  
  FOR BCnt = 1 TO NumOfBillRec
    GET CHandle, BCnt, UBCustRec(1)
    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
    ZipIndex(BCnt).SName = UBCustRec(1).SEARCH
    ZipIndex(BCnt).RecNum = BCnt
    ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
  NEXT
  
  CLOSE
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT ZipIndex(1), NumOfBillRec, 0, 32, 0, 10
  
  First = 1
  Last = 1
  
  SZip$ = ZipIndex(1).ZIPCODE
  
  FOR ZCnt = 2 TO NumOfBillRec
    EZip$ = ZipIndex(ZCnt).ZIPCODE
    IF SZip$ <> EZip$ THEN
      Last = ZCnt - 1
      GOSUB SortThisZip
      First = ZCnt
      SZip$ = EZip$
    END IF
    ShowPctComp ZCnt, NumOfBillRec              'show user percentage complete
  NEXT
  Last = ZCnt - 1
  GOSUB SortThisZip
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  
  IHandle = FREEFILE
  OPEN TempIndexName FOR OUTPUT AS IHandle
  CLOSE IHandle
  
  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO NumOfBillRec
    Prec& = ZipIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumOfBillRec               'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, ZipIndex
  
  EXIT SUB
  
SortThisZip:
  IF First < Last THEN
    SortT ZipIndex(First), Last - First + 1, 0, 32, 10, 10
  END IF
RETURN
  
END SUB
