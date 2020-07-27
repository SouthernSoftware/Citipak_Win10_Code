DECLARE SUB ClearBack ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)

DEFINT A-Z
SUB RelinkTransactions

  ClearBack
  ShowProcessingScrn "Relinking Customer Transactions."

  REDIM ARCust(1) AS ARCustRecType
  CustRecLen = LEN(ARCust(1))

  REDIM ARTran(1 TO 2) AS ARTransRecType
  TranRecLen = LEN(ARTran(1))

  CustFile = FREEFILE
  OPEN "TaxCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = CustRecLen
  NumCRec& = LOF(CustFile) / CustRecLen

  TranFile = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TranRecLen
  NumTRec& = LOF(TranFile) / TranRecLen

  FOR Cnt& = 1 TO NumCRec&
    GET CustFile, Cnt&, ARCust(1)
    ARCust(1).FirstTrans = 0
    ARCust(1).LastTrans = 0
    PUT CustFile, Cnt&, ARCust(1)
    ShowPctComp Cnt&, NumCRec&
  NEXT
  
  FOR Cnt& = 1 TO NumTRec&
    GET TranFile, Cnt&, ARTran(1)
    CustRec& = VAL(ARTran(1).CustomerNumber)
    IF (CustRec& > 0) AND (CustRec& <= NumCRec&) THEN
      GET CustFile, CustRec&, ARCust(1)
      IF ARCust(1).FirstTrans = 0 THEN
        ARCust(1).FirstTrans = Cnt&
        PUT CustFile, CustRec&, ARCust(1)
        ARTran(1).NextTrans = 0
        PUT TranFile, Cnt&, ARTran(1)

      ELSE

      ARTran(1).LastTrans = ARCust(1).LastTrans
      ARCust(1).LastTrans = Cnt&
      PUT CustFile, CustRec&, ARCust(1)
      PUT TranFile, Cnt&, ARTran(1)
      ShowPctCompL Cnt&, NumTRec&
    ELSE
      'LPRINT ARTran(1).CustomerRec
      BadTran = BadTran + 1
    END IF
  NEXT

  CLOSE
'LOCATE 5, 1: PRINT "Bad:"; BadCnt
'WaitForAction
END SUB

