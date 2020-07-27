DEFINT A-Z
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNum#)


'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'newcust.bi'
'$INCLUDE: 'ubTrans.BI'
'$INCLUDE: 'ubEPP.BI'

SUB MakeEPPReconTrans (RecNo&, ClearEPPFlag%)

  REDIM UBEPPRec(1) AS UBCustEPPRecType
  UBEPPRecLen = LEN(UBEPPRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  EPFile = FREEFILE
  OPEN "UBCUSTEP.DAT" FOR RANDOM SHARED AS EPFile LEN = UBEPPRecLen
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  TRFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS TRFile LEN = UBTransRecLen
  NextTrans& = LOF(TRFile) \ UBTransRecLen + 1

  GET EPFile, RecNo&, UBEPPRec(1)
  GET UBFile, RecNo&, UBCustRec(1)

  IF UBEPPRec(1).EPBalTot = 0 THEN
    GOTO DontDoIt
  END IF

  UBTransRec(1).TransDate = Date2Num(DATE$)
  UBTransRec(1).TransType = TranEPPReconcile
  UBTransRec(1).TransDesc = "EPP RECONCILE ACCT"
  UBTransRec(1).TransAmt = UBEPPRec(1).EPBalTot

  FOR zz = 1 TO 15
    UBTransRec(1).RevAmt(zz) = Round#(UBEPPRec(1).EPInfo(zz).RevBal)
    UBCustRec(1).CurrRevAmts(zz) = Round#(UBCustRec(1).CurrRevAmts(zz) + UBEPPRec(1).EPInfo(zz).RevBal)
  NEXT
  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + UBEPPRec(1).EPBalTot)

  IF ClearEPPFlag THEN
    UBCustRec(1).EPPFlag = "N"
  END IF

  FOR zz = 1 TO 15
    UBEPPRec(1).EPInfo(zz).RevBal = 0
  NEXT
  UBEPPRec(1).EPBalTot = 0
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

  UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
  UBCustRec(1).LastTrans = NextTrans&

  PUT EPFile, RecNo&, UBEPPRec(1)
  PUT UBFile, RecNo&, UBCustRec(1)
  PUT TRFile, NextTrans&, UBTransRec(1)

DontDoIt:
  CLOSE EPFile, UBFile, TRFile

END SUB

