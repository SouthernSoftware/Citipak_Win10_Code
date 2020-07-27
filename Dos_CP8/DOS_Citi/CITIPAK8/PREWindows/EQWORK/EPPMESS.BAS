DEFINT A-Z
DECLARE FUNCTION MakeBucks$ (Bucksin#)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNum#)


''''INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'newcust.bi'
'$INCLUDE: 'ubTrans.BI'
'$INCLUDE: 'ubEPP.BI'


  DIM UBCustRecLen AS INTEGER
  DIM UBEPPRecLen AS INTEGER
  DIM UBFile AS INTEGER, UBeFile AS INTEGER
  DIM TotCust AS LONG
  DIM tCnt AS LONG
  DIM rCnt AS INTEGER

  DIM EPPCnt AS INTEGER
  DIM p AS STRING
  DIM rAmt AS DOUBLE
  DIM CurBal AS DOUBLE
  DIM PreBal AS DOUBLE
  DIM zCnt%
  DIM bGotCur AS INTEGER
  DIM bGotPre AS INTEGER

  p = "|"

  REDIM UBEPPCust(1) AS UBCustEPPRecType
  UBEPPRecLen = LEN(UBEPPCust(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  CLS
  OPEN "NegLIST.TXT" FOR OUTPUT AS #1

  UBeFile = FREEFILE
  OPEN "UBCUSTEP.DAT" FOR RANDOM SHARED AS UBeFile LEN = UBEPPRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  TotCust = LOF(UBFile) \ UBCustRecLen

'  GOTO OverHere
  PRINT #1, "Acct|EPB|ED1|EB1|ED2|EB2|ED3|EB3|ED4|EB4|ED5|EB5|ED6|EB6|";
  PRINT #1, "ED7|EB7|ED8|EB8|ED9|EB9|ED10|EB10|ED11|EB11|ED12|EB12|";
  PRINT #1, "ED13|EB13|ED14|EB14|ED15|EB15"
  FOR tCnt = 1 TO TotCust
    GET UBFile, tCnt, UBCustRec(1)
    GET UBeFile, tCnt, UBEPPCust(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO NotThisOne:
    END IF

    IF UBCustRec(1).EPPFlag = "Y" THEN 'UBEPPCust(1).EPBalTot <> 0 AND
      PRINT #1, QPTrim$(STR$(tCnt))
      '; p; MakeBucks(UBEPPCust(1).EPBalTot); p;
      'FOR zCnt = 1 TO 14
      '  PRINT #1, MakeBucks(UBEPPCust(1).EPInfo(zCnt).RevDis); p; MakeBucks(UBEPPCust(1).EPInfo(zCnt).Revbal); p;
      'NEXT
      'PRINT #1, MakeBucks(UBEPPCust(1).EPInfo(zCnt).RevDis); p; MakeBucks(UBEPPCust(1).EPInfo(zCnt).Revbal)
      EPPCnt = EPPCnt + 1
    END IF

NotThisOne:
  NEXT
  CLOSE

  'PRINT "Total Count:"; TotCust
  PRINT "  EPP Count:"; EPPCnt
  GOTO DoneNow:


OverHere:

  PRINT #1, "Account       EP          Status         Current         Previous"
  FOR tCnt = 1 TO TotCust
    GET UBFile, tCnt, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThis1:
    END IF
    GOTO GotoSkip:
    bGotCur = 0: bGotPre = 0

    IF UBCustRec(1).CurrBalance < 0 THEN
      IF UBCustRec(1).PrevBalance > 0 THEN
        zCnt% = zCnt% + 1
        PRINT #1, tCnt, UBCustRec(1).EPPFlag, UBCustRec(1).Status, ;
        PRINT #1, MakeBucks(UBCustRec(1).CurrBalance), ;
        PRINT #1, MakeBucks(UBCustRec(1).PrevBalance)
        GOTO SkipThis1:
      END IF
    END IF
    'this checks in the other direction
    IF UBCustRec(1).PrevBalance < 0 THEN
      IF UBCustRec(1).CurrBalance > 0 THEN
        zCnt% = zCnt% + 1
        PRINT #1, tCnt, UBCustRec(1).EPPFlag, UBCustRec(1).Status, ;
        PRINT #1, MakeBucks(UBCustRec(1).CurrBalance), ;
        PRINT #1, MakeBucks(UBCustRec(1).PrevBalance)
      END IF
    END IF
GotoSkip:

    FOR rCnt = 1 TO 15
      rAmt = UBCustRec(1).CurrRevAmts(rCnt)
      IF rAmt < 0 THEN
        'STOP
      END IF
    NEXT

SkipThis1:
  NEXT

DoneNow:
  PRINT "Count:"; zCnt%

  CLOSE

FUNCTION MakeBucks$ (Bucksin#)
'ON ERROR RESUME NEXT

DIM BucksOut$, Cash$, Cent$, TCash$
DIM Cpos AS INTEGER, pPos%

TCash$ = QPTrim(STR$(Bucksin#))

FOR Cpos = 1 TO LEN(TCash$)
  IF INSTR("0123456789.", MID$(TCash$, Cpos, 1)) = 0 THEN
    GOTO BadVal:
  END IF
NEXT

IF Bucksin# = 0 THEN
  BucksOut$ = "0.00"
  GOTO BucksExit
END IF

BucksOut$ = QPTrim$(STR$(Round(Bucksin#)))

Cpos = INSTR(BucksOut$, ".")

IF Cpos > 1 THEN   'format is ?.? possiable ????.??
  Cash$ = LEFT$(BucksOut$, Cpos - 1)
  Cent$ = MID$(BucksOut$, Cpos)
ELSEIF Cpos = 1 THEN ' format is .?"
  Cash$ = "0"
  Cent$ = BucksOut$
ELSE 'no cents
  Cash$ = BucksOut$
  Cent$ = ".00"
END IF

IF LEN(Cent$) = 2 THEN 'format is .5 must append 0
  Cent$ = Cent$ + "0"
END IF

BucksOut$ = Cash$ + Cent$

BucksExit:

  MakeBucks$ = BucksOut$

  EXIT FUNCTION

BadVal:

  MakeBucks$ = TCash$

END FUNCTION

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
    UBTransRec(1).RevAmt(zz) = Round#(UBEPPRec(1).EPInfo(zz).Revbal)
    UBCustRec(1).CurrRevAmts(zz) = Round#(UBCustRec(1).CurrRevAmts(zz) + UBEPPRec(1).EPInfo(zz).Revbal)
  NEXT
  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + UBEPPRec(1).EPBalTot)

  IF ClearEPPFlag THEN
    UBCustRec(1).EPPFlag = "N"
  END IF

  FOR zz = 1 TO 15
    UBEPPRec(1).EPInfo(zz).Revbal = 0
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

