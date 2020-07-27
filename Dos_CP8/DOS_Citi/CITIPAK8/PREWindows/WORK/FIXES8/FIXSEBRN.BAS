DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Invalid Transactions. . ."
  'END

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen

  UBTFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  
  FOR CCnt& = 1 TO CNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; CCnt&; " of"; CNumOfRecs&;
    GET UBCFile, CCnt&, UBCustRec(1)
    IF UBCustRec(1).BILLCYCL = 4 THEN
      IF UBCustRec(1).LastTrans > 0 THEN
        GET UBTFile, UBCustRec(1).LastTrans, UBTranRec(1)
        IF UBTranRec(1).TransDate = 0 THEN
          FOR Cnt = 1 TO 15
             UBCustRec(1).CurrRevAmts(Cnt) = uRound#(UBCustRec(1).CurrRevAmts(Cnt) - UBTranRec(1).RevAmt(Cnt))
          NEXT
          CurTotal# = 0
          FOR Cnt = 1 TO 15
            CurTotal# = uRound#(CurTotal# + UBCustRec(1).CurrRevAmts(Cnt))
          NEXT
          UBCustRec(1).CurrBalance = CurTotal#
          UBCustRec(1).PrevBalance = 0
          PUT UBCFile, CCnt&, UBCustRec(1)
          FOR Cnt = 1 TO 15
             UBTranRec(1).RevAmt(Cnt) = 0
          NEXT
          UBTranRec(1).TransAmt = 0
          UBTranRec(1).CustAcctNo = -1
          UBTranRec(1).PrevTrans = -1
          PUT UBTFile, UBCustRec(1).LastTrans, UBTranRec(1)
          Found = Found + 1
        END IF
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Removed:"; Found
PRINT
PRINT
PRINT "Press any key. . .";
a$ = INPUT$(1)
SHELL "ubmisc relink"
KILL "FIXSEBRN.EXE"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

