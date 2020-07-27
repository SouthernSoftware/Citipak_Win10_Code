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


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Customer Information. . ."
  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen
  
  FOR CCnt& = 1 TO NumCust&
    GET UBCFile, CCnt&, UBCustRec(1)
    LOCATE , 1
    PRINT "Processing:"; CCnt&; "of"; NumCust&;
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    FOR zz = 1 TO 15
      UBCustRec(1).CurrRevAmts(zz) = 0
      UBCustRec(1).PrevRevAmts(zz) = 0
    NEXT
    UBCustRec(1).DepositAmt = 0
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).OldRec = 0

    SELECT CASE UBCustRec(1).Book
    CASE "01"
      UBCustRec(1).Book = "11"
    CASE "02"
      UBCustRec(1).Book = "12"
    CASE ELSE
    END SELECT
    FOR Cnt = 1 TO 15
      ThisCode$ = QPTrim$(UBCustRec(1).Serv(Cnt).RateCode)
      IF LEN(ThisCode$) > 0 THEN
        SELECT CASE ThisCode$
        CASE "10"
          LSET UBCustRec(1).Serv(Cnt).RateCode = "WRI1"
        CASE "20"
          LSET UBCustRec(1).Serv(Cnt).RateCode = "WRO1"
        CASE "98", "SWAC"
          LSET UBCustRec(1).Serv(Cnt).RateCode = "SRI1"
        CASE "97"
          LSET UBCustRec(1).Serv(5).RateCode = "REC"
        CASE "99"
          LSET UBCustRec(1).Serv(Cnt).RateCode = ""
        END SELECT
      END IF
    NEXT
    PUT UBCFile, CCnt&, UBCustRec(1)
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Processing Complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

