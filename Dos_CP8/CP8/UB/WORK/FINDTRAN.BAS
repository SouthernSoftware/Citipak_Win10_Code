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

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  TargDate = Date2Num%("10-30-1998")

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Billing Information. . ."

  PRINT
  PRINT
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    GET UBTFile, Cnt&, UBTranRec(1)
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    'PRINT "Account:"; UBTranRec(1).CustAcctNo
    CFlag = 0
    PFlag = 0
    FixIt = 0

    IF UBTranRec(1).Transtype = TranUtilityBill THEN
      'IF UBTranRec(1).TransDate >= TargDate THEN
        FOR MCnt = 1 TO 7
          SELECT CASE UBTranRec(1).MtrTypes(MCnt)
          CASE 7
            CRead$ = STR$(UBTranRec(1).CurRead(MCnt))
            PRead$ = STR$(UBTranRec(1).PrevRead(MCnt))
            IF LEN(CRead$) >= 4 THEN
              IF RIGHT$(CRead$, 3) = "000" THEN
                CFlag = -1
                UBTranRec(1).CurRead(MCnt) = UBTranRec(1).CurRead(MCnt) / 1000
              END IF
            END IF
            IF LEN(PRead$) >= 4 THEN
              IF RIGHT$(PRead$, 3) = "000" THEN
                PFlag = -1
                UBTranRec(1).PrevRead(MCnt) = UBTranRec(1).PrevRead(MCnt) / 1000
              END IF
            END IF
            IF CFlag OR PFlag THEN
              FixIt = -1
            END IF
          CASE ELSE
          END SELECT
        NEXT
      'END IF
      IF FixIt THEN
        PUT UBTFile, Cnt&, UBTranRec(1)
        Fixed& = Fixed& + 1
      END IF
    END IF
JumpOut:
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Fixed:"; Fixed&

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

