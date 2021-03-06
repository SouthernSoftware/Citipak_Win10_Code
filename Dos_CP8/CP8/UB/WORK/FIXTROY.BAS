DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)


  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ub2wind.BI'

  CONST False = 0, True = NOT False

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  TargDate = Date2Num%("10-30-1998")
  ReadDate = Date2Num%("05-13-2004")

  CLS
  'LOCATE 3, 1, 0
  PRINT "Processing Transactions. . ."

  PRINT
  PRINT
  'END

  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    GET UBTFile, Cnt&, UBTranRec(1)
    LOCATE 4, 1
    PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    CFlag = 0
    PFlag = 0
    FixIt = 0

    IF UBTranRec(1).TransType = TranUtilityBill THEN
      'IF UBTranRec(1).TransDate >= TargDate THEN
        FOR MCnt = 1 TO 7
          UBTranRec(1).EstRead(MCnt) = ""
        NEXT
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
      IF FixIt THEN
        Fixed& = Fixed& + 1
      END IF
      PUT UBTFile, Cnt&, UBTranRec(1)
    END IF
JumpOut:
  NEXT

SkiptoNext:
  CLOSE

  LOCATE 6, 1
  PRINT "Fixed:"; Fixed&

  PRINT
  PRINT
  PRINT "Processing Customer Readings. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    DidEm = 0
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    UBCustRec(1).EstFlag = ""
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).MtrType = "T" THEN
        UBCustRec(1).LocMeters(MCnt).CurDate = ReadDate
        UBCustRec(1).LocMeters(MCnt).MTRMulti = 1000
        UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).PrevRead / 1000
        UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).CurRead / 1000
        DidEm = -1
      END IF
    NEXT
    IF DidEm THEN
      Fxed& = Fxed& + 1
    END IF
    PUT UBFile, Cnt&, UBCustRec(1)

SkipThisOne:
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; Fxed&
  CALL KILLfile("FIXTROY.EXE")

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

