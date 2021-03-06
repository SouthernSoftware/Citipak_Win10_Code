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
  PRINT "Scanning. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    GET UBCFile, Cnt&, UBCustRec(1)
    'IF Cnt& = 4663 THEN STOP
    IF UBCustRec(1).Status = "A" THEN
      OKFlag = 0
      FOR SCnt = 1 TO 14
        RateCode$ = QPTrim$(UBCustRec(1).Serv(SCnt).RateCode)
        IF LEN(RateCode$) > 0 THEN
          OKFlag = -1
          EXIT FOR
        END IF
      NEXT
      IF OKFlag = 0 THEN
        FOR SCnt = 1 TO 4
          IF UBCustRec(1).FlatRates(SCnt).FRAmt > 0 THEN
            OKFlag = -1
            EXIT FOR
          END IF
        NEXT
      END IF
      IF OKFlag = 0 THEN
        LPRINT "Acct:"; Cnt&
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Scanning Complete."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

