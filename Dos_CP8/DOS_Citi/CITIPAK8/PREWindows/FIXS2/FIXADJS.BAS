DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION uRound# (DblNum#)

  '$INCLUDE: 'ubtrans.BI'
  PDate = Date2Num("01-01-1997")

  CLS
  PRINT "WARNING   WARNING   WARNING!!!!!!!"
  PRINT
  PRINT
  PRINT "ALL UTILITY BILLING OPERATIONS MUST HALT BEFORE CONTINUING!!!!"
  PRINT
  PRINT "THIS INCLUDES TAKING PAYMENTS OR ANY OTHER UTILITY ACTION!!!!!"
  PRINT
  PRINT "YOU MUST NOT DO ANYTHING IN THE UTILITY PROGRAM UNTIL THIS PROGRAM"
  PRINT
  PRINT "IS COMPLETED!!!!!!"
  PRINT
  PRINT
  PRINT "HAS ALL UTILITY ACTIVITY STOPPED?"
  PRINT
  PRINT "PRESS ANY KEY WHEN READY";

  DO
    aa$ = INKEY$
  LOOP WHILE LEN(aa$) = 0

  CLS
  PRINT "DO NOTHING UNTIL THIS PROGRAM HAS COMPLETED"
  PRINT
  PRINT
  PRINT "Correcting Adjustment Transaction Amounts. . ."
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBFile LEN = UBTransRecLen

  NumOfRecs& = LOF(UBFile) \ UBTransRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBTrans(1)
    IF UBTrans(1).TransType = TranUpwardAdjustment OR UBTrans(1).TransType = TranDownwardAdjustment THEN
      'IF UBTrans(1).TransDate >= PDate THEN
        IF UBTrans(1).TransAmt = 0 THEN
          Total# = 0
          FOR zCnt = 1 TO 15
            Total# = uRound#(Total# + UBTrans(1).RevAmt(zCnt))
          NEXT
          UBTrans(1).TransAmt = Total#
          Fixed& = Fixed& + 1
          PUT UBFile, Cnt&, UBTrans(1)
        END IF
      'END IF
    END IF

putem:
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; Fixed&

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

