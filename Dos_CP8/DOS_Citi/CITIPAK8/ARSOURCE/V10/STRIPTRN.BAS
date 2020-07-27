DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION uRound# (N#)

  BDate1 = Date2Num("07-06-1900")
  BDate2 = Date2Num("07-05-1901")
  
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS

  CLS
  PRINT "Processing. . ."

  CALL KillFile("ARTrans.Da2")

  REDIM ARTransRec(1) AS ARTransRecType
  TranLen = LEN(ARTransRec(1))
  REDIM ARCust(1) AS ARCustRecType
  CustLen = LEN(ARCust(1))

  NAME "ARTrans.Dat" AS "ARTrans.Da2"

  CFile = FREEFILE
  OPEN "arcust.dat" FOR RANDOM AS #CFile LEN = CustLen
  
  ARTRFile = FREEFILE
  OPEN "ARTrans.Da2" FOR RANDOM SHARED AS #ARTRFile LEN = TranLen
  TrNumRecs = LOF(ARTRFile) / TranLen

  ARTRFile2 = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM SHARED AS #ARTRFile2 LEN = TranLen

  FOR Cnt& = 1 TO TrNumRecs
    GET ARTRFile, Cnt&, ARTransRec(1)
    TRDate = ARTransRec(1).TransDate
    IF (TRDate = BDate1) OR (TRDate = BDate2) THEN
      GET CFile, VAL(ARTransRec(1).CustomerNumber), ARCust(1)
      ARCust(1).AcctBal = uRound#(ARCust(1).AcctBal - ARTransRec(1).TransAmount)
      ARCust(1).IssueLicense = "Y"
      PUT CFile, VAL(ARTransRec(1).CustomerNumber), ARCust(1)
      Fixed = Fixed + 1
    ELSE
      PUT ARTRFile2, , ARTransRec(1)
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Corrected:"; Fixed
  CALL KILLFILE("STRIPTRN.EXE")



'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .500000001#) / 100
END FUNCTION

