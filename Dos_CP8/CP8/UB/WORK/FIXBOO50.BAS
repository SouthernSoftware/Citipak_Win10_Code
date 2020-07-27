DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Moving readings. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumCust& = LOF(UBFile) \ UBCustRecLen

  FOR CCnt& = 1 TO NumCust&
    GET UBFile, CCnt&, UBCustRec(1)
    IF UBCustRec(1).Book = "50" THEN
      FOR zz = 1 TO 7
        UBCustRec(1).locmeters(zz).CurRead = UBCustRec(1).locmeters(zz).PrevRead
        UBCustRec(1).locmeters(zz).ReadFlag = "N"
        UBCustRec(1).EstFlag = ""
      NEXT
      RestCnt = RestCnt + 1
    END IF
    PUT UBFile, CCnt&, UBCustRec(1)
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Corrected:"; RestCnt
CALL KillFile("fixboo50.exe")

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

