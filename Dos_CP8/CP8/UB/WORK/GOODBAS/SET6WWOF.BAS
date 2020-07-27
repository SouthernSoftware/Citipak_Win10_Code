DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
  '$INCLUDE: 'newcust.bi'

  CLS
  PRINT
  PRINT "Adding Rate Code. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  OPEN "ubcust.dat" FOR RANDOM AS #1 LEN = UBCustRecLen
  NumCust& = LOF(1) / UBCustRecLen

  FOR Cnt& = 1 TO NumCust&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumCust&;
    GET #1, Cnt&, UBCustRec(1)
    UBCustRec(1).Serv(6).RateCode = "FIRE"
    Fixed = Fixed + 1
    PUT #1, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  LOCATE 9, 1
  PRINT "Added:"; Fixed
  CALL KillFile("set6wwof.exe")

