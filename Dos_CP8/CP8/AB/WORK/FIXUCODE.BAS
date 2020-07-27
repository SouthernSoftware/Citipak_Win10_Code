DEFINT A-Z
  '$INCLUDE: 'newcust.bi'

  CLS
  PRINT
  PRINT "Changing (UserCode2) to '7'"
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  OPEN "ubcust.dat" FOR RANDOM AS #1 LEN = UBCustRecLen
  NumCust& = LOF(1) / UBCustRecLen

  FOR Cnt& = 1 TO NumCust&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumCust&;
    GET #1, Cnt&, UBCustRec(1)
    UBCustRec(1).UserCode2 = "7"
    PUT #1, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  LOCATE 9, 1
  PRINT "Processing Complete."
  CALL killfile("fixucode.exe")

