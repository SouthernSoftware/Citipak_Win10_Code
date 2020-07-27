DEFINT A-Z
  '$INCLUDE: 'newcust.bi'

  CLS
  PRINT
  PRINT "Changing. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  OPEN "ubcust.dat" FOR RANDOM AS #1 LEN = UBCustRecLen
  NumCust& = LOF(1) / UBCustRecLen

  FOR Cnt& = 1 TO NumCust&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumCust&;
    GET #1, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Book = "04" THEN
      Fixed& = Fixed& + 1
      UBCustRec(1).Serv(6).RateCode = "DRED"
      PUT #1, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE
  LOCATE 9, 1
  PRINT "Changed:"; Fixed&
  CALL KillFile("fixsunbh.exe")

