DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  CONST True = -1, False = NOT True

  '$INCLUDE: 'ub2wind.bi'

  DIM UBCust AS NewUBCustRecType
  RecLen = LEN(UBCust)

  CLS
  PRINT "Coping account number to 911. . ."

  UBCustFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBCustFile LEN = RecLen
  NumORecs& = LOF(UBCustFile) / RecLen
  FOR Cnt& = 1 TO NumORecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumORecs&;
    GET #UBCustFile, Cnt&, UBCust
    UBCust.Addr911 = QPTrim$(STR$(Cnt&))
    PUT #UBCustFile, Cnt&, UBCust
  NEXT
  CLOSE
  LOCATE 10, 1
  PRINT "Done."
  CALL KillFile("ubthing.exe")
  END

