DEFINT A-Z

  '$INCLUDE: 'CACust.BI'

  PRINT "Processing. . ."
  DIM CACust AS CACUSTType
  CustLen = LEN(CACust)
  CustFile = FREEFILE
  OPEN "camaster.dat" FOR RANDOM AS CustFile LEN = CustLen
  NumCust& = LOF(CustFile) / CustLen
  FOR Cnt& = 1 TO NumCust&
    GET CustFile, Cnt&, CACust
    CACust.WATERDATE = 8073
    PUT CustFile, Cnt&, CACust
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  CALL KILLFILE("setwdate.EXE")

