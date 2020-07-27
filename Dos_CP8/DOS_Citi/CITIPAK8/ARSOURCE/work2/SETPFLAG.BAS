DEFINT A-Z

'$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS
  
  DIM ARCust AS ARCustRecType
  CustLen = LEN(ARCust)
  CLS
  PRINT "Setting Print Next Flag. . ."
  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = CustLen
  FOR Cnt& = 1 TO LOF(1) \ CustLen

    GET #1, Cnt&, ARCust
    ARCust.IssueLicense = "Y"
    PUT #1, Cnt&, ARCust
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT "Done."
  CALL KillFile("SETPFLAG.EXE")



