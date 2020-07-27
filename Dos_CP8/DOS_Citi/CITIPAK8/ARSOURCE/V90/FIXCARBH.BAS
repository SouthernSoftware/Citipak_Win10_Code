DEFINT A-Z

  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS

  CLS
  PRINT "Converting. . ."
  DIM SHARED ARCust(1) AS ARCustRecTypeOLDV85

  ARCustLen = LEN(ARCust(1))
  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = ARCustLen
  NumCust = LOF(1) / ARCustLen
  FOR cnt = 611 TO 620
    GET #1, cnt, ARCust(1)
    ARCust(1).AcctBal = 0
    PUT #1, cnt, ARCust(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Done"
  CALL KillFile("fixcarbh.exe")
  END


