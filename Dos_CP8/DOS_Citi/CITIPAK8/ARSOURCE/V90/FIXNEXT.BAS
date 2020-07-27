DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS

  DIM SHARED ARCust(1) AS ARCustRecType
  CLS
  PRINT "Processing. . ."
  ARCustLen = LEN(ARCust(1))
  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = ARCustLen
  FOR cnt = 1 TO (LOF(1) / ARCustLen)
  GET #1, cnt, ARCust(1)
    IF ARCust(1).Deleted <> "Y" THEN
      ARCust(1).IssueLicense = "Y"
      PUT #1, cnt, ARCust(1)
    END IF
  NEXT
  CLOSE
  CALL KILLFILE("FIXNEXT.EXE")
  PRINT
  PRINT
  PRINT
  PRINT "Done."
  END


