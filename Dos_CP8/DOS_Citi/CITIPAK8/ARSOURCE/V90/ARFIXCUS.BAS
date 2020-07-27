DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS

  REDIM ARCust(1) AS ARCustRecType
  ARCustLen = LEN(ARCust(1))

  CLS
  PRINT
  PRINT " Processing Accounts. . ."

  Cmd$ = COMMAND$
  IF INSTR(Cmd$, "UNDELETE") > 0 THEN
    UNDel = -1
  ELSE
    UNDel = 0
  END IF
  PRINT
  PRINT
  SELECT CASE UNDel
  CASE -1
    PRINT "Undeleting. ."
  CASE ELSE
    PRINT "  Deleting. ."
  END SELECT

  ARCust = FREEFILE
  OPEN "arcust.dat" FOR RANDOM AS ARCust LEN = ARCustLen
  NumOfRec& = LOF(ARCust) / ARCustLen

  FOR zz& = 1 TO NumOfRec&
    GET ARCust, zz&, ARCust(1)
    IF UNDel THEN
      ARCust(1).Deleted = ""
    ELSE
      IF INSTR(ARCust(1).SortName, "DELETED") > 0 THEN
        ARCust(1).Deleted = "Y"
      ELSE
        ARCust(1).Deleted = ""
      END IF
    END IF
    PUT ARCust, zz&, ARCust(1)
  NEXT

  CLOSE
  LOCATE 9, 1
  PRINT "Conversion Complete."
  PRINT
  CALL KillFile("ARFIXCUS.EXE")
  END



