DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS

  REDIM ARCust(1) AS ARCustRecType
  ARCustLen = LEN(ARCust(1))

  CLS
  PRINT
  PRINT " Printing Accounts. . ."

  ARCust = FREEFILE
  OPEN "arcust.dat" FOR RANDOM AS ARCust LEN = ARCustLen
  NumOfRec& = LOF(ARCust) / ARCustLen

  FOR zz& = 1 TO NumOfRec&
    GET ARCust, zz&, ARCust(1)
    IF INSTR(ARCust(1).SortName, "DELETED") > 0 THEN
      LPRINT USING "#####"; zz&;
      LPRINT TAB(15); ARCust(1).SortName, ARCust(1).BILLNAME
'    IF INSTR(ARCust(1).SortName, "DELETED") > 0 THEN
'      ARCust(1).Deleted = "Y"
'    ELSE
'      ARCust(1).Deleted = ""
'    END IF
 '   IF INSTR(ARCust(1).SortName, "DELETED") > 0 THEN
 '     ARCust(1).Deleted = "Y"
 '   ELSE
 '     ARCust(1).Deleted = ""
    END IF
 '
 '   PUT ARCust, zz&, ARCust(1)
  NEXT

  CLOSE
  LOCATE 9, 1
  PRINT " Complete."
  PRINT
  END



