DEFINT A-Z
DECLARE FUNCTION CheckInterestInfo% ()

DEFSNG A-Z
DEFINT A-Z
FUNCTION CheckInterestInfo
  IF Exist("TAXINTCK.DAT") THEN
    TickFile = FREEFILE
    OPEN "TAXINTCK.DAT" FOR INPUT AS #TickFile
    LINE INPUT #TickFile, LastMonth$
    CLOSE TickFile
    ThisMonth$ = LEFT$(DATE$, 2)
    IF ThisMonth$ <> LastMonth$ THEN
      ClearBack

      Ok = MsgBox("TAX", "TAXINTCK")
      SELECT CASE Ok
      CASE 1   'no do it later
        CheckInterestInfo = False
      CASE 2   'yep set flag to load interest program
        CheckInterestInfo = True
        FromFile = FREEFILE
        OPEN "TAXFROMM.DAT" FOR OUTPUT AS #FromFile
        PRINT #FromFile, "FROM MAIN"
        CLOSE FromFile
      END SELECT
    END IF
  ELSE
    UpDateTicklerFile
    CheckInterestInfo = False
  END IF

END FUNCTION

