DEFINT A-Z


'$INCLUDE: 'oschk.bi'
  CLS
  PRINT
  PRINT "Correcting Invalid Dates. . ."

  REDIM OSChk(1) AS OSChkRecType
  RecLen = LEN(OSChk(1))

'

  CRHandle = FREEFILE
  OPEN "crchk.dat" FOR RANDOM SHARED AS CRHandle LEN = RecLen
  NumOfRec& = LOF(CRHandle) \ RecLen

  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET CRHandle, Cnt&, OSChk(1)
    IF OSChk(1).Src = 1 THEN
      SELECT CASE LEFT$(OSChk(1).ChkDate, 2)
      CASE "01", "02", "03"
        NewDate$ = LEFT$(OSChk(1).ChkDate, 6) + "02"
        OSChk(1).ChkDate = NewDate$
        PUT CRHandle, Cnt&, OSChk(1)
        Fixed = Fixed + 1
      END SELECT
    END IF
  NEXT
  CLOSE
  LOCATE 9, 1
  PRINT "Corrected:"; Fixed
  CALL KillFile("fixcdate.exe")

