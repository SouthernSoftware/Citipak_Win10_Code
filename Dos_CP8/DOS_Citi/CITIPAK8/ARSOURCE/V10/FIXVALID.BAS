DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS

  CLS
  LOCATE 3, 1
  PRINT "Correcting Valid Thru Dates. . ."
  
  REDIM ARCustRec(1) AS ARCustRecType
  ARCustRecLen = LEN(ARCustRec(1))

  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  FOR Cnt = 1 TO NumOfArRecs
    LOCATE 6, 1
    PRINT "Processing:"; Cnt; " of"; NumOfArRecs;
    GET ARFile, Cnt, ARCustRec(1)
    WhatDate$ = Num2Date$(ARCustRec(1).VALID)
    Year$ = RIGHT$(WhatDate$, 4)
    IF Year$ = "1900" THEN
      MID$(WhatDate$, 7, 4) = "2000"
      ARCustRec(1).VALID = Date2Num(WhatDate$)
      Fixed = Fixed + 1
      PUT ARFile, Cnt, ARCustRec(1)
    END IF
  NEXT
  CLOSE

  LOCATE 8, 1
  PRINT " Corrected:"; Fixed
  LOCATE 10, 1
  PRINT "Processing Complete."

