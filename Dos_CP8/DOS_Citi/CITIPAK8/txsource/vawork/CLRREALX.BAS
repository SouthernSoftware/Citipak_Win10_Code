DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'PROPAbst.BI'
  
  DIM SHARED RealRec(1) AS PropertyRecType
  RealRecLen = LEN(RealRec(1))

  TargetDate = Date2Num("04-15-2005")
  CLS
  PRINT "Correcting Property. . ."
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM SHARED AS #PropTaxFile LEN = RealRecLen
  NumPropRecs& = LOF(PropTaxFile) / RealRecLen
  FOR Srch& = 1 TO NumPropRecs&
    LOCATE 4, 1
    PRINT "Processing:"; Srch&; "of"; NumPropRecs&;
    GET PropTaxFile, Srch&, RealRec(1)
    IF RealRec(1).PROPDATE = TargetDate THEN
      IF RealRec(1).EXMPOTHR > 0 THEN
        RealRec(1).EXMPOTHR = 0
        Fixed& = Fixed& + 1
        PUT PropTaxFile, Srch&, RealRec(1)
      END IF
    END IF
  NEXT
  CLOSE

  LOCATE 12, 1
  PRINT "Fixed:"; Fixed&; "of"; NumPropRecs&

  CALL KillFile("clrrealx.exe")

