DEFINT A-Z

  DECLARE FUNCTION QPTrim$ (TheText$)

  '$INCLUDE: 'PROPAbst.BI'
  CLS
  DIM PersonalRec AS PersonalRecType
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM SHARED AS #PersTaxFile LEN = LEN(PersonalRec)
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersonalRec)
  PRINT
  PRINT "Processing. . ."
  FOR zz = 1 TO NumOfPersRecs
    LOCATE 5, 1
    PRINT "Processing:"; zz; "of"; NumOfPersRecs;
    GET PersTaxFile, zz, PersonalRec
    PersonalRec.PropPIN = QPTrim$(PersonalRec.DESC1)
    PUT PersTaxFile, zz, PersonalRec
  NEXT
  PRINT
  PRINT
  PRINT
  PRINT "Processing Complete"
  CALL killfile("copy2pin.exe")

  CLOSE

