DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ubschlum.bi'

  REDIM UBSchlumHHRec(1) AS SchlumHHType
  UBSchlumHHRecLen = LEN(UBSchlumHHRec(1))
  
  UBSchlFile = FREEFILE
  OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBSchlFile LEN = UBSchlumHHRecLen

  'Open and Initialize the PC3000 Read Information File
  NumTRGetRecs = LOF(UBSchlFile) / UBSchlumHHRecLen
  FOR Cnt = 1 TO NumTRGetRecs
    GET UBSchlFile, Cnt, UBSchlumHHRec(1)
    PrevRead# = VAL(UBSchlumHHRec(1).PrevRead)
    High# = VAL(UBSchlumHHRec(1).HighLimit)
    IF High# > 0 THEN
      PrevRead# = PrevRead# + (High# \ 2)
    ELSE
      PrevRead# = PrevRead# + 1000
    END IF
    UBSchlumHHRec(1).MtrRead = QPTrim$(STR$(PrevRead#))
    UBSchlumHHRec(1).ReadDate = LEFT$(UBSchlumHHRec(1).PrevDate, 4) + "01"
    UBSchlumHHRec(1).ReadType = "M"
    PUT UBSchlFile, Cnt, UBSchlumHHRec(1)
  NEXT
  CLOSE

