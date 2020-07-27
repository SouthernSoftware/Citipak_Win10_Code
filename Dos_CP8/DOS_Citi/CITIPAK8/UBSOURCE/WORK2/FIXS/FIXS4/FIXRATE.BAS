DEFINT A-Z
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ubrate.BI'
CLS
  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))

  'IF Exist("UBRATE.ORG") THEN
  '  KILL "UBRATE.ORG"
  'END IF
  NAME "UBRATE.DAT" AS "UBRATE.ORG"

  OFile = FREEFILE
  OPEN "UBRATE.ORG" FOR RANDOM SHARED AS OFile LEN = UBRateTblRecLen

  'OFile = FREEFILE
  'OPEN "UBRATE.DAT" FOR RANDOM SHARED AS OFile LEN = UBRateTblRecLen

  NFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS NFile LEN = UBRateTblRecLen

  NumOfRec = LOF(OFile) \ UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRec
    GET OFile, Cnt, UBRateTblRec(1)
    Code$ = QPTrim$(UBRateTblRec(1).RateCode)
    IF LEN(Code$) = 0 THEN
      Desc$ = QPTrim$(UBRateTblRec(1).RateDesc)
      IF LEN(Desc$) = 0 THEN
        GOTO SkipCode
      END IF
    END IF
    PUT NFile, , UBRateTblRec(1)
    PRINT Code$
SkipCode:
  NEXT

  CLOSE


