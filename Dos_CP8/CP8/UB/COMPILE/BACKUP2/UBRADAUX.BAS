DEFINT A-Z
DECLARE FUNCTION GetAccount& (Account$)
DECLARE FUNCTION GetReading& (Reading$)
DECLARE FUNCTION GetBadgerPath$ ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION QPTrim$ (Text2Trim$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'Newcust.bi'
'$INCLUDE: 'ubradio.bi'

FUNCTION GetAccount& (Account$)

  Account$ = QPTrim$(Account$)

  IF LEN(Account$) = 0 THEN
    GetAccount& = -1
    GOTO AccountExit
  END IF

  DO
    IF LEFT$(Account$, 1) = "0" THEN
      IF LEN(Account$) > 0 THEN
        Account$ = MID$(Account$, 2)
      ELSE
        GetAccount& = -1
        GOTO AccountExit
      END IF
    ELSE
      GetAccount& = QPValL(Account$)
      EXIT DO
    END IF
  LOOP

AccountExit:

END FUNCTION

FUNCTION GetBadgerPath$

  ON LOCAL ERROR GOTO BadFile

  IF NOT Exist%(PathData$) THEN
    PathFile = FREEFILE
    OPEN PathData$ FOR OUTPUT AS PathFile
    PRINT #PathFile, "C:\mvrs50\xfer"
    CLOSE PathFile
  END IF
  PathFile = FREEFILE
  OPEN PathData$ FOR INPUT AS PathFile
  LINE INPUT #PathFile, Path$
  CLOSE PathFile

  Path$ = QPTrim$(Path$)
  IF LEN(Path$) > 0 THEN
    IF RIGHT$(Path$, 1) <> "\" THEN
      Path$ = Path$ + "\"
    END IF
  END IF

'ok lets see if we can create a file in this path...
  PathTest = FREEFILE
  OPEN Path$ + "UBTMPMES.MES" FOR OUTPUT AS PathTest
  CLOSE PathTest
'if we get to here the path is ok...
  KillFile Path$ + "UBTMPMES.MES"

  GetBadgerPath$ = Path$

BadExit:
  EXIT FUNCTION

BadFile:
  GetBadgerPath$ = "BAD"
  RESUME BadExit

END FUNCTION

FUNCTION GetReading& (Reading$)

  Reading$ = QPTrim$(Reading$)

  IF LEN(Reading$) = 0 THEN
    GetReading& = -1
    GOTO ReadingExit
  END IF

  DO
    IF LEFT$(Reading$, 1) = "0" THEN
      IF LEN(Reading$) > 0 THEN
        Reading$ = MID$(Reading$, 2)
      ELSE
        GetReading& = -1
        GOTO ReadingExit
      END IF
    ELSE
      GetReading& = QPValL(Reading$)
      EXIT DO
    END IF
  LOOP

ReadingExit:


END FUNCTION

