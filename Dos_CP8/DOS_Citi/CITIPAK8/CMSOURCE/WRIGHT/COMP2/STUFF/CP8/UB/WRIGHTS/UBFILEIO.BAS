DEFINT A-Z



SUB OpenUBFile (FileName$, Handle%, RecLen%)

  ON ERROR RESUME NEXT

  TryCnt = 0

  DO
    ERR = 0
    TryCnt = TryCnt + 1
    Handle = FREEFILE
    OPEN FileName$ FOR RANDOM SHARED AS Handle LEN = RecLen
    IF ERR = 0 THEN
      EXIT DO
    END IF
  LOOP UNTIL TryCnt >= 10
  IF TryCnt = 10 THEN
    CLS
    LOCATE 5, 1
    PRINT "ERROR opening file: "; FileName$; "   ERROR CODE:"; ERR
    END
  END IF

END SUB

