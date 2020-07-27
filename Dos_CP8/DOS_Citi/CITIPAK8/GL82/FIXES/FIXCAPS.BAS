DEFINT A-Z
DECLARE FUNCTION GetFileCount& (filespec$)

    CLS

    filespec$ = "*.cap"
    REDIM FileName$(1 TO 500)
    FileName$(1) = DIR$(filespec$)
    FCnt = 1
    DO
      FileN$ = DIR$
      IF LEN(FileN$) > 0 THEN
        FCnt = FCnt + 1
        FileName$(FCnt) = FileN$
      ELSE
        EXIT DO
      END IF
    LOOP
    PRINT
    PRINT
    PRINT "Files to process:"; FCnt
    PRINT
    PRINT
    FOR Cnt = 1 TO FCnt
      LOCATE , 1: PRINT "Processing:"; Cnt;
      TempFile$ = LEFT$(FileName$(Cnt), LEN(FileName$(Cnt)) - 4) + ".$$$"
      NAME FileName$(Cnt) AS TempFile$
      OPEN TempFile$ FOR INPUT AS #1
      OPEN FileName$(Cnt) FOR OUTPUT AS #2
      DO UNTIL EOF(1)
        LINE INPUT #1, a$
        WordPos = INSTR(a$, "NumTrans")
        IF WordPos > 0 THEN
          MID$(a$, WordPos) = "NumTran&"
          Fixed = Fixed + 1
        END IF
        PRINT #2, a$
      LOOP
      CLOSE

    NEXT
PRINT
PRINT
PRINT "Fixed:"; Fixed

SUB GetFileNames (filespec$, FileName$())
  'DIM FileCount AS LONG
END SUB

