DEFINT A-Z
DECLARE FUNCTION GetFileCount& (filespec$)
    REDIM FileName$(1 TO 1)

    filespec$ = "*.bas"
    FileCnt = 1                    'It is, so count files.
    FileName$(1) = DIR$(filespec$)

    IF LEN(FileName$(1)) = 0 THEN
      GOTO ExitCnt
    END IF

    DO
      TFile$ = DIR$
      IF LEN(TFile$) > 0 THEN
        FileCnt = FileCnt + 1
        REDIM PRESERVE FileName$(1 TO FileCnt)
        FileName$(FileCnt) = TFile$
      ELSE
        EXIT DO
      END IF
    LOOP
    FOR Cnt = 1 TO FileCnt
      PRINT FileName$(Cnt)
    NEXT
    PRINT "Count:"; FileCnt
ExitCnt:

FUNCTION GetFileCount& (filespec$)
'This function evaluates a file specification and returns the
'number of files that match the specification.  Wild card characters
'("*" and "?") are permitted. Drive and directory path specifications
'may also be included in filespec$.
END FUNCTION

