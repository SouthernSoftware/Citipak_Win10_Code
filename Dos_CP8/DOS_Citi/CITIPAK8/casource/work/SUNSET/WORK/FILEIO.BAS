DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Exist (FileName$)
DECLARE FUNCTION FLof& (FileHandle%)
DECLARE SUB UBLog (LogText$)
DECLARE SUB FClose (Handle%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FOpenAll (FileName$, AccessMode%, SharedMode%, Handle%)
DECLARE FUNCTION DosError% ()
DECLARE FUNCTION WhichError% ()
DECLARE SUB WaitForAction ()

'This function is a replacement for the QuickPak FileSize function.
'Due to the way Windows NT updates a file's size in the directory, an
'error can occur using DOS Function 4Eh (Find first file service) to
'read a file's size from the Directory. You can force Windows NT to
'commit the directory info by just opening the file again.
FUNCTION FileSize& (FileName$)
  IF Exist(FileName$) THEN
    FOpenS FileName$, TempHandle
    FileSize& = FLof(TempHandle)
    FClose TempHandle
  ELSE
    FileSize& = 0
  END IF
END FUNCTION

'This sub is a replacement for the QuickPak FOpenS Subroutine.
SUB FOpenS (FileName$, Handle) STATIC
  'Access Mode 2 = Open for reading and writing
  'Shared Mode 4 = Deny none  (Full Access)
  BadCnt = 0
  DO
    FOpenAll FileName$, 2, 4, Handle
    IF Handle <= 0 THEN
      BadCnt = BadCnt + 1
      IF BadCnt >= 100 THEN
        EXIT DO
      END IF
    ELSE
      EXIT DO
    END IF
  LOOP
  IF BadCnt > 0 AND BadCnt < 100 THEN
    UBLog "I/O ERROR: " + FileName$ + " Count:" + STR$(BadCnt)
  ELSEIF BadCnt >= 100 THEN
    CLS
    LOCATE 5, 1
    PRINT "UNABLE TO OPEN: "; UCASE$(FileName$)
    PRINT
    PRINT "    ERROR CODE: "; WhichError
    PRINT
    PRINT "PLEASE WRITE DOWN THE ERROR CODE AND CALL SOFTWARE SUPPORT"
    PRINT
    PRINT "PROGRAM TERMINATED!"
    PRINT
    PRINT "Press any key to Exit.";
    FOR Cnt = 6 TO 20
      FClose Cnt
    NEXT
    CLOSE
    WaitForAction
    END
  END IF
END SUB

