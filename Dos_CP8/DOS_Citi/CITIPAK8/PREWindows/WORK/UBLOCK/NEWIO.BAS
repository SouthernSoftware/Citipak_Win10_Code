DEFINT A-Z
DECLARE FUNCTION NFileSize& (FileHand)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB NFOpenS (FileName$, Handle%)
DECLARE SUB OpenFile (FileName$, FileHandle%, LockMode%, RecLen%, ErrCode%)
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
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB DisplayCMNScrn (ScrnName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION PDQTimer& ()

CONST False = 0, True = NOT False

'This function is a replacement for the QuickPak FileSize function.
'Due to the way Windows NT updates a file's size in the directory, an
'error can occur using DOS Function 4Eh (Find first file service) to
'read a file's size from the Directory. You can force Windows NT to
'commit the directory info by just opening the file again.
FUNCTION NFileSize& (FileHand)
  NFileSize& = LOF(FileHand)
END FUNCTION

'This sub is a replacement for the QuickPak FOpenS Subroutine.
SUB NFOpenS (FileName$, Handle) STATIC
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

SUB OpenFile (FileName$, FileHandle%, LockMode%, RecLen%, ErrCode%)
  ON LOCAL ERROR GOTO ErrHandler

  Row = CSRLIN
  Col = POS(0)
  
  REDIM TempScrn(1)
  SaveScrn TempScrn()

  TryCnt = 0
  
  IF LEN(FileName$) = 0 THEN
    STOP
  END IF


Label1:
  IF LockMode <> 0 THEN
    FileHandle% = FREEFILE
    OPEN FileName$ FOR RANDOM ACCESS READ WRITE LOCK READ WRITE AS FileHandle% LEN = RecLen%
  ELSE
    FileHandle% = FREEFILE
    OPEN FileName$ FOR RANDOM SHARED AS FileHandle% LEN = RecLen%
  END IF

  IF ErrCode THEN
    ERR = 0
    GOSUB GetUserChoice
    IF RetryFlag THEN
      GOTO Label1
    END IF
  ELSE
    ErrCode = 0
  END IF

  LOCATE Row, Col

EXIT SUB

ErrHandler:
  ErrCode = ERR
RESUME NEXT

GetUserChoice:
  IF TryCnt < 5 THEN
    TryCnt = TryCnt + 1
    RetryFlag = True
    GOSUB PauseABit
  ELSE
    OK = MsgBox%("UBSETUP", "OPENERR")
    SELECT CASE OK
    CASE 1
      TryCnt = 0
      RetryFlag = True
      ErrCode = 0
    CASE 2
      RetryFlag = False
      CLOSE
    END SELECT
    RestScrn TempScrn()
    LOCATE Row, Col
  END IF

RETURN

PauseABit:
  EndTime& = PDQTimer + 18

  DO UNTIL PDQTimer > EndTime&
    z$ = INKEY$
    IF LEN(z$) > 0 THEN
      EXIT DO
    END IF
  LOOP

RETURN

END SUB

