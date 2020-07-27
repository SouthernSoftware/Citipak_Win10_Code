DEFINT A-Z
DECLARE FUNCTION OpenLockCRFile% (CRHandle)

CONST False = 0, True = NOT False

'$INCLUDE: 'oschk.bi'

FUNCTION OpenLockCRFile (CRHandle)

  ON LOCAL ERROR GOTO HaveError

  LockErrCnt = 0

  REDIM OSChk(1) AS OSChkRecType
  ChkRecLen = LEN(OSChk(1))
  ERASE OSChk

ErrorOpenRetry:
  CRHandle = FREEFILE
  OPEN "crchk.dat" FOR RANDOM LOCK READ WRITE AS CRHandle LEN = ChkRecLen

  OpenLockCRFile = True

ErrorExit:
  EXIT FUNCTION

HaveError:
  LockErrCnt = LockErrCnt + 1
  IF LockErrCnt < 4 THEN
    RESUME ErrorOpenRetry
  ELSE
    OpenLockCRFile = False
    CRHandle = 0
    RESUME ErrorExit
  END IF


END FUNCTION

