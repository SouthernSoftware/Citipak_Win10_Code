DECLARE SUB DisplayCFScrn (ScrnName$)
DEFINT A-Z

DECLARE FUNCTION GetValidateDataPath% (DataPath$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB WaitForAction ()

CONST False = 0, True = NOT False

FUNCTION GetValidateDataPath% (DataPath$)
  ON LOCAL ERROR RESUME NEXT

  GetValidateDataPath% = 0
  REDIM TempScrn(0)

  PathFile = FREEFILE
  OPEN "CUSTFIND.DAT" FOR INPUT SHARED AS #PathFile
  LINE INPUT #PathFile, DataPath$

  IF ERR THEN
    SaveScrn TempScrn()
    DisplayCFScrn "BADDPATH"
    WaitForAction
    GOTO ExitGetDataPath
  END IF

  CLOSE PathFile

  DataPath$ = QPTrim$(DataPath$)
  IF LEN(DataPath$) = 0 THEN
    SaveScrn TempScrn()
    DisplayCFScrn "BADDPATH"
    WaitForAction
    GOTO ExitGetDataPath
  END IF

  IF RIGHT$(DataPath$, 1) <> "\" THEN
    DataPath$ = DataPath$ + "\"
  END IF

  DataPath$ = DataPath$ + "WBCUST.DAT"
  WBFile = FREEFILE
  OPEN DataPath$ FOR BINARY ACCESS READ SHARED AS WBFile
  IF ERR THEN
    SaveScrn TempScrn()
    DisplayCFScrn "BADDPATH"
    WaitForAction
    GOTO ExitGetDataPath
  END IF

  GetValidateDataPath% = True

ExitGetDataPath:
  CLOSE
  ON LOCAL ERROR GOTO 0
END FUNCTION

