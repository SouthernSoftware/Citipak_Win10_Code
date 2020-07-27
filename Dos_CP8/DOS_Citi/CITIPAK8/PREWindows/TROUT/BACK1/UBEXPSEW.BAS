DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (thedate$)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB BlockClear ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB WaitForAction ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB QPrintRC (Text$, Row, col, WhatColor)
DECLARE FUNCTION Monitor% ()
DECLARE SUB CursorOff ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB FCreate (FileName$)
DECLARE FUNCTION DosError% ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'UBIMPEXP.BI'

  CONST False = 0, True = NOT False
  MaxLines = 55

  BlockClear
  ShowCursor

  ImpSewFile$ = "A:\UBSEWRED.DAT"

  IF Exist("UBCUST.DAT") THEN
    FilesOKFlag = True
  END IF

  FCreate ImpSewFile$
  IF DosError <> 0 THEN
    FilesOKFlag = False
  END IF

  IF NOT FilesOKFlag THEN
    CursorOff
    LibFile2Scrn "UBSETUP", "EXPNOTOK", MonoCode%, Attribute%, ErrorCode%
    WaitForAction
    GOTO AbortExit
  END IF
  
  Ok = MsgBox%("UBSETUP", "EXPSURE")

  CursorOff
  SELECT CASE Ok
  CASE 2
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Export Sewer Readings Aborted!", 10, 25, -1
    QPrintRC "Press any key to continue.", 13, 27, -1
    WaitForAction
    GOTO AbortExit
  END SELECT
  
  BlockClear
  ShowProcessingScrn "Exporting Meter Readings."
  ShowCursor

  'KillFile "UBSEWRED.DAT"

  REDIM ExpImpRec(1) AS SewImpExpRecType
  UBImpExpRecLen = LEN(ExpImpRec(1))

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  WaterFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS WaterFile LEN = UBCustRecLen
  WNumOfRecs& = LOF(WaterFile) \ UBCustRecLen

  ExpFile = FREEFILE
  OPEN ImpSewFile$ FOR RANDOM SHARED AS ExpFile LEN = UBImpExpRecLen

  FOR WCnt& = 1 TO WNumOfRecs&
    GET WaterFile, WCnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      WBookSeq$ = QPTrim$(UBCustRec(1).BOOK + UBCustRec(1).SEQNUMB)
      IF LEN(WBookSeq$) > 0 THEN
        REDIM ExpImpRec(1) AS SewImpExpRecType
        ExpImpRec(1).Location = WBookSeq$
        FOR MCnt = 1 TO 7
          ExpImpRec(1).Reads(MCnt).CurrRead = UBCustRec(1).LocMeters(MCnt).CurRead
          ExpImpRec(1).Reads(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).PrevRead
          ExpImpRec(1).Reads(MCnt).ReadDate = UBCustRec(1).LocMeters(MCnt).CurDate
          ExpImpRec(1).Reads(MCnt).ReadFlag = UBCustRec(1).LocMeters(MCnt).ReadFlag
        NEXT
        PUT ExpFile, , ExpImpRec(1)
      END IF
    END IF
    ShowPctComp WCnt&, WNumOfRecs&
  NEXT
  CLOSE

  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

AbortExit:
  IF Exist("UBMENU.EXE") AND INSTR(COMMAND$, "TEST") = 0 THEN
    BlockClear
    RUN "UBMENU"
  END IF

HideCursor
CLS
END


