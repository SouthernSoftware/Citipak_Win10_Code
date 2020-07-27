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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'UBIMPEXP.BI'

  TYPE ImpLookUpType
    Location AS STRING * 8
  END TYPE

  CONST False = 0, True = NOT False

  MaxLines = 55

'  IF INSTR(COMMAND$, "TEST") = 0 THEN
'    ImpSewFile$ = "UBSEWRED.DAT"
'  ELSE
    ImpSewFile$ = "A:\UBSEWRED.DAT"
'  END IF
  
  BlockClear
  ShowCursor

  IF Exist("UBCUST.DAT") THEN
    IF Exist(ImpSewFile$) THEN
      FilesOKFlag = True
    END IF
  END IF

  IF NOT FilesOKFlag THEN
    CursorOff
    LibFile2Scrn "UBSETUP", "NOTSERVR", MonoCode%, Attribute%, ErrorCode%
    WaitForAction
    GOTO AbortExit
  END IF
  
  Ok = MsgBox%("UBSETUP", "IMPSURE")

  CursorOff
  SELECT CASE Ok
  CASE 2
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Import Meter Readings Aborted!", 10, 25, -1
    QPrintRC "Press any key to continue.", 13, 27, -1
    WaitForAction
    GOTO AbortExit
  END SELECT
  
  BlockClear
  ShowCursor

  ShowProcessingScrn "Importing Sewer Meter Readings."

  QPrintRC "Reading Import File.", 11, 31, -1

  REDIM LookRec(1 TO 1) AS ImpLookUpType

  REDIM ExpImpRec(1) AS SewImpExpRecType
  UBImpExpRecLen = LEN(ExpImpRec(1))

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  RptFile = FREEFILE
  OPEN "BADMETER.RPT" FOR OUTPUT AS RptFile LEN = 255
  GOSUB PrintHeader

  ImpFile = FREEFILE
  OPEN ImpSewFile$ FOR RANDOM SHARED AS ImpFile LEN = UBImpExpRecLen
  ImpRecCnt = LOF(ImpFile) \ UBImpExpRecLen
  REDIM LookRec(1 TO ImpRecCnt) AS ImpLookUpType

  FOR Cnt = 1 TO ImpRecCnt
    GET ImpFile, Cnt, ExpImpRec(1)
    LookRec(Cnt).Location = ExpImpRec(1).Location
    ShowPctComp Cnt, ImpRecCnt
  NEXT

  QPrintRC "Updating Meter Readings.", 11, 30, -1

  SewerFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS SewerFile LEN = UBCustRecLen
  SNumOfRecs& = LOF(SewerFile) \ UBCustRecLen
  FOR SCnt& = 1 TO SNumOfRecs&
    GET SewerFile, SCnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipImport
    END IF
    SBookSeq$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
    SBookSeq$ = QPTrim$(SBookSeq$)
    IF LEN(SBookSeq$) > 0 THEN
      FoundIt = False
      FOR WCnt& = 1 TO ImpRecCnt
        WBookSeq$ = LookRec(WCnt&).Location
        IF SBookSeq$ = WBookSeq$ THEN
          GET ImpFile, WCnt&, ExpImpRec(1)
          FoundIt = True
          FOR MCnt = 1 TO 7
            UBCustRec(1).LocMeters(MCnt).CurRead = ExpImpRec(1).Reads(MCnt).CurrRead
            UBCustRec(1).LocMeters(MCnt).PrevRead = ExpImpRec(1).Reads(MCnt).PrevRead
            UBCustRec(1).LocMeters(MCnt).PastDate = UBCustRec(1).LocMeters(MCnt).CurDate
            UBCustRec(1).LocMeters(MCnt).CurDate = ExpImpRec(1).Reads(MCnt).ReadDate
            UBCustRec(1).LocMeters(MCnt).ReadFlag = ExpImpRec(1).Reads(MCnt).ReadFlag
          NEXT
          PUT SewerFile, SCnt&, UBCustRec(1)
          EXIT FOR
        END IF
      NEXT
      IF NOT FoundIt THEN
        BadCnt = BadCnt + 1
        GotABadOne = True
        PRINT #RptFile, UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb; TAB(12); USING "#####"; SCnt&;
        PRINT #RptFile, TAB(24); UBCustRec(1).CustName
        LineCnt = LineCnt + 1
        IF LineCnt > MaxLines THEN
          PRINT #RptFile, CHR$(12)
          GOSUB PrintHeader
        END IF
      END IF
    END IF
SkipImport:
    ShowPctComp SCnt&, SNumOfRecs&
  NEXT
  CLOSE

  IF GotABadOne THEN
    PrintRptFile "Accounts Not Found in Water Data", "BADMETER.RPT", LPTPort%, RetCode%, Entry
  ELSE
    BlockClear
    DisplayUBScrn "UPDATEOK"
    WaitForAction
  END IF

AbortExit:
  IF Exist("UBMENU.EXE") AND INSTR(COMMAND$, "TEST") = 0 THEN
    BlockClear
    RUN "UBMENU"
  END IF

  HideCursor
  CLS
  END


PrintHeader:
 PageNo = PageNo + 1
 PRINT #RptFile, "Accounts not found in Sewer Data."; TAB(70); "Page:"; PageNo
 PRINT #RptFile, "Location      Account       Customer"
 PRINT #RptFile, STRING$(80, "-")
 LineCnt = 4
RETURN

