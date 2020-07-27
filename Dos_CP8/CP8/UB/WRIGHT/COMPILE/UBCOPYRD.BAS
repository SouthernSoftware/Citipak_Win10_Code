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

  CONST False = 0, True = NOT False
  MaxLines = 55

  BlockClear
  ShowCursor

  IF Exist("F:\UBCUST.dat") THEN
    IF Exist("G:\UBCUST.dat") THEN
      FilesOKFlag = True
    END IF
  END IF

  IF NOT FilesOKFlag THEN
    CursorOff
    LibFile2Scrn "UBSETUP", "NOTSERVR", MonoCode%, Attribute%, ErrorCode%
    WaitForAction
    GOTO AbortExit
  END IF
  
  Ok = MsgBox%("UBSETUP", "MAKESURE")

  CursorOff
  SELECT CASE Ok
  CASE 2
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Copy Meter Readings Aborted!", 10, 26, -1
    QPrintRC "Press any key to continue.", 13, 27, -1
    WaitForAction
    GOTO AbortExit
  END SELECT
  
  BlockClear
  ShowProcessingScrn "Copying Water Readings to Sewer."
  ShowCursor

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  RptFile = FREEFILE
  OPEN "BADMETER.RPT" FOR OUTPUT AS RptFile LEN = 255
  GOSUB PrintHeader

  WaterFile = FREEFILE
  OPEN "F:\UBCUST.dat" FOR RANDOM SHARED AS WaterFile LEN = UBCustRecLen
  WNumOfRecs& = LOF(WaterFile) \ UBCustRecLen

  SewerFile = FREEFILE
  OPEN "G:\UBCUST.dat" FOR RANDOM SHARED AS SewerFile LEN = UBCustRecLen
  SNumOfRecs& = LOF(SewerFile) \ UBCustRecLen

  FOR SCnt& = 1 TO SNumOfRecs&
    GET SewerFile, SCnt&, UBCustRec(1)
    SBookSeq$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
    SBookSeq$ = QPTrim$(SBookSeq$)
    IF LEN(SBookSeq$) > 0 THEN
      FoundIt = False
      FOR WCnt& = 1 TO WNumOfRecs&
        GET WaterFile, WCnt&, UBCustRec(2)
        WBookSeq$ = UBCustRec(2).Book + UBCustRec(2).SeqNumb
        IF SBookSeq$ = WBookSeq$ THEN
          FoundIt = True
          FOR MCnt = 1 TO 7
            UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(2).LocMeters(MCnt).CurRead
            UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(2).LocMeters(MCnt).PrevRead
            UBCustRec(1).LocMeters(MCnt).CurDate = UBCustRec(2).LocMeters(MCnt).CurDate
            UBCustRec(1).LocMeters(MCnt).PastDate = UBCustRec(2).LocMeters(MCnt).PastDate
            UBCustRec(1).LocMeters(MCnt).ReadFlag = UBCustRec(2).LocMeters(MCnt).ReadFlag
          NEXT
          PUT SewerFile, SCnt&, UBCustRec(1)
          EXIT FOR
        END IF
      NEXT
      IF NOT FoundIt THEN
        BadCnt = BadCnt + 1
        'LOCATE 25, 1: PRINT "BAD:"; BadCnt;
        GotABadOne = True
        PRINT #RptFile, UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb; TAB(12); USING "#####"; SCnt&;
        PRINT #RptFile, TAB(24); UBCustRec(1).CustName
        LineCnt = LineCnt + 1
        FOR Mtr = 1 TO 7
          IF UBCustRec(1).LocMeters(Mtr).CurRead > 0 OR UBCustRec(1).LocMeters(Mtr).PrevRead > 0 THEN
            PRINT #RptFile, UBCustRec(1).LocMeters(Mtr).MtrNum;
            PRINT #RptFile, TAB(16); USING "##########"; UBCustRec(1).LocMeters(Mtr).CurRead; TAB(32); UBCustRec(1).LocMeters(Mtr).PrevRead
            LineCnt = LineCnt + 1
          END IF
        NEXT
        IF LineCnt > MaxLines THEN
          PRINT #RptFile, CHR$(12)
          GOSUB PrintHeader
        END IF
      END IF
    END IF
    ShowPctComp SCnt&, SNumOfRecs&
  NEXT
  CLOSE

  IF GotABadOne THEN
    PrintRptFile "Invalid Accounts From Water Data", "BADMETER.RPT", LPTPort%, RetCode%, Entry
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
 PRINT #RptFile, "Matching Accounts not found in Water Data."; TAB(70); "Page:"; PageNo
 PRINT #RptFile, "Location      Account       Customer"
 PRINT #RptFile, "  Mtr No.         Current        Previous"
 PRINT #RptFile, STRING$(80, "-")
 LineCnt = 4
RETURN

