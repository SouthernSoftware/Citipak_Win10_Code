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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  BlockClear
  ShowProcessingScrn "Copying Water Readings to Sewer."

  RDate = Date2Num("01-22-1999")

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  RptFile = FREEFILE
  OPEN "BADMETER.RPT" FOR OUTPUT AS RptFile LEN = 255

  WaterFile = FREEFILE
  OPEN "F:UBCUST.dat" FOR RANDOM SHARED AS WaterFile LEN = UBCustRecLen
  WNumOfRecs& = LOF(WaterFile) \ UBCustRecLen

  SewerFile = FREEFILE
  OPEN "G:UBCUST.dat" FOR RANDOM SHARED AS SewerFile LEN = UBCustRecLen
  SNumOfRecs& = LOF(SewerFile) \ UBCustRecLen

  FOR WCnt& = 1 TO WNumOfRecs&
    GET WaterFile, WCnt&, UBCustRec(2)
    WBookSeq$ = UBCustRec(2).Book + UBCustRec(2).SeqNumb

    FoundIt = False
    FOR SCnt& = 1 TO SNumOfRecs&
      GET SewerFile, SCnt&, UBCustRec(1)
      SBookSeq$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
      IF SBookSeq$ = WBookSeq$ THEN
        FoundIt = True
        FOR MCnt = 1 TO 7
          UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(2).LocMeters(MCnt).CurRead
          UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(2).LocMeters(MCnt).PrevRead
          UBCustRec(1).LocMeters(MCnt).CurDate = UBCustRec(2).LocMeters(MCnt).CurDate
          UBCustRec(1).LocMeters(MCnt).PastDate = UBCustRec(2).LocMeters(MCnt).PastDate
          UBCustRec(1).LocMeters(MCnt).ReadFlag = UBCustRec(2).LocMeters(MCnt).ReadFlag
        NEXT
        'PUT SeterFile, SCnt&, UBCustRec(1)
        EXIT FOR
      END IF
    NEXT
    IF NOT FoundIt THEN
      BadCnt = BadCnt + 1
      LOCATE 25, 1: PRINT "BAD:"; BadCnt;
      GotABadOne = True
      PRINT #RptFile, UBCustRec(2).Book + "-" + UBCustRec(2).SeqNumb; TAB(12); USING "#####"; WCnt&;
      PRINT #RptFile, TAB(24); UBCustRec(2).CustName
    END IF
    ShowPctComp WCnt&, WNumOfRecs&
  NEXT
  CLOSE
  IF GotABadOne THEN
    PrintRptFile "Invalid Accounts From Water Data", "BADMETER.RPT", LPTPort%, RetCode%, Entry
  ELSE
    BlockClear
    DisplayUBScrn "UPDATEOK"
    WaitForAction
  END IF

CLS
END


PrintHeader:
 PRINT #RptFile, "Matching Accounts not found in Sewer Data."
 PRINT #RptFile, "Location   Account    Customer"
 PRINT #RptFile,
RETURN

