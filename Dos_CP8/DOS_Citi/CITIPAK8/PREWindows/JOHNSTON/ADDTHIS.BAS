DEFINT A-Z

SUB NoSerialNumRPT
  REDIM TempScrn(0)

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  MaxLines = 52
  FF$ = CHR$(12)

  Action = 1
  BlockClear

  '--Check for Key presses

  BegRoute = 1
  EndRoute = 99

  'Open Report File
  ReportFile$ = "WBMTRLST.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

  CustName$ = SPACE$(30)

  ' Location Order ********************************************************
  'if UBSetupRec(1).

  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen

  Cnt& = 1
  GOSUB PrnNOSerHeader
  ShowProcessingScrn "Reading Meter Information"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF INSTR(UBCustRec(1).HHMSG1, "NOREAD") = 0 THEN
        LSET CustName$ = UBCustRec(1).CustName
        IF UBCustRec(1).Status > "" THEN
          FOR MtrCnt& = 1 TO 7                'find last active meter
            IF (UBCustRec(1).LocMeters(MtrCnt&).CurRead >= 0) OR (UBCustRec(1).LocMeters(MtrCnt&).PrevRead) > 0 THEN
              IF (LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRNum)) = 0) AND LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)) > 0 THEN
                PRINT #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).seqnumb; "   "; UBCustRec(1).Status; "   "; CustName$; LEFT$(UBCustRec(1).SERVADDR, 30); "  "; MtrCnt&
                LineCnt = LineCnt + 1
                EXIT FOR
              END IF
            END IF
          NEXT
        END IF
      END IF
    END IF
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrnNOSerHeader
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT

  PRINT #RptHandle, FF$
  CLOSE

  Header$ = "NO SERIAL NUMBER Report"

  ERASE IdxBuff
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  EXIT SUB

PrnNOSerHeader:
  page = page + 1
  PRINT #RptHandle, TAB(30); "NO Meter Number Report"; TAB(65); "Date: "; DATE$
  PRINT #RptHandle, TAB(70); "Page #"; page
  PRINT #RptHandle, "Location Status Customer Name"; TAB(41); "Service Address"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
RETURN



END SUB

