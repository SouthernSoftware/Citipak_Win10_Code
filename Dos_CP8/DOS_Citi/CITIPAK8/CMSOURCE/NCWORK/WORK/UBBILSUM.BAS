DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Round# (N#)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)  'load setup file
DECLARE SUB BlockClear ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB UBLog (Text$)
DECLARE SUB MakePostalIndex (IdxTypeText$)
DECLARE SUB MakeSequenceIndex (IdxTypeText$)


  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  ''$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  ''$INCLUDE: 'ubrate.BI'

  '$INCLUDE: 'ubpinfo.bi'

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

SUB BillSummaryRpt

  SHARED Choice$()

  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")

  LPIFlag = False

'  IF NOT ChkBillFile% THEN
'    UBLog "ERROR: NO BILL FILE."
'    CursorOff
'    BlockClear
'    DisplayUBScrn "NON2PRNT"
'    WaitForAction
'    GOTO ExitBillSum
'  END IF

  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM ScrnArray(0)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen      'load setup file
  TownName$ = UBSetUpRec(1).UTILNAME

  REDIM RevDesc(1 TO MaxRevsCnt) AS STRING * 12
  FOR Cnt = 1 TO MaxRevsCnt     'find last active revenue
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = Cnt - 1       'set actual number of revenues
      EXIT FOR
    ELSE    'build revenue description lines
      LSET RevDesc(Cnt) = UCASE$(TempRev$)
    END IF
  NEXT
  REDIM RevTotals(1 TO NumOfRevs) AS DOUBLE

  REDIM Choice$(0 TO 6, 3)

  Choice$(0, 0) = "1"
  Choice$(0, 1) = "2"
  Choice$(0, 2) = "3"

  Choice$(1, 0) = "Customer Name Order"
  Choice$(2, 0) = "Account Number Order"
  Choice$(3, 0) = "Location Number Order"
  Choice$(4, 0) = "Postal Carrier Route Order"
  Choice$(5, 0) = "ZipCode Order"

  IF UBSetUpRec(1).UseSeq = "Y" THEN
    Choice$(6, 0) = "Sequence Number Order"
  END IF

  Choice$(1, 1) = "Detailed"
  Choice$(2, 1) = "Summary"

  Choice$(1, 2) = "Printer 1"
  Choice$(2, 2) = "Printer 2"
  Choice$(3, 2) = "Screen   "

  LibName$ = "UB"
  ScrnName$ = "BILSUMRP"
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0

  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True

  FirstTime = True
  '--Set screen number to one and display screen

  BlockClear

  DisplayUBScrn ScrnName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = Choice$(1, 1)
      Action = 1
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckReqFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT

    '--Check screen page
  LOOP UNTIL ExitFlag

  IF AbortFlag THEN
    UBLog "ABORTED: Billing Summary"
    GOTO ExitBillSum
  END IF

  SELECT CASE BillOrder$
  CASE "C"
    IndexName$ = NameIndexFile
    UsingName = True
    OKFlag = True
  CASE "A"
    IndexName$ = ""
    UsingAcct = True
    OKFlag = True
  CASE "L"
    IndexName$ = BookIndexFile
    UsingBook = True
    OKFlag = True
  CASE "P", "Z"
    IF BillOrder$ = "P" THEN
      IdxTypeText$ = "Postal Route"
    ELSE
      IdxTypeText$ = "Zip-Code"
    END IF
    MakePostalIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  CASE "S"
    IdxTypeText$ = "Sequence Number"
    MakeSequenceIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  END SELECT

  'IF AbortFlag GOTO ExitBillSum

  PastDate$ = Form$(PastDFld, 0)

'do bill printing here
'**************************************************************************

  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))

  IF UsingAcct THEN             'load the index
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF

  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen

  UBRpt = FREEFILE
  OPEN "UBBILSUM.RPT" FOR OUTPUT AS UBRpt

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  ShowProcessingScrn "Creating Billing Summary."

  GOSUB DOBillSumHeader

  '-----------------------------------------
  PrintedCnt = 0
  NotDone = True

  FOR Cnt = 1 TO NumOfRecs
    IF UsingAcct THEN
      CustAcctNo& = Cnt
    ELSE
      CustAcctNo& = IndexArray(Cnt).RecNum
    END IF

    GET UBBill, CustAcctNo&, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN
      GET UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
      PRINT #UBRpt, USING "########"; CustAcctNo&;
      PRINT #UBRpt, TAB(12); LEFT$(UBCustRec(1).CUSTNAME, 25); TAB(70); USING "###,###.##"; UBBillRec(1).TransAmt
      LineCnt = LineCnt + 1
      IF ShowDetail THEN
        GOSUB ShowSumDetail
      END IF
    END IF

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DOBillSumHeader
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  GOSUB DOBillSumTotals

  CLOSE

  ERASE UBCustRec

  IF NOT AbortFlag THEN
    PrintRptFile "Billing Summary Report.", "UBBILSUM.RPT", 1, RetCode, EntryPoint
  END IF
  
ExitBillSum:

EXIT SUB

ShowSumDetail:
  TCnt = 0
  Detail$ = SPACE$(18)
  First = 1
  Last = NumOfRevs
  FOR RCnt = First TO Last
    TCnt = TCnt + 1
    TabStop = (TCnt * 21) - 20
    IF TabStop > 81 THEN
      TCnt = 1
      TabStop = (TCnt * 21) - 20
      LineCnt = LineCnt + 1
    END IF
    LSET Detail$ = RevDesc(RCnt)
    MID$(Detail$, 10) = "#####.##"
    RevTotals(RCnt) = Round#(RevTotals(RCnt) + UBCustRec(1).CurrRevAmts(RCnt))
    PRINT #UBRpt, TAB(TabStop); USING Detail$; UBBillRec(1).RevAmt(RCnt);
  NEXT
  PRINT #UBRpt,
  LineCnt = LineCnt + 1

DoneDetail:
  PRINT #UBRpt, Dash80$
  LineCnt = LineCnt + 1
RETURN

RETURN

DOBillSumHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing Report      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "           Acct "; TAB(72); "Customer"
  PRINT #UBRpt, "Location    No.   Customer Name             Service Address             Status"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN

DOBillSumTotals:
  PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: ####"; Active
  PRINT #UBRpt, USING "   Final: ####"; Final
  PRINT #UBRpt, USING "Inactive: ####"; InActive
  PRINT #UBRpt, USING " Balance: ####"; Balance
  PRINT #UBRpt, USING " Unknown: ####"; UnKnown
  PRINT #UBRpt, USING " Deleted: ####"; DeletedCnt
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: ####"; CustCnt
  PRINT #UBRpt, CHR$(12)
RETURN

CheckReqFields:
  BillOrder$ = QPTrim$(LEFT$(Form$(1, 0), 1))
  RptType$ = QPTrim$(LEFT$(Form$(2, 0), 1))
  RptDev$ = RIGHT$(QPTrim$(Form$(2, 0)), 1)
  SELECT CASE RptDev$
  CASE "1"
    LptPort = 1
  CASE "2"
    LptPort = 2
  CASE ELSE
    LptPort = 1
    ScrnFlag = True
  END SELECT
  IF LEFT$(RptType$, 1) = "D" THEN
    ShowDetail = True
  END IF

  IF LEN(BillOrder$) = 0 THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "NOBORDER"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = 4
  ELSE
    ReqFldsOK = True
  END IF


RETURN


END SUB

