DEFINT A-Z
DECLARE SUB CustHistoryRpt ()
DECLARE SUB ClearBack ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNumber#)
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB ShowCursor ()
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Monitor% ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  TYPE HistRecInfoType
    TranRec    AS LONG
    TranType   AS INTEGER
    TranDate   AS INTEGER
    BelongTo   AS LONG
    Printed    AS INTEGER
  END TYPE

  CustHistoryRpt

SUB CustHistoryRpt
  
  SHARED Choice$()
  
  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM HistRecs(1 TO 1) AS HistRecInfoType
  
  HistRecLen = LEN(HistRecs(1))
  
  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTran(1))
  
  RptFile$ = "TaxCHIST.RPT"     'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 56
  LineCnt = 0
  DidCnt = 0
  
  ClearBack
  LookUp RecNo&, "Customer Transaction Report", False, True, False
  IF RecNo& <= 0 THEN
    GOTO ExitHistRpt
  END IF

  RptHandle = FREEFILE
  OPEN RptFile$ FOR OUTPUT AS #RptHandle
  
  TaxFile = FREEFILE
  OPEN "TaxCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  GET TaxFile, RecNo&, TaxCustRec(1)
  CLOSE TaxFile
  
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  
  PrevTranRec& = TaxCustRec(1).LastTrans
  
  GOSUB CustHistRptHeader
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET TaxTran, PrevTranRec&, TaxTran(1)
      DidCnt = DidCnt + 1
      REDIM PRESERVE HistRecs(1 TO DidCnt) AS HistRecInfoType
      HistRecs(DidCnt).TranRec = PrevTranRec&
      HistRecs(DidCnt).TranType = TaxTran(1).TranType
      HistRecs(DidCnt).TranDate = TaxTran(1).TransDate
      HistRecs(DidCnt).BelongTo = TaxTran(1).BelongTo
      PrevTranRec& = TaxTran(1).LastTrans
    LOOP
  END IF
  
  'Sort in trans type order, Bills will be first
  'SortT HistRecs(1), DidCnt, 0, HistRecLen, 3, -1
  
  FOR Cnt = 1 TO DidCnt
    IF HistRecs(Cnt).TranType = 1 THEN
      GET TaxTran, HistRecs(Cnt).TranRec, TaxTran(1)

      GOSUB GetTransInfo
      GOSUB PrintBillInfo
      REDIM THistRecs(1 TO 1) AS HistRecInfoType
      PCnt = 0
      ThisRec& = HistRecs(Cnt).TranRec
      FOR ZCnt = 1 TO DidCnt
        IF HistRecs(ZCnt).TranType <> 1 THEN
          IF HistRecs(ZCnt).BelongTo = ThisRec& THEN
            PCnt = PCnt + 1
            REDIM PRESERVE THistRecs(1 TO PCnt) AS HistRecInfoType
            LSET THistRecs(PCnt) = HistRecs(ZCnt)
          END IF
        END IF
      NEXT
      IF PCnt > 0 THEN
        FOR ZCnt = 1 TO PCnt
          GET TaxTran, THistRecs(ZCnt).TranRec, TaxTran(1)
          GOSUB GetTransInfo
          PRINT #RptHandle, ToPrint$ ' "  "; BillType$, TaxYear$, Post2GL$, BillNum$, Num2Date(TaxTran(1).TransDate), FUsing$(STR$(TaxTran(1).Amount), "#####.##")
          LineCnt = LineCnt + 1
        NEXT
      END IF
      ToPrint$ = SPACE$(79)
      MID$(ToPrint$, 23) = "Totals:"
      MID$(ToPrint$, 35) = FUsing$(STR$(TOwed#), "#######.##")
      MID$(ToPrint$, 50) = FUsing$(STR$(TPaid#), "#######.##")
      MID$(ToPrint$, 65) = FUsing$(STR$(TOwed# - TPaid#), "#######.##")
      PRINT #RptHandle, ToPrint$
      PRINT #RptHandle, Dash80$
      LineCnt = LineCnt + 2

    END IF
    IF LineCnt > MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB CustHistRptHeader
    END IF
  NEXT
  GOSUB CustHistRptFooter:
  CLOSE

  PrintRptFile "Customer Transaction History", RptFile$, LPTPort%, RetCode%, 1
  
  EXIT SUB
  
PrintBillInfo:
' PRINT #RptHandle, TransDate$
  PRINT #RptHandle, ToPrint$
  LineCnt = LineCnt + 1
  
RETURN

CustHistRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(27); "Customer Transaction Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "   Customer: "; QPTrim$(TaxCustRec(1).FNAME); " "; QPTrim$(TaxCustRec(1).LName)
  PRINT #RptHandle, Dash80$
  LineCnt = 4
RETURN
  
CustHistRptFooter:
  ToPrint$ = SPACE$(79)
  MID$(ToPrint$, 23) = "Grand Totals:"
  MID$(ToPrint$, 35) = FUsing$(STR$(GTOwed#), "#######.##")
  MID$(ToPrint$, 50) = FUsing$(STR$(GTPaid#), "#######.##")
  MID$(ToPrint$, 65) = FUsing$(STR$(GTOwed# - GTPaid#), "#######.##")
  PRINT #RptHandle, ToPrint$
  PRINT #RptHandle, FF$
RETURN
  
GetTransInfo:
  TransDate$ = Num2Date(TaxTran(1).TransDate)
  BillType$ = ""
  TaxYear$ = ""
  Post2GL$ = "N"
  IF TaxTran(1).Posted2GL = "Y" THEN
    Post2GL$ = "Y"
  END IF
  ToPrint$ = SPACE$(79)
  SELECT CASE TaxTran(1).TranType
  CASE 1
    BillType$ = "Bill #" + ParseBillNum$(TaxTran(1).Description)
    LSET ToPrint$ = BillType$
    SELECT CASE TaxTran(1).BillType
    CASE "R"
      BillType$ = "Real-Estate"
    CASE "P"
      BillType$ = "Personal Property"
    CASE "C"
      BillType$ = "Combined"
    CASE "M"
      BillType$ = "Manual"
    END SELECT
    TaxYear$ = QPTrim$(STR$(TaxTran(1).TaxYear))
    'MID$(ToPrint$, 16) = TaxYear$
    'MID$(ToPrint$, 23) = BillType$
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 28) = TaxYear$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TPaid# = 0
    TOwed# = TaxTran(1).Amount
    GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
  CASE 2
    MID$(ToPrint$, 3) = "Payment"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 50) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TPaid# = Round#(TPaid# + TaxTran(1).Amount)
    GTPaid# = Round#(GTPaid# + TaxTran(1).Amount)
  CASE 3
    BillType$ = "Release"
  CASE 4
    MID$(ToPrint$, 3) = "Interest"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# + TaxTran(1).Amount)
    GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
  CASE 6
    BillType$ = "Collection/Ad Cost"
    MID$(ToPrint$, 3) = "Col/Ad Cost"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# + TaxTran(1).Amount)
    GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
  CASE 7
    MID$(ToPrint$, 3) = "Adjustment"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 50) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TPaid# = Round#(TPaid# + TaxTran(1).Amount)
    GTPaid# = Round#(GTPaid# + TaxTran(1).Amount)
  CASE 8      'This will be the misc addcost adjustment
    BillType$ = "Miscellaneous Cost"
    MID$(ToPrint$, 3) = "Misc. Cost"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# + TaxTran(1).Amount)
    GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
  CASE ELSE
    BillType$ = "?????"
  END SELECT
RETURN
  
ExitHistRpt:
  
END SUB

