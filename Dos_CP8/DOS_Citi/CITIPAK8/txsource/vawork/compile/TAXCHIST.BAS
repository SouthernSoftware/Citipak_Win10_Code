DEFINT A-Z
DECLARE SUB ClearBack ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNumber#)
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB ShowCursor ()
DECLARE FUNCTION FUsing$ (Number$, Image$)
  
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

SUB CustHistoryRpt1
  
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
  TaxType = 1
  ClearBack
  LookUp RecNo&, "Customer Transaction Report", False, True, False, TaxType
  IF RecNo& <= 0 THEN
    GOTO ExitHistRpt
  END IF
  GOTO Skip2Here
  LibName$ = "TAX"
  ScrnName$ = "CUSTHRPT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  REDIM Choice$(0 TO 3, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Summary"
  Choice$(2, 0) = "Detail"
  
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "LPT1"
  Choice$(2, 1) = "LPT2"
  Choice$(3, 1) = "Screen"
  
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "D" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
      Where$ = QPTrim$(Form$(2, 0))
      FOR Cnt = 1 TO 3
        IF Where$ = Choice$(Cnt, 1) THEN
          SELECT CASE Cnt
          CASE 1
            LPTPort = 1
            EntryPoint = 1
          CASE 2
            LPTPort = 2
            EntryPoint = 1
          CASE 3
            LPTPort = 1
            EntryPoint = 2
          END SELECT
        END IF
      NEXT
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN
    GOTO ExitHistRpt
  END IF

Skip2Here:
  RptHandle = FREEFILE
  OPEN RptFile$ FOR OUTPUT AS #RptHandle
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
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
  
  'first sort in trans type order
  SortT HistRecs(1), DidCnt, 0, HistRecLen, 3, -1
  
  FOR Cnt = 1 TO DidCnt
    IF HistRecs(Cnt).TranType = 1 THEN
      GET TaxTran, HistRecs(Cnt).TranRec, TaxTran(1)
      TOwed# = TaxTran(1).Amount
      TPaid# = 0
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
        SortT THistRecs(1), PCnt, 0, HistRecLen, 6, -1
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
  CLOSE
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  PrintRptFile Header$, RptFile$, LPTPort%, RetCode%, EntryPoint
  
  EXIT SUB
  

PrintBillInfo:
  PRINT #RptHandle, ToPrint$
  LineCnt = LineCnt + 1
  RETURN

CustHistRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(27); "Customer Transaction Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "   Customer: "; QPTrim$(TaxCustRec(1).FNAME); " "; QPTrim$(TaxCustRec(1).LName)
  PRINT #RptHandle, "Trans Desc"; TAB(15); "   Date"; TAB(37); "Bill Amt"; TAB(52); "Paid Amt"; TAB(67); "Balance"; TAB(78); "GL?"
  PRINT #RptHandle, Dash80$
  LineCnt = 4
  RETURN
  
CustHistRptFooter:
  PRINT #RptHandle, Dash80$
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
  CASE 2
    MID$(ToPrint$, 3) = "Payment"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 50) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TPaid# = Round#(TPaid# + TaxTran(1).Amount)
  CASE 3
    BillType$ = "Abatement"
    MID$(ToPrint$, 3) = "Abatement"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$((TaxTran(1).Amount * -1)), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# - TaxTran(1).Amount)

  CASE 4
    MID$(ToPrint$, 3) = "Interest"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# + TaxTran(1).Amount)
  CASE 5
    MID$(ToPrint$, 3) = "Penalty"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# + TaxTran(1).Amount)
  
  CASE ELSE
    BillType$ = "?????"
  END SELECT
RETURN
  
ExitHistRpt:
  
END SUB

SUB CustHistoryRpt2

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
  TaxType = 2
  ClearBack
  LookUp RecNo&, "Customer Transaction Report", False, True, False, TaxType
  IF RecNo& <= 0 THEN
    GOTO ExitHistRpt1
  END IF
  GOTO Skip2Here1
  LibName$ = "TAX"
  ScrnName$ = "CUSTHRPT"

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  REDIM Choice$(0 TO 3, 0 TO 2)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Summary"
  Choice$(2, 0) = "Detail"

  Choice$(0, 1) = "2"
  Choice$(1, 1) = "LPT1"
  Choice$(2, 1) = "LPT2"
  Choice$(3, 1) = "Screen"

  Action = 1
  ClearBack

  ShowCursor

  DisplayTaxScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "D" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
      Where$ = QPTrim$(Form$(2, 0))
      FOR Cnt = 1 TO 3
        IF Where$ = Choice$(Cnt, 1) THEN
          SELECT CASE Cnt
          CASE 1
            LPTPort = 1
            EntryPoint = 1
          CASE 2
            LPTPort = 2
            EntryPoint = 1
          CASE 3
            LPTPort = 1
            EntryPoint = 2
          END SELECT
        END IF
      NEXT
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag

  IF AbortFlag THEN
    GOTO ExitHistRpt1
  END IF

Skip2Here1:
  RptHandle = FREEFILE
  OPEN RptFile$ FOR OUTPUT AS #RptHandle

  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  GET TaxFile, RecNo&, TaxCustRec(1)
  CLOSE TaxFile

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  PrevTranRec& = TaxCustRec(1).LastTrans

  GOSUB CustHistRptHeader1

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

  'first sort in trans type order
  SortT HistRecs(1), DidCnt, 0, HistRecLen, 3, -1

  FOR Cnt = 1 TO DidCnt
    IF HistRecs(Cnt).TranType = 1 THEN
      GET TaxTran, HistRecs(Cnt).TranRec, TaxTran(1)
      TOwed# = TaxTran(1).Amount
      TPaid# = 0
      GOSUB GetTransInfo1
      GOSUB PrintBillInfo1
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
        SortT THistRecs(1), PCnt, 0, HistRecLen, 6, -1
        FOR ZCnt = 1 TO PCnt
          GET TaxTran, THistRecs(ZCnt).TranRec, TaxTran(1)
          GOSUB GetTransInfo1
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
      GOSUB CustHistRptHeader1
    END IF

  NEXT
  CLOSE

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF

  PrintRptFile Header$, RptFile$, LPTPort%, RetCode%, EntryPoint

  EXIT SUB


PrintBillInfo1:
  PRINT #RptHandle, ToPrint$
  LineCnt = LineCnt + 1

RETURN

CustHistRptHeader1:
  Page = Page + 1
  PRINT #RptHandle, TAB(27); "Customer Transaction Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "   Customer: "; QPTrim$(TaxCustRec(1).FNAME); " "; QPTrim$(TaxCustRec(1).LName)
  PRINT #RptHandle, "Trans Desc"; TAB(15); "   Date"; TAB(37); "Bill Amt"; TAB(52); "Paid Amt"; TAB(67); "Balance"; TAB(78); "GL?"
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN

CustHistRptFooter1:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, FF$
RETURN

GetTransInfo1:
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
  CASE 2
    MID$(ToPrint$, 3) = "Payment"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 50) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TPaid# = Round#(TPaid# + TaxTran(1).Amount)
  CASE 3
    MID$(ToPrint$, 3) = "Abatement"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$((TaxTran(1).Amount * -1)), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# - TaxTran(1).Amount)

CASE 4
    MID$(ToPrint$, 3) = "Interest"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# + TaxTran(1).Amount)
  CASE 5
    MID$(ToPrint$, 3) = "Penalty"
    MID$(ToPrint$, 15) = TransDate$
    MID$(ToPrint$, 35) = FUsing$(STR$(TaxTran(1).Amount), "#######.##")
    MID$(ToPrint$, 78) = Post2GL$
    TOwed# = Round#(TOwed# + TaxTran(1).Amount)
    
  CASE ELSE
    BillType$ = "?????"
  END SELECT
RETURN

ExitHistRpt1:


END SUB

