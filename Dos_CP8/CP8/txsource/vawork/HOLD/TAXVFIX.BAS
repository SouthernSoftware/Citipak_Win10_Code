DECLARE SUB CustHistoryRpt2 ()
DECLARE SUB CustHistoryRpt1 ()
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

CustHistoryRpt2


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

  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  FOR Rec& = 1 TO LOF(TaxFile) / TaxCustRecLen

  GET TaxFile, Rec&, TaxCustRec(1)
  DidCnt = 0
  PrevTranRec& = TaxCustRec(1).LastTrans


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
      BillRecord& = HistRecs(Cnt).TranRec
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
        NEXT
      END IF
      ToPrint$ = SPACE$(79)
      MID$(ToPrint$, 23) = "Totals:"
      MID$(ToPrint$, 35) = FUsing$(STR$(TOwed#), "#######.##")
      MID$(ToPrint$, 50) = FUsing$(STR$(TPaid#), "#######.##")
      MID$(ToPrint$, 65) = FUsing$(STR$(TOwed# - TPaid#), "#######.##")
      Balance@ = TOwed# - TPaid#
      Balance@ = INT((Balance@ * 100) + .5) / 100

      IF Balance@ = 0 OR Balance@ = .01 OR Balance@ = -.01 THEN
       GET TaxTran, BillRecord&, TaxTran(1)
       TaxTran(1).Revenue.Principle1Pd = TaxTran(1).Revenue.Principle1
       TaxTran(1).Revenue.Principle2Pd = TaxTran(1).Revenue.Principle2
       TaxTran(1).Revenue.Principle3Pd = TaxTran(1).Revenue.Principle3
       TaxTran(1).Revenue.Principle4Pd = TaxTran(1).Revenue.Principle4
       TaxTran(1).Revenue.Principle5Pd = TaxTran(1).Revenue.Principle5
       TaxTran(1).Revenue.InterestPd = TaxTran(1).Revenue.Interest
       TaxTran(1).Revenue.PenaltyPd = TaxTran(1).Revenue.Penalty
       TaxTran(1).Revenue.CollectionPd = TaxTran(1).Revenue.Collection
       PUT TaxTran, BillRecord&, TaxTran(1)
      END IF
      
    END IF

  NEXT

  NEXT Rec&
  CLOSE
  END

  EXIT SUB


PrintBillInfo:
  RETURN

CustHistRptHeader:
  RETURN

CustHistRptFooter:
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
  '  MID$(ToPrint$, 30) = "Trans #" + STR$(HistRecs(Cnt).TranRec)
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
  TaxType = 1

  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  FOR Rec& = 1 TO LOF(TaxFile) / TaxCustRecLen

  GET TaxFile, Rec&, TaxCustRec(1)
  DidCnt = 0
  PrevTranRec& = TaxCustRec(1).LastTrans


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
      BillRecord& = HistRecs(Cnt).TranRec
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
        NEXT
      END IF
      ToPrint$ = SPACE$(79)
      MID$(ToPrint$, 23) = "Totals:"
      MID$(ToPrint$, 35) = FUsing$(STR$(TOwed#), "#######.##")
      MID$(ToPrint$, 50) = FUsing$(STR$(TPaid#), "#######.##")
      MID$(ToPrint$, 65) = FUsing$(STR$(TOwed# - TPaid#), "#######.##")
      Balance@ = TOwed# - TPaid#
      Balance@ = INT((Balance@ * 100) + .5) / 100
      
      IF Balance@ = 0 OR Balance@ = .01 OR Balance@ = -.01 THEN
       GET TaxTran, BillRecord&, TaxTran(1)
       TaxTran(1).Revenue.Principle1Pd = TaxTran(1).Revenue.Principle1
       TaxTran(1).Revenue.Principle2Pd = TaxTran(1).Revenue.Principle2
       TaxTran(1).Revenue.Principle3Pd = TaxTran(1).Revenue.Principle3
       TaxTran(1).Revenue.Principle4Pd = TaxTran(1).Revenue.Principle4
       TaxTran(1).Revenue.Principle5Pd = TaxTran(1).Revenue.Principle5
       TaxTran(1).Revenue.InterestPd = TaxTran(1).Revenue.Interest
       TaxTran(1).Revenue.PenaltyPd = TaxTran(1).Revenue.Penalty
       TaxTran(1).Revenue.CollectionPd = TaxTran(1).Revenue.Collection
       PUT TaxTran, BillRecord&, TaxTran(1)
      END IF

    END IF

  NEXT

  NEXT Rec&
  CLOSE
  END

  EXIT SUB


PrintBillInfo1:
  RETURN

CustHistRptHeader1:
  RETURN

CustHistRptFooter1:
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
  '  MID$(ToPrint$, 30) = "Trans #" + STR$(HistRecs(Cnt).TranRec)
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

ExitHistRpt1:


END SUB

