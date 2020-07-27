DEFINT A-Z
DECLARE SUB SortServiceAddrs (IndexText$)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB FCreate (FileName$)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB BlockClear ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (array())
DECLARE SUB SaveScrn (array())
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WaitForAction ()
DECLARE SUB HideCursor ()
DECLARE SUB ShowCursor ()
DECLARE SUB CursorOff ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION GetNumOfCust% ()

  CONST False = 0, True = NOT False

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'

   FF$ = CHR$(12)

SUB PaymentSumReport
  
  REDIM TempScrn(0)
  REDIM RevenueName$(15)
  REDIM Revenues(1 TO 15) AS DOUBLE
  REDIM TaxRates(1 TO 15) AS SINGLE
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  FOR RCnt = 1 TO 15
    TRevName$ = QPTrim$(UBSetup(1).Revenues(RCnt).RevName)
    IF LEN(TRevName$) > 0 THEN
      RevenueName$(RCnt) = TRevName$
      TaxRates(RCnt) = UBSetup(1).Revenues(RCnt).TAXRATE
    ELSE
      MaxRevenue = RCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  TownName$ = UBSetup(1).UTILNAME
  TownLen = LEN(RTRIM$(TownName$))
  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1
  
  '*******************************
  
SelectFromList3:
  LibName$ = "UB"
  ScrnName$ = "UBPAYSUM"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Set Defaults as Anticpated Response
  Form$(1, 0) = DATE$
  Form$(2, 0) = DATE$
  Action = 1
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      BegDate = Date2Num(Form$(1, 0))
      
      EndDate = Date2Num(Form$(2, 0))
      IF EndDate < BegDate THEN
        Done = False
        BEEP
        SaveScrn TempScrn()
        CursorOff
        DisplayUBScrn "UBBADDAT"
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        Done = True
      END IF
      
    CASE ESC
      GOTO ExitDetailedListing3
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  '***************
  ' Set Up Specifications from Input Screen
  Operator$ = Form$(3, 0)
  Operator = VAL(Operator$)
  FromDate$ = Form$(1, 0)
  ToDate$ = Form$(2, 0)
  
  IF Operator = 0 THEN
    BegOperator = 0
    EndOperator = 99
  ELSE
    BegOperator = Operator
    EndOperator = Operator
  END IF
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  UBRpt = FREEFILE
  OPEN "UBPAYSUM.RPT" FOR OUTPUT AS UBRpt
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  
  NumOfRecs& = LOF(UBTrans) \ UBTransRecLen
  
  BlockClear
  ShowProcessingScrn "Payment Summary Report."
  
  GOSUB DoDetailedRptHeader3
  
  FOR Cnt& = 1 TO NumOfRecs&
    GET UBTrans, Cnt&, UBTrans(1)
    TransOK = False
    IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) AND (UBTrans(1).OperatorNumber >= BegOperator AND UBTrans(1).OperatorNumber <= EndOperator) THEN
      SELECT CASE UBTrans(1).TransType
      CASE TranBillPayment, TranBillPayment + 100
        TransOK = True
      CASE TranDraftPayment, TranDraftPayment + 100
        TransOK = True
      'CASE TranDepositPayment, TranDepositPayment + 100
      '  TransOK = True
      END SELECT
      
      IF TransOK THEN
        IF UBTrans(1).TaxExempt = "Y" THEN
          TaxExempt = True
          GOTO SkipEm
          'LPRINT UBTrans(1).CustAcctNo
        ELSE
          TaxExempt = False
        END IF
        
        FOR RevCnt = 1 TO 15
          IF NOT TaxExempt THEN
            IF TaxRates(RevCnt) > 0 THEN
              Diff# = Round#(UBTrans(1).RevAmt(RevCnt) / (1 + TaxRates(RevCnt)))
              Tax# = Round#(UBTrans(1).RevAmt(RevCnt) - Diff#)
              TaxAmt(RevCnt) = Round#(TaxAmt(RevCnt) + Tax#)
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + (UBTrans(1).RevAmt(RevCnt) - Tax#))
            ELSE
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTrans(1).RevAmt(RevCnt))
            END IF
          ELSE
            Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTrans(1).RevAmt(RevCnt))
          END IF
        NEXT
        TransCnt& = TransCnt& + 1
      END IF
    END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
SkipEm:
    ShowPctCompL Cnt&, NumOfRecs&
  NEXT
  
  GOSUB DoDetailedRptFooter3
  
  CLOSE
  
  IF NOT AbortFlag THEN
    PrintRptFile "Payment Summary Report.", "UBPAYSUM.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBPAYSUM.RPT"
  
ExitDetailedListing3:
  
  EXIT SUB
  
DoDetailedRptHeader3:
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(29); "Payment Summary Report"
  PRINT #UBRpt, "Beginning Date: "; FromDate$;
  IF VAL(Operator$) = 0 THEN
    PRINT #UBRpt, TAB(65); " Operator #: ALL"
  ELSE
    PRINT #UBRpt, TAB(65); " Operator #: "; Operator$
  END IF
  PRINT #UBRpt, "   Ending Date: "; ToDate$
  PRINT #UBRpt,
  PRINT #UBRpt, "    Source                           Revenue Amt                 Tax"
RETURN
  
DoDetailedRptFooter3:
  PRINT #UBRpt, Dash80$
  FOR Cnt = 1 TO MaxRevenue
    PRINT #UBRpt, TAB(5); RevenueName$(Cnt); TAB(35); USING "$$#######,#.##"; Revenues(Cnt); TAB(55); TaxAmt(Cnt)
    TotalTrans# = Round#(TotalTrans# + Revenues(Cnt))
    TaxTotal# = Round#(TaxTotal# + TaxAmt(Cnt))
  NEXT
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Total Payments: "; TAB(20); USING "######"; TransCnt&
  PRINT #UBRpt, "Revenue Totals: "; TAB(35); USING "$$#######,#.##"; TotalTrans#; TAB(55); TaxTotal#
  PRINT #UBRpt, FF$
RETURN
  
END SUB

SUB SortServiceAddrs (IndexText$)
  
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs = GetNumOfCust%
  
  REDIM ServIndex(1 TO NumCustRecs)  AS UBServiceAddressIndexType
  IndexRecLen = LEN(ServIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs
    GET CHandle, Cnt, UBCustRec(1)
    ServIndex(Cnt).ServiceAddress = UBCustRec(1).ServAddr
    ServIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT ServIndex(1), NumCustRecs, 0, 16, 0, 14
  ' SortT (Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IHandle = FREEFILE
  IndexName$ = TempIndexName
  FCreate IndexName$
  OPEN IndexName$ FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO NumCustRecs
    CRec& = ServIndex(Cnt).RecNum
    PUT IHandle, Cnt, CRec&
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, ServIndex
  
  
END SUB

