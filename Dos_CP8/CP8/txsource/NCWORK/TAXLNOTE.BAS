DEFINT A-Z
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE FUNCTION MakeDay$ (DayNum%)
DECLARE FUNCTION MakeLongDate$ (PDate$)
DECLARE FUNCTION Num2Day% (DayNum%)
DECLARE FUNCTION Date2Num% (TDate$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB PrintLateNotices ()
DECLARE SUB RePrintLateNotices ()
DECLARE SUB MakeCustIndex (IdxType%)
DECLARE SUB BillsPrintedRpt ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB PrintTaxBills ()
DECLARE SUB ClearBack ()
DECLARE SUB ClearScrn ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB WaitForAction ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE FUNCTION ParseBillNum$ (Text$)
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxBINFO.BI'
  '$INCLUDE: 'Taxconst.BI'
  '$INCLUDE: 'TaxBill.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  STACK 5000
  
  
  DIM SHARED TaxSetUpRec AS TaxMasterType
  DIM SHARED TaxTransRec AS TaxTransactionType
  DIM SHARED TaxCustRec AS TaxCustType
  
  
  PrintLateNotices
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF INSTR(COMMAND$, "TEST") <= 0 THEN
    RUN "TAXBILL"
  END IF
  
  END

FUNCTION MakeDay$ (DayNum)
  SELECT CASE DayNum
  CASE 1
    MakeDay$ = "Sunday"
  CASE 2
    MakeDay$ = "Monday"
  CASE 3
    MakeDay$ = "Tuesday"
  CASE 4
    MakeDay$ = "Wednesday"
  CASE 5
    MakeDay$ = "Thursday"
  CASE 6
    MakeDay$ = "Friday"
  CASE 7
    MakeDay$ = "Saturday"
  END SELECT
  
END FUNCTION

FUNCTION MakeLongDate$ (PDate$)
  DayNum = Num2Day(Date2Num(PDate$))
  DayName$ = MakeDay$(DayNum)
  MonthName$ = MakeMonth$(PDate$)
  MakeLongDate$ = DayName$ + ", " + MonthName$ + " " + MID$(PDate$, 4, 2) + ", " + RIGHT$(PDate$, 4)
END FUNCTION

FUNCTION MakeMonth$ (TDate$)
  Month = VAL(LEFT$(TDate$, 2))
  SELECT CASE Month
  CASE 1
    MakeMonth$ = "January"
  CASE 2
    MakeMonth$ = "February"
  CASE 3
    MakeMonth$ = "March"
  CASE 4
    MakeMonth$ = "April"
  CASE 5
    MakeMonth$ = "May"
  CASE 6
    MakeMonth$ = "June"
  CASE 7
    MakeMonth$ = "July"
  CASE 8
    MakeMonth$ = "August"
  CASE 9
    MakeMonth$ = "September"
  CASE 10
    MakeMonth$ = "October"
  CASE 11
    MakeMonth$ = "November"
  CASE 12
    MakeMonth$ = "December"
  END SELECT
END FUNCTION

SUB PrintLateNotices
  
  SHARED Choice$()

  'Get Bill Format Here
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  GET TxFile, 1, TaxSetUpRec
  BillFormat$ = LEFT$(TaxSetUpRec.TaxForm, 1)
  CLOSE TxFile
  IF (INSTR(TaxSetUpRec.Name, "MURPHY") > 0 AND INSTR(TaxSetUpRec.TaxSt, "NC") > 0) THEN
    MurphyFlag = True
  END IF
  
  REDIM TempScrn(0)
  
  'Open Customer File
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = LEN(TaxCustRec)
  NumOfCRecs& = LOF(TaxFile) \ LEN(TaxCustRec)
  CLOSE TaxFile
  
  REDIM Choice$(0 TO 2, 0)
  Choice$(0, 0) = "4"
  Choice$(1, 0) = "1) Account Number Order"
  Choice$(2, 0) = "2) Customer Name Order"
  
  REDIM PINTemp(1) AS STRING * 9
  
  LibName$ = "TAX"
  FormName$ = "PRNTLN"
  
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  FirstTime = True
  
  ClearBack
  DisplayTaxScrn FormName$
  
  Form$(2, 0) = "1"
  Form$(3, 0) = STR$(NumOfCRecs&)
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      GOSUB CheckBillParm
      IF OK2Print THEN
        EXIT DO
      END IF
    CASE F5KEY
      CursorOff
      SaveScrn TempScrn()
      PrintRptFile "", "TAXLNMSK.DAT", 1, RetCode%, 4
      RestScrn TempScrn()
      Action = 2
    CASE EscKey
      ExitFlag = True
      EXIT DO
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 20
        SELECT CASE Frm(1).MCol
        CASE 22 TO 33           'ESC Cancel button
          PressButton 27, 20, 22, 33
        CASE 34 TO 48           'F5 Align Mask
          PressButton -63, 20, 34, 48
        CASE 49 TO 61           'F10 Save Button
          PressButton -68, 20, 49, 61
        END SELECT
      END SELECT
    END IF
  LOOP          'proper key not set
  
  IF ExitFlag THEN
    GOTO ExitBillPrint
  END IF
  
  IF Order$ = "2" THEN
    UsingIdx = True
    MakeCustIndex 1
  END IF
  
  IF UsingIdx THEN
    NumIdxRecs = FileSize&("TAXTEMP.IDX") / 2
    REDIM IdxBuff(1 TO NumIdxRecs) AS INTEGER
    FGetAH "TAXTEMP.IDX", IdxBuff(1), 2, NumIdxRecs
  END IF
  
  REDIM RealRec(1)  AS PropertyRecType
  REDIM PersRec(1)  AS PersonalRecType
  
  RealRecLen = LEN(RealRec(1))
  PersRecLen = LEN(PersRec(1))
  
  ClearBack
  ShowProcessingScrn "Creating Late Notices"
  
  
  RptFile = FREEFILE
  OPEN "TAXLN.PRN" FOR OUTPUT AS RptFile
  
  'Open Transaction File
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTransRec)
  'Open Customer File
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = LEN(TaxCustRec)
  NumOfCRecs& = LOF(TaxFile) \ LEN(TaxCustRec)
  
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  
  FOR Cnt& = 1 TO NumOfCRecs&
    
    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF
    
    IF CustAcct& >= FirstAcct& AND CustAcct& <= LastAcct& THEN
      
      GET #TaxFile, CustAcct&, TaxCustRec       'get cust record
      
      DoneHeader = False
      DoFooter = False
      TBalance# = 0
      
      CustName$ = RTRIM$(TaxCustRec.FName) + " " + RTRIM$(TaxCustRec.LNAME)
      CustName$ = LTRIM$(CustName$)
      
      'Loop Thru and Get any Current Bill And Print Late Notice For That Bill Only
      
      TransRecord& = TaxCustRec.LastTrans
      
      WHILE TransRecord& > 0
        
        GET TransFile, TransRecord&, TaxTransRec
        IF MurphyFlag THEN
          IF TaxTransRec.TranType <> 1 THEN
            GOTO NotABill
          END IF
        ELSE
          IF (TaxTransRec.TranType <> 1 AND TaxTransRec.TaxYear <> WhatYear) THEN
            'IF TaxTransRec.TranType = 1 THEN 'AND TaxTransRec.TaxYear = WhatYear THEN
            GOTO NotABill
          END IF
        END IF
        
        Balance# = 0
        TxIntP# = Round#(TaxTransRec.Revenue.InterestPd + TaxTransRec.Revenue.PenaltyPd + TaxTransRec.Revenue.CollectionPd)
        TxIntC# = Round#(TaxTransRec.Revenue.Interest + TaxTransRec.Revenue.Penalty + TaxTransRec.Revenue.Collection)
        Balance# = Round#(TaxTransRec.Revenue.Principle1 + TaxTransRec.Revenue.Principle2 + TaxTransRec.Revenue.Principle3 + TaxTransRec.Revenue.Principle4 + TaxTransRec.Revenue.Principle5)
        Balance# = Round#(Balance# - (TaxTransRec.Revenue.Principle1Pd + TaxTransRec.Revenue.Principle2Pd + TaxTransRec.Revenue.Principle3Pd + TaxTransRec.Revenue.Principle4Pd + TaxTransRec.Revenue.Principle5Pd))
        Balance# = Round#(Balance# + TaxTransRec.Revenue.Interest + TaxTransRec.Revenue.Penalty + TaxTransRec.Revenue.Collection)
        Balance# = Round#(Balance# - (TaxTransRec.Revenue.InterestPd + TaxTransRec.Revenue.PenaltyPd + TaxTransRec.Revenue.CollectionPd))
        
        IF Balance# > 0 THEN
          TBalance# = Round#(Balance# + TBalance#)
          'IF BillFormat$ = "P" THEN
          ''$INCLUDE: 'TxTROYLN.BI'    'Troy
          ''$INCLUDE: 'TaxSTDLN.BI'
          ''$INCLUDE: 'TaxSUNLN.BI'      'Sun Beech
          ''$INCLUDE: 'TaxSUGLN.BI'
          ''$INCLUDE: 'TXNWSTLN.BI'     'Northwest
          ''$INCLUDE: 'TXMAGNLN.BI'     'Magnolia
          ''$INCLUDE: 'TXCLAKLN.BI'     'Clarkton NC
          '$INCLUDE: 'TXmrphln.BI'     '
          'END IF
          PrnCnt = PrnCnt + 1
        END IF
NotABill:
        TransRecord& = TaxTransRec.LastTrans
      WEND
      IF MurphyFlag THEN
        IF DoneHeader THEN
          DoFooter = True
          '$INCLUDE: 'TXmrphln.BI'
          'PrnCnt = PrnCnt + 1
        END IF
      END IF
    END IF      'Test for Valid Acct
SkipEm:
    IF PrnCnt > 2 AND INSTR(COMMAND$, "TEST") > 0 THEN EXIT FOR
    ShowPctComp Cnt&, NumOfCRecs&
  NEXT Cnt&
  
  CLOSE
  
  ERASE Frm, Form$, Fld, TempScrn
  
  PrintRptFile "Late Notice Printing.", "TAXLN.PRN", 1, RetCode%, 0
  
ExitBillPrint:
  EXIT SUB
  
CheckBillParm:
  WhatYear = Value#(Form$(1, 0), Ecode)
  FirstAcct& = Value#(Form$(2, 0), Ecode)
  LastAcct& = Value#(Form$(3, 0), Ecode)
  Order$ = LEFT$(Form$(4, 0), 1)
  AdvDate$ = Form$(5, 0)
  PayDate$ = Form$(6, 0)
  PDate$ = UCASE$(MakeLongDate$(PayDate$))
  ADate$ = UCASE$(MakeLongDate$(AdvDate$))
  IF (WhatYear > 0) OR (MurphyFlag = True) THEN
    OK2Print = True
  END IF

  RETURN
  
GetRealInfo:    'DO NOT REMOVE
  'This is called from Sugar Mountains include file
  
  ThisProp& = TaxCustRec.FirstPropRec
  RealCalcVal# = 0
  CalcVal# = 0
  PersCalcVal# = 0
  LSET PINTemp(1) = ""
  
  IF ThisProp& > 0 THEN
    GET #RealFile, ThisProp&, RealRec(1)
    PINTemp(1) = RealRec(1).REALPIN
    RealValue# = RealRec(1).PROPVALU
    RealExmp# = Round#(RealRec(1).EXMPSENI + RealRec(1).EXMPOTHR)
    RealCalcVal# = Round#(RealValue# - RealExmp#)
  END IF
  RETURN
  
GetPersInfo:    'DO NOT REMOVE
  'This is called from Sugar Mountains include file
  
  PersExmp# = 0
  PersValue# = 0
  IF TaxCustRec.FirstPersRec > 0 THEN
    GET #PersFile, TaxCustRec.FirstPersRec, PersRec(1)
    PersValue# = Round#(PersRec(1).PersVal + PersRec(1).MHVALUE + PersRec(1).MCVALUE)
    PersValue# = Round#(PersValue# + PersRec(1).CVALUE + PersRec(1).MTVALUE)
    PersExmp# = Round#(PersRec(1).EXMPSENI + PersRec(1).EXMPOTHR)
    PersCalcVal# = Round#(PersValue# - PersExmp#)
    IF PersCalcVal# < 0 THEN
      PersCalcVal# = 0
    END IF
  END IF
  
  RETURN
  
END SUB

