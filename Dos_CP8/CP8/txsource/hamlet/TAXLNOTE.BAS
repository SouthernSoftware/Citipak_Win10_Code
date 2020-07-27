DEFINT A-Z
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
'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
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
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxBINFO.BI'
  '$INCLUDE: 'Taxconst.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TaxBill.BI'
  
  STACK 5000


  DIM SHARED TaxSetUpRec AS TaxMasterType
  DIM SHARED TaxTransRec AS TaxTransactionType
  DIM SHARED TaxCustRec AS TaxCustType
  DIM SHARED TaxPropRec AS PropertyRecType

  
  PrintLateNotices
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  
'  RUN "TAXBILL"
  
  END

SUB PrintLateNotices
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  DIM TaxYear$(11), TaxAmt#(11), OthAmt#(11), TaxRecpt$(11)
  
 'Open Customer File
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = LEN(TaxCustRec)
  NumOfCRecs& = LOF(TaxFile) \ LEN(TaxCustRec)
  CLOSE TaxFile

  REDIM Choice$(0 TO 2, 0)
  Choice$(0, 0) = "6"
  Choice$(1, 0) = "1) Account Number Order"
  Choice$(2, 0) = "2) Customer Name Order"
  
  REDIM PINTemp(1) AS STRING * 9
  
  LibName$ = "TAX"
  FormName$ = "PRNTLNHT"
  
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
  
  ClearBack
  ShowProcessingScrn "Creating Late Notices"

 'Get Bill Format Here
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  GET TxFile, 1, TaxSetUpRec
   BillFormat$ = LEFT$(TaxSetUpRec.TaxForm, 1)
  CLOSE TxFile

  RptFile = FREEFILE
  OPEN "TAXLN.PRN" FOR OUTPUT AS RptFile
  
 'Open Transaction File
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTransRec)
 'Open Customer File
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = LEN(TaxCustRec)
  NumOfCRecs& = LOF(TaxFile) \ LEN(TaxCustRec)
 'Open TaxProp File
  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = LEN(TaxPropRec)

  FOR Cnt& = 1 TO NumOfCRecs&

    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF

    IF CustAcct& >= FirstAcct& AND CustAcct& <= LastAcct& THEN

    GET #TaxFile, CustAcct&, TaxCustRec         'get cust record

    IF TaxCustRec.Bankrupt = "Y" THEN
      GOTO LateNSkip
    END IF
    IF TaxCustRec.LateNotice <> "Y" THEN
      GOTO LateNSkip
    END IF

    'Loop Thru and Get any Current Bill And Print Late Notice For That Bill Only

     TransRecord& = TaxCustRec.LastTrans
     Tnt = 0

      WHILE TransRecord& <> 0

      GET TransFile, TransRecord&, TaxTransRec
      IF TaxTransRec.TranType = 1 THEN
        Balance# = 0
        Balance# = Round#(TaxTransRec.Revenue.Principle1 + TaxTransRec.Revenue.Principle2 + TaxTransRec.Revenue.Principle3 + TaxTransRec.Revenue.Principle4 + TaxTransRec.Revenue.Principle5)
        Balance# = Round#(Balance# + TaxTransRec.Revenue.interest + TaxTransRec.Revenue.Penalty + TaxTransRec.Revenue.Collection)
        Balance# = Round#(Balance# - (TaxTransRec.Revenue.Principle1Pd + TaxTransRec.Revenue.Principle2Pd + TaxTransRec.Revenue.Principle3Pd + TaxTransRec.Revenue.Principle4Pd + TaxTransRec.Revenue.Principle5Pd))
        Balance# = Round#(Balance# - (TaxTransRec.Revenue.InterestPd + TaxTransRec.Revenue.PenaltyPd + TaxTransRec.Revenue.CollectionPd))
     
      IF Balance# > 0 THEN
       Tnt = Tnt + 1
       k$ = ParseBillNum$(TaxTransRec.Description)
       TaxYear$(Tnt) = LTRIM$(STR$(TaxTransRec.TaxYear))
       TaxRecpt$(Tnt) = k$
       TaxAmt#(Tnt) = Round#(TaxTransRec.Revenue.Principle1 + TaxTransRec.Revenue.Principle2 + TaxTransRec.Revenue.Principle3 + TaxTransRec.Revenue.Principle4 + TaxTransRec.Revenue.Principle5)
       TaxAmt#(Tnt) = TaxAmt#(Tnt) - Round#(TaxTransRec.Revenue.Principle1Pd + TaxTransRec.Revenue.Principle2Pd + TaxTransRec.Revenue.Principle3Pd + TaxTransRec.Revenue.Principle4Pd + TaxTransRec.Revenue.Principle5Pd)
       OthAmt#(Tnt) = Round#(TaxTransRec.Revenue.interest + TaxTransRec.Revenue.Penalty + TaxTransRec.Revenue.Collection - (TaxTransRec.Revenue.InterestPd + TaxTransRec.Revenue.PenaltyPd + TaxTransRec.Revenue.CollectionPd))
      END IF
      END IF         'End Test for Bill / Year
      TransRecord& = TaxTransRec.LastTrans
      WEND

       IF Tnt > 0 THEN

          CustName$ = RTRIM$(TaxCustRec.FName) + " " + RTRIM$(TaxCustRec.LName)
          CustName$ = LTRIM$(CustName$)
          PropertyRecord& = TaxCustRec.FirstPropRec
          IF PropertyRecord& > 0 THEN
           GET PropFile, PropertyRecord&, TaxPropRec
           PDesc$ = TaxPropRec.RealPIN
          END IF

       
       '$INCLUDE: 'TaxHMTLN.BI'    'Hamlet Late Notice
       
      'Clear Array for Next Bill
       FOR C = 1 TO Tnt
        TaxYear$(C) = ""
        TaxAmt#(C) = 0
        TaxAmt#(C) = 0
        OthAmt#(C) = 0
       NEXT C


       END IF
   END IF     'Test for Valid Acct
    PrnCnt = PrnCnt + 1
LateNSkip:
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
  Mess1$ = Form$(4, 0)
  Mess2$ = Form$(5, 0)
  Order$ = LEFT$(Form$(6, 0), 1)
  IF WhatYear > 0 THEN
   OK2Print = True
  END IF
RETURN
  
END SUB

