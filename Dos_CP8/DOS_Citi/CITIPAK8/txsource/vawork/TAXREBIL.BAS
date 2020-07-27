DEFINT A-Z
DECLARE FUNCTION MakeDay$ (DayNum%)
DECLARE FUNCTION MakeLongDate$ (PDate$)
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE FUNCTION Num2Day% (DayNum%)
DECLARE SUB MakeCustIndexRE (IdxType%)
DECLARE SUB LateNotices ()
DECLARE SUB PostTaxBills ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB MakeCustIndex (IdxType)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB TAXPrebilling ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
  
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION ParseBillNum$ (Text$)
  
  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxBill.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXBINFO.BI'
  
  DIM SHARED TaxSetUpRec AS TaxMasterType
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = "Print Pre-Billing Register  "
  MChoice$(2) = "Print/Reprint Tax Bills "
  MChoice$(3) = "Post Tax Bills"
  MChoice$(4) = "Apply Penalty"
  MChoice$(5) = "Apply Interest"
  MChoice$(6) = "Print Late Notice's"
  MChoice$(7) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((24 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Real Estate Tax Billing Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      TAXPrebilling
    CASE 2
      RUN "TXPRNTRB"
    CASE 3
      PostTaxBills
    CASE 4
      RUN "TAXREPEN"
    CASE 5
      RUN "TAXREINT"
    CASE 6
      LateNotices
    CASE 7
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("Taxmenu" + Ext$) THEN
    RUN "TaxMenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB LateNotices

  SHARED Choice$()
  
  'Get Bill Format Here
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  GET TxFile, 1, TaxSetUpRec
  CLOSE TxFile

  BillFormat$ = LEFT$(TaxSetUpRec.TaxForm, 1)
  TownName$ = QPTrim$(TaxSetUpRec.Name)

  IF INSTR(TownName$, "MARKET") > 0 THEN
    MrktFlag = True
  END IF

  IF (INSTR(TaxSetUpRec.Name, "POCAHO")) > 0 THEN
    SpecialFlag = True
  END IF

  IF (INSTR(TaxSetUpRec.Name, "REMINGTON")) > 0 THEN
    SpecialFlag = True
  END IF

  IF (INSTR(TaxSetUpRec.Name, "BROOKNEAL")) > 0 THEN
    SpecialFlag = True
  END IF
  
  IF INSTR(TownName$, "MONTROSS") > 0 THEN
    MontFlag = True
  END IF

  SpecialFlag = True


  TownLen = LEN(TownName$)
  ADD1$ = QPTrim$(TaxSetUpRec.ADD1)
  Add1Len = LEN(ADD1$)
  ADD2$ = QPTrim$(TaxSetUpRec.ADD2)
  Add2Len = LEN(ADD2$)
  ADD3$ = QPTrim$(TaxSetUpRec.ADD3)
  Add3Len = LEN(ADD3$)

  t1 = 5
  s1 = t1
  s2 = t1
  s3 = t1

  IF TownLen > 0 THEN
    t1 = 40 - (TownLen / 2)
  END IF
  IF Add1Len > 0 THEN
    s1 = 40 - (Add1Len / 2)
  END IF
  IF Add2Len > 0 THEN
    s2 = 40 - (Add2Len / 2)
  END IF
  IF Add3Len > 0 THEN
    s3 = 40 - (Add3Len / 2)
  END IF
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM Fmt$(1 TO 3)

  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  RealRecLen = LEN(RealRec(1))
  TransLen = LEN(TaxTrans(1))
  
  REDIM TempScrn(0)
  REDIM Choice$(0 TO 2, 0)

  Choice$(0, 0) = "2"
  Choice$(1, 0) = "1- Account Number Order"
  Choice$(2, 0) = "2- Customer Name Order"
  
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  LNDate$ = MakeLongDate$(TheDate$)
  Fmt$(1) = "$$#########,#"
  Fmt$(2) = "#####"
  Fmt$(3) = "########,#"
  
  MaxLines = 60
  LineCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "RELATE"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  FirstTime = True
  
  '--Clear all fields
  ClearBack
  DisplayTaxScrn ScrnName$
  IF MrktFlag THEN
    QPrintRC "     Due Date:", 9, 24, -1
  END IF
  ShowCursor
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = 0
      IF MontFlag THEN
        QPrintRC "   Notice Date", 9, 23, -1
      END IF
      IF WEnvTest% THEN
        LSET Form$(1, 0) = "2001"
        LSET Form$(2, 0) = Choice$(1, 0)
        LSET Form$(3, 0) = DATE$
        'FOR zz = 1 TO 5
        '  LSET Form$(zz + 3, 0) = STRING$(50, 64 + zz)
        'NEXT
        Action = 1
      END IF
    END IF
    
    IF (Frm(1).KeyCode = -72) OR (Frm(1).KeyCode = -15) THEN
      FOR zz = Frm(1).FldNo TO 1 STEP -1        'loop backwards
        IF Fld(zz).Protected = False THEN
          Frm(1).FldNo = zz
          Action = 1
          EXIT FOR
        END IF
      NEXT
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      CursorOff
      SaveScrn TempScrn()
      PrintRptFile "", "TAXRELN.MSK", 1, RetCode%, 4
      RestScrn TempScrn()
      Action = 2
    CASE -68    'F10Key    Proceed with look up
      GOSUB CheckPreInfo1
      IF PreInfoOK THEN
        OKFlag = True
      END IF
    CASE 27
      ExitFlag = True
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 20
        SELECT CASE Frm(1).MCol
        CASE 29 TO 40           'ESC Cancel button
          PressButton 27, 20, 29, 40
        CASE 42 TO 54           'F10 Proceed Button
          PressButton -68, 20, 42, 54
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  IF ExitFlag THEN
    GOTO ExitLateNot
  END IF
  
  ClearBack
  
  IF Order$ = "2" THEN
    UsingIdx = True
    MakeCustIndexRE 1
  END IF
  IF UsingIdx THEN
    NumIdxRecs = FileSize&("TAXTEMP.IDX") / 2
    REDIM IdxBuff(1 TO NumIdxRecs) AS INTEGER
    FGetAH "TAXTEMP.IDX", IdxBuff(1), 2, NumIdxRecs
  END IF
  
  ShowProcessingScrn "Creating Tax Late Notices"
  
  RptFile = FREEFILE
  OPEN "TAXLN.REP" FOR OUTPUT AS #RptFile
  
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfCRecs& = LOF(TaxFile) \ TaxCustLen
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM SHARED AS TransFile LEN = TransLen

  'Late Notice calc start *******************************

  FOR Cnt& = 1 TO NumOfCRecs&
    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF
    
    GET #TaxFile, CustAcct&, TaxCust(1)         'get cust record

    DoneHeader = False
    DoFooter = False
    TBalance# = 0

    'IF CustAcct& = 10 THEN STOP
    'here
    'IF TaxCust(1).FirstPropRec <= 0 THEN
    '  GOTO LateNotSkip
    'END IF
    'Process Each Trans Looking for Outstanding Tax Owed
    
    TRecord& = TaxCust(1).LastTrans
    
    WHILE TRecord& > 0
      GET TransFile, TRecord&, TaxTrans(1)
      IF (TaxTrans(1).TranType = 1) AND (TaxTrans(1).BillType = "R") AND (TaxTrans(1).TaxYear <= WhatYear) THEN
      'IF (TaxTrans(1).TranType = 1) AND (TaxTrans(1).BillType = "R") THEN
      'IF (TaxTrans(1).TranType = 1) AND (TaxTrans(1).BillType = "R") AND (WhatYear = TaxTrans(1).TAXYEAR) THEN
        Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
        Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection
        Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
        Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
        Balance# = Round#(Balance#)

        IF Balance# > 0 THEN
          TBalance# = Round#(Balance# + TBalance#)

          TaxDue# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
          TaxDue# = TaxDue# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
          PenaltyDue# = TaxTrans(1).Revenue.Penalty - TaxTrans(1).Revenue.PenaltyPd
          InterestDue# = TaxTrans(1).Revenue.Interest - TaxTrans(1).Revenue.InterestPd
          BillNumber$ = ParseBillNum$(TaxTrans(1).Description)
          TaxYear$ = QPTrim$(STR$(TaxTrans(1).TaxYear))
          GOSUB PrintLateNotice
        END IF
      END IF    'end test for trans=bill
      TRecord& = TaxTrans(1).LastTrans
    WEND

'do header/footer here
    IF SpecialFlag THEN
      IF DoneHeader THEN
        DoFooter = True
        '$INCLUDE: 'TXPOCPLN.BI'            'Pocahontas
        PrnCnt = PrnCnt + 1
      END IF
    END IF

    ShowPctComp Cnt&, NumOfCRecs&
    IF INSTR(COMMAND$, "TEST") > 0 THEN
      IF PrnCnt& > 1 THEN
        EXIT FOR
      END IF
    END IF
    
LateNotSkip:
  NEXT
  CLOSE
  
  PrintRptFile "Tax Late Notices", "TAXLN.REP", 1, RetCode%, 0
  
ExitLateNot:
  EXIT SUB
  
PrintLateNotice:
  RealFlag = True
  'Get PropDesc From Record 1
  PRec& = TaxCust(1).FirstPropRec
  DESC1$ = ""
  DESC2$ = ""
  DESC3$ = ""
  IF PRec& > 0 THEN
    GET RealFile, PRec&, RealRec(1)
    DESC1$ = QPTrim$(RealRec(1).PROPNOT1)
    DESC2$ = QPTrim$(RealRec(1).PROPNOT2)
    DESC3$ = QPTrim$(RealRec(1).PROPNOT3)
  END IF
  
  CustName$ = QPTrim$(TaxCust(1).FName) + " " + QPTrim$(TaxCust(1).LNAME)
  CitySt$ = QPTrim$(TaxCust(1).City) + " " + TaxCust(1).State
  ''$INCLUDE: 'TXPOCPLN.BI'            '
  ''$INCLUDE: 'TXNMRELN.BI'            'New market VA
  ''$INCLUDE: 'TAXELKLN.BI'            'Elkton
  ''$INCLUDE: 'TAXHALLN.BI'            'Halfix
  ''$INCLUDE: 'TAXRESLN.BI'            'Regular Tax Bill We Sell
  ''$INCLUDE: 'TXASHLN.BI'             'Ashland Special
  ''$INCLUDE: 'TxMONTRS.BI'            '
  ''$INCLUDE: 'TaxSTDLN.BI'
  ''$INCLUDE: 'TaxSUGLN.BI'
  ''$INCLUDE: 'TXNWSTLN.BI'            'Northwest
  
  PrnCnt& = PrnCnt& + 1
RETURN
  
CheckPreInfo1:
  WhatYear = Value#(Form$(1, 0), ECode)
  Order$ = LEFT$(Form$(2, 0), 1)
  InterestDate$ = Form$(3, 0)

  Mess1$ = QPTrim$(Form$(4, 0))
  Mess2$ = QPTrim$(Form$(5, 0))
  Mess3$ = QPTrim$(Form$(6, 0))
  Mess4$ = QPTrim$(Form$(7, 0))
  Mess5$ = QPTrim$(Form$(8, 0))
  
  IF WhatYear < 1990 OR WhatYear > 2015 THEN
    GOSUB ShowInfoErr1
    Frm(1).FldNo = 1
    GOTO InfoErrorExit1
  END IF
  
  IF LEN(QPTrim$(Order$)) = 0 THEN
    GOSUB ShowInfoErr1
    Frm(1).FldNo = 2
    GOTO InfoErrorExit1
  END IF
  
  IF LEN(QPTrim$(InterestDate$)) = 0 THEN
    GOSUB ShowInfoErr1
    Frm(1).FldNo = 3
    GOTO InfoErrorExit1
  'ELSE
  '  LNDate$ = MakeLongDate$(InterestDate$)
  END IF
  
  PreInfoOK = True
  
InfoErrorExit1:
  
  RETURN
  
ShowInfoErr1:
  SaveScrn TempScrn()
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "Invalid Late Note Information!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn TempScrn()
RETURN
  
GetPastBalance1:
  Balance# = 0
  ABalance# = 0
  IF TaxCust(1).LastTrans > 0 THEN
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
    TransRecord& = TaxCust(1).LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN
        Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
        Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection
        Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
        Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
        Balance# = Round#(Balance#)
      END IF
      ABalance# = ABalance# + Balance#
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  END IF
RETURN
  
END SUB

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

SUB PostTaxBills
  
  REDIM TaxBill(1)  AS TaxBillType
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM PersRec(1)  AS PersonalRecType
  REDIM BillInfo(1) AS TaxBillInfoType
  REDIM Revenue(1) AS RevSourceType
  
  BillLen = LEN(TaxBill(1))
  BillInfoLen = LEN(BillInfo(1))
  RealRecLen = LEN(RealRec(1))
  PersRecLen = LEN(PersRec(1))
  
  FGetAH "TAXBINFO.DAT", BillInfo(1), BillInfoLen, 1
  
  REDIM TempScrn(0)
  
  ClearBack
  ShowProcessingScrn "Verifying Billing Information"
  BillFile = FREEFILE
  OPEN RETaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      Printed = Printed + 1
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  
  IF Printed < 1 THEN
    GOSUB ShowPostError
    GOTO PostExit
  END IF
  
  PostFlag = MsgBox%("TAX", "POSTBILL")
  
  SELECT CASE PostFlag
  CASE 1
    GOTO Postem
  CASE 2
    CursorOff
    ClearBack
    DisplayTaxScrn "POSTCAN"
    WaitForAction
    GOTO PostExit
  END SELECT
  
Postem:
  ClearBack
  ShowProcessingScrn "Posting Tax Billing Transactions"
  
  'Open All Needed Files Here
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  
  BillFile = FREEFILE
  OPEN RETaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  CustFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM AS CustFile LEN = LEN(TaxCust(1))
  
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      'Update the Transaction File First
      TaxTrans(1).TransDate = Date2Num%(DATE$)
      TaxTrans(1).TaxYear = TaxBill(1).TaxYear
      TaxTrans(1).TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "R"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TaxBill(1).TotalBillDue              'Total Transaction Amount
      TaxTrans(1).Revenue.Principle1 = TaxBill(1).TotalBillDue
      TaxTrans(1).Revenue.Principle2 = 0
      TaxTrans(1).Revenue.Principle3 = 0
      TaxTrans(1).Revenue.Principle4 = 0
      TaxTrans(1).Revenue.Principle5 = 0
      TaxTrans(1).Revenue.Interest = 0
      TaxTrans(1).Revenue.Penalty = 0
      TaxTrans(1).Revenue.Collection = 0
      TaxTrans(1).Revenue.Future1 = 0
      TaxTrans(1).Revenue.Future2 = 0
      TaxTrans(1).Revenue.Principle1Pd = 0
      TaxTrans(1).Revenue.Principle2Pd = 0
      TaxTrans(1).Revenue.Principle3Pd = 0
      TaxTrans(1).Revenue.Principle4Pd = 0
      TaxTrans(1).Revenue.Principle5Pd = 0
      TaxTrans(1).Revenue.InterestPd = 0
      TaxTrans(1).Revenue.PenaltyPd = 0
      TaxTrans(1).Revenue.CollectionPd = 0
      TaxTrans(1).Revenue.Future1Pd = 0
      TaxTrans(1).Revenue.Future2Pd = 0
      TaxTrans(1).Description = "Tax Bill #" + STR$(TaxBill(1).BillNumber)
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).CustomerRec = TaxBill(1).CustRec
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = 0
      TaxTrans(1).Padding = ""
      
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
      PUT #TransFile, NextRecord&, TaxTrans(1)
      
      'Update the Customer Pointers Now
      GET #CustFile, TaxBill(1).CustRec, TaxCust(1)
      
      IF TaxCust(1).LastTrans = 0 THEN
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxBill(1).CustRec, TaxCust(1)
      ELSE
        Previous& = TaxCust(1).LastTrans
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxBill(1).CustRec, TaxCust(1)
        
        GET #TransFile, NextRecord&, TaxTrans(1)
        TaxTrans(1).LastTrans = Previous&
        PUT #TransFile, NextRecord&, TaxTrans(1)
      END IF
      
      'Now Update the Property Records with the Tax Year to prevent duplicate billing per year
      RealRecord& = TaxBill(1).RealPropRecord
      WHILE RealRecord& <> 0
        GET RealFile, RealRecord&, RealRec(1)
        RealRec(1).LastYrPrinted = TaxBill(1).TaxYear
        PUT RealFile, RealRecord&, RealRec(1)
        RealRecord& = RealRec(1).NextRec
      WEND
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  
  CLOSE
  
  DisplayTaxScrn "UPDATEOK"
  
  WaitForAction
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
  S$ = "DEL TAXRBILL.DAT"
  SHELL S$
  
PostExit:
  EXIT SUB

ShowPostError:
  ClearBack
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "THERE ARE NO TAX BILLS TO POST!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RETURN
END SUB

SUB TAXPrebilling
  
  SHARED Choice$()
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM TBillRec(1) AS TaxBillType
  REDIM BillInfo(1) AS VARETaxBillInfoType

  DiscoveryFlag$ = "N"
  
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetUpRec
  END IF
  CLOSE TxFile

  'Check For New Market (and others) Which Bills Twice Per Year

  IF INSTR(TaxSetUpRec.Name, "DUMFRIES") > 0 THEN
    NewMarketFlag = True
    DumFriesFlag = True
  END IF

  IF INSTR(TaxSetUpRec.Name, "NEW MARKET") > 0 THEN
    NewMarketFlag = True
  END IF

  IF INSTR(TaxSetUpRec.Name, "STEPHENS CITY") > 0 THEN
    NewMarketFlag = True
  END IF
  
  REDIM Fmt$(1 TO 3)
  
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  RealRecLen = LEN(RealRec(1))
  
  TBillLen = LEN(TBillRec(1))
  BillInfoLen = LEN(BillInfo(1))
  
  REDIM TempScrn(0)
  
  REDIM Choice$(0 TO 4, 0)
  Choice$(0, 0) = "6"
  Choice$(1, 0) = "1) Account Number Order"
  Choice$(2, 0) = "2) Customer Name Order"
  Choice$(3, 0) = "3) Search Name Order"
  Choice$(4, 0) = "4) Social Security Order"
  
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  Fmt$(1) = "$$#########,#"
  Fmt$(2) = "#####"
  Fmt$(3) = "########,#"
  
  MaxLines = 55
  LineCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VRPREBIL"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  ClearBack
  
  DisplayTaxScrn ScrnName$
  
  ShowCursor
  Form$(4, 0) = "0"
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF (Frm(1).KeyCode = -72) OR (Frm(1).KeyCode = -15) THEN
      FOR zz = Frm(1).FldNo TO 1 STEP -1        'loop backwards
        IF Fld(zz).Protected = False THEN
          Frm(1).FldNo = zz
          Action = 1
          EXIT FOR
        END IF
      NEXT
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key    Proceed with look up
      GOSUB CheckPreInfo
      IF PreInfoOK THEN
        OKFlag = True
      END IF
    CASE 27
      ExitFlag = True
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 20
        SELECT CASE Frm(1).MCol
        CASE 29 TO 40           'ESC Cancel button
          PressButton 27, 20, 29, 40
        CASE 42 TO 54           'F10 Proceed Button
          PressButton -68, 20, 42, 54
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  IF ExitFlag THEN
    GOTO ExitPreBill
  END IF
  
  ClearBack
  
  IF Order$ = "2" THEN
    UsingIdx = True
    MakeCustIndexRE 1
  ELSEIF Order$ = "3" THEN
    UsingIdx = True
    MakeCustIndexRE 2
  ELSEIF Order$ = "4" THEN
    UsingIdx = True
    MakeCustIndexRE 3
  END IF

  IF UsingIdx THEN
    NumIdxRecs = FileSize&("TAXTEMP.IDX") / 2
    REDIM IdxBuff(1 TO NumIdxRecs) AS INTEGER
    FGetAH "TAXTEMP.IDX", IdxBuff(1), 2, NumIdxRecs
  END IF
  
  IF NewMarketFlag THEN
     LOCATE 8, 15: PRINT "      VIRGINIA : SPECIAL DISCOVERY ROUTINE"
     LOCATE 10, 20: PRINT "Print Supplemental Bills Now (Y/N) .. "; :
     INPUT dd$
     dd$ = UCASE$(dd$)
     IF dd$ = "Y" THEN DiscoveryFlag$ = "Y" ELSE DiscoveryFlag$ = "N"
   END IF
  
  ShowProcessingScrn "Calculating Tax Prebilling"
  
  IF Exist%(RETaxBillFile) THEN
    KILL RETaxBillFile          'kill any old work file
  END IF
  
  TBillFile = FREEFILE
  OPEN RETaxBillFile FOR RANDOM SHARED AS TBillFile LEN = TBillLen
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfCRecs& = LOF(TaxFile) \ TaxCustLen
  
  'Prebill calc start *******************************
  FOR Cnt& = 1 TO NumOfCRecs&
    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF
    REDIM TBillRec(1) AS TaxBillType            'make a new empty record each time
    GET #TaxFile, CustAcct&, TaxCust(1)         'get cust record

    IF (TaxCust(1).Deleted <> 0) OR (TaxCust(1).TaxExempt = "Y") THEN

    ELSE
      LateFlag$ = "N"           'Initialize Late Flag
      PastFlagSet = 0           'Initialize Past Balance Flag
      IF TaxCust(1).FirstPropRec <= 0 THEN
        NoProp = 1
        GOSUB SetCustInfo
        GOSUB WriteIt2Disk
        GOTO PreBillSkip
      END IF
      
      NoProp = 0
      GOSUB SetCustInfo
      GOSUB GetRealInfo
    END IF
    ShowPctComp Cnt&, NumOfCRecs&
PreBillSkip:
  NEXT
  CLOSE
  'Prebill Calc END   *******************************
  'Prebill Print Start #############
  TotalPers# = 0
  TotalReal# = 0
  TotalEx# = 0
  NumBills& = 0
  
  ShowProcessingScrn "Creating Tax Prebilling Report"
  RptFile = FREEFILE
  OPEN RETaxPreRptFile FOR OUTPUT AS #RptFile
  
  TBillFile = FREEFILE
  OPEN RETaxBillFile FOR RANDOM SHARED AS TBillFile LEN = TBillLen
  NumOfRec& = LOF(TBillFile) / TBillLen
  GOSUB PreBillHeading
  FOR Cnt& = 1 TO NumOfRec&
    GET #TBillFile, Cnt&, TBillRec(1)
    IF Discovery$ = "Y" AND TBillRec(1).BillNumber = -1 THEN
      
    ELSE
      PRINT #RptFile, USING Fmt$(2); TBillRec(1).CustRec;
      PRINT #RptFile, TAB(8); LEFT$(TBillRec(1).CustName, 30);
      IF DumFriesFlag THEN
        IF TBillRec(1).CarShore > 0 THEN
          PRINT #RptFile, TAB(45); USING "####.##"; TBillRec(1).CarShore;
        END IF
      END IF

      IF TBillRec(1).PriorYrBalance > 0 THEN
        PRINT #RptFile, TAB(53); USING "#####.##"; TBillRec(1).PriorYrBalance;
      END IF

      IF TBillRec(1).BillNumber = -1 THEN
        PRINT #RptFile, TAB(62); "N/A";
      ELSE
        PRINT #RptFile, TAB(62); USING Fmt$(2); TBillRec(1).BillNumber;
      END IF
      
      PRINT #RptFile, TAB(68); USING "$$#####,#.##"; TBillRec(1).TotalBillDue
      PRINT #RptFile, USING Fmt$(3); TBillRec(1).RealValue - TBillRec(1).PersValue;
      PRINT #RptFile, USING Fmt$(3); TAB(15); TBillRec(1).PersValue; TAB(30); TBillRec(1).ExptValue;

      TValue# = Round#(TBillRec(1).RealValue - TBillRec(1).ExptValue)
      IF TValue# < 0 THEN TValue# = 0

      PRINT #RptFile, TAB(45); USING "########,#"; TValue#;

      IF TBillRec(1).LateTaxDue > 0 THEN
        PRINT #RptFile, TAB(58); "Late="; USING "#####.##"; TBillRec(1).LateTaxDue
      ELSE
        PRINT #RptFile, ""
      END IF

      PRINT #RptFile, STRING$(79, "-")
      LineCnt = LineCnt + 3

      TotalReal# = Round#(TotalReal# + TBillRec(1).RealValue - TBillRec(1).PersValue)
      TotalBldg# = Round#(TotalBldg# + TBillRec(1).PersValue)
      TotalEx# = Round#(TotalEx# + TBillRec(1).ExptValue)

      IF DumFriesFlag THEN
        TotalSFee# = Round#(TotalSFee# + TBillRec(1).CarShore)
      END IF

      TotalBills# = Round#(TotalBills# + TBillRec(1).TotalBillDue)
      TotalLate# = Round#(TotalLate# + TBillRec(1).LateTaxDue)
      TotalPast# = Round#(TotalPast# + TBillRec(1).PriorYrBalance)

      'IF TBillRec(1).PriorYrBalance < 0 THEN STOP

      IF TBillRec(1).TotalBillDue > 0 THEN
        NumBills& = NumBills& + 1
      END IF
      
      IF LineCnt >= MaxLines THEN
        PRINT #RptFile, CHR$(12);
        GOSUB PreBillHeading
      END IF
    END IF      'Test for Discovery Bills
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  
  PRINT #RptFile, Dash$
  PRINT #RptFile, "Number of Bills to Process: "; USING "#####,#"; NumBills&
  PRINT #RptFile, "      Total Real Valuation: "; USING Fmt$(1); TotalReal#
  PRINT #RptFile, "      Total Bldg Valuation: "; USING Fmt$(1); TotalBldg#
  PRINT #RptFile, "          Total Exemptions: "; USING Fmt$(1); TotalEx#
  PRINT #RptFile, "     Net Valuation to Bill: "; USING Fmt$(1); TotalReal# + TotalBldg# - TotalEx#

  IF DumFriesFlag THEN
    PRINT #RptFile, "   Total Storm Fee to Bill: "; USING "$$########,#.##"; TotalSFee#
  END IF

  PRINT #RptFile, "      Total Amount to Bill: "; USING "$$########,#.##"; TotalBills#
  PRINT #RptFile, "    Total Past Amt to Bill: "; USING "$$########,#.##"; TotalPast#
  PRINT #RptFile, "Grand Total Amount to Bill: "; USING "$$########,#.##"; Round#(TotalPast# + TotalBills#)
  PRINT #RptFile, ""
  PRINT #RptFile, "         Total Late Amount: "; USING "$$########,#.##"; TotalLate#

  IF NewMarketFlag THEN
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, "Note: All Taxes are Billed Using the 1/2 Year Billing"
  END IF
  PRINT #RptFile, CHR$(12);
  CLOSE
  
  ERASE TaxCust, RealRec, TBillRec, BillInfo
  ERASE Fmt$, Frm, Form$, Fld, TempScrn, IdxBuff
  
  PrintRptFile "Tax Prebilling Report.", RETaxPreRptFile, 1, RetCode%, 0
  KILL RETaxPreRptFile
  'PreBill Print END   #############
  
ExitPreBill:
  
  EXIT SUB
  
PreBillHeading:
  PAGE = PAGE + 1
  PRINT #RptFile, TAB(20); "Property Tax Billing : Pre-Billing Register"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; PAGE
  IF DumFriesFlag THEN
    PRINT #RptFile, "Acct #"; TAB(8); "Customer Name"; TAB(45); "  Stm Fee"
  ELSE
    PRINT #RptFile, "Acct #"; TAB(8); "Customer Name"
  END IF
  PRINT #RptFile, "Real Value"; TAB(15); "Bldg Value"; TAB(30); "Discnt Value"; TAB(45); "Net Valuation"; TAB(62); "Bill #"; TAB(73); "Tax Due"
  PRINT #RptFile, Dash$
  LineCnt = 6
RETURN
  
GetRealInfo:

  IF NewMarketFlag THEN
    GOSUB GetRealInfoNewMarket
    RETURN
  END IF

  ThisProp& = TaxCust(1).FirstPropRec
  IF ThisProp& > 0 THEN
    GET #RealFile, ThisProp&, RealRec(1)

    IF RealRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"

    IF RealRec(1).LastYrPrinted <> WhatYear THEN
    IF RealRec(1).LATELIST = "Y" THEN LateFlag$ = "Y"
      RealValue# = Round#(RealRec(1).PROPVALU + RealRec(1).EXMPSENI)
      RealExmp# = Round#(RealRec(1).EXMPOTHR)
      RealCalcVal# = Round#((RealValue# - RealExmp#) / 100)
      RealTaxDue# = Round#(RealCalcVal# * REALRATE#)
      TBillRec(1).ExptValue = Round#(TBillRec(1).ExptValue + RealExmp#)
      TBillRec(1).RealValue = RealValue#
      TBillRec(1).PersValue = RealRec(1).EXMPSENI               'Holds the Bldg Value
      TBillRec(1).RealTaxDue = RealTaxDue#
      TBillRec(1).RealPropRecord = ThisProp&
      TBillRec(1).RealTaxRate = REALRATE#
      TBillRec(1).RDesc1 = RealRec(1).PROPNOT1
      TBillRec(1).RDesc2 = RealRec(1).PROPNOT2
      TBillRec(1).RDesc3 = RealRec(1).PROPNOT3

      'IF LEN(QPTrim$(RealRec(1).PROPNOT3)) > 0 THEN
      '  STOP
      'END IF

      TBillRec(1).REALPIN = RealRec(1).REALPIN
      IF LateFlag$ = "Y" THEN
        LateAmt# = Round#((RealTaxDue#) * (LATELIST# / 100))
      ELSE
        LateAmt# = 0
      END IF
      TBillRec(1).LateTaxDue = Round#(TBillRec(1).LateTaxDue + LateAmt#)
      TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + RealTaxDue# + LateAmt#)
      
      '****** put the first rec to disk
      GOSUB WriteIt2Disk
      END IF
      NextRealRec& = RealRec(1).NextRec

    IF NextRealRec& > 0 THEN
      DO
        REDIM TBillRec(1) AS TaxBillType        'make a new empty record each time
        GOSUB SetCustInfo
        GET #RealFile, NextRealRec&, RealRec(1)
        IF RealRec(1).LastYrPrinted = WhatYear THEN Discovery$ = "Y"
        IF RealRec(1).LastYrPrinted <> WhatYear THEN
          IF RealRec(1).LATELIST = "Y" THEN LateFlag$ = "Y"
          RealValue# = Round#(RealRec(1).PROPVALU + RealRec(1).EXMPSENI)
          RealExmp# = Round#(RealRec(1).EXMPOTHR)
          RealCalcVal# = Round#((RealValue# - RealExmp#) / 100)
          RealTaxDue# = Round#(RealCalcVal# * REALRATE#)
          TBillRec(1).ExptValue = RealExmp#
          TBillRec(1).RealValue = RealValue#
          TBillRec(1).PersValue = RealRec(1).EXMPSENI           'Holds the Bldg Value
          TBillRec(1).RealTaxDue = RealTaxDue#
          TBillRec(1).RealPropRecord = NextRealRec&
          TBillRec(1).RealTaxRate = REALRATE#
          TBillRec(1).RDesc1 = RealRec(1).PROPNOT1
          TBillRec(1).RDesc2 = RealRec(1).PROPNOT2
          TBillRec(1).REALPIN = RealRec(1).REALPIN
          IF LateFlag$ = "Y" THEN
            LateAmt# = Round#(RealTaxDue# * (LATELIST# / 100))
            TBillRec(1).LateTaxDue = LateAmt#
          ELSE
            LateAmt# = 0
          END IF
          
          TBillRec(1).TotalBillDue = Round#(RealTaxDue# + LateAmt#)
          
          GOSUB WriteIt2Disk
          END IF
        NextRealRec& = RealRec(1).NextRec
      LOOP WHILE NextRealRec& > 0
    END IF
  ELSE
    GOSUB WriteIt2Disk
  END IF
  
  RETURN

GetRealInfoNewMarket:
  ThisProp& = TaxCust(1).FirstPropRec
  IF ThisProp& > 0 THEN
    GET #RealFile, ThisProp&, RealRec(1)
      IF RealRec(1).LATELIST = "Y" THEN LateFlag$ = "Y"
     'Special Routine for New Market Supplemental Bills
      IF DiscoveryFlag$ = "Y" AND NewMarketFlag AND RealRec(1).PROPDISC = "N" THEN
        GOTO skip2here:
      END IF

      RealValue# = Round#(RealRec(1).PROPVALU + RealRec(1).EXMPSENI)
      RealExmp# = Round#(RealRec(1).EXMPOTHR)
      RealCalcVal# = Round#((RealValue# - RealExmp#) / 100)
      RealTaxDue# = Round#(RealCalcVal# * REALRATE#)

      IF NewMarketFlag THEN
        RealTaxDue# = Round#(RealTaxDue# / 2)
      END IF

      ThisSFee# = 0 'must clear for each recs
      IF DumFriesFlag THEN
        IF RealRec(1).StormFee > 0 THEN
          ThisSFee# = Round#(0# + RealRec(1).StormFee)
          ThisSFee# = Round#(ThisSFee# / 2)
          RealStormFee# = Round#(RealStormFee# + ThisSFee#)
          RealTaxDue# = Round#(RealTaxDue# + ThisSFee#)
        END IF
      END IF

      'IF ThisSFee# > 0 THEN
      '  RealRec(1).StormFee = ThisSFee#
      'END IF

      TBillRec(1).ExptValue = Round#(TBillRec(1).ExptValue + RealExmp#)
      TBillRec(1).RealValue = RealValue#
      TBillRec(1).PersValue = RealRec(1).EXMPSENI               'Holds the Bldg Value
      TBillRec(1).RealTaxDue = RealTaxDue#
      TBillRec(1).RealPropRecord = ThisProp&
      TBillRec(1).RealTaxRate = REALRATE#
      TBillRec(1).RDesc1 = RealRec(1).PROPNOT1
      TBillRec(1).RDesc2 = RealRec(1).PROPNOT2

      TBillRec(1).RDesc3 = RealRec(1).PROPNOT3

    '  IF LEN(QPTrim$(RealRec(1).PROPNOT3)) > 0 THEN
    '    STOP
    '  END IF

      TBillRec(1).REALPIN = RealRec(1).REALPIN

      TBillRec(1).CarShore = ThisSFee#

      IF LateFlag$ = "Y" THEN
        LateAmt# = Round#((RealTaxDue#) * (LATELIST# / 100))
      ELSE
        LateAmt# = 0
      END IF
      TBillRec(1).LateTaxDue = Round#(TBillRec(1).LateTaxDue + LateAmt#)
      TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + RealTaxDue# + LateAmt#)
'LOOK FOR ME HERE DALE
      'TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + ThisSFee#)

      'if This

      '****** put the first rec to disk
      GOSUB WriteIt2Disk


skip2here:
    NextRealRec& = RealRec(1).NextRec
    IF NextRealRec& > 0 THEN
      DO
        REDIM TBillRec(1) AS TaxBillType        'make a new empty record each time
        GOSUB SetCustInfo
        GET #RealFile, NextRealRec&, RealRec(1)
          IF RealRec(1).LATELIST = "Y" THEN LateFlag$ = "Y"
         'New Market Special Routine
          IF DiscoveryFlag$ = "Y" AND NewMarketFlag AND RealRec(1).PROPDISC = "N" THEN
            GOTO Skip2Here1:
          END IF
          RealValue# = Round#(RealRec(1).PROPVALU + RealRec(1).EXMPSENI)
          RealExmp# = Round#(RealRec(1).EXMPOTHR)
          RealCalcVal# = Round#((RealValue# - RealExmp#) / 100)
          RealTaxDue# = Round#(RealCalcVal# * REALRATE#)
          IF NewMarketFlag THEN
            RealTaxDue# = Round#(RealTaxDue# / 2)
          END IF

          TBillRec(1).ExptValue = RealExmp#
          TBillRec(1).RealValue = RealValue#
          TBillRec(1).PersValue = RealRec(1).EXMPSENI           'Holds the Bldg Value
          TBillRec(1).RealTaxDue = RealTaxDue#
          TBillRec(1).RealPropRecord = NextRealRec&
          TBillRec(1).RealTaxRate = REALRATE#
          TBillRec(1).RDesc1 = RealRec(1).PROPNOT1
          TBillRec(1).RDesc2 = RealRec(1).PROPNOT2
          TBillRec(1).REALPIN = RealRec(1).REALPIN
          IF LateFlag$ = "Y" THEN
            LateAmt# = Round#(RealTaxDue# * (LATELIST# / 100))
            TBillRec(1).LateTaxDue = LateAmt#
          ELSE
            LateAmt# = 0
          END IF

          TBillRec(1).TotalBillDue = Round#(RealTaxDue# + LateAmt#)

          GOSUB WriteIt2Disk
Skip2Here1:
        NextRealRec& = RealRec(1).NextRec
      LOOP WHILE NextRealRec& > 0
    END IF
  ELSE
    GOSUB WriteIt2Disk
  END IF

  RETURN
  
  
WriteIt2Disk:
  'write the info out to disk here.
  TBillRec(1).BillPrinted = False
  IF TBillRec(1).TotalBillDue > 0 THEN
    TBillRec(1).BillNumber = BILLNUM&
  ELSE
    TBillRec(1).BillNumber = -1
  END IF
  
  PUT #TBillFile, , TBillRec(1)
  IF TBillRec(1).TotalBillDue > 0 THEN
    BILLNUM& = BILLNUM& + 1
  END IF
  RETURN
  
SetCustInfo:
  TBillRec(1).CustRec = CustAcct&               'cust acct rec
  CustName$ = QPTrim$(TaxCust(1).FName) + " " + QPTrim$(TaxCust(1).LNAME)
  TBillRec(1).CustName = CustName$
  TBillRec(1).CustAdd1 = TaxCust(1).Addr1
  TBillRec(1).CustAdd2 = TaxCust(1).Addr2
  CitySt$ = QPTrim$(TaxCust(1).City) + " " + TaxCust(1).State
  TBillRec(1).CustAdd3 = CitySt$
  TBillRec(1).CustZip = TaxCust(1).Zip
  TBillRec(1).CustPin = TaxCust(1).PIN
  TBillRec(1).TaxYear = WhatYear
  
  'Set Prior Balance if any
  
  GOSUB GetPastBalance
  IF ABalance# <> 0 THEN
    IF PastFlagSet = 0 THEN
      TBillRec(1).PriorYrBalance = ABalance#
    END IF
    PastFlagSet = 1
  END IF
  RETURN
  
CheckPreInfo:
  WhatYear = Value#(Form$(1, 0), ECode)
  BILLNUM& = Value#(Form$(2, 0), ECode)
  REALRATE# = Value#(Form$(3, 0), ECode)
  PERSRATE# = 0 'Not Used
  LATELIST# = Value#(Form$(5, 0), ECode)
  Order$ = LEFT$(Form$(6, 0), 1)
  
  IF WhatYear < 1997 OR WhatYear > 2020 THEN
    GOSUB ShowInfoErr
    Frm(1).FldNo = 1
    GOTO InfoErrorExit
  END IF
  
  IF BILLNUM& = 0 THEN
    GOSUB ShowInfoErr
    Frm(1).FldNo = 2
    GOTO InfoErrorExit
  END IF
  
  IF LEN(QPTrim$(Order$)) = 0 THEN
    GOSUB ShowInfoErr
    Frm(1).FldNo = 6
    GOTO InfoErrorExit
  END IF
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(BillInfo(1)), VARPTR(BillInfo(1)), BillInfoLen, 0
  FPutAH "TAXBINFO.DAT", BillInfo(1), BillInfoLen, 1
  
  PreInfoOK = True
  
InfoErrorExit:
  
  RETURN
  
ShowInfoErr:
  SaveScrn TempScrn()
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "Invalid Prebilling Information!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn TempScrn()
  RETURN
  
GetPastBalance:
  
  Balance# = 0
  ABalance# = 0
  IF TaxCust(1).LastTrans > 0 THEN
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
    TransRecord& = TaxCust(1).LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      
      IF TaxTrans(1).TranType = 1 THEN
        Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
        Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection
        Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
        Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
        Balance# = Round#(Balance#)
        ABalance# = ABalance# + Balance#
      END IF
      
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  END IF
  RETURN
  
END SUB

