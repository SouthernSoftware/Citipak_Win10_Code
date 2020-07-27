DEFINT A-Z
DECLARE SUB smallpause ()
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB IntLookUp (RecNo&, Text$, CLSFlag%)
DECLARE SUB Search4IntCust (Search$, RecNo&, CLSFlag%)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB UpDateTicklerFile ()
DECLARE SUB PostTransactions ()
DECLARE SUB CalcInterest ()
DECLARE SUB EditTrans ()
DECLARE SUB PrintTrans ()
DECLARE SUB PostTaxBills ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB MakeCustIndex (IdxType)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION IsCustDeleted% (AcctNum&, TaxType%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB UpDateTicklerFile ()
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxBill.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXBINFO.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE

  DIM SHARED TaxSetUpRec AS TaxMasterType

  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Calculate Interest"
  MChoice$(2) = "Edit Interest Transactions"
  MChoice$(3) = "Print Interest Edit Report    "
  MChoice$(4) = "Post Interest to Accounts"
  MChoice$(5) = "Exit to OS"
  
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
  
  IF Exist("TAXFROMM.DAT") THEN
    'KillFile "TAXFROMM.DAT"
    MainFlag = True
  END IF
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing 'PP Interest' Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      CalcInterest
    CASE 2
      EditTrans
    CASE 3
      PrintTrans
    CASE 4
      PostTransactions
    CASE 5
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
  
  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  END IF
  
  RUN "TaxPPBIL"

SUB CalcInterest
  
  SHARED Choice$()
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM IntTrans(1) AS InterestRecType
  
  DIM Year!(99), IAmt#(99)

  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, TaxSetUpRec
    'Check For New Market Which Bills Twice Per Year
    NewMarketFlag = INSTR(TaxSetUpRec.Name, "NEW MARKET")
  END IF
  CLOSE TxFile
  
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  TIntLen = LEN(IntTrans(1))
  
  REDIM Fmt$(3)
  
  REDIM TempScrn(0)
  
  REDIM Choice$(0 TO 2, 0)
  Choice$(0, 0) = "6"
  Choice$(1, 0) = "1) Account Number Order"
  Choice$(2, 0) = "2) Customer Name Order"
  
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  
  Fmt$(1) = "$$########,#"
  Fmt$(2) = "#####"
  Fmt$(3) = "#######,#"
  
  MaxLines = 54
  LineCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "CALCINTV"
  
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
  IF NewMarketFlag = 0 THEN
   Form$(3, 0) = ""
   Form$(5, 0) = ""
   Fld(3).Protected = True
   Fld(5).Protected = True
  END IF

  
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
    GOTO ExitCalcInt
  END IF
  
  ClearBack
  
  IF Order$ = "2" THEN
    UsingIdx = True
    MakeCustIndex 1
  END IF
  IF UsingIdx THEN
    NumIdxRecs = FileSize&("TAXTEMP.IDX") / 2
    REDIM IdxBuff(1 TO NumIdxRecs) AS INTEGER
    FGetAH "TAXTEMP.IDX", IdxBuff(1), 2, NumIdxRecs
  END IF
  
  ShowProcessingScrn "Calculating Tax CalcInting"
  
  IF Exist%(PPTaxIntFile) THEN
    KILL PPTaxIntFile             'kill any old work file
  END IF
  
  TIntFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM SHARED AS TIntFile LEN = TIntLen
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
  
  NumOfCRecs& = LOF(TaxFile) \ TaxCustLen
  
  'CalcInt calc start *******************************
  FOR Cnt& = 1 TO NumOfCRecs&
    IF UsingIdx THEN
      CustAcct& = IdxBuff(Cnt&)
    ELSE
      CustAcct& = Cnt&
    END IF
    
    GET #TaxFile, CustAcct&, TaxCust(1)         'get cust record
    IF TaxCust(1).Interest <> "Y" THEN GOTO CalcIntSkip

 

    TransRecord& = TaxCust(1).LastTrans
    
    WHILE TransRecord& > 0
      GET TransFile, TransRecord&, TaxTrans(1)
      CurYearFlag = False
      PastYearFlag = False
      IF TaxTrans(1).TranType = 1 THEN
        IF NewMarketFlag > 0 THEN                       'Figure on Principle Only For New Market
         Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
         Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
        ELSE
         Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5 + TaxTrans(1).Revenue.Penalty
         Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd + TaxTrans(1).Revenue.PenaltyPd)
        END IF
        Balance# = Round#(Balance#)
        IF Balance# > 0 THEN
          IF TaxTrans(1).TAXYEAR = WhatYear THEN
            CurYearFlag = True
           'Check For New Market Special Condition
            IF NewMarketFlag > 0 AND TaxTrans(1).TransDate <> CurBillDate THEN CurYearFlag = False
          END IF
          


          IF TaxTrans(1).TAXYEAR <> WhatYear THEN
           PastYearFlag = True
           'Check For New Market Special Condition
            IF NewMarketFlag > 0 AND TaxTrans(1).TransDate <> PastBillDate THEN PastYearFlag = False
           END IF


          'Get Bill Number
          BillNumber$ = TaxTrans(1).Description
          sb = INSTR(BillNumber$, "Bill #")
          IF sb > 0 THEN
            BillNumber$ = MID$(TaxTrans(1).Description, sb + 6, 10)
          END IF
          
          IF CurYearFlag THEN
            IntAmount# = Round#(Balance# * (CurRate# / 100))
          END IF
          IF PastYearFlag THEN
            IntAmount# = Round#(Balance# * (PastRate# / 100))
          END IF

          IF Round#(IntAmount#) > 0 THEN
            Nme$ = RTRIM$(TaxCust(1).FNAME) + " " + RTRIM$(TaxCust(1).LName)
            Nme$ = LTRIM$(Nme$)
            IntTrans(1).CustRec = CustAcct&
            IntTrans(1).CustName = Nme$
            IntTrans(1).TAXYEAR = TaxTrans(1).TAXYEAR
            IntTrans(1).Amount = Round#(IntAmount#)
            IntTrans(1).BillNumber = BillNumber$
            IntTrans(1).BillRec = TransRecord&
            IntTrans(1).CurYear = WhatYear
            IntRecord& = IntRecord& + 1
            PUT TIntFile, IntRecord&, IntTrans(1)
            IntAmount# = 0
         END IF
        END IF  'Check For Balance
      END IF    'Check for Tax Bill
      TransRecord& = TaxTrans(1).LastTrans
    WEND
    
CalcIntSkip:
    ShowPctComp Cnt&, NumOfCRecs&
  NEXT
ExitWend:
  
  'CalcInt Calc END   *******************************
  CLOSE TaxFile
  CLOSE TransFile
  'CalcInt Print Start #############
  
  ShowProcessingScrn "Creating Interest Report"
  RptFile = FREEFILE
  OPEN InterestReportFile FOR OUTPUT AS #RptFile
  
  GOSUB CalcIntHeading
  
  FOR Cnt& = 1 TO IntRecord&
    GET #TIntFile, Cnt&, IntTrans(1)
    PRINT #RptFile, USING "######"; IntTrans(1).CustRec;
    PRINT #RptFile, TAB(8); IntTrans(1).CustName;
    PRINT #RptFile, TAB(50); IntTrans(1).TAXYEAR;
    PRINT #RptFile, TAB(60); QPTrim$(IntTrans(1).BillNumber);
    PRINT #RptFile, TAB(70); USING "$####,#.##"; IntTrans(1).Amount
    LineCnt = LineCnt + 1
    TotalInterest# = TotalInterest# + IntTrans(1).Amount
    IF IntTrans(1).TAXYEAR = WhatYear THEN
      CurInterest# = CurInterest# + IntTrans(1).Amount
    ELSE
      PastInterest# = PastInterest# + IntTrans(1).Amount
    END IF
    
    'Now Book to Specific Year
    IF YearCount = 0 THEN
      YearCount = 1
      Year!(YearCount) = IntTrans(1).TAXYEAR
      IAmt#(YearCount) = IntTrans(1).Amount
    ELSE
      IFlag = 0
      FOR SCnt = 1 TO YearCount
        IF Year!(SCnt) = IntTrans(1).TAXYEAR THEN
          IAmt#(SCnt) = IAmt#(SCnt) + IntTrans(1).Amount
          IFlag = 1
        END IF
      NEXT SCnt
      IF IFlag = 0 THEN
        YearCount = YearCount + 1
        Year!(YearCount) = IntTrans(1).TAXYEAR
        IAmt#(YearCount) = IntTrans(1).Amount
      END IF
    END IF
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptFile, CHR$(12);
      GOSUB CalcIntHeading
    END IF
    
    ShowPctComp Cnt&, IntRecord&
  NEXT
  
  PRINT #RptFile, Dash$
  PRINT #RptFile, "Total Interest Charged .. "; USING "$$######,#.##"; TotalInterest#
  PRINT #RptFile, "Total Current Interest .. "; USING "$$######,#.##"; CurInterest#
  PRINT #RptFile, "Total Past Interest ..... "; USING "$$######,#.##"; PastInterest#
  PRINT #RptFile, CHR$(12);
  Page = Page + 1
  PRINT #RptFile, "Interest Breakdown by Tax Year"; TAB(65); "Page #"; Page
  PRINT #RptFile, ""
  PRINT #RptFile, " Year"; TAB(10); "Interest Calc."
  PRINT #RptFile, "--------------------------------"
  FOR LN = 1 TO YearCount
    PRINT #RptFile, Year!(LN); TAB(10); USING "$$######,#.##"; IAmt#(LN)
  NEXT LN
  
  CLOSE
  
  PrintRptFile "Tax Interest Report.", InterestReportFile, 1, RetCode%, 0
  'CalcInt Print END   #############
  
ExitCalcInt:
  
  EXIT SUB
  
CalcIntHeading:
  Page = Page + 1
  PRINT #RptFile, TAB(15); "Property Tax Billing : Interest Calculation Register"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; Page
  PRINT #RptFile, "Current Tax Year is "; WhatYear
  IF NewMarketFlag > 0 THEN
   PRINT #RptFile, "Current Bill Date Charged: "; Num2Date$(CurBillDate)
   PRINT #RptFile, "   Past Bill Date Charged: "; Num2Date$(PastBillDate)
  END IF
  PRINT #RptFile, "Acct #"; TAB(8); "Customer Name"; TAB(50); "Tax Yr"; TAB(60); "Bill #"; TAB(71); "Interest"
  PRINT #RptFile, Dash$
  LineCnt = 6
  RETURN
  
CheckPreInfo1:
  WhatYear = Value#(Form$(1, 0), ECode)
  CurRate# = Value#(Form$(2, 0), ECode)
  CurBillDate = Date2Num(Form$(3, 0))
  PastRate# = Value#(Form$(4, 0), ECode)
  PastBillDate = Date2Num(Form$(5, 0))
  Order$ = LEFT$(Form$(6, 0), 1)
  
  IF WhatYear < 1990 OR WhatYear > 2010 THEN
    GOSUB ShowInfoErr1
    Frm(1).FldNo = 1
    GOTO InfoErrorExit1
  END IF
  
  IF LEN(QPTrim$(Order$)) = 0 THEN
    GOSUB ShowInfoErr1
    Frm(1).FldNo = 4
    GOTO InfoErrorExit1
  END IF
  
  PreInfoOK = True
  
InfoErrorExit1:
  
  RETURN
  
ShowInfoErr1:
  SaveScrn TempScrn()
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "Invalid CalcInting Information!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn TempScrn()
RETURN
  
GetPastBalance1:
  Balance# = 0
  ABalance# = 0
  IF TaxCust(1).LastTrans > 0 THEN
    TransRecord& = TaxCust(1).LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN
      END IF
      ABalance# = ABalance# + Balance#
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
  END IF
  RETURN
  
END SUB

SUB EditTrans
  
  RecDisp$ = SPACE$(14)
  SHARED Choice$()
  
  REDIM EdScrn(0)
  
  REDIM TaxIntRec(1 TO 1) AS InterestRecType
  TaxIntRecLen = LEN(TaxIntRec(1))
  
  TaxFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM AS TaxFile LEN = TaxIntRecLen      'open data file
  NumIntRec& = LOF(TaxFile) / TaxIntRecLen
  CLOSE TaxFile
  
  CursorOff
  ClearBack
  
  IF NumIntRec& < 1 THEN
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "No Interest Transaction to Edit!", 10, 25, -1
    QPrintRC "Press any key to continue.", 12, 28, -1
    WaitForAction
    GOTO EditIntExit
  ELSE
    GOSUB CheckAnyNotDeleted
    IF NoneToEdit THEN
      DisplayTaxScrn "ERRSCRN1"
      QPrintRC "No Active Interest Transactions!", 10, 25, -1
      QPrintRC "Press any key to continue.", 12, 28, -1
      WaitForAction
      GOTO EditIntExit
    END IF
  END IF
  
  LibName$ = "TAX"
  ScrnName$ = "TAXINTED"
  
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
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  DisplayTaxScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      Forward = True
      IntRecNo& = 1
      GOSUB GetIntRec
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key ', EnterKey    'F10   Save
      GOSUB UpDateIntRec

    CASE F3Key
      Ok = MsgBox%("TAX", "DELINTRN")
      SELECT CASE Ok
      CASE 2    'if button 2 was pressed
        GOSUB DelIntTrans
        GOSUB GetIntRec
        'FirstTime = True
      CASE ELSE
      END SELECT
      Action = 1
    CASE F5KEY  'F5 Look-up
      SaveScrn EdScrn()
      MPaintBox 6, 4, 18, 76, 8
      IntLookUp FindRecNo&, "Interest Edit Customer", CLSFlag
      RestScrn EdScrn()
      IF FindRecNo& <> IntRecNo& AND FindRecNo& > 0 THEN
        IntRecNo& = FindRecNo&
        GOSUB GetIntRec
      END IF
      Action = 1
      
    CASE PgUpKey                'previous rec
      Forward = False
      IF IntRecNo& > 1 THEN
        IntRecNo& = IntRecNo& - 1
        GOSUB GetIntRec
      END IF
      
    CASE PgDnKey                'next rec
      Forward = True
      IF IntRecNo& < NumIntRec& THEN
        IntRecNo& = IntRecNo& + 1
        GOSUB GetIntRec
      END IF
    CASE EscKey
      RecNo& = 0
      ExitFlag = True
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 17
        SELECT CASE Frm(1).MCol
        CASE 8 TO 17            'ESC Cancel button
          PressButton EscKey, 17, 8, 17
        CASE 18 TO 28           'F3 Delete
          PressButton F3Key, 17, 18, 28
        CASE 29 TO 39           'F5 Search
          PressButton F5KEY, 17, 29, 39
        CASE 40 TO 49           'F10 Save Button
          PressButton F10Key, 17, 40, 49
        CASE 50 TO 60           'PgUp Prev Rec
          PressButton PgUpKey, 17, 50, 60
        CASE 61 TO 71           'PgDn Next Rec
          PressButton PgDnKey, 17, 61, 71
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  GOTO EditIntExit
  
GetIntRec:
  DidFirst = False
  DidLast = False
  TaxFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM AS TaxFile LEN = TaxIntRecLen      'open data file
  DO
    GET TaxFile, IntRecNo&, TaxIntRec(1)
    IF TaxIntRec(1).DelFlag <> 0 THEN
      SELECT CASE Forward
      CASE True
        IF IntRecNo& < NumIntRec& THEN
          IntRecNo& = IntRecNo& + 1
        ELSE
          Forward = False
          DidLast = True
        END IF
      CASE False
        IF IntRecNo& > 1 THEN
          IntRecNo& = IntRecNo& - 1
        ELSE
          Forward = True
          DidFirst = True
        END IF
      END SELECT
    ELSE
      EXIT DO
    END IF
    IF DidFirst AND DidLast THEN
      CursorOff
      CLOSE TaxFile
      DisplayTaxScrn "ERRSCRN1"
      QPrintRC "No Active Interest Transactions!", 10, 25, -1
      QPrintRC "Press any key to continue.", 12, 28, -1
      WaitForAction
      GOTO EditIntExit
    END IF
  LOOP
  CLOSE TaxFile
  
  BCopy VARSEG(TaxIntRec(1)), VARPTR(TaxIntRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
  GOSUB DisplayRecCnt
RETURN
  
CheckAnyNotDeleted:
  ShowProcessingScrn "Scanning for Active Transactions"
  NoneToEdit = True
  TaxFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM AS TaxFile LEN = TaxIntRecLen      'open data file
  NumIntRec& = LOF(TaxFile) / TaxIntRecLen
  FOR Cnt& = 1 TO NumIntRec&
    GET TaxFile, Cnt&, TaxIntRec(1)
'IF Cnt& <= 60 THEN
'  TaxIntRec(1).DelFlag = -1
'  PUT TaxFile, Cnt&, TaxIntRec(1)
'END IF
    IF TaxIntRec(1).DelFlag = 0 THEN
      NoneToEdit = False
    END IF
    ShowPctCompL Cnt&, NumIntRec&
  NEXT
  CLOSE TaxFile
  smallpause
RETURN
  
UpDateIntRec:
  SaveScrn EdScrn()
  TaxFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM AS TaxFile LEN = TaxIntRecLen      'open data file
  GET TaxFile, IntRecNo&, TaxIntRec(1)
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TaxIntRec(1)), VARPTR(TaxIntRec(1)), LEN(Form$(0, 0)), 0
  PUT TaxFile, IntRecNo&, TaxIntRec(1)
  CLOSE TaxFile
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn EdScrn()
RETURN
  
DelIntTrans:
  SaveScrn EdScrn()
  TaxFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM AS TaxFile LEN = TaxIntRecLen      'open data file
  GET TaxFile, IntRecNo&, TaxIntRec(1)
  TaxIntRec(1).DelFlag = True
  PUT TaxFile, IntRecNo&, TaxIntRec(1)
  CLOSE TaxFile
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn EdScrn()
RETURN
  
DisplayRecCnt:
  LSET RecDisp$ = QPTrim$(STR$(IntRecNo&)) + " of " + QPTrim$(STR$(NumIntRec&))
  QPrintRC RecDisp$, 8, 56, -1
RETURN
  
EditIntExit:
  
END SUB

SUB IntLookUp (RecNo&, Text$, CLSFlag%)
  
LookUpTop:
  
  REDIM Hlp$(1 TO 2)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Name here."
  
  SName$ = ""
  AcctNum& = 0
  LScrn = 2
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 2) AS STRING * 17
  
  MScrn = 2
  
  LText(1) = " Account Number:"
  LText(2) = "    Search Name:"
  
  LibName$ = "TAX"
  ScrnName$ = "LUPACCT"
  
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
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  Text$ = Text$ + " Look-Up"
  TextLen = LEN(Text$)
  TCol = ((80 - TextLen) \ 2)
  DisplayTaxScrn ScrnName$
  
  QPrintRC Text$, 8, TCol, -1
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      SELECT CASE LScrn
      CASE 1, 2
        LSET Form$(1, 0) = ""
        Fld(1).FType = 2
      END SELECT
      Form$(1, 1) = Hlp$(LScrn)
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13                'F10Key    Proceed with look up
      CursorOff
      SELECT CASE LScrn
      CASE 1    'account lookup
        AcctNum& = QPValL(Form$(1, 0))
        IF AcctNum& < 1 OR AcctNum& > GetTaxCustCnt& THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")
        ELSE
          RecNo& = AcctNum&
          OKFlag = True
        END IF
        Action = 1
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4IntCust SName$, RecNo&, CLSFlag
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      END SELECT
    CASE -65    'F7Key
      IF LScrn < MScrn THEN
        LScrn = LScrn + 1
      ELSE
        LScrn = 1
      END IF
      LSET Form$(1, 0) = ""
      Action = 1
      FirstTime = True
      SaveField 0, Form$(), Fld(), BadField
      GOSUB DisplayLookupText
    CASE 27
      RecNo& = 0
      ExitFlag = True
    END SELECT
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16
        SELECT CASE Frm(1).MCol
        CASE 22 TO 33           'ESC Cancel button
          PressButton 27, 16, 22, 33
        CASE 35 TO 45           'F7 Toggle Choice
          PressButton -65, 16, 35, 45
        CASE 47 TO 59           'F10 Save Button
          PressButton -68, 16, 47, 59
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  RestScrn ScrnArray()
  
  ERASE ScrnArray, ScrnArray2
  ERASE Frm, Form$, Fld, LText, Hlp$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText(LScrn), 12, 15, -1
  RETURN
  
  
END SUB

SUB PostTransactions

  REDIM TempScrn(0)
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM TaxIntRec(1 TO 1) AS InterestRecType
  TaxIntRecLen = LEN(TaxIntRec(1))
  IntRecLen = LEN(TaxIntRec(1))
  
  ClearBack
  ShowProcessingScrn "Verifying Interest Transactions"

  IntFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM AS IntFile LEN = IntRecLen
  NumOfRec& = LOF(IntFile) / IntRecLen
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #IntFile, Cnt&, TaxIntRec(1)
    IF TaxIntRec(1).DelFlag = 0 THEN
      DidSome = DidSome + 1
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  smallpause

  IF DidSome < 1 THEN
    GOSUB ShowPostError
    GOTO PostExit
  END IF
  
  PostFlag = MsgBox%("TAX", "POSTINTR")
  
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
  ShowProcessingScrn "Posting Tax Interest Transactions"
  
  'Open All Needed Files Here

  IntFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM AS IntFile LEN = IntRecLen
  NumOfRec& = LOF(IntFile) / IntRecLen
  
  CustFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM AS CustFile LEN = LEN(TaxCust(1))
  
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #IntFile, Cnt&, TaxIntRec(1)
    IF TaxIntRec(1).DelFlag = 0 THEN
      'Update the Bill transaction first
      'TaxIntRec(1).BillRec
      GET #TransFile, TaxIntRec(1).BillRec, TaxTrans(1) 'get bill trans
      TaxTrans(1).Revenue.Interest = Round#(TaxTrans(1).Revenue.Interest + TaxIntRec(1).Amount)
      PUT #TransFile, TaxIntRec(1).BillRec, TaxTrans(1) 'put it back
      'Now make a new clean transaction
      REDIM TaxTrans(1) AS TaxTransactionType
      TaxTrans(1).TransDate = Date2Num%(DATE$)
      TaxTrans(1).TAXYEAR = TaxIntRec(1).TAXYEAR
      TaxTrans(1).TranType = 4       '4=Interest
      TaxTrans(1).BillType = "P"     'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TaxIntRec(1).Amount  'Total Transaction Amount
      TaxTrans(1).Revenue.Interest = TaxIntRec(1).Amount
      TaxTrans(1).Description = "Tax Int on Bill# " + QPTrim$(TaxIntRec(1).BillNumber)
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).CustomerRec = TaxIntRec(1).CustRec
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = TaxIntRec(1).BillRec
      LSET TaxTrans(1).Padding = ""
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
      PUT #TransFile, NextRecord&, TaxTrans(1)
      'Update the Customer Pointers Now
      GET #CustFile, TaxIntRec(1).CustRec, TaxCust(1)
      
      IF TaxCust(1).LastTrans = 0 THEN
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxIntRec(1).CustRec, TaxCust(1)
      ELSE
        Previous& = TaxCust(1).LastTrans
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxIntRec(1).CustRec, TaxCust(1)
        GET #TransFile, NextRecord&, TaxTrans(1)
        TaxTrans(1).LastTrans = Previous&
        PUT #TransFile, NextRecord&, TaxTrans(1)
      END IF
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  
  UpDateTicklerFile

  ClearBack
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
  KillFile PPTaxIntFile
  
PostExit:
  EXIT SUB
  
  
ShowPostError:
  ClearBack
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "No Active Interest Transactions!", 10, 25, -1
  QPrintRC "Press any key to continue.", 12, 28, -1
  WaitForAction
RETURN
END SUB

SUB PrintTrans
  
  REDIM IntTrans(1) AS InterestRecType
  
  DIM Year!(99), IAmt#(99)
  TIntLen = LEN(IntTrans(1))
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  MaxLines = 54
  LineCnt = 0
  
  TIntFile = FREEFILE
  OPEN PPTaxIntFile FOR RANDOM SHARED AS TIntFile LEN = TIntLen
  IntRecord& = LOF(TIntFile) / TIntLen
  IF IntRecord& = 0 THEN
    EXIT SUB
  END IF
  'Get Current Year
  GET #TIntFile, 1, IntTrans(1)
  WhatYear = IntTrans(1).CurYear
  
  'CalcInt Print Start #############
  
  ClearBack
  ShowProcessingScrn "Creating Interest Report"
  RptFile = FREEFILE
  OPEN InterestReportFile FOR OUTPUT AS #RptFile
  
  GOSUB CalcIntHeading1
  
  FOR Cnt& = 1 TO IntRecord&
    GET #TIntFile, Cnt&, IntTrans(1)
    IF IntTrans(1).DelFlag <> 0 THEN
      GOTO SkipThisIntTrans
    END IF
    PRINT #RptFile, USING "######"; IntTrans(1).CustRec;
    PRINT #RptFile, TAB(8); IntTrans(1).CustName;
    PRINT #RptFile, TAB(50); IntTrans(1).TAXYEAR;
    PRINT #RptFile, TAB(60); QPTrim$(IntTrans(1).BillNumber);
    PRINT #RptFile, TAB(70); USING "$####,#.##"; IntTrans(1).Amount;
    PRINT #RptFile, USING "####"; Cnt&
    LineCnt = LineCnt + 1
    TotalInterest# = TotalInterest# + IntTrans(1).Amount
    IF IntTrans(1).TAXYEAR = WhatYear THEN
      CurInterest# = CurInterest# + IntTrans(1).Amount
    ELSE
      PastInterest# = PastInterest# + IntTrans(1).Amount
    END IF
    
    'Now Book to Specific Year
    IF YearCount = 0 THEN
      YearCount = 1
      Year!(YearCount) = IntTrans(1).TAXYEAR
      IAmt#(YearCount) = IntTrans(1).Amount
    ELSE
      IFlag = 0
      FOR SCnt = 1 TO YearCount
        IF Year!(SCnt) = IntTrans(1).TAXYEAR THEN
          IAmt#(SCnt) = IAmt#(SCnt) + IntTrans(1).Amount
          IFlag = 1
        END IF
      NEXT SCnt
      IF IFlag = 0 THEN
        YearCount = YearCount + 1
        Year!(YearCount) = IntTrans(1).TAXYEAR
        IAmt#(YearCount) = IntTrans(1).Amount
      END IF
    END IF
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptFile, CHR$(12);
      GOSUB CalcIntHeading1
    END IF
SkipThisIntTrans:
    ShowPctComp Cnt&, IntRecord&
  NEXT
  
  PRINT #RptFile, Dash$
  PRINT #RptFile, "Total Interest Charged .. "; USING "$$######,#.##"; TotalInterest#
  PRINT #RptFile, "Total Current Interest .. "; USING "$$######,#.##"; CurInterest#
  PRINT #RptFile, "Total Past Interest ..... "; USING "$$######,#.##"; PastInterest#
  PRINT #RptFile, CHR$(12);
  Page = Page + 1
  PRINT #RptFile, "Interest Breakdown by Tax Year"; TAB(65); "Page #"; Page
  PRINT #RptFile, ""
  PRINT #RptFile, " Year"; TAB(10); "Interest Calc."
  PRINT #RptFile, "--------------------------------"
  FOR LN = 1 TO YearCount
    PRINT #RptFile, Year!(LN); TAB(10); USING "$$######,#.##"; IAmt#(LN)
  NEXT LN
  
  
  CLOSE
  
  PrintRptFile "Tax Interest Report.", InterestReportFile, 1, RetCode%, 0
  'CalcInt Print END   #############
  
  EXIT SUB
  
CalcIntHeading1:
  Page = Page + 1
  PRINT #RptFile, TAB(10); "Va. Personal Property Tax Billing : Interest Calculation Register"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; Page
  PRINT #RptFile, "Current Tax Year is "; WhatYear
  PRINT #RptFile, "Acct #"; TAB(8); "Customer Name"; TAB(50); "Tax Yr"; TAB(60); "Bill #"; TAB(71); "Interest"
  PRINT #RptFile, Dash$
  LineCnt = 5
  RETURN
  
  
END SUB

SUB Search4IntCust (Search$, RecNo&, CLSFlag%)
  
  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(5)
  
  BNum$ = SPACE$(7)
  
  IF LastSEARCH$ <> Search$ THEN
    LastSEARCH$ = Search$
    Choice = 1
  END IF
  
  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
  '091598 Found a bug that caused the lookup to get erroneous records
  '       if this was changed to greater than 32???
  CustBlock = 1
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM ITaxCust(1 TO 1) AS InterestRecType
  
  ITaxCustLen = LEN(ITaxCust(1))
  
  SearchLen = LEN(Search$)
  Match = False
  
  TaxFile = FREEFILE
  
  FOpenS PPTaxIntFile, TaxFile    'open data file
  NumOfCust& = FLof&(TaxFile) / ITaxCustLen
  
  FOR CCnt& = 1 TO NumOfCust&
    FGetRTA TaxFile, ITaxCust(1), CCnt&, ITaxCustLen
    IF INSTR(ITaxCust(1).CustName, Search$) > 0 AND ITaxCust(1).DelFlag = 0 THEN
      GOSUB CustLoadEM2
    END IF
DelSkip3:
    ShowSearchPCT CCnt&, NumOfCust&
  NEXT
  
  FClose TaxFile
  
  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2
  ELSE
    
    'SortT MChoice(1), DCnt, Direction%, LEN(MChoice(1).V), 0, 18
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer Name                  Year      Amt    BillNo  Acct"
    '--Find max menu width
    '--Center Menu within Screen
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      ClearBack
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
LoopRestart:
    LOCATE Row, Col, 0
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  END IF
  
ExitSearch2:
  RestScrn ScrnArray()
  
  ERASE ScrnArray, MChoice, ITaxCust
  
  EXIT SUB
  
CustLoadEM2:
  
  DCnt = DCnt + 1
  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
  LSET MChoice(DCnt).V = LEFT$(QPTrim$(ITaxCust(1).CustName), 30)
  
  MID$(MChoice(DCnt).V, 32) = FUsing$(STR$(ITaxCust(1).TAXYEAR), "####")
  MID$(MChoice(DCnt).V, 38) = FUsing$(STR$(ITaxCust(1).Amount), ",####.##")
  RSET BNum$ = QPTrim$(ITaxCust(1).BillNumber)
  MID$(MChoice(DCnt).V, 48) = BNum$
  RSET Acct$ = QPTrim$(STR$(ITaxCust(1).CustRec))
  MID$(MChoice(DCnt).V, 55, 5) = Acct$
  MID$(MChoice(DCnt).V, 61) = MKL$(CCnt&)
  RETURN
  
  
END SUB

