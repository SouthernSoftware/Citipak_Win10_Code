DEFINT A-Z
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE SUB Search4IntAcct (AcctNum&, RecNo&)
DECLARE SUB SmallPause ()
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
DECLARE FUNCTION GetTaxCustCnt& ()
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
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
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
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Calculate Interest"
  MChoice$(2) = "Edit Interest Transactions"
  MChoice$(3) = "Print Interest Edit Report "
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
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing 'Interest' Menu ", Cnf
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
  
  IF Exist("Taxbill" + Ext$) THEN
    IF MainFlag THEN
      KillFile "TAXFROMM.DAT"
      RUN "TaxMenu"
    ELSE
      RUN "TaxBILL"
    END IF
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB CalcInterest
  
  SHARED Choice$()
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM IntTrans(1) AS InterestRecType
  
  DIM Year!(99), IAmt#(99)
  
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  TIntLen = LEN(IntTrans(1))
  
  REDIM Fmt$(3)
  
  REDIM TempScrn(0)
  
  REDIM Choice$(0 TO 4, 0)
  Choice$(0, 0) = "4"
  Choice$(1, 0) = "1) Account Number Order"
  Choice$(2, 0) = "2) Customer Name Order"
  
  Dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  
  Fmt$(1) = "$$########,#"
  Fmt$(2) = "#####"
  Fmt$(3) = "#######,#"
  
  MaxLines = 60
  LineCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "CALCINT"
  
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
  
  IF Exist%(TaxIntFile) THEN
    KILL TaxIntFile             'kill any old work file
  END IF
  
  TIntFile = FREEFILE
  OPEN TaxIntFile FOR RANDOM SHARED AS TIntFile LEN = TIntLen
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
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

    IF TaxCust(1).Deleted <> 0 THEN GOTO CalcIntSkip

    IF TaxCust(1).interest = "N" THEN GOTO CalcIntSkip
    
    TransRecord& = TaxCust(1).LastTrans
    
    WHILE TransRecord& > 0
      GET TransFile, TransRecord&, TaxTrans(1)
      CurYearFlag = False
      PastYearFlag = False
      IF TaxTrans(1).TranType = 1 THEN
        Balance# = TaxTrans(1).Revenue.Principle1
        Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd)
        Balance# = Round#(Balance#)
        IF Balance# > 0 THEN
          IF TaxTrans(1).TAXYEAR = WhatYear THEN CurYearFlag = True
          IF TaxTrans(1).TAXYEAR <> WhatYear THEN PastYearFlag = True
          'Get Bill Number
          BillNumber$ = TaxTrans(1).Description

          BillNumber$ = ParseBillNum$(BillNumber$)
          'sb = INSTR(BillNumber$, "Bill #")
          'IF sb > 0 THEN
          '  BillNumber$ = MID$(TaxTrans(1).Description, sb + 6, 10)
          'END IF
          
          IF CurYearFlag THEN
            IntAmount# = Round#(Balance# * (CurRate# / 100))
          END IF
          IF PastYearFlag THEN
            IntAmount# = Round#(Balance# * (PastRate# / 100))
          END IF

          IF Round#(IntAmount#) > 0 THEN
            NME$ = RTRIM$(TaxCust(1).FName) + " " + RTRIM$(TaxCust(1).LName)
            NME$ = LTRIM$(NME$)
            IntTrans(1).CustRec = CustAcct&
            IntTrans(1).CustName = NME$
            IntTrans(1).TAXYEAR = TaxTrans(1).TAXYEAR
            IntTrans(1).Amount = Round#(IntAmount#)
            IntTrans(1).BillNumber = BillNumber$
            IntTrans(1).billrec = TransRecord&
            IntTrans(1).CurYear = WhatYear
            IntRecord& = IntRecord& + 1
            PUT TIntFile, IntRecord&, IntTrans(1)
            'IF IntRecord& = 5 THEN GOTO ExitWend   'For testing
          END IF
        END IF  'Check For Balance
      END IF    'Check for Tax Bill
      TransRecord& = TaxTrans(1).LastTrans
    WEND
    
CalcIntSkip:
    ShowPctComp Cnt&, NumOfCRecs&
  NEXT

  'CalcInt Calc END   *******************************
  CLOSE

  'CalcInt Print Start #############

   PrintTrans
ExitCalcInt:

   EXIT SUB
  
'  ShowProcessingScrn "Creating Interest Report"
'  RptFile = FREEFILE
'  OPEN InterestReportFile FOR OUTPUT AS #RptFile
'
'  GOSUB CalcIntHeading
'
'  FOR Cnt& = 1 TO IntRecord&
'    GET #TIntFile, Cnt&, IntTrans(1)
'    PRINT #RptFile, USING "######"; IntTrans(1).CustRec;
'    PRINT #RptFile, TAB(8); IntTrans(1).CustName;
'    PRINT #RptFile, TAB(50); IntTrans(1).TAXYEAR;
'    PRINT #RptFile, TAB(60); QPTrim$(IntTrans(1).BillNumber);
'    PRINT #RptFile, TAB(70); USING "######.##"; IntTrans(1).Amount
'    LineCnt = LineCnt + 1
'    TotalInterest# = TotalInterest# + IntTrans(1).Amount
'    IF IntTrans(1).TAXYEAR = WhatYear THEN
'      CurInterest# = CurInterest# + IntTrans(1).Amount
'    ELSE
'      PastInterest# = PastInterest# + IntTrans(1).Amount
'    END IF
'
'    'Now Book to Specific Year
'    IF YearCount = 0 THEN
'      YearCount = 1
'      Year!(YearCount) = IntTrans(1).TAXYEAR
'      IAmt#(YearCount) = IntTrans(1).Amount
'    ELSE
'      IFlag = 0
'      FOR SCnt = 1 TO YearCount
'        IF Year!(SCnt) = IntTrans(1).TAXYEAR THEN
'          IAmt#(SCnt) = IAmt#(SCnt) + IntTrans(1).Amount
'          IFlag = 1
'        END IF
'      NEXT SCnt
'      IF IFlag = 0 THEN
'        YearCount = YearCount + 1
'        Year!(YearCount) = IntTrans(1).TAXYEAR
'        IAmt#(YearCount) = IntTrans(1).Amount
'      END IF
'    END IF
'
'    IF LineCnt >= MaxLines THEN
'      PRINT #RptFile, CHR$(12);
'      GOSUB CalcIntHeading
'    END IF
'
'    ShowPctComp Cnt&, IntRecord&
'  NEXT
'
'  PRINT #RptFile, Dash$
'  PRINT #RptFile, "Total Interest Charged .. "; USING "$$######,#.##"; TotalInterest#
'  PRINT #RptFile, "Total Current Interest .. "; USING "$$######,#.##"; CurInterest#
'  PRINT #RptFile, "Total Past Interest ..... "; USING "$$######,#.##"; PastInterest#
'  PRINT #RptFile, CHR$(12);
'  Page = Page + 1
'  PRINT #RptFile, "Interest Breakdown by Tax Year"; TAB(65); "Page #"; Page
'  PRINT #RptFile, ""
'  PRINT #RptFile, " Year"; TAB(10); "Interest Calc."
'  PRINT #RptFile, "--------------------------------"
'  FOR LN = 1 TO YearCount
'    PRINT #RptFile, Year!(LN); TAB(10); USING "$$######,#.##"; IAmt#(LN)
'  NEXT LN
'
'  CLOSE
'  PrintRptFile "Tax Interest Report.", InterestReportFile, 1, RetCode%, 0
  
  
'CalcIntHeading:
'  Page = Page + 1
'  PRINT #RptFile, TAB(15); "Property Tax Billing : Interest Calculation Register"
'  PRINT #RptFile,
'  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; Page
'  PRINT #RptFile, "Current Tax Year is "; WhatYear
'  PRINT #RptFile, "Acct #"; TAB(8); "Customer Name"; TAB(50); "Tax Yr"; TAB(60); "Bill #"; TAB(71); "Interest"
'  PRINT #RptFile, Dash$
'  LineCnt = 5
'  RETURN
  
CheckPreInfo1:
  WhatYear = Value#(Form$(1, 0), ECode)
  CurRate# = Value#(Form$(2, 0), ECode)
  PastRate# = Value#(Form$(3, 0), ECode)
  Order$ = LEFT$(Form$(4, 0), 1)
  
  IF WhatYear < 1997 OR WhatYear > 2010 THEN
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
  
'GetPastBalance1:
'  Balance# = 0
'  ABalance# = 0
'  IF TaxCust(1).LastTrans > 0 THEN
'    TransRecord& = TaxCust(1).LastTrans
'    DO WHILE TransRecord& <> 0
'      GET TransFile, TransRecord&, TaxTrans(1)
'      IF TaxTrans(1).TranType = 1 THEN
'      END IF
'      ABalance# = ABalance# + Balance#
'      TransRecord& = TaxTrans(1).LastTrans
'    LOOP
'    CLOSE TransFile
'  END IF
'  RETURN
  
END SUB

