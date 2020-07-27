DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE SUB LateLabels ()
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

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = "Print Late Notice"
  MChoice$(2) = "Print Late Notice Labels"
  MChoice$(3) = "Exit to OS"

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

    TitleBox 2, Col, MaxLen + 3, "Late Notice Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      PrintLateNotices
    CASE 2
      LateLabels
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  IF INSTR(COMMAND$, "TEST") > 0 THEN
    END
  END IF

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("Taxmenu" + Ext$) THEN
    RUN "TaxBill"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  
'  PrintLateNotices
'
'  IF WEnvTest THEN
'    Ext$ = ".bas"
'  ELSE
'    Ext$ = ".exe"
'  END IF
'  IF INSTR(COMMAND$, "TEST") <= 0 THEN
'    RUN "TAXBILL"
'  END IF
  
  END

SUB LateLabels

  SHARED Choice$()

  ReportFile$ = "TAXLABEL.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  MaxLines = 56
  LineCnt = 0
  CustCnt = 0

  LibName$ = "TAX"
  ScrnName$ = "VCUSTLAB"

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F

  REDIM Choice$(0 TO 2, 0 TO 2)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  Choice$(0, 2) = "4"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  Form$(2, 0) = "R"             'Default to Real
  Form$(3, 0) = "N"             'Default to No
  Action = 1
  ClearBack

  ShowCursor
  FirstTime = True

  DisplayTaxScrn ScrnName$

  DO
'
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FOR zz = 10 TO 15
        QPrintRC SPACE$(25), zz, 20, 0
      NEXT

      QPrintRC "Late Notice Mailing Labels.", 11, 27, 14
      QPrintRC "Print's Labes for Last run only!", 13, 24, 14
      FirstTime = False
      Fld(1).Protected = True
      Fld(2).Protected = True
      Fld(3).Protected = True
      Action = 1
    END IF


    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      TaxType$ = Form$(2, 0)
      IF LEFT$(Form$(3, 0), 1) = "Y" THEN
        DetailFlag = True
      ELSE
        DetailFlag = False
      END IF
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      ExitFlag = True
    CASE F5KEY
      GOSUB PrintAlign
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag

  IF AbortFlag THEN EXIT SUB

  TaxLLHandle = FREEFILE
  OPEN "TAXLNLST.DAT" FOR RANDOM AS TaxLLHandle LEN = 4
  LCnt& = LOF(TaxLLHandle) / 4
  REDIM IdxBuff(1 TO LCnt&) AS LONG
  FOR zz& = 1 TO LCnt&
    GET TaxLLHandle, , IdxBuff(zz&)
  NEXT
  CLOSE TaxLLHandle


  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec)
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec)


  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndexML
  END IF

  ClearBack
  ShowProcessingScrn "Mailing Labels"

  FOR Cnt = 1 TO LCnt&
    CustRecNo = IdxBuff(Cnt)

    GET TaxFile, CustRecNo, TaxCustRec

    IF NOT TaxCustRec.Deleted THEN
      'Mortcode Test Here

      IF DetailFlag AND TaxType$ = "R" AND LEN(MC$) > 0 THEN
        NME$ = QPTrim$(TaxCustRec.FName) + " " + QPTrim$(TaxCustRec.LNAME)
        NME$ = QPTrim$(NME$)    'this one cleans up those with only last name
        PRINT #RptHandle, USING "#####"; CustRecNo;
        PRINT #RptHandle, TAB(20); "MC="; MC$
        PRINT #RptHandle, NME$
        PRINT #RptHandle, TaxCustRec.Addr1
        PRINT #RptHandle, TaxCustRec.Addr2
        PRINT #RptHandle, TaxCustRec.City; " "; TaxCustRec.State; " "; TaxCustRec.Zip
        PRINT #RptHandle,
      END IF
      IF NOT (DetailFlag) THEN
        NME$ = QPTrim$(TaxCustRec.FName) + " " + QPTrim$(TaxCustRec.LNAME)
        NME$ = QPTrim$(NME$)    'this one cleans up those with only last name
        PRINT #RptHandle, USING "#####"; CustRecNo
        PRINT #RptHandle, NME$
        PRINT #RptHandle, TaxCustRec.Addr1
        PRINT #RptHandle, TaxCustRec.Addr2
        PRINT #RptHandle, TaxCustRec.City; " "; TaxCustRec.State; " "; TaxCustRec.Zip
        PRINT #RptHandle,
      END IF
      CustCnt = CustCnt + 1
    END IF
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT

  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF

  ERASE Frm, Form$, Fld

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB

PrintAlign:
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT
RETURN


GetNameIndexML:
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

SUB PrintLateNotices
  
  SHARED Choice$()

  'Get Bill Format Here
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  GET TxFile, 1, TaxSetUpRec
  BillFormat$ = LEFT$(TaxSetUpRec.TaxForm, 1)
  CLOSE TxFile
  TownName$ = QPTrim$(TaxSetUpRec.Name)

  IF (INSTR(TownName$, "MURPHY") > 0 AND INSTR(TaxSetUpRec.TaxSt, "NC") > 0) THEN
    MurphyFlag = True
  END IF
  IF (INSTR(TownName$, "LANDIS")) > 0 THEN
    LandisFlag = True
  END IF
  IF (INSTR(TownName$, "MAGGIE")) > 0 THEN
    MaggFlag = True
  END IF

  IF (INSTR(TownName$, "CANTON")) > 0 THEN
    SpecialFlag = True
  END IF

  IF (INSTR(TownName$, "COLUMBIA")) > 0 THEN
    SpecialFlag = True
  END IF

  SpecialFlag = True

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
  
  REDIM PINTemp(1) AS STRING * 12
  
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
  
  CALL KillFile("TAXLNLST.DAT")

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
  
  TaxLLHandle = FREEFILE
  OPEN "TAXLNLST.DAT" FOR RANDOM AS TaxLLHandle LEN = 4

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
    'SmallPause
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
          IF NOT ((TaxTransRec.TranType = 1) AND (TaxTransRec.TaxYear < WhatYear)) THEN
            'IF (TaxTransRec.TranType <> 1 AND TaxTransRec.TaxYear < WhatYear) THEN
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

        'Balance# = Round#(Balance# - (TaxTransRec.Revenue.InterestPd + TaxTransRec.Revenue.PenaltyPd + TaxTransRec.Revenue.CollectionPd))

        IntBal# = Round(TxIntC# - TxIntP#)

        'IF CustAcct& = 2741 THEN STOP
        IF Balance# > 0 THEN

          BillYear = TaxTransRec.TaxYear
          TBalance# = Round#(Balance# + TBalance#)
          PrnCnt = PrnCnt + 1
          ''$INCLUDE: 'TxCLMBLN.BI'     'Columbia
          ''$INCLUDE: 'TxWARNCL.BI'     'warsaw nc
          ''$INCLUDE: 'TxCANTLN.BI'     'Canton
          ''$INCLUDE: 'TxTROYLN.BI'     'Troy
          ''$INCLUDE: 'TaxSTDLN.BI'
          '$INCLUDE: 'TxLNCALB.BI'     'Calabash
          ''$INCLUDE: 'TaxSUNLN.BI'     'Sun Beech
          ''$INCLUDE: 'TaxSUGLN.BI'
          ''$INCLUDE: 'TXNWSTLN.BI'     'Northwest
          ''$INCLUDE: 'TXMAGNLN.BI'     'Magnolia
          ''$INCLUDE: 'TXCLAKLN.BI'     'Clarkton NC
          ''$INCLUDE: 'TXmrphln.BI'     '
          ''$INCLUDE: 'TXlndsln.BI'     '
          ''$INCLUDE: 'TXlnMAGG.BI'     'Maggie Valley
          ''$INCLUDE: 'TXlelsln.BI'     '
          ''$INCLUDE: 'TXSpruln.BI'     '
          PrnCnt = PrnCnt + 1
          PUT #TaxLLHandle, , CustAcct&
        END IF
NotABill:
        TransRecord& = TaxTransRec.LastTrans
      WEND
      IF MurphyFlag OR LandisFlag OR MaggFlag OR SpecialFlag THEN
        IF DoneHeader THEN
          DoFooter = True
          ''$INCLUDE: 'TxCLMBLN.BI'     'Columbia
          ''$INCLUDE: 'TxCANTLN.BI'     'Canton
          ''$INCLUDE: 'TXlelsln.BI'     '
          ''$INCLUDE: 'TXlndsln.BI'     '
          ''$INCLUDE: 'TXmrphln.BI'
          ''$INCLUDE: 'TXlnMAGG.BI'     'Maggie Valley
          ''$INCLUDE: 'TXSpruln.BI'     '
          ''$INCLUDE: 'TXSpruln.BI'     '
          'PrnCnt = PrnCnt + 1
          PUT #TaxLLHandle, , CustAcct&
        END IF
      END IF
    END IF      'Test for Valid Acct
SkipEm:
    IF PrnCnt > 1 AND INSTR(COMMAND$, "TEST") > 0 THEN EXIT FOR
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
                'This label is called from include files
  
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

