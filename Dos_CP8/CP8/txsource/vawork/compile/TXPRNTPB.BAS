DEFINT A-Z
DECLARE SUB BillsPrintedRpt ()
DECLARE SUB RePrintTaxBills ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB PrintTaxBills ()
DECLARE SUB ClearBack ()
DECLARE SUB ClearScrn ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)

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

CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxBINFO.BI'
  '$INCLUDE: 'Taxconst.BI'
  '$INCLUDE: 'TaxBill.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  STACK 6000
  
  DIM SHARED TaxSetUpRec AS TaxMasterType
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = " Print Tax Bills"
  MChoice$(2) = " Reprint Tax Bills"
  MChoice$(3) = " Tax Bills Printed Report "
  MChoice$(4) = " Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Personal Property Bill Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      PrintTaxBills
    CASE 2
      RePrintTaxBills
    CASE 3
      BillsPrintedRpt
    CASE 4
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
  
  RUN "TAXPPBIL"
  
  END

SUB BillsPrintedRpt
  
  FF$ = CHR$(12)
  Dash$ = STRING$(80, "-")
  TheDate$ = DATE$
  MaxLines = 55
  PCnt = 0
  REDIM Fmt$(1 TO 3)
  Fmt$(1) = "#####.##"
  Fmt$(2) = "#######,.##"
  Fmt$(3) = "$$######,.##"
  
  REDIM TaxBill(1)  AS VAPPTaxBillType
  BillLen = LEN(TaxBill(1))
  
  ClearBack
  ShowProcessingScrn "Processing Billing Report."
  
  RptFile = FREEFILE
  OPEN "TBPRNRPT.PRN" FOR OUTPUT AS RptFile
  
  BillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  GOSUB BillPrintHeader
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      PCnt = PCnt + 1
      PRINT #RptFile, USING "#####"; TaxBill(1).BillNumber;
      PRINT #RptFile, TAB(10); QPTrim$(LEFT$(TaxBill(1).CustName, 29));
      PRINT #RptFile, TAB(40); USING Fmt$(1); (TaxBill(1).PersTaxDue + TaxBill(1).MTTaxDue + TaxBill(1).FETaxDue + TaxBill(1).MHTaxDue + TaxBill(1).MCTaxDue);
      PRINT #RptFile, TAB(55); USING Fmt$(1); TaxBill(1).PPTRADiscnt;
      PRINT #RptFile, TAB(69); USING Fmt$(2); TaxBill(1).TotalBillDue
      PTotal# = Round#(PTotal# + TaxBill(1).PersTaxDue + TaxBill(1).MTTaxDue# + TaxBill(1).MHTaxDue + FETaxDue)
      DTotal# = Round#(DTotal# + TaxBill(1).PPTRADiscnt)
      GTotal# = Round#(GTotal# + TaxBill(1).TotalBillDue)
      LineCnt = LineCnt + 1
      IF LineCnt >= MaxLines THEN
        PRINT #RptFile, Dash$
        PRINT #RptFile, FF$
        GOSUB BillPrintHeader
      END IF
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  PRINT #RptFile, Dash$
  PRINT #RptFile, "Billing Totals:"; TAB(20); PCnt
  PRINT #RptFile, "Taxes Levied ........ "; USING Fmt$(3); PTotal#
  PRINT #RptFile, "PPTRA Discount ...... "; USING Fmt$(3); DTotal#
  PRINT #RptFile, "Grand Total Tax Due.. "; USING Fmt$(3); GTotal#
  CLOSE
  
  ERASE Fmt$, TaxBill
  
  PrintRptFile "Bills Printed Report.", "TBPRNRPT.PRN", 1, RetCode%, 0
  
  EXIT SUB
  
BillPrintHeader:
  PAGE = PAGE + 1
  PRINT #RptFile, TAB(20); "Property Tax Billing : Bills Printed Report"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(70); "Page #"; PAGE
  PRINT #RptFile, "Bill No."; TAB(10); "Customer Name"; TAB(40); "Total Tax"; TAB(55); "PPTRA Discount"; TAB(72); "Total Due"
  PRINT #RptFile, Dash$
  LineCnt = 5
RETURN
  
END SUB

SUB PrintTaxBills
  
  SHARED Choice$()
  
  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))

  REDIM TaxBill(1)  AS VAPPTaxBillType
  BillLen = LEN(TaxBill(1))
  
  REDIM BillInfo(1) AS VAPPTaxBillInfoType
  BillInfoLen = LEN(BillInfo(1))
  FGetAH "TAXBINF2.DAT", BillInfo(1), BillInfoLen, 1
  
  REDIM TempScrn(0)
  
  REDIM Choice$(0 TO 4, 0)
  Choice$(0, 0) = "6"
  
  REDIM PINTemp(1) AS STRING * 9
  
  LibName$ = "TAX"
  FormName$ = "VPRNTBIL"
  
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
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      BCopy VARSEG(BillInfo(1)), VARPTR(BillInfo(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), BillInfoLen, 0
      UnPackBuffer 0, 0, Form$(), Fld()
      FirstTime = False
      Fld(9).Protected = True
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      GOSUB CheckBillParm
      IF OK2Print THEN
        EXIT DO
      END IF
    CASE F5KEY
      CursorOff
      SaveScrn TempScrn()
      PrintRptFile "", "TAXPPMSK.DAT", 1, RetCode%, 4
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
  
  ClearBack
  ShowProcessingScrn "Creating Tax Bills"

 'Get Bill Format Here
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  GET TxFile, 1, TaxSetUpRec
   BillFormat$ = LEFT$(TaxSetUpRec.TaxForm, 1)
   TownName$ = TaxSetUpRec.Name
   Add1$ = TaxSetUpRec.Add1
   Add2$ = TaxSetUpRec.Add2
   Add3$ = TaxSetUpRec.Add3
  CLOSE TxFile
 
  RptFile = FREEFILE
  OPEN "TAXPPBIL.PRN" FOR OUTPUT AS RptFile
  
  BillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  FOR Cnt& = 1 TO NumOfRec&
    WhatRec& = Cnt&
    GET #BillFile, WhatRec&, TaxBill(1)
    IF TaxBill(1).BillNumber > 0 THEN
    
        RSET PINTemp(1) = TaxBill(1).REALPIN
        CustName$ = QPTrim$(TaxBill(1).CustName)
        TaxBill(1).BillNumber = BillNo&

         'Ashland Tax Bills
          TYear$ = Form$(1, 0)
          PersTaxRate# = VAL(Form$(3, 0))
          MHTaxRate# = VAL(Form$(4, 0))
          MCTaxRate# = VAL(Form$(5, 0))
          FETaxRate# = VAL(Form$(6, 0))
          MTTaxRate# = VAL(Form$(7, 0))
          DueDate$ = Form$(11, 0)
          DueDate$ = LEFT$(DueDate$, 6) + RIGHT$(DueDate$, 2)

'          '$INCLUDE: 'TXPPWARS.BI'              'Warsaw
'          '$INCLUDE: 'TXPPNMRK.BI'              'Standard VA Pers Prop Bill w/Remington VA Prorating Shown
'          '$INCLUDE: 'TXASHLD.BI'               'Ashland VA Bill
'          '$INCLUDE: 'TXMIDTWN.BI'              'Middletown VA Bill
'          '$INCLUDE: 'TXINDPP.BI'               'Independence VA Bill
'          '$INCLUDE: 'TXPPHALX.BI'              'Halifax VA Bill
'          '$INCLUDE: 'TXPPSTD.BI'               'Standard VA Pers Prop Bill
          '$INCLUDE: 'TXPPREM.BI'               'Standard VA Pers Prop Bill w/Remington VA Prorating Shown
'          '$INCLUDE: 'TXPPGDY.BI'               'Standard VA Pers w/Grundy Special
'          '$INCLUDE: 'TXPPHILL.BI'              'Standard VA Pers w/Hillsville Special
'          '$INCLUDE: 'TXSTPAUL.BI'              'ST Paul VA Pers Prop Bill
'          '$INCLUDE: 'TXPPCDBF.BI'              'Cedar Bluff (fag boy special)

          TaxBill(1).BillPrinted = True
          PUT #BillFile, WhatRec&, TaxBill(1)
          BillNo& = BillNo& + 1
          PrnCnt = PrnCnt + 1
    
    END IF
    ShowPctComp WhatRec&, NumOfRec&
  NEXT
  
  CLOSE
  
  ERASE PINTemp, TaxBill, BillInfo
  ERASE Frm, Form$, Fld, TempScrn

  PrintRptFile "Tax Bill Printing.", "TAXPPBIL.PRN", 1, RetCode%, 0
  
ExitBillPrint:
  EXIT SUB
  
CheckBillParm:
  WhatYear = Value#(Form$(1, 0), ECode)
  BillNo& = Value#(Form$(2, 0), ECode)
  DueDate$ = Form$(10, 0)

  IF BillNo& <= 0 THEN
    OK2Print = False
    SaveScrn TempScrn()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "Invalid Starting Bill Number!", 10, 26, -1
    QPrintRC "Please correct and try again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
  ELSE
    PrnCnt = 1
    OK2Print = True
    BillInfo(1).BILLNUM = BillNo&
    'update the bill info file with the starting bill number
    BInfoFile = FREEFILE
    OPEN "TAXBINF2.DAT" FOR RANDOM AS BInfoFile LEN = BillInfoLen
    PUT #BInfoFile, 1, BillInfo(1)
    CLOSE BInfoFile
  END IF
  
RETURN
  
END SUB

SUB RePrintTaxBills

  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))

  REDIM TaxBill(1)  AS VAPPTaxBillType
  BillLen = LEN(TaxBill(1))
  
  REDIM BillInfo(1) AS VAPPTaxBillInfoType
  BillInfoLen = LEN(BillInfo(1))
  FGetAH "TAXBINF2.DAT", BillInfo(1), BillInfoLen, 1
  
  REDIM Fmt$(1 TO 2)
  Fmt$(1) = "#####"
  Fmt$(2) = "#.###"
  
  REDIM TempScrn(0)
  REDIM PINTemp(1) AS STRING * 9
  
  ClearBack
  ShowProcessingScrn "Reading Billing Information"
  BillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      Printed = Printed + 1
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  
  LibName$ = "TAX"
  FormName$ = "VRPRNBIL"
  
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
  GOSUB ShowPrintInfo
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
  IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = STR$(Printed)
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      GOSUB CheckBillReParm
      IF OK2Print THEN
        EXIT DO
      END IF
    CASE F5KEY
      CursorOff
      SaveScrn TempScrn()
      PrintRptFile "", "TAXPPMSK.DAT", 1, RetCode%, 4
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
    GOTO ExitReprintBills
  END IF
  
  ClearBack
  ShowProcessingScrn "Reprinting Tax Bills"
 'Get Bill Format Here
  TxFile = FREEFILE
  OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
  GET TxFile, 1, TaxSetUpRec
   BillFormat$ = LEFT$(TaxSetUpRec.TaxForm, 1)
   TownName$ = TaxSetUpRec.Name
   Add1$ = TaxSetUpRec.Add1
   Add2$ = TaxSetUpRec.Add2
   Add3$ = TaxSetUpRec.Add3
  CLOSE TxFile

  RptFile = FREEFILE
  OPEN "TAXBILRE.PRN" FOR OUTPUT AS RptFile
  
  PrnCnt = 0
  BillFile = FREEFILE
  OPEN PPTaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  FOR Cnt& = 1 TO NumOfRec&
    WhatRec& = Cnt&
    GET #BillFile, WhatRec&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      PrnCnt = PrnCnt + 1
      IF PrnCnt >= FBill AND PrnCnt <= LBill THEN
        RSET PINTemp(1) = TaxBill(1).REALPIN
        CustName$ = QPTrim$(TaxBill(1).CustName)

  ' Ashland Format
   PersTaxRate# = BillInfo(1).PERSRATE
   MHTaxRate# = BillInfo(1).MHRate
   MCTaxRate# = BillInfo(1).MCRate
   FETaxRate# = BillInfo(1).FERate
   MTTaxRate# = BillInfo(1).MTRate
   DueDate$ = Num2Date$(BillInfo(1).DueDate)
   DueDate$ = LEFT$(DueDate$, 6) + RIGHT$(DueDate$, 2)
   WhatYear = BillInfo(1).TaxYear
   TYear$ = LTRIM$(STR$(WhatYear))
         
'          '$INCLUDE: 'TXPPWARS.BI'              'Warsaw
'           '$INCLUDE: 'TXMIDTWN.BI'           'Middletown Va Bill
'           '$INCLUDE: 'TXPPHALX.BI'           'Halifax VA Bill
'           '$INCLUDE: 'TXINDPP.BI'            'Independence VA Bill
'           '$INCLUDE: 'TXPPSTD.BI'            'Standard Va Pers Prop Bill We Sell in Virgina
'           '$INCLUDE: 'TXPPGDY.BI'            'Standard VA Pers w/Grundy Special
'           '$INCLUDE: 'TXPPHILL.BI'           'Standard VA Pers w/Hillsville Special
           '$INCLUDE: 'TXPPREM.BI'            'Standard VA Pers Prop Bill w/Remington VA Prorating Shown
'           '$INCLUDE: 'TXSTPAUL.BI'           'ST Paul VA Pers Prop Bill
'           '$INCLUDE: 'TXPPCDBF.BI'           'Cedar Bluff (fag boy special)
'           '$INCLUDE: 'TXPPNMRK.BI'           'Standard VA Pers Prop Bill w/Remington VA Prorating Shown
        PUT #BillFile, WhatRec&, TaxBill(1)
      END IF
    END IF
    ShowPctComp WhatRec&, NumOfRec&
  NEXT
  
  CLOSE
  
  ERASE PINTemp, TaxBill, BillInfo, Fmt$
  ERASE Frm, Form$, Fld, TempScrn
  
  PrintRptFile "Tax Bill Reprinting.", "TAXBILRE.PRN", 1, RetCode%, 0
  
ExitReprintBills:
  EXIT SUB
  
CheckBillReParm:
  
  FBill = Value#(Form$(1, 0), ECode)
  LBill = Value#(Form$(2, 0), ECode)
  
  IF FBill < 1 OR FBill > LBill THEN
    ParmErr = 1
    Frm(1).FldNo = 1
    GOSUB ShowParmError
    GOTO ParmErrorExit
  ELSEIF LBill > Printed THEN
    ParmErr = 2
    Frm(1).FldNo = 2
    LSET Form$(2, 0) = STR$(Printed)
    GOSUB ShowParmError
    GOTO ParmErrorExit
  ELSE
    OK2Print = True
  END IF
  
ParmErrorExit:
RETURN
  
ShowParmError:
  SaveScrn TempScrn()
  DisplayTaxScrn "ERRSCRN1"
  SELECT CASE ParmErr
  CASE 1
    QPrintRC "Invalid Starting Bill Number!", 10, 26, -1
  CASE 2
    QPrintRC "Invalid Ending Bill Number!", 10, 27, -1
  END SELECT
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn TempScrn()
  Action = 1
RETURN
  
ShowPrintInfo:
  
  QPrintRC FUsing$(STR$(BillInfo(1).TaxYear), Fmt$(1)), 9, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).BILLNUM), Fmt$(1)), 10, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).PERSRATE), Fmt$(2)), 11, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).MHRate), Fmt$(2)), 12, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).MCRate), Fmt$(2)), 13, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).FERate), Fmt$(2)), 14, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).MTRate), Fmt$(2)), 15, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).LATEPCT), Fmt$(1)), 16, 39, -1
  
  QPrintRC BillInfo(1).PRNORDER, 17, 39, -1
  QPrintRC Num2Date$(BillInfo(1).DueDate), 18, 39, -1
RETURN
  
END SUB

