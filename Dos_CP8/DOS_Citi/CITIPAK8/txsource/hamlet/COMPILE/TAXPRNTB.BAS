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
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
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
  
  STACK 5000



  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE




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
    
    TitleBox 2, Col, MaxLen + 3, "Bill Printing Menu ", Cnf
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
  
  RUN "TAXBILL"
  
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

  DIM ZipCnt&(500), ZIP$(500)
  
  REDIM TaxBill(1)  AS TaxBillType
  BillLen = LEN(TaxBill(1))
  
  ClearBack
  ShowProcessingScrn "Processing Billing Report."
  
  RptFile = FREEFILE
  OPEN "TBPRNRPT.PRN" FOR OUTPUT AS RptFile
  
  BillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  GOSUB BillPrintHeader
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #BillFile, Cnt&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
      PCnt = PCnt + 1
      PRINT #RptFile, USING "#####"; TaxBill(1).BillNumber;
      PRINT #RptFile, TAB(12); QPTrim$(LEFT$(TaxBill(1).CustName, 32));
      PRINT #RptFile, TAB(45); USING Fmt$(1); TaxBill(1).RealTaxDue; TAB(59); TaxBill(1).PersTaxDue;
      PRINT #RptFile, TAB(69); USING Fmt$(2); TaxBill(1).TotalBillDue
      RTotal# = Round#(RTotal# + TaxBill(1).RealTaxDue)
      PTotal# = Round#(PTotal# + TaxBill(1).PersTaxDue)
      GTotal# = Round#(GTotal# + TaxBill(1).TotalBillDue)

      GOSUB ZipSort
      
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
  PRINT #RptFile, "Billing Totals:"; TAB(20); PCnt; TAB(41); USING Fmt$(3); RTotal#; TAB(55); PTotal#; TAB(68); GTotal#

  PRINT #RptFile, ""
  PRINT #RptFile, "Bills By Zip Code"

   'MiniSort Here
    GOSUB sortthem


    
   FOR ZC = 1 TO ZCnt&
    PRINT #RptFile, "Zip Code "; ZIP$(ZC); "  Bills Printed = "; USING "####,#"; ZipCnt&(ZC)
    TZ# = TZ# + ZipCnt&(ZC)
    LC = LC + 1
      IF LC > 55 THEN
        PRINT #RptFile, CHR$(12);
        LC = 1
        PRINT #RptFile, "Bills By Zip Code"
      END IF
   
  

  NEXT ZC
  PRINT #RptFile, "Total Bills Counted by Zip: "; TZ#
  PRINT #RptFile, CHR$(12);
  CLOSE
  
  ERASE Fmt$, TaxBill
  
  PrintRptFile "Bills Printed Report.", "TBPRNRPT.PRN", 1, RetCode%, 0
  
  EXIT SUB
  
BillPrintHeader:
  page = page + 1
  PRINT #RptFile, TAB(20); "Property Tax Billing : Bills Printed Report"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(70); "Page #"; page
  PRINT #RptFile, "Bill No."; TAB(12); "Customer Name                    Real Due      Pers Due        Total"
  PRINT #RptFile, Dash$
  LineCnt = 5
RETURN

ZipSort:
  ZIP& = VAL(LEFT$(TaxBill(1).CustZip, 5))
IF ZCnt& = 0 THEN
    ZCnt& = 1
    ZipCnt&(ZCnt&) = 1
    ZIP$(Znt&) = LTRIM$(LEFT$(TaxBill(1).CustZip, 5))
    RETURN
END IF

FOR ZZCnt& = 1 TO ZCnt&
 IF ZIP& = VAL(ZIP$(ZZCnt&)) THEN
  ZipCnt&(ZZCnt&) = ZipCnt&(ZZCnt&) + 1
  RETURN
 END IF
NEXT ZZCnt&
 ZCnt& = ZCnt& + 1
 ZipCnt&(ZCnt&) = 1
 ZIP$(ZCnt&) = LTRIM$(LEFT$(TaxBill(1).CustZip, 5))
 RETURN



sortthem:
26000 REM sort
26020 M = ZCnt&
26030 M = INT(M / 2)
26040 IF M = 0 THEN 26190
26050 FOR st = 1 TO M
26060 I = st
26070 J = st + M
26080 SW = 0
26090 IF ZIP$(I) <= ZIP$(J) THEN 26120
26100 SW = 1
26110 SWAP ZIP$(I), ZIP$(J): SWAP ZipCnt&(I), ZipCnt&(J)
26120 I = J
26130 J = J + M
26140 IF J <= ZCnt& THEN 26090
26150 IF SW = 0 THEN 26170
26160 GOTO 26060
26170 NEXT st
26180 GOTO 26030
26190 RETURN


  
END SUB

SUB PrintTaxBills
  
  SHARED Choice$()
  
  REDIM TaxBill(1)  AS TaxBillType
  BillLen = LEN(TaxBill(1))

  REDIM MortCodeRec(1) AS MortCodeRecType
  MortCodeRecLen = LEN(MortCodeRec(1))
  REDIM Array(1 TO 100) AS Struct 'Template for the sort array
  DIM MCode$(100)
  
  
  REDIM BillInfo(1) AS TaxBillInfoType
  BillInfoLen = LEN(BillInfo(1))
  FGetAH "TAXBINFO.DAT", BillInfo(1), BillInfoLen, 1
  
  REDIM TempScrn(0)
  
  REDIM Choice$(0 TO 4, 0)
  Choice$(0, 0) = "6"
  
  REDIM PINTemp(1) AS STRING * 9
  
  LibName$ = "TAX"
  FormName$ = "PRNTBIL"
  
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
      UnPackBuffer 1, 6, Form$(), Fld()
      FirstTime = False
      Fld(6).Protected = True
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
      PrintRptFile "", "TAXBLMSK.DAT", 1, RetCode%, 4
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
  CLOSE TxFile

  RptFile = FREEFILE
  OPEN "TAXBIL.PRN" FOR OUTPUT AS RptFile
  
  BillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
 
 'MasterLoop=1 Means Print All Non Mort Code Customers First
  FOR Cnt& = 1 TO NumOfRec&
    WhatRec& = Cnt&
    GET #BillFile, WhatRec&, TaxBill(1)
    IF TaxBill(1).BillNumber > 0 THEN
      IF TaxBill(1).TotalBillDue > 0 THEN
        IF VAL(QPTrim$(TaxBill(1).MortCode)) = 0 THEN
          RSET PINTemp(1) = TaxBill(1).RealPin
          CustName$ = QPTrim$(TaxBill(1).CustName)
          TaxBill(1).BillNumber = BillNo&
          '$INCLUDE: 'TAXHMLET.BI'
          TaxBill(1).BillPrinted = True
          PUT #BillFile, WhatRec&, TaxBill(1)
          BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
       END IF
      END IF
    END IF
    ShowPctComp WhatRec&, NumOfRec&
  NEXT
  
 'MasterLoop=2 Means Print All Mort Code Customers 2nd
 'Get Mort Codes
 GOSUB MCodes
 FOR MCnt& = 1 TO NumOfCatRecs
  QPrintRC "Processing Mortgage Code: " + MCode$(MCnt&), 11, 27, -1
  FOR Cnt& = 1 TO NumOfRec&
    WhatRec& = Cnt&
    GET #BillFile, WhatRec&, TaxBill(1)
    IF TaxBill(1).BillNumber > 0 THEN
      IF TaxBill(1).TotalBillDue > 0 THEN
        IF VAL(QPTrim$(TaxBill(1).MortCode)) = VAL(MCode$(MCnt&)) THEN
          RSET PINTemp(1) = TaxBill(1).RealPin
          CustName$ = QPTrim$(TaxBill(1).CustName)
          TaxBill(1).BillNumber = BillNo&
          '$INCLUDE: 'TAXHMLET.BI'
          TaxBill(1).BillPrinted = True
          PUT #BillFile, WhatRec&, TaxBill(1)
          BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
       END IF
      END IF
    END IF
    ShowPctComp MCnt&, NumOfCatRecs
  NEXT Cnt&
  NEXT MCnt&
  CLOSE

  ERASE PINTemp, TaxBill, BillInfo
  ERASE Frm, Form$, Fld, TempScrn
  
  PrintRptFile "Tax Bill Printing.", "TAXBIL.PRN", 1, RetCode%, 0
  
ExitBillPrint:
  EXIT SUB
  
CheckBillParm:
  WhatYear = Value#(Form$(1, 0), ECode)
  BillNo& = Value#(Form$(2, 0), ECode)
  
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
    BillInfo(1).BillNum = BillNo&
    'update the bill info file with the starting bill number
    BInfoFile = FREEFILE
    OPEN "TAXBINFO.DAT" FOR RANDOM AS BInfoFile LEN = BillInfoLen
    PUT #BInfoFile, 1, BillInfo(1)
    CLOSE BInfoFile
  END IF
  
RETURN


MCodes:
  MortFile = FREEFILE
  OPEN "TAXMORT.DAT" FOR RANDOM AS MortFile LEN = MortCodeRecLen
  NumOfCatRecs = LOF(MortFile) \ MortCodeRecLen
  IF NumOfCatRecs > 0 THEN
    GOSUB SortMortCodes
    REDIM MChoice$(1 TO NumOfCatRecs)
    FOR Cnt = 1 TO NumOfCatRecs
      GET MortFile, Array(Cnt).RecNum, MortCodeRec(1)
      MCode$(Cnt) = RTRIM$(MortCodeRec(1).MortCode)
    NEXT Cnt
  END IF
  RETURN


SortMortCodes:

  FOR Cnt = 1 TO NumOfCatRecs
    GET MortFile, Cnt, MortCodeRec(1)
    Array(Cnt).who = MortCodeRec(1).MortCode + STRING$(11, " ")
    Array(Cnt).RecNum = Cnt
  NEXT Cnt
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN



  
END SUB

SUB RePrintTaxBills
  
  REDIM TaxBill(1)  AS TaxBillType
  BillLen = LEN(TaxBill(1))
  
  REDIM BillInfo(1) AS TaxBillInfoType
  BillInfoLen = LEN(BillInfo(1))
  FGetAH "TAXBINFO.DAT", BillInfo(1), BillInfoLen, 1
  
  REDIM Fmt$(1 TO 3)
  Fmt$(1) = "#####"
  Fmt$(2) = "#.####"
  Fmt$(3) = "##.##"
  
  REDIM TempScrn(0)
  REDIM PINTemp(1) AS STRING * 9
  
  ClearBack
  ShowProcessingScrn "Reading Billing Information"
  BillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM AS BillFile LEN = BillLen
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
  FormName$ = "RPRNTBIL"
  
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
      LSET Form$(1, 0) = STR$(BillInfo(1).BillNum)
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
      PrintRptFile "", "TAXBLMSK.DAT", 1, RetCode%, 4
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
  CLOSE TxFile
  
  RptFile = FREEFILE
  OPEN "TAXBILRE.PRN" FOR OUTPUT AS RptFile
  
  PrnCnt = 0
  BillFile = FREEFILE
  OPEN TaxBillFile FOR RANDOM AS BillFile LEN = BillLen
  NumOfRec& = LOF(BillFile) / BillLen
  
  FOR Cnt& = 1 TO NumOfRec&
    WhatRec& = Cnt&
    GET #BillFile, WhatRec&, TaxBill(1)
    IF TaxBill(1).BillPrinted THEN
    '  PrnCnt = PrnCnt + 1
      'IF PrnCnt >= FBill AND PrnCnt <= LBill THEN
      IF TaxBill(1).BillNumber >= FBill AND TaxBill(1).BillNumber <= LBill THEN
        RSET PINTemp(1) = TaxBill(1).RealPin
        CustName$ = QPTrim$(TaxBill(1).CustName)
        PrnCnt = PrnCnt + 1
        '$INCLUDE: 'TaxHMLET.BI'
        PUT #BillFile, WhatRec&, TaxBill(1)
'      ELSE
'        STOP
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
  QPrintRC FUsing$(STR$(BillInfo(1).TaxYear), Fmt$(1)), 11, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).BillNum), Fmt$(1)), 12, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).REALRATE), Fmt$(2)), 13, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).PERSRATE), Fmt$(2)), 14, 39, -1
  QPrintRC FUsing$(STR$(BillInfo(1).LATEPCT), Fmt$(3)), 15, 39, -1
  QPrintRC BillInfo(1).PRNORDER, 17, 39, -1
RETURN
  
END SUB

