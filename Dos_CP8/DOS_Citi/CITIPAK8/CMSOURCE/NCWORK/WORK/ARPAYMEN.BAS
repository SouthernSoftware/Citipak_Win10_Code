DEFINT A-Z
DECLARE SUB OperatorEntry (OperOKFlag%, OperNum%, PostDate$)
DECLARE FUNCTION AskSavePayment% (DefaultFlag$)
DECLARE SUB PrintReceipt (OperNum, TRRec%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ClearBack ()
DECLARE SUB CursorOff ()
DECLARE SUB PostPayments (OperNum)
DECLARE SUB PrintEditList (OpenNum)
DECLARE SUB EditPayment (OperNum, PostDate$)
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
DECLARE SUB OpenARCustFile (NumOfArRecs, ARFile)
DECLARE SUB EnterPayment (OperNum, PostDate$)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB DisplayARScrn (ScrnName$)
DECLARE SUB WaitForAction ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE FUNCTION GetDEFPort% ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'CMFILES.BI'

  CONST False = 0, True = NOT False
  
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
  DIM SHARED EditPaymentRec(1) AS AREditPaymentRecType
  
  '  GOTO StartHere
  
  STACK 8000

SUB DisplayARScrn (ScrnName$)
  LibFile2Scrn "AR", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB EnterPayment (OperNum, PostDate$)
  SHARED Choice$()
  
  REDIM TempScrn(0)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  TOWNNAME$ = UBSetUpRec(1).UTILNAME

  Oper$ = QPTrim$(STR$(OperNum))
  PayFileName$ = "ARCMPY" + Oper$ + ".DAT"

MainBody:
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0
  
  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  help$ = "Enter Payments"
  CursorOff
  
  
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
  
  Action = 1
  
  Form$(1, 0) = "Payment"
  IF LEN(PrevDate$) = 0 THEN
    Form$(2, 0) = DATE$
  ELSE
    Form$(2, 0) = PrevDate$
  END IF
  
  Fld(1).Protected = True
  Frm(1).FldNo = 2
  Form$(16, 0) = "N"
  Fld(17).Protected = True
  Form$(17, 0) = "N"
  
  TenderFld = 10
  CashAmtFld = 11
  ChkAmtFld = 12
  
  'FirstTime = True
  REDIM Choice$(0 TO 3, 0 TO 0)
  
  Choice$(0, 0) = "10"
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"
  
  DisplayARScrn "BAKCLEAR"
  DisplayARScrn ScrnName$

KeepEditing:

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).FldNo = TenderFld AND Frm(1).PrevFld <> TenderFld THEN
      GOSUB FixCashChkFlds
    ELSEIF Frm(1).PrevFld = TenderFld AND Frm(1).FldNo <> TenderFld THEN
      GOSUB FixCashChkFlds
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      SELECT CASE Frm(1).PrevFld
      CASE ChkAmtFld, CashAmtFld, 22, 24
      'IF Frm(1).PrevFld = ChkAmtFld OR Frm(1).PrevFld = CashAmtFld THEN
        SaveField ChkAmtFld, Form$(), Fld(), BadField
        SaveField CashAmtFld, Form$(), Fld(), BadField
        CalcFields 0, ChkAmtFld + 1, Form$(), Fld()
        'PrintArray 1, NumFlds - 1, Form$(), Fld()
        AmtOwed# = Value#(Form$(9, 0), ECode)
        AmtPaid# = Value#(Form$(13, 0), ECode)

        TotDist# = Value#(Form$(26, 0), ECode)
        IF AmtPaid# > TotDist# THEN
          Change# = ABS(TotDist# - AmtPaid#)
        ELSE
          Change# = 0
        END IF

        'IF AmtPaid# > AmtOwed# THEN
        '  Change# = ABS(AmtOwed# - AmtPaid#)
        'ELSE
        '  Change# = 0
        'END IF
        LSET Form$(14, 0) = STR$(Change#)
        Action = 2
      'END IF
      END SELECT
    END IF
    
    IF Frm(1).PrevFld = 4 AND CustomerGrabed = 0 THEN
      GOSUB GetCustomer
      IF AccountRecord = 0 THEN
        CLOSE
        Action = 1
        GOTO MainBody
      END IF
      
      REM check for existing transaction
      CLOSE ARFile
      AREditRecLen = LEN(EditPaymentRec(1))
      ARFile = FREEFILE

      OPEN PayFileName$ FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
      NumPayRecs = LOF(ARFile) \ AREditRecLen
      IF NumPayRecs > 0 THEN
        Cnt = 0
        WHILE Cnt < NumPayRecs
          Cnt = Cnt + 1
          GET ARFile, Cnt, EditPaymentRec(1)
          IF (VAL(EditPaymentRec(1).CustNumber) = VAL(Form$(3, 0))) THEN        'AND (EditPaymentRec(1).Amount > 0) THEN
            SaveScrn TempScrn()
            CursorOff
            BEEP
            DisplayARScrn "ARDUPCUS"
            WaitForAction
            RestScrn TempScrn()
            CLOSE
            GOTO MainBody
          END IF
        WEND
      END IF
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckSavePayment
      IF BadPay THEN
        ok = MsgBox("AR", "BADPAYMN")
        IF BadField <= 0 THEN
          BadField = 19
        END IF
        Frm(1).FldNo = BadField
        Action = 1
      ELSE
        SaveTR = AskSavePayment%("Y")
        SELECT CASE SaveTR
        CASE -1  'save no reciept

          GOSUB SaveARPayInfo
          PostPayments OperNum
          GOSUB SaveCMTrans
          DisplayARScrn "PRESSKEY"
          WaitForAction

        CASE 0   'oops continue editing
          Action = 2
          GOTO KeepEditing
        CASE 1   'save do reciept
          GOSUB SaveARPayInfo
          PostPayments OperNum
          GOSUB SaveCMTrans
          SaveScrn TempScrn()
          GOSUB PrintReceiptFunction
          RestScrn TempScrn()
          DisplayARScrn "PRESSKEY"

          WaitForAction

        END SELECT
        'IF SaveTR THEN
        '  gosub
        Done = True
        GOTO MainBody
      END IF
      
    CASE EscKey
      EXIT SUB
    END SELECT
  LOOP
  
GetCustomer:
  FirstTime = True
  CustomerGrabed = 0
  AccountRecord = VAL(Form$(3, 0))
  
  REM **************************************************************************
  
  IF AccountRecord = 0 THEN
    
    MaxLen = 50 'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
    TText$ = SPACE$(MaxLen + 4)
    LSET TText$ = " Cust #    Customer Sort Name"
    
    '--Center Menu within Screen
    Row = 8
    Col = 15
    
    OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
    OpenARCustFile NumOfArRecs, ARFile
    
    REDIM Mchoice$(1 TO NumOfARIdxRecs)
    
    ChoiceCounter = 0
    FOR Cnt = 1 TO NumOfARIdxRecs
      GET ARIdxFile, Cnt, ARCustIdxRec(1)
      IF LEFT$(ARCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
        GET ARFile, ARCustIdxRec(1).IDXRECORD, ARCustRec(1)
        ChoiceCounter = ChoiceCounter + 1
        Mchoice$(ChoiceCounter) = SPACE$(35)
        LSET Mchoice$(ChoiceCounter) = STR$(ARCustIdxRec(1).IDXRECORD)
        MID$(Mchoice$(ChoiceCounter), 10) = LEFT$(ARCustIdxRec(1).IDXNAME, 25)
      END IF
    NEXT Cnt
    CLOSE ARIdxFile
    
    DO
      
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      'ClearBack
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        AccountRecord = 0
        ExitFlag = True
      ELSE
        AccountRecord = VAL(LEFT$(Mchoice$(Choice), 8))
        ExitFlag = True
      END IF
      
    LOOP UNTIL ExitFlag
    
    LibName$ = "AR"
    ScrnName$ = "ARTRANEN"
    CursorOff
    DisplayARScrn ScrnName$
  END IF
  
  REM ************************************************************************
  ' OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
  OpenARCustFile NumOfArRecs, ARFile
  
  IF AccountRecord > 0 AND AccountRecord <= NumOfArRecs THEN
    GET ARFile, AccountRecord, ARCustRec(1)
    CLOSE ARFile
    
    IF (ARCustRec(1).Deleted = "Y") THEN
      GOSUB CustomerDeleted
      GOTO MainBody
    END IF
    'IF (ARCustRec(1).AcctBal <= 0) THEN
    '  GOSUB NoBalance
    '  GOTO MainBody
    'END IF
    
    Form$(3, 0) = ARCustRec(1).CustNumb
    Form$(4, 0) = ARCustRec(1).CustName
    Form$(5, 0) = ARCustRec(1).ADDRESS1
    Form$(6, 0) = ARCustRec(1).CITY
    Form$(7, 0) = ARCustRec(1).STATE
    Form$(8, 0) = ARCustRec(1).ZIPCODE
    
    Form$(9, 0) = FUsing$(STR$(ARCustRec(1).AcctBal), "#####.##")
    
    'Form$(14, 0) = FUsing$(STR$(ARCustRec(1).IssuanceFee), "#####.##")
    'Form$(17, 0) = FUsing$(STR$(ARCustRec(1).FeeBal), "#####.##")
    Form$(19, 0) = FUsing$(STR$(ARCustRec(1).LicBal), "#####.##")
    Form$(21, 0) = FUsing$(STR$(ARCustRec(1).PenBal), "#####.##")
    Form$(25, 0) = Form$(9, 0)
    
    Fld(3).Protected = True
    Fld(14).Protected = True
    Frm(1).FldNo = 10
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    SaveField 0, Form$(), Fld(), BadField
    'RETURN
    
  ELSE
    
    LibName$ = "AR"
    ScrnName$ = "ARBADCUS"
    help$ = "Set Customer Balances"
    CursorOff
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
    BEEP
    
    DisplayARScrn ScrnName$
    
    Done = False
    Action = 1
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        Done = True
      END SELECT
      IF Done = True THEN GOTO MainBody
    LOOP
  END IF
  RETURN
  
CustomerDeleted:
  ok = MsgBox("AR", "ARDELCUS")
  Action = 1
RETURN
  
NoBalance:
  DisplayARScrn "BAKCLEAR"
  ok = MsgBox("AR", "ARNOBAL")
  Action = 1
  
RETURN
  
  
CheckSavePayment:
  
  BadPay = False                'Assume all is well
  PrevDate$ = Form$(2, 0):      'Keep Default Date Here
  EditPaymentRec(1).TranDate = Date2Num(Form$(2, 0))
  TranDate = EditPaymentRec(1).TranDate
  EditPaymentRec(1).CustNumber = Form$(3, 0)
  CustNumb& = VAL(EditPaymentRec(1).CustNumber)
  CustName$ = QPTrim$(Form$(4, 0))
  EditPaymentRec(1).CustName = CustName$

  Address$ = QPTrim$(Form$(5, 0))

  EditPaymentRec(1).Add1 = Address$
  EditPaymentRec(1).CITY = Form$(6, 0)
  EditPaymentRec(1).STATE = Form$(7, 0)
  EditPaymentRec(1).ZIPCODE = Form$(8, 0)
  
  AmtOwed# = Value#(Form$(9, 0), ECode)         'total owed
  EditPaymentRec(1).CashChk = Form$(10, 0)
  EditPaymentRec(1).CashAmt = Value(Form$(11, 0), A%)           'Cash Amount
  TotalCash# = EditPaymentRec(1).CashAmt

  EditPaymentRec(1).ChkAmt = Value(Form$(12, 0), A%)            'Cash Amount
  TotalCheck# = EditPaymentRec(1).ChkAmt

  AmtPaid# = Value(Form$(13, 0), A%)            'amt tendered
  TotDist# = Value#(Form$(26, 0), ECode)

  IF TotDist# <= 0 THEN
  'IF AmtPaid# <= 0 THEN
    BadPay = True
    BadField = 10
    GOTO CheckReturn
  ELSE
    EditPaymentRec(1).AmtPaid = AmtPaid#
  END IF
  IF AmtPaid# > TotDist# THEN
    Change# = ABS(TotDist# - AmtPaid#)
  ELSE
    Change# = 0
  END IF

'  IF AmtPaid# > AmtOwed# THEN
'    Change# = ABS(AmtOwed# - AmtPaid#)
'  ELSE
'    Change# = 0
'  END IF
  
  EditPaymentRec(1).Change = Change#
  IssLic$ = "N"
  IF IssLic$ = "" THEN
    BadPay = True
    BadField = 12
    GOTO CheckReturn
  END IF
  EditPaymentRec(1).ISSUELIC = Form$(15, 0)     'issue lic flag
  Desc$ = QPTrim$(Form$(18, 0))
  EditPaymentRec(1).Desc = Desc$
  EditPaymentRec(1).LICDUE = Value(Form$(19, 0), A%)            'lic due
  EditPaymentRec(1).LICPAID = Value(Form$(22, 0), A%)           'amt to lic
  
'  IF EditPaymentRec(1).LICPAID > EditPaymentRec(1).LICDUE THEN
'    BadPay = True
'    BadField = 19
'    GOTO CheckReturn
'  END IF
  
  EditPaymentRec(1).PENPAID = Value(Form$(24, 0), A%)           'amt to penalty
  EditPaymentRec(1).PENDUE = Value(Form$(21, 0), A%)            'pen due
  
'  IF EditPaymentRec(1).PENPAID > EditPaymentRec(1).PENDUE THEN
'    BadPay = True
'    BadField = 21
'    GOTO CheckReturn
'  END IF
  
  EditPaymentRec(1).TotDue = Value(Form$(25, 0), A%)            'sum of (due)
  EditPaymentRec(1).TotPaid = Value(Form$(26, 0), A%)           'sum of (paid)
  
  TotPaid# = Round#(AmtPaid# - Change#)
  
  IF TotPaid# <> EditPaymentRec(1).TotPaid THEN
    BadPay = True
    BadField = 19
    GOTO CheckReturn
  END IF
  
  EditPaymentRec(1).Amount = TotPaid# 'EditPaymentRec(1).TOTDUE           'Transaction Amount

CheckReturn:

  RETURN

SaveARPayInfo:
  CLOSE
  AREditRecLen = LEN(EditPaymentRec(1))
  ARFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
  NextRec = (LOF(ARFile) \ AREditRecLen) + 1
  PUT ARFile, NextRec, EditPaymentRec(1)
  CLOSE ARFile

RETURN

  
FixCashChkFlds:
  
  SELECT CASE QPTrim$(Form$(TenderFld, 0))
  CASE Choice$(1, 0)            'CASH               this sets the cash or check
    Fld(TenderFld + 1).Protected = False        'amount fields protected or
    Fld(TenderFld + 2).Protected = True         'unprotected based of the
    LSET Form$(ChkAmtFld, 0) = "0"
  CASE Choice$(2, 0)            'CHECK              tender type field selection
    Fld(TenderFld + 1).Protected = True
    Fld(TenderFld + 2).Protected = False
    LSET Form$(CashAmtFld, 0) = "0"
  CASE ELSE     'BOTH
    Fld(TenderFld + 1).Protected = False
    Fld(TenderFld + 2).Protected = False
    '  LSET Form$(ChkAmtFld, 0) = "0"
  END SELECT
  
  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField
  '
  '  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
  '  CalcFields 0, CashAmtFld, Form$(), Fld()
  '
  PrintArray 1, NumFlds - 1, Form$(), Fld()
  
  'GOSUB AutoDistribute:
  RETURN

PrintReceiptFunction:
'  REDIM CMTRRec(1) AS CMTransRecType
'  CMTrRecLen = LEN(CMTRRec(1))
'  CHandle = FREEFILE
'  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
'  RecNumber = (LOF(CHandle) \ CMTrRecLen) + 1
'  CLOSE CHandle

  'Set Printer Port for Receipts Printer if none set to 1

  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))

  OPEN "O", 15, "RECPT.PRN": WIDTH #15, 255

  PRINT #15, ""
  PRINT #15, TOWNNAME$
  PRINT #15, "LICENSE PAYMENT"
  PRINT #15, "Date: "; PostDate$
  PRINT #15,
  PRINT #15, "Account #"; CustNumb$
  PRINT #15, CustName$
  PRINT #15, Address$
  PRINT #15, Desc$
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; AmtOwed#
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; TotalCash#
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; TotalCheck#
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; AmtPaid#
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; Change#
  PRINT #15,
  PRINT #15,
  PRINT #15, "Operator: "; OperNum
  PRINT #15, "Receipt#: "; USING "######"; RecptNum&
  PRINT #15,
  PRINT #15, TAB(7); "T H A N K   Y O U !"
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  CLOSE
  PrintRptFile Header$, "RECPT.PRN", GetDEFPort%, RetCode%, 5

RETURN

SaveCMTrans:
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  
  CMTRRec(1).TransDate = TranDate
  CMTRRec(1).TransAmount = AmtPaid#
  CMTRRec(1).TransCash = TotalCash#
  CMTRRec(1).TransCheck = TotalCheck#
  CMTRRec(1).TransAmtOwed = TotDist#
  CMTRRec(1).TransDesc = Desc$
  CMTRRec(1).TransSource = 41
  CMTRRec(1).TransName = CustName$
  CMTRRec(1).TransAcctNum = CustNumb&
  CMTRRec(1).TransDetNum = DetailRecord!
  CMTRRec(1).TransOperNum = OperNum
  CMTRRec(1).TransPad = ""
  FldFactor = 0
  FOR Cnt = 1 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = 0
  NEXT Cnt
  RecptNum& = 0& + ((LOF(CHandle) / CMTrRecLen) + 1)
  PUT CHandle, RecptNum&, CMTRRec(1)
  CLOSE CHandle
RETURN
  
END SUB

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
END SUB

SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

SUB PostPayments (OperNum)

  Oper$ = QPTrim$(STR$(OperNum))
  'PayFileName$ = "AREDPY" + Oper$ + ".DAT"

  PayFileName$ = "ARCMPY" + Oper$ + ".DAT"

  IF FileSize(PayFileName$) <= 0 THEN
    EXIT SUB
  END IF
  
  REDIM TempScrn(0)
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  
PostTrans:
  
  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
  OpenARCustFile NumOfArRecs, ARFile
  
  AREditRecLen = LEN(EditPaymentRec(1))
  AREdFile = FREEFILE
  
  OPEN PayFileName$ FOR RANDOM ACCESS READ WRITE SHARED AS AREdFile LEN = AREditRecLen
  NumOfArRecs = LOF(AREdFile) \ AREditRecLen
  
  ' See if any records to post
  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
  NextTransRec = NumOfTransRecs + 1
  
  DO
    Cnt = Cnt + 1
    GET AREdFile, Cnt, EditPaymentRec(1)
    
    IF EditPaymentRec(1).Amount <> 0 THEN
      GET ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
      'Get Catagory Code Record # Here
      Code$ = QPTrim$(ARCustRec(1).BILLCAT1)
      CatFile = FREEFILE
      OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
      NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
      CatRecord = 0
      FOR SCnt! = 1 TO NumOfCatRecs
        GET CatFile, SCnt!, ARCatCodeRec(1)
        ARCode$ = QPTrim$(ARCatCodeRec(1).CATCODE)
        IF ARCode$ = Code$ THEN
          CatRecord = SCnt!
          EXIT FOR
        END IF
      NEXT SCnt!
      CLOSE CatFile
      'Set New Balance
      NewBalance# = ARCustRec(1).AcctBal - EditPaymentRec(1).Amount
      
      ' Post Transaction Record First
      ARTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
      ARTransRec(1).TransDate = EditPaymentRec(1).TranDate
      ARTransRec(1).TransAmount = EditPaymentRec(1).Amount
      ARTransRec(1).TransType = 2               ' Type 2 = Payment
      ARTransRec(1).TransDesc = "Payment"
      ARTransRec(1).CashAmount = EditPaymentRec(1).CashAmt
      ARTransRec(1).ChkAmount = EditPaymentRec(1).ChkAmt
      ARTransRec(1).BalanceAfterTrans = NewBalance#
      ARTransRec(1).ExtraRoom = ""
      ARTransRec(1).Posted2GL = "N"
      ARTransRec(1).CatCodeRec1 = CatRecord
      
      ARTransRec(1).NextTrans = 0
      PUT ARTransFile, NextTransRec, ARTransRec(1)
      
      'Update Customer
      CustRec = VAL(EditPaymentRec(1).CustNumber)
      GET ARFile, CustRec, ARCustRec(1)
      ARCustRec(1).IssueLicense = EditPaymentRec(1).ISSUELIC
      ARCustRec(1).AcctBal = Round(ARCustRec(1).AcctBal - EditPaymentRec(1).Amount)
      ARCustRec(1).LicBal = Round(ARCustRec(1).LicBal - EditPaymentRec(1).LICPAID)
      ARCustRec(1).PenBal = Round(ARCustRec(1).PenBal - EditPaymentRec(1).PENPAID)
      
      ARCustRec(1).IssuanceFee = EditPaymentRec(1).ISSueFEE
      
      IF EditPaymentRec(1).SETFEE = "Y" THEN
        ARCustRec(1).FeeAmt = EditPaymentRec(1).Amount
      END IF
      
      PUT ARFile, CustRec, ARCustRec(1)
      
      IF ARCustRec(1).FirstTrans = 0 THEN
        ARCustRec(1).FirstTrans = NextTransRec
        ARCustRec(1).LastTrans = NextTransRec
        PUT ARFile, CustRec, ARCustRec(1)
      ELSE
        Prev& = ARCustRec(1).LastTrans
        ARCustRec(1).LastTrans = NextTransRec
        PUT ARFile, CustRec, ARCustRec(1)
        GET ARTransFile, Prev&, ARTransRec(1)
        ARTransRec(1).NextTrans = NextTransRec
        PUT ARTransFile, Prev&, ARTransRec(1)
      END IF
      NextTransRec = NextTransRec + 1
    END IF
    
  LOOP UNTIL Cnt > NumOfArRecs
  CLOSE

  KILL PayFileName$
  
END SUB

