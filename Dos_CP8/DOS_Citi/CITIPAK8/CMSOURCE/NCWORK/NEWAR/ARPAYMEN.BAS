DEFINT A-Z
DECLARE SUB DisplayCMScrn (ScrnName$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ClearBack ()
DECLARE SUB CursorOff ()
DECLARE SUB PostPayments ()
DECLARE SUB PrintEditList ()
DECLARE SUB EditPayment ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
DECLARE SUB OpenARCustFile (NumOfArRecs, ARFile)
DECLARE SUB EnterPayment (OperRecNum%, PostDate$)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  ''$INCLUDE: 'GL.bi'
  '$INCLUDE: 'ubsetup.BI'
  '$INCLUDE: 'CMFILES.BI'

  CONST False = 0, True = NOT False
  
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
  DIM SHARED EditPaymentRec(1) AS AREditPaymentRecType
  
  '  GOTO StartHere
  
  STACK 8000
  
  'StartHere:
  '  CustLen = LEN(ARCustRec(1))
  '  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = CustLen
  '  FOR zz = 1 TO LOF(1) / CustLen
  '    GET #1, zz, ARCustRec(1)
  '    ARCustRec(1).LicBal = ARCustRec(1).AcctBal
  '    ARCustRec(1).PenBal = 0
  '    PUT #1, zz, ARCustRec(1)
  '  NEXT
  '  CLOSE

SUB DisplayARScrn (ScrnName$)
  LibFile2Scrn "AR", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  IF ErrorCode% <> 0 THEN
    PRINT "Screen Error: "; ScrnName$
    END
  END IF
END SUB

SUB EnterPayment (OperRecNum%, PostDate$)
  SHARED Choice$()
  
  REDIM TempScrn(0)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  TownName$ = UBSetUpRec(1).UTILNAME
  
MainBody:
  REDIM CMOperRec(1) AS CMOperRecType
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0
  
  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  help$ = "Enter Payments"
  CursorOff
'  PrintHelp help$

  GOSUB SetOperatorName

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
  
  'DisplayARScrn "BAKCLEAR"
  DisplayARScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).FldNo = TenderFld AND Frm(1).PrevFld <> TenderFld THEN
      GOSUB FixCashChkFlds
    ELSEIF Frm(1).PrevFld = TenderFld AND Frm(1).FldNo <> TenderFld THEN
      GOSUB FixCashChkFlds
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      IF Frm(1).PrevFld = ChkAmtFld OR Frm(1).PrevFld = CashAmtFld THEN
        SaveField ChkAmtFld, Form$(), Fld(), BadField
        SaveField CashAmtFld, Form$(), Fld(), BadField
        CalcFields 0, ChkAmtFld + 1, Form$(), Fld()
        'PrintArray 1, NumFlds - 1, Form$(), Fld()
        AmtOwed# = Value#(Form$(9, 0), ECode)
        AMTPAID# = Value#(Form$(13, 0), ECode)
        
        IF AMTPAID# > AmtOwed# THEN
          CHANGE# = ABS(AmtOwed# - AMTPAID#)
        ELSE
          CHANGE# = 0
        END IF
        LSET Form$(14, 0) = STR$(CHANGE#)
        Action = 2
      END IF
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
      OPEN "AREDPYT.DAT" FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
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

        ok = MsgBox("AR", "UPDATEOK")
        done = True
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
    IF (ARCustRec(1).AcctBal <= 0) THEN
      GOSUB NoBalance
      GOTO MainBody
    END IF
    
    Form$(3, 0) = ARCustRec(1).CUSTNUMB
    Form$(4, 0) = ARCustRec(1).CustName
    Form$(5, 0) = ARCustRec(1).Address1
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
    PrintHelp help$
    
    done = False
    Action = 1
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        done = True
      END SELECT
      IF done = True THEN GOTO MainBody
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
  EditPaymentRec(1).CustNumber = Form$(3, 0)
  EditPaymentRec(1).CustName = Form$(4, 0)
  EditPaymentRec(1).ADD1 = Form$(5, 0)
  EditPaymentRec(1).CITY = Form$(6, 0)
  EditPaymentRec(1).STATE = Form$(7, 0)
  EditPaymentRec(1).ZIPCODE = Form$(8, 0)
  
  AmtOwed# = Value#(Form$(9, 0), ECode)         'total owed
  EditPaymentRec(1).CashChk = Form$(10, 0)
  EditPaymentRec(1).CashAmt = Value(Form$(11, 0), A%)           'Cash Amount
  EditPaymentRec(1).ChkAmt = Value(Form$(12, 0), A%)            'Cash Amount
  
  AMTPAID# = Value(Form$(13, 0), A%)            'amt tendered
  
  IF AMTPAID# <= 0 THEN
    BadPay = True
    BadField = 10
    GOTO CheckReturn
  ELSE
    EditPaymentRec(1).AMTPAID = AMTPAID#
  END IF
  
  IF AMTPAID# > AmtOwed# THEN
    CHANGE# = ABS(AmtOwed# - AMTPAID#)
  ELSE
    CHANGE# = 0
  END IF
  
  EditPaymentRec(1).CHANGE = CHANGE#
  IssLic$ = QPTrim$(Form$(15, 0))
  IF IssLic$ = "" THEN
    BadPay = True
    BadField = 12
    GOTO CheckReturn
  END IF
  EditPaymentRec(1).ISSUELIC = Form$(15, 0)     'issue lic flag
  EditPaymentRec(1).DESC = Form$(18, 0)
  
  EditPaymentRec(1).LICDUE = Value(Form$(19, 0), A%)            'lic due
  EditPaymentRec(1).LICPAID = Value(Form$(22, 0), A%)           'amt to lic
  
  IF EditPaymentRec(1).LICPAID > EditPaymentRec(1).LICDUE THEN
    BadPay = True
    BadField = 19
    GOTO CheckReturn
  END IF
  
  EditPaymentRec(1).PENPAID = Value(Form$(24, 0), A%)           'amt to penalty
  EditPaymentRec(1).PENDUE = Value(Form$(21, 0), A%)            'pen due
  
  IF EditPaymentRec(1).PENPAID > EditPaymentRec(1).PENDUE THEN
    BadPay = True
    BadField = 21
    GOTO CheckReturn
  END IF
  
  EditPaymentRec(1).TOTDUE = Value(Form$(25, 0), A%)            'sum of (due)
  EditPaymentRec(1).TOTPAID = Value(Form$(26, 0), A%)           'sum of (paid)
  
  TOTPAID# = Round#(AMTPAID# - CHANGE#)
  
  IF TOTPAID# <> EditPaymentRec(1).TOTPAID THEN
    BadPay = True
    BadField = 19
    GOTO CheckReturn
  END IF
  
  EditPaymentRec(1).Amount = EditPaymentRec(1).TOTDUE           'Transaction Amount
  
  CLOSE

'dale

  GOSUB PrintReceipt

  GOSUB PostPayment

  
CheckReturn:
  
RETURN
  
FixCashChkFlds:
  
  SELECT CASE QPTrim$(Form$(TenderFld, 0))
  CASE Choice$(1, 0)            'CASH          this sets the cash or check
    Fld(TenderFld + 1).Protected = False       'amount fields protected or
    Fld(TenderFld + 2).Protected = True        'unprotected based of the
    LSET Form$(ChkAmtFld, 0) = "0"
  CASE Choice$(2, 0)            'CHECK         tender type field selection
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

SetOperatorName:
  Action = 1
  OperRecNumber = OperRecNum%
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  OperName$ = LEFT$(CMOperRec(1).OperatorName, 18)
  Operator = CMOperRec(1).OperatorNumber
  Operator$ = STR$(Operator)
  Operator$ = RIGHT$(Operator$, LEN(Operator$) - 1)
  CLOSE CMFile
RETURN

PrintReceipt:
'STOP
  REDIM TmpScrn(0)
  SaveScrn TmpScrn()
  help$ = "Citipak Cash Management System v8.0"
  ShowCursor
  ScrnName$ = "CMPOST"

  ' Define Fields
  NumFlds = LibNumberOfFields("CM", ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef "CM", ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F

  DisplayCMScrn "CMPOST"

  'PrintTitle User$
  'PrintHelp help$

  Form$(1, 0) = "Y"
  Action = 1

PrtReceipt:
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        GOSUB PrintReceiptFunction
        RestScrn TmpScrn()
        done = True
        'Form$(1, 0) = "N"
        'Frm(1).FldNo = 1
        'Action = 1
        'GOTO PrtReceipt
      END IF
      Posted = True
      done = True
    CASE ESC
      Posted = False
      done = True
    CASE ELSE
      done = False
    END SELECT
  LOOP UNTIL done
RETURN

PrintReceiptFunction:
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CHandle LEN = CMTrRecLen
  IF LOF(CHandle) > 0 THEN
    CMRecNumb& = (LOF(CHandle) \ CMTrRecLen) + 1
  ELSE
    CMRecNumb& = 1
  END IF
  CLOSE CHandle

  'Set Printer Port for Receipts Printer if none set to 1

  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))

  OPEN "O", 15, "RECPT.PRN": WIDTH #15, 255

  PRINT #15, ""
  PRINT #15, TownName$
  PRINT #15, "LICENSE PAYMENT"
  PRINT #15, "Date: "; PostDate$
  PRINT #15,
  PRINT #15, "Account #"; Acct$
  PRINT #15, QPTrim$(ARCustRec(1).CustName)
  PRINT #15, QPTrim$(ARCustRec(1).Address1)
  PRINT #15,
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; EditPaymentRec(1).TOTDUE
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; EditPaymentRec(1).CashAmt
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; EditPaymentRec(1).ChkAmt
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; EditPaymentRec(1).TOTPAID
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; EditPaymentRec(1).CHANGE
  PRINT #15,
  PRINT #15,
  PRINT #15, "Operator: "; OperRecNum%
  PRINT #15, "Receipt#: "; USING "######"; CMRecNumb&
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

PostPayment:

  REDIM TempScrn(0)
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  'REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  CatCodeRecLen = LEN(ARCatCodeRec(1))

  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
  OpenARCustFile NumOfArRecs, ARFile

  AREditRecLen = LEN(EditPaymentRec(1))

  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
  NextTransRec = NumOfTransRecs + 1

  IF EditPaymentRec(1).Amount <> 0 THEN
    GET ARFile, AccountRecord, ARCustRec(1)
    help$ = "Posting: " + LEFT$(ARCustRec(1).BILLNAME, 30)
    'PrintHelp help$
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

    NewBalance# = ARCustRec(1).AcctBal - EditPaymentRec(1).TOTPAID

    ' Post Transaction Record First
    ARTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
    ARTransRec(1).TransDate = EditPaymentRec(1).TranDate
    ARTransRec(1).TransAmount = EditPaymentRec(1).Amount
    ARTransRec(1).TransType = 2               ' Type 2 = Payment
    ARTransRec(1).TransDesc = "Payment (CM)"
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

  CLOSE

  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CMTRRec(1).TransDate = EditPaymentRec(1).TranDate

  CMTRRec(1).TransAmount = AMTPAID# 'EditPaymentRec(1).Amount
  CMTRRec(1).TransCash = EditPaymentRec(1).CashAmt
  CMTRRec(1).TransCheck = EditPaymentRec(1).ChkAmt
  CMTRRec(1).TransAmtOwed = EditPaymentRec(1).TOTDUE
  CMTRRec(1).TransDesc = "AR-PAYMENT"
  CMTRRec(1).TransSource = 41
  CMTRRec(1).TransName = CustName$
  CMTRRec(1).TransAcctNum = AccountRecord
  CMTRRec(1).TransDetNum = DetailRecord!
  CMTRRec(1).TransOperNum = Operator
  CMTRRec(1).TransPad = ""
  FldFactor = 0
  FOR Cnt = 1 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = 0
  NEXT Cnt
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CHandle LEN = CMTrRecLen
  CMRecNumb& = (LOF(CHandle) / CMTrRecLen) + 1
  PUT CHandle, CMRecNumb&, CMTRRec(1)
  CLOSE CHandle


  ' Show All Posted
  BEEP
  'SaveScrn TempScrn()
  'DisplayARScrn "ARPOSTED"
  'WaitForAction
  'RestScrn TempScrn()
  'LOCATE , , 1
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

