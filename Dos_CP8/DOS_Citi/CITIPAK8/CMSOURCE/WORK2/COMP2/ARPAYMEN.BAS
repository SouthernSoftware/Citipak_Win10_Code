
DEFINT A-Z
DECLARE SUB PostPayments ()
DECLARE SUB PrintEditList ()
DECLARE SUB EditPayment ()
DECLARE SUB OpenARCustIdxFile (NumOfArIdxRecs, ARIdxFile)
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
DECLARE FUNCTION FUsing$ (Number$, Format$)
'DECLARE SUB WaitForAction
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
DECLARE FUNCTION GetDEFPort% ()
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB WaitForAction ()

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  '$INCLUDE: 'CMFILES.BI'
  '$INCLUDE: 'ubsetup.BI'

   CONST False = 0, True = NOT False
   DIM SHARED ARCustRec(1) AS ARCustRecType
   DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
   DIM SHARED EditPaymentRec(1) AS AREditPaymentRecType


 STACK 8000

   '--Dim the choice array to the number of menu items
   REDIM Mchoice$(1 TO 5)

   Mchoice$(1) = " Enter Payment/Transaction  "
   Mchoice$(2) = " Edit Payment/Transaction "
   Mchoice$(3) = " Print Edit List "
   Mchoice$(4) = " Post Payments "
   Mchoice$(5) = " Exit to OS "

   MaxLen = 0     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

SUB DisplayARScrn (ScrnName$)
 LibFile2Scrn "AR", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB EnterPayment (OperRecNum%, PostDate$)
  SHARED Choice$()
  REDIM TempScrn(0)

mainbody:
  CLOSE
  REDIM CMOperRec(1) AS CMOperRecType
  REDIM UBSetUpRec(1) AS UBSetupRecType

  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0

  LibName$ = "CM"
  ScrnName$ = "CMBLPAY"
  help$ = "Enter Payments"
  LOCATE 1, 1, 0

  GOSUB SetOperatorName


  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  'Rem Set Choice
  REDIM Choice$(3, 0)
  Choice$(0, 0) = "8"
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F

  Form$(1, 0) = OperName$
  Form$(2, 0) = PostDate$
  Fld(1).Protected = True
  Fld(2).Protected = True

  Frm(1).StayOnField = -1
  Action = 1

  OpenARCustIdxFile NumOfArIdxRecs, ARIdxFile
  OpenARCustFile NumOfArRecs, ARFile


  'Trap for No Customers Here
  IF NumOfArIdxRecs = 0 THEN
           SaveScrn TempScrn()
           LOCATE , , 0
           PRINT CHR$(7);
           DisplayARScrn "ARNOCUST"
           WaitForAction
           RestScrn TempScrn()
           LOCATE , , 1
           CLOSE
           EXIT SUB
  END IF

  'Set Field Defaults
  CustAcctFld = 3
  OwedAmtFld = 7

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    'This Gets the Customer
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
     LSET Form$(7, 0) = "0.00"
     LSET Form$(8, 0) = ""
     LSET Form$(9, 0) = "0.00"
     LSET Form$(10, 0) = "0.00"
     LSET Form$(11, 0) = "0.00"
     LSET Form$(12, 0) = "0.00"


      GOSUB GetCustomer
      IF AccountRecord = 0 THEN
        CLOSE : Action = 1: GOTO mainbody
      END IF
    Action = 1
   END IF

   'This Makes Sure that the Amount Owed Field is Not Equal than Zero
      IF Frm(1).FldNo > OwedAmtFld AND Frm(1).PrevFld = OwedAmtFld THEN
       IF Value(Form$(7, 0), ecode) = 0 THEN Frm(1).FldNo = 7
       Action = 1
      END IF


   'Now Get Money Tendered

    IF LEFT$(Form$(8, 0), 6) = "Cash  " THEN
     Fld(10).Protected = True
     LSET Form$(10, 0) = "0.00"
     Fld(9).Protected = False
     IF PolledIt = 0 THEN
      Action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(8, 0), 2)
     END IF
    END IF

    IF LEFT$(Form$(8, 0), 2) = "Ch" THEN
     Fld(9).Protected = True
     LSET Form$(9, 0) = "0.00"
     Fld(10).Protected = False
     IF PolledIt = 0 THEN
      Action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(8, 0), 2)
     END IF
    END IF

    IF LEFT$(Form$(8, 0), 6) = "Cash &" THEN
     Fld(9).Protected = False
     Fld(10).Protected = False
     IF PolledIt = 0 THEN
      Action = 1: PolledIt = 1: PayHow$ = LEFT$(Form$(8, 0), 2)
     END IF
    END IF

    IF PolledIt = 1 AND LEFT$(Form$(8, 0), 2) <> PayHow$ THEN
     PolledIt = 0
     Action = 1
    END IF

    IF Frm(1).PrevFld = 8 AND Frm(1).FldEdited = -1 THEN
     LSET Form$(9, 0) = "0.00"
     LSET Form$(10, 0) = "0.00"
     LSET Form$(11, 0) = "0.00"
     LSET Form$(12, 0) = "0.00"
     Action = 1
    END IF
   
    'Calculate and Display the Change Due
    IF (Frm(1).FldNo <> 7) AND Frm(1).PrevFld = 7 THEN
     TotalRec# = Value(Form$(9, 0), ecode) + Value(Form$(10, 0), ecode)
     LSET Form$(11, 0) = STR$(TotalRec#)
     ChangeDue# = TotalRec# - Value(Form$(7, 0), ecode)
     IF ChangeDue# < 0 THEN ChangeDue# = 0
     LSET Form$(12, 0) = STR$(ChangeDue#)
     Action = 1
     CalcCode = 1
    END IF
    IF (Frm(1).FldNo <> 10) AND Frm(1).PrevFld = 10 THEN
     TotalRec# = Value(Form$(9, 0), ecode) + Value(Form$(10, 0), ecode)
     LSET Form$(11, 0) = STR$(TotalRec#)
     ChangeDue# = TotalRec# - Value(Form$(7, 0), ecode)
     IF ChangeDue# < 0 THEN ChangeDue# = 0
     LSET Form$(12, 0) = STR$(ChangeDue#)
     Action = 1
     CalcCode = 1
    END IF

    IF Value(Form$(7, 0), ecode) <> AmountOwed# THEN
     CalcCode = 0
     AmountOwed# = Value(Form$(7, 0), ecode)
    END IF

    IF Value(Form$(9, 0), ecode) <> CashIn# THEN
     CalcCode = 0
     CashIn# = Value(Form$(9, 0), ecode)
    END IF

    IF Value(Form$(10, 0), ecode) <> ChkIn# THEN
     CalcCode = 0
     ChkIn# = Value(Form$(10, 0), ecode)
    END IF
    
    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      IF Value(Form$(7, 0), ecode) <> 0 THEN
        Acct$ = Form$(3, 0)
        CustName$ = Form$(4, 0)
        Address$ = Form$(5, 0)
        Desc$ = Form$(6, 0)
        TotalOwed# = Value(Form$(7, 0), ecode%)
        TotalCash# = Value(Form$(9, 0), ecode%)
        TotalCheck# = Value(Form$(10, 0), ecode%)
        TotalPaid# = Value(Form$(11, 0), ecode%)
        TotalChange# = Value(Form$(12, 0), ecode%)
        SetLicense$ = Form$(13, 0)
        SetFee$ = Form$(14, 0)
        GOSUB PrintReceipt
        GOSUB PostTransaction
      END IF
      CLOSE
      Done = True
      GOTO mainbody
    CASE EscKey
      EXIT SUB

    END SELECT

  LOOP

GetCustomer:
  CustomerGrabed = 0
  AccountRecord = VAL(Form$(3, 0))

IF AccountRecord = 0 THEN

   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = " Cust #    Customer Sort Name"

   '--Center Menu within Screen
   Row = 8
   Col = 15

   REDIM Mchoice$(1 TO NumOfArIdxRecs)

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfArIdxRecs
    GET ARIdxFile, Cnt, ARCustIdxRec(1)
    IF LEFT$(ARCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
     ChoiceCounter = ChoiceCounter + 1
     Mchoice$(ChoiceCounter) = SPACE$(50)
     LSET Mchoice$(ChoiceCounter) = ARCustIdxRec(1).IDXNAME
     MID$(Mchoice$(ChoiceCounter), 40) = STR$(ARCustIdxRec(1).IDXRECORD)
    END IF
  NEXT Cnt

   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        AccountRecord = 0
        ExitFlag = True
      ELSE
        AccountRecord = VAL(RIGHT$(Mchoice$(Choice), 10))
        ExitFlag = True
      END IF

   LOOP UNTIL ExitFlag

  LibName$ = "CM"
  ScrnName$ = "CMBLPAY"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%

  END IF

REM ************************************************************************
IF AccountRecord > 0 AND AccountRecord <= NumOfArRecs THEN
    GET ARFile, AccountRecord, ARCustRec(1)
    IF ARCustRec(1).Deleted = "Y" THEN
     GOSUB CustomerDeleted
     GOTO mainbody
    END IF
    Form$(3, 0) = ARCustRec(1).CUSTNUMB
    Form$(4, 0) = ARCustRec(1).CustName
    Form$(5, 0) = ARCustRec(1).ADDRESS1
    Form$(7, 0) = STR$(ARCustRec(1).AcctBal)
    Form$(13, 0) = "N"
    Form$(14, 0) = "N"
    Frm(1).FldNo = 6
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    RETURN

   ELSE

    LibName$ = "AR"
    ScrnName$ = "ARBADCUS"
    help$ = "Set Customer Balances"
    LOCATE 1, 1, 0

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

    PRINT CHR$(7);

    ShowCursor
    LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    PrintHelp help$

    Done = False
    Action = 1

  DO

   EditForm Form$(), Fld(), Frm(1), Cnf, Action

   SELECT CASE Frm(1).KeyCode
    CASE EscKey
     Done = True
     END SELECT
     IF Done = True THEN GOTO mainbody
  LOOP

 END IF

CustomerDeleted:
    LibName$ = "AR"
    ScrnName$ = "ARDELCUS"
    help$ = "Payment Entry"
    LOCATE 1, 1, 0

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

    PRINT CHR$(7);

    ShowCursor
    LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    PrintHelp help$

    Done = False
    Action = 1


  DO

   EditForm Form$(), Fld(), Frm(1), Cnf, Action

   SELECT CASE Frm(1).KeyCode
    CASE F10Key
     RETURN
    END SELECT
  LOOP

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
   help$ = "Citipak Cash Management System v8.0"
  ShowCursor
  LibName$ = "CM.QSL"
  ScrnName$ = "CMPOST"

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

  LibFile2Scrn "CM.QSL", "CMPOST", MonoCode%, Attribute%, ErrorCode%
  PrintTitle User$
  PrintHelp help$

  Form$(1, 0) = "Y"
                   Action = 1
PrtReceipt:
    REDIM TmpScrn(0)
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        SaveScrn TmpScrn()
        GOSUB PrintReceiptFunction
        RestScrn TmpScrn()
        Form$(1, 0) = "N"
        Frm(1).FldNo = 1
        Action = 1
        GOTO PrtReceipt
      END IF
      Posted = True
      Done = True
    CASE ESC
      Posted = False
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN

PrintReceiptFunction:
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  RecNumber = (LOF(CHandle) \ CMTrRecLen) + 1
  CLOSE CHandle

  
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  TOWNNAME$ = UBSetUpRec(1).UTILNAME
  'Set Printer Port for Receipts Printer if none set to 1

  IF INSTR(UBSetUpRec(1).UTILNAME, "SPENCER") THEN
    SpencerFlag = True
  END IF
  Work$ = SPACE$(19)
  
  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))

  OPEN "RECPT.PRN" FOR OUTPUT SHARED AS #15
  WIDTH #15, 255

  PRINT #15, ""
  PRINT #15, TOWNNAME$
  PRINT #15, "LICENSE PAYMENT"
  PRINT #15, "Date: "; PostDate$
  PRINT #15,
  PRINT #15, "Account #"; Acct$
  PRINT #15, CustName$
  PRINT #15, Address$
  PRINT #15, Desc$
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; TotalOwed#
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; TotalCash#
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; TotalCheck#
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; TotalPaid#
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; TotalChange#
  PRINT #15,
'put spencer trap here
  IF SpencerFlag THEN
    PRINT #15, CHR$(27); "="; CHR$(2)   'only to display
    LSET Work$ = "THANK YOU!"
    PRINT #15, Work$
    LSET Work$ = "Change:" + FUsing$(STR$(TotalChange#), "$$####.##")
    PRINT #15, Work$; CHR$(10)
    PRINT #15, CHR$(27); "="; CHR$(1)  'only to printer
  END IF
''
  PRINT #15,
  PRINT #15, "Operator: "; OperRecNum%
  PRINT #15, "Receipt#: "; USING "######"; RecNumber
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
   
PostTransaction:
IF Posted <> True THEN RETURN
  PostDate = Date2Num%(PostDate$)

  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  CMTRRec(1).TransDate = PostDate
  CMTRRec(1).TransAmount = TotalPaid#
  CMTRRec(1).TransCash = TotalCash#
  CMTRRec(1).TransCheck = TotalCheck#
  CMTRRec(1).TransAmtOwed = TotalOwed#
  CMTRRec(1).TransDesc = Desc$
  CMTRRec(1).TransSource = 41
  CMTRRec(1).TransName = CustName$
  CMTRRec(1).TransAcctNum = AcctRecord
  CMTRRec(1).TransDetNum = DetailRecord!
  CMTRRec(1).TransOperNum = Operator
  CMTRRec(1).TransPad = ""
  FldFactor = 0
  FOR Cnt = 1 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = 0
  NEXT Cnt
  PUT CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTRRec(1)
  CLOSE CHandle

  ' Now Post to BL System
 OpenARCustFile NumOfArRecs, ARFile

 REDIM ARTransRec(1) AS ARTransRecType
 ARTransRecLen = LEN(ARTransRec(1))
 ARTransFile = FREEFILE
 OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
 NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
 NextTransRec = NumOfTransRecs + 1

 GET ARFile, AccountRecord, ARCustRec(1)
 help$ = "Posting: " + LEFT$(ARCustRec(1).BILLNAME, 30)
 PrintHelp help$

 NewBalance# = ARCustRec(1).AcctBal - EditPaymentRec(1).Amount

   ' Post Transaction Record First
  ARTransRec(1).CustomerNumber = Acct$
  ARTransRec(1).TransDate = Date2Num%(PostDate$)
  ARTransRec(1).TransAmount = TotalPaid#
  ARTransRec(1).TransType = 2               ' Type 2 = Payment
  ARTransRec(1).TransDesc = "Payment " + Desc$
  ARTransRec(1).CashAmount = TotalCash#
  ARTransRec(1).ChkAmount = TotalCheck#
  ARTransRec(1).BalanceAfterTrans = NewBalance#
  ARTransRec(1).ExtraRoom = ""
  ARTransRec(1).NextTrans = 0
  PUT ARTransFile, NextTransRec, ARTransRec(1)

  GET ARFile, AccountRecord, ARCustRec(1)
   ARCustRec(1).IssueLicense = SetLicense$
   ARCustRec(1).AcctBal = ARCustRec(1).AcctBal - TotalPaid#

  IF SetFee$ = "Y" THEN
   ARCustRec(1).FeeAmt = TotalOwed#
  END IF

  PUT ARFile, AccountRecord, ARCustRec(1)

  IF ARCustRec(1).FirstTrans = 0 THEN
   ARCustRec(1).FirstTrans = NextTransRec
   ARCustRec(1).LastTrans = NextTransRec
   PUT ARFile, AccountRecord, ARCustRec(1)
  ELSE
   Prev! = ARCustRec(1).LastTrans
   ARCustRec(1).LastTrans = NextTransRec
   PUT ARFile, AccountRecord, ARCustRec(1)
   GET ARTransFile, Prev!, ARTransRec(1)
   ARTransRec(1).NextTrans = NextTransRec
   PUT ARTransFile, Prev!, ARTransRec(1)
  END IF

  CLOSE

  RETURN

END SUB

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
END SUB

SUB OpenARCustIdxFile (NumOfArIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfArIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

