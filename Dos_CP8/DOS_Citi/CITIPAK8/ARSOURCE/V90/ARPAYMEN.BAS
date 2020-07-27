DEFINT A-Z
DECLARE SUB OperatorEntry (OperOKFlag%, OperNum%, PostDate$)
DECLARE FUNCTION AskSavePayment% (DefaultFlag$, OperNum%, NextRec%)
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
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
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
  
  CONST False = 0, True = NOT False
  
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
  DIM SHARED EditPaymentRec(1) AS AREditPaymentRecType
  
  '  GOTO StartHere
  
  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM Mchoice$(1 TO 5)
  '
  Mchoice$(1) = " Enter Payment/Transaction  "
  Mchoice$(2) = " Edit Payment/Transaction "
  Mchoice$(3) = " Print Edit List "
  Mchoice$(4) = " Post Payments "
  Mchoice$(5) = " Exit to OS "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(Mchoice$)
    TLen = LEN(Mchoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(Mchoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1
  
  'help$ = "Accounts Receivable Payment Menu"
  help$ = "Payment Menu"
  
  IF NOT WEnvTest% OR COMMAND$ = "TEST" THEN
    DO
      OperatorEntry OperOKFlag, OperNum, PostDate$
    LOOP WHILE OperOKFlag = 0
  ELSE
    OperNum = 1
    PostDate$ = DATE$
  END IF

  IF OperOKFlag < 0 THEN GOTO ExitPaymentProg


  ShowCursor
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ClearBack
    
    TitleBox 3, Col, MaxLen + 3, "Payment Entry/Edit Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    PrintTitle User$
    PrintHelp help$
    
    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      EnterPayment OperNum, PostDate$
    CASE 2
      EditPayment OperNum, PostDate$
    CASE 3
      PrintEditList OperNum
    CASE 4
      PostPayments OperNum
    CASE 5
      HideCursor
      CLS
      END
    END SELECT
  LOOP

ExitPaymentProg:
  IF INSTR(COMMAND$, "TEST") <= 0 THEN
    RUN "armenu"
  END IF
  
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

FUNCTION AskSavePayment% (DefaultFlag$, OperNum, NextRec)

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  LibName$ = "AR"
  ScrnName$ = "ASKR2PST"

  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  DisplayARScrn ScrnName$

  ShowCursor

  Action = 1
  FirstTime = True

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = NOT FirstTime
      LSET Form$(1, 0) = DefaultFlag$
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Receipt$ = Form$(1, 0)
      IF Receipt$ = "Y" THEN
        PrintReceipt OperNum, NextRec
      END IF
      EXIT DO
    END SELECT

    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14
        SELECT CASE Frm(1).MCol
        CASE 36 TO 45           'F10 Save Button
          PressButton -68, 14, 36, 45
        END SELECT
      END SELECT
    END IF
  LOOP

  RestScrn TempScrn()

END FUNCTION

SUB EditPayment (OperNum, PostDate$)
  
  SHARED Choice$()

  Oper$ = QPTrim$(STR$(OperNum))
  PayFileName$ = "AREDPY" + Oper$ + ".DAT"
  
  REDIM Choice$(0 TO 3, 0 TO 0)
  
  Choice$(0, 0) = "10"
  Choice$(1, 0) = "Cash"
  Choice$(2, 0) = "Check"
  Choice$(3, 0) = "Cash & Check"
  
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0
  
  TenderFld = 10
  CashAmtFld = 11
  ChkAmtFld = 12
  
Continue:
  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  help$ = "Edit Customer Balance Entry's"
  CursorOff
  
  PrintHelp help$
  
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
  
  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
  OpenARCustFile NumOfArRecs, ARFile
  
  Form$(1, 0) = "Payment"
  IF AccountRecord = 0 THEN
    GOSUB GetEditRecord
  END IF
  
  IF AccountRecord = 0 THEN
    CLOSE
    EXIT SUB
  END IF
  
  GET AREdFile, AccountRecord, EditPaymentRec(1)
  GET ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
  
  Form$(2, 0) = Num2Date$(EditPaymentRec(1).TranDate)
  Form$(3, 0) = EditPaymentRec(1).CustNumber
  Form$(4, 0) = EditPaymentRec(1).CustName
  Form$(5, 0) = EditPaymentRec(1).Add1
  Form$(6, 0) = EditPaymentRec(1).CITY
  Form$(7, 0) = EditPaymentRec(1).STATE
  Form$(8, 0) = EditPaymentRec(1).ZIPCODE
  Form$(9, 0) = STR$(EditPaymentRec(1).Amount)
  Form$(10, 0) = EditPaymentRec(1).CashChk
  Form$(11, 0) = STR$(EditPaymentRec(1).CashAmt)
  Form$(12, 0) = STR$(EditPaymentRec(1).ChkAmt)
  
  Form$(13, 0) = STR$(EditPaymentRec(1).AmtPaid)
  Form$(14, 0) = STR$(EditPaymentRec(1).Change)
  Form$(15, 0) = EditPaymentRec(1).ISSUELIC
  
  'Form$(13, 0) = EditPaymentRec(1).SETFEE
  'Form$(17, 0) = STR$(EditPaymentRec(1).ISSueFEE)
  Form$(18, 0) = EditPaymentRec(1).Desc
  Form$(19, 0) = STR$(EditPaymentRec(1).LICDUE)
  'Form$(17, 0) = STR$(EditPaymentRec(1).FEEDUE)
  Form$(21, 0) = STR$(EditPaymentRec(1).PENDUE)
  Form$(22, 0) = STR$(EditPaymentRec(1).LicPAID)
  'Form$(20, 0) = STR$(EditPaymentRec(1).FEEPAID)
  Form$(24, 0) = STR$(EditPaymentRec(1).PenPAID)
  Form$(25, 0) = STR$(EditPaymentRec(1).TotDue)
  Form$(26, 0) = STR$(EditPaymentRec(1).TotPaid)
  
  '  Form$(16, 0) = FUsing$(STR$(ARCustRec(1).AcctBal), "#####.##")
  
  Fld(1).Protected = True
  
  FOR Fld = 3 TO 9
    Fld(Fld).Protected = True
  NEXT Fld
  
  DisplayARScrn "BAKCLEAR"
  DisplayARScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    '***************************************************
    IF Frm(1).FldNo = TenderFld AND Frm(1).PrevFld <> TenderFld THEN
      GOSUB EDFixCashChkFlds
    ELSEIF Frm(1).PrevFld = TenderFld AND Frm(1).FldNo <> TenderFld THEN
      GOSUB EDFixCashChkFlds
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      SELECT CASE Frm(1).PrevFld
      CASE ChkAmtFld, CashAmtFld, 22, 24
      'IF Frm(1).PrevFld = ChkAmtFld OR Frm(1).PrevFld = CashAmtFld THEN
        SaveField ChkAmtFld, Form$(), Fld(), BadField
        SaveField CashAmtFld, Form$(), Fld(), BadField
        CashAmt# = Value#(Form$(CashAmtFld, 0), ECode)
        ChekAmt# = Value#(Form$(ChkAmtFld, 0), ECode)
        'CalcFields 0, ChkAmtFld + 1, Form$(), Fld()
        CalcFields 0, 13, Form$(), Fld()
        LSET Form$(13, 0) = STR$(CashAmt# + ChekAmt#)
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
    '************************************************
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F3Key
      GOSUB DeleteRecord
      IF Deleted THEN
        EditPaymentRec(1).CustNumber = ""
        EditPaymentRec(1).Amount = 0
        'ERASE EditPaymentRec
        PUT AREdFile, AccountRecord, EditPaymentRec(1)
        CLOSE AREdFile
        EXIT SUB
      ELSE
        GOTO Continue
      END IF
      
    CASE F10Key
      GOSUB EDCheckSavePayment
      IF BadPay THEN
        ok = MsgBox("AR", "BADPAYMN")
        IF BadField <= 0 THEN
          BadField = 19
        END IF
        Frm(1).FldNo = BadField
        Action = 1
      ELSE
        dodo = AskSavePayment%("Y", OperNum, AccountRecord)
        DisplayARScrn "PRESSKEY"
        WaitForAction
        Done = True
        EXIT SUB
      END IF
      
    CASE EscKey
      EXIT SUB
    CASE -72   'Uparrow

     IF Frm(1).PrevFld = 12 AND Frm(1).FldNo = 11 THEN
       Frm(1).FldNo = 10
       'Action
     END IF
    CASE IS <> 0

    END SELECT
    
  LOOP
  
GetEditRecord:
  
  MaxLen = 50   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = " Cust #    Customer Business Name"
  
  '--Center Menu within Screen
  Row = 8
  Col = 15
  
  AREditRecLen = LEN(EditPaymentRec(1))
  AREdFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS AREdFile LEN = AREditRecLen
  
  NumOfArRecs = LOF(AREdFile) \ AREditRecLen
  
  IF NumOfArRecs = 0 THEN
    CLOSE AREdFile
    RETURN
  END IF
  
  REDIM Mchoice$(1 TO NumOfArRecs)
  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfArRecs
    GET AREdFile, Cnt, EditPaymentRec(1)
    IF EditPaymentRec(1).Amount <> 0 THEN
      ChoiceCounter = ChoiceCounter + 1
      Mchoice$(ChoiceCounter) = SPACE$(50)
      LSET Mchoice$(ChoiceCounter) = EditPaymentRec(1).CustNumber
      MID$(Mchoice$(ChoiceCounter), 10, 30) = EditPaymentRec(1).CustName
      MID$(Mchoice$(ChoiceCounter), 45, 5) = STR$(Cnt)
    END IF
  NEXT Cnt
  
  IF ChoiceCounter = 0 THEN     'No active transactions. Get out.
    GOTO NoneRet
  END IF
  
  DO            '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ClearBack
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      AccountRecord = 0
      ExitFlag = True
    ELSE
      AccountRecord = VAL(RIGHT$(Mchoice$(Choice), 5))
      ExitFlag = True
    END IF
  LOOP UNTIL ExitFlag
NoneRet:
  RETURN
  
DeleteRecord:
  LibName$ = "AR"
  ScrnName$ = "OK2DEL"
'  help$ = "Delete Edit Entry"
  CursorOff
  
  DisplayARScrn ScrnName$
  
'  PrintHelp help$
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
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      EXIT DO
    CASE F10Key
      Deleted = True
      EXIT DO
    END SELECT
    
  LOOP
  
  RETURN
  
EDCheckSavePayment:
  
  BadPay = False                'Assume all is well
  PrevDate$ = Form$(2, 0)       'Keep Default Date Here
  EditPaymentRec(1).TranDate = Date2Num(Form$(2, 0))
  EditPaymentRec(1).CustNumber = Form$(3, 0)
  EditPaymentRec(1).CustName = Form$(4, 0)
  EditPaymentRec(1).Add1 = Form$(5, 0)
  EditPaymentRec(1).CITY = Form$(6, 0)
  EditPaymentRec(1).STATE = Form$(7, 0)
  EditPaymentRec(1).ZIPCODE = Form$(8, 0)
  
  AmtOwed# = Value#(Form$(9, 0), ECode)         'total owed
  
  EditPaymentRec(1).CashChk = Form$(10, 0)
  
  EditPaymentRec(1).CashAmt = Value(Form$(11, 0), A%)           'Cash Amount
  EditPaymentRec(1).ChkAmt = Value(Form$(12, 0), A%)            'Cash Amount
  
  AmtPaid# = Value(Form$(13, 0), A%)            'amt tendered
  TotDist# = Value#(Form$(26, 0), ECode)
  IF TotDist# <= 0 THEN
  'IF AmtPaid# <= 0 THEN
    BadPay = True
    BadField = 10
    GOTO EDCheckReturn
  ELSE
    EditPaymentRec(1).AmtPaid = AmtPaid#
  END IF
  IF AmtPaid# > TotDist# THEN
    Change# = ABS(TotDist# - AmtPaid#)
  ELSE
    Change# = 0
  END IF
  
  'IF AmtPaid# > AmtOwed# THEN
  '  Change# = ABS(Round#(AmtOwed# - AmtPaid#))
  'ELSE
  '  Change# = 0
  'END IF
  
  EditPaymentRec(1).Change = Change#
  IssLic$ = "N"
  IF IssLic$ = "" THEN
    BadPay = True
    BadField = 12
    GOTO EDCheckReturn
  END IF
  EditPaymentRec(1).ISSUELIC = Form$(15, 0)     'issue lic flag
  EditPaymentRec(1).Desc = Form$(18, 0)
  
  EditPaymentRec(1).LICDUE = Value(Form$(19, 0), A%)            'lic due
  EditPaymentRec(1).LicPAID = Value(Form$(22, 0), A%)           'amt to lic
  
  'IF EditPaymentRec(1).LICPAID > EditPaymentRec(1).LICDUE THEN
  '  BadPay = True
  '  BadField = 19
  '  GOTO EDCheckReturn
  'END IF
  
  EditPaymentRec(1).PenPAID = Value(Form$(24, 0), A%)           'amt to penalty
  EditPaymentRec(1).PENDUE = Value(Form$(21, 0), A%)            'pen due
  
  'IF EditPaymentRec(1).PENPAID > EditPaymentRec(1).PENDUE THEN
  '  BadPay = True
  '  BadField = 21
  '  GOTO EDCheckReturn
  'END IF
  
  EditPaymentRec(1).TotDue = Value(Form$(25, 0), A%)            'sum of (due)
  EditPaymentRec(1).TotPaid = Value(Form$(26, 0), A%)           'sum of (paid)
  
  TotPaid# = Round#(AmtPaid# - Change#)
  
  IF TotPaid# <> EditPaymentRec(1).TotPaid THEN
    BadPay = True
    BadField = 19
    GOTO EDCheckReturn
  END IF
  
  EditPaymentRec(1).Amount = TotPaid#  'EditPaymentRec(1).TOTDUE           'Transaction Amount
  
  CLOSE
  AREditRecLen = LEN(EditPaymentRec(1))
  ARFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
  PUT ARFile, AccountRecord, EditPaymentRec(1)
  CLOSE ARFile
  
EDCheckReturn:
  RETURN
  
EDFixCashChkFlds:
  
  SELECT CASE QPTrim$(Form$(TenderFld, 0))
  CASE Choice$(1, 0)            'CASH           this sets the cash or check
    Fld(TenderFld + 1).Protected = False        'amount fields protected or
    Fld(TenderFld + 2).Protected = True         'unprotect based of the
    LSET Form$(ChkAmtFld, 0) = "0"
  CASE Choice$(2, 0)            'CHECK          tender type field selection
    Fld(TenderFld + 1).Protected = True
    Fld(TenderFld + 2).Protected = False
    LSET Form$(CashAmtFld, 0) = "0"
  CASE ELSE     'BOTH
    Fld(TenderFld + 1).Protected = False
    Fld(TenderFld + 2).Protected = False
    '  LSET Form$(ChkAmtFld, 0) = "0"
  END SELECT
  'LSET Form$(13, 0) = "0"'set totalpaid field
  'LSET Form$(14, 0) = STR$(CHANGE#)
  Action = 2
  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField
  SaveField 13, Form$(), Fld(), BadField
  '
  '  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
  '  CalcFields 0, CashAmtFld, Form$(), Fld()
  '
  PrintArray 1, NumFlds - 1, Form$(), Fld()
  RETURN
END SUB

SUB EnterPayment (OperNum, PostDate$)
  SHARED Choice$()
  
  REDIM TempScrn(0)

  Oper$ = QPTrim$(STR$(OperNum))
  PayFileName$ = "AREDPY" + Oper$ + ".DAT"

MainBody:
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0
  
  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  help$ = "Enter Payments"
  CursorOff
  
  PrintHelp help$
  
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

  QPrintRC "                ", 22, 12, 15

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
        dodo = AskSavePayment%("Y", OperNum, NextRec)
        DisplayARScrn "PRESSKEY"
        WaitForAction
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
    
    Form$(3, 0) = ARCustRec(1).CUSTNUMB
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
    PrintHelp help$
    
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
  EditPaymentRec(1).CustNumber = Form$(3, 0)
  EditPaymentRec(1).CustName = Form$(4, 0)
  EditPaymentRec(1).Add1 = Form$(5, 0)
  EditPaymentRec(1).CITY = Form$(6, 0)
  EditPaymentRec(1).STATE = Form$(7, 0)
  EditPaymentRec(1).ZIPCODE = Form$(8, 0)
  
  AmtOwed# = Value#(Form$(9, 0), ECode)         'total owed
  EditPaymentRec(1).CashChk = Form$(10, 0)
  EditPaymentRec(1).CashAmt = Value(Form$(11, 0), A%)           'Cash Amount
  EditPaymentRec(1).ChkAmt = Value(Form$(12, 0), A%)            'Cash Amount
  
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
  EditPaymentRec(1).Desc = Form$(18, 0)
  
  EditPaymentRec(1).LICDUE = Value(Form$(19, 0), A%)            'lic due
  EditPaymentRec(1).LicPAID = Value(Form$(22, 0), A%)           'amt to lic
  
  'IF EditPaymentRec(1).LICPAID > EditPaymentRec(1).LICDUE THEN
  '  BadPay = True
  '  BadField = 19
  '  GOTO CheckReturn
  'END IF
  
  EditPaymentRec(1).PenPAID = Value(Form$(24, 0), A%)           'amt to penalty
  EditPaymentRec(1).PENDUE = Value(Form$(21, 0), A%)            'pen due
  
  'IF EditPaymentRec(1).PENPAID > EditPaymentRec(1).PENDUE THEN
  '  BadPay = True
  '  BadField = 21
  '  GOTO CheckReturn
  'END IF
  
  EditPaymentRec(1).TotDue = Value(Form$(25, 0), A%)            'sum of (due)
  EditPaymentRec(1).TotPaid = Value(Form$(26, 0), A%)           'sum of (paid)
  
  TotPaid# = Round#(AmtPaid# - Change#)
  
  IF TotPaid# <> EditPaymentRec(1).TotPaid THEN
    BadPay = True
    BadField = 19
    GOTO CheckReturn
  END IF
  
  EditPaymentRec(1).Amount = TotPaid# 'EditPaymentRec(1).TOTDUE           'Transaction Amount
  
  CLOSE
  AREditRecLen = LEN(EditPaymentRec(1))
  ARFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
  NextRec = (LOF(ARFile) \ AREditRecLen) + 1
  PUT ARFile, NextRec, EditPaymentRec(1)
  CLOSE ARFile
  
CheckReturn:
  
  RETURN
  
FixCashChkFlds:
  
  SELECT CASE QPTrim$(Form$(TenderFld, 0))
  CASE Choice$(1, 0)            'CASH            this sets the cash or check
    Fld(TenderFld + 1).Protected = False        'amount fields protected or
    Fld(TenderFld + 2).Protected = True         'unprotected based of the
    LSET Form$(ChkAmtFld, 0) = "0"
  CASE Choice$(2, 0)            'CHECK           tender type field selection
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
  PayFileName$ = "AREDPY" + Oper$ + ".DAT"
  
  IF FileSize(PayFileName$) <= 0 THEN
    EXIT SUB
  END IF
  
  REDIM TempScrn(0)
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  'REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  CatCodeRecLen = LEN(ARCatCodeRec(1))
  ClearBack
  
MainPostBody:
  
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0
  
  LibName$ = "AR"
  ScrnName$ = "AROKPOST"
  help$ = "Set Customer Balances"
  CursorOff
  DisplayARScrn ScrnName$
  PrintHelp help$
  
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
  
  Form$(1, 0) = "Y"
  Action = 1
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB PostTrans
      EXIT SUB
    CASE EscKey
      EXIT SUB
      
    END SELECT
    
  LOOP
  
  
PostTrans:
  
  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
  OpenARCustFile NumOfArRecs, ARFile
  
  AREditRecLen = LEN(EditPaymentRec(1))
  AREdFile = FREEFILE
  
  OPEN PayFileName$ FOR RANDOM ACCESS READ WRITE SHARED AS AREdFile LEN = AREditRecLen
  NumOfArRecs = LOF(AREdFile) \ AREditRecLen
  
  ' See if any records to post
  IF NumOfArRecs = 0 THEN
    BEEP
    SaveScrn TempScrn()
    DisplayARScrn "ARNOTRAN"
    WaitForAction
    RestScrn TempScrn()
    LOCATE , , 1
    CLOSE
    EXIT SUB
  END IF
  
  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
  NextTransRec = NumOfTransRecs + 1
  
  DO
    Cnt = Cnt + 1
    GET AREdFile, Cnt, EditPaymentRec(1)
    CRec = VAL(EditPaymentRec(1).CustNumber)
    IF CRec > 0 THEN
      GET ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
      help$ = "Posting: " + LEFT$(ARCustRec(1).BILLNAME, 30)
      PrintHelp help$
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
      ARTransRec(1).CashAmount = EditPaymentRec(1).Amount
      ARTransRec(1).ChkAmount = 0
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
      ARCustRec(1).LicBal = Round(ARCustRec(1).LicBal - EditPaymentRec(1).LicPAID)
      ARCustRec(1).PenBal = Round(ARCustRec(1).PenBal - EditPaymentRec(1).PenPAID)
      
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
  
  ' Show All Posted
  BEEP
  SaveScrn TempScrn()
  DisplayARScrn "ARPOSTED"
  WaitForAction
  RestScrn TempScrn()
  LOCATE , , 1
  RETURN
  
END SUB

SUB PrintEditList (OperNum)
  
  SHARED Choice$()

  Oper$ = QPTrim$(STR$(OperNum))
  PayFileName$ = "AREDPY" + Oper$ + ".DAT"
  
  ReportFile$ = "AREDPY" + Oper$ + ".PRN" 'Report File Name
  FF$ = CHR$(12)
  LPTPort% = 1
  MaxLines = 53
  LineCnt = 0
  ClearBack
  Header$ = "Payment Edit List"
  ShowProcessingScrn Header$
  
  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
  
  AREditRecLen = LEN(EditPaymentRec(1))
  ARFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
  NumOfArRecs = LOF(ARFile) \ AREditRecLen
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO NumOfArRecs
    GET ARFile, Cnt, EditPaymentRec(1)
    CRec = VAL(EditPaymentRec(1).CustNumber)
    IF CRec <= 0 THEN
      GOTO SkipDeleted
    END IF
    GET TrHandle, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)

    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader
    END IF
    IF EditPaymentRec(1).Amount <> 0 THEN
      CustNo& = VAL(QPTrim$(ARCustRec(1).CUSTNUMB))
      Change# = Round#(EditPaymentRec(1).AmtPaid - Round#(EditPaymentRec(1).LicPAID + EditPaymentRec(1).PenPAID))
      TChange# = Round#(TChange# + Change#)
      TLic# = Round#(TLic# + EditPaymentRec(1).LicPAID)
      TPen# = Round#(TPen# + EditPaymentRec(1).PenPAID)
      
      PRINT #RptHandle, USING "#####"; CustNo&;
      PRINT #RptHandle, TAB(8); LEFT$(ARCustRec(1).BILLNAME, 25); TAB(35); USING "#####.##"; EditPaymentRec(1).AmtPaid; TAB(45); EditPaymentRec(1).LicPAID; TAB(55); EditPaymentRec(1).PenPAID; TAB(65); Change#
      
      TotalCust = TotalCust + 1
      'TotalValue# = Round#(TotalValue# + EditPaymentRec(1).Amount)
      LineCnt = LineCnt + 1
    END IF
SkipDeleted:
    ShowPctComp Cnt, NumOfArRecs
  NEXT
  
  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(18); "Business License : Payment Edit Listing"
  PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Cust#"; TAB(8); "Billing Name"; TAB(35); "Tendered"; TAB(50); "Lic"; TAB(60); "Pen"; TAB(67); "Change"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, STRING$(80, "-")
  PRINT #RptHandle, "Number of Entries .. "; USING "####"; TotalCust;
  PRINT #RptHandle, TAB(45); USING "#####.##"; TLic#; TAB(55); TPen#; TAB(65); TChange#
  'PRINT #RptHandle, TAB(55); USING "$$#####,#.##"; TotalValue#
  PRINT #RptHandle, FF$
  RETURN
  
  
END SUB

SUB PrintReceipt (OperNum, TRRec%)

  Oper$ = QPTrim$(STR$(OperNum))
  PayFileName$ = "AREDPY" + Oper$ + ".DAT"

  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  TownName$ = UBSetUpRec(1).UTILNAME

  REDIM PayRec(1)  AS AREditPaymentRecType
  AREditRecLen = LEN(PayRec(1))
  ARFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
  GET ARFile, TRRec, PayRec(1)
  CLOSE ARFile

  'Set Printer Port for Receipts Printer if none set to 1

  OPEN "O", 15, "RECPT.PRN"
  WIDTH #15, 255

  PRINT #15, ""
  PRINT #15, TownName$
  PRINT #15, "LICENSE PAYMENT"
  PRINT #15, "Date: "; Num2Date(PayRec(1).TranDate)
  PRINT #15,
  PRINT #15, "Account #"; PayRec(1).CustNumber
  PRINT #15, QPTrim$(PayRec(1).CustName)
  PRINT #15, QPTrim$(PayRec(1).Add1)
  PRINT #15, QPTrim$(PayRec(1).Desc)
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; PayRec(1).TotDue
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; PayRec(1).CashAmt
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; PayRec(1).ChkAmt
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; PayRec(1).AmtPaid#
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; PayRec(1).Change
  PRINT #15,
  PRINT #15,
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
  PrintRptFile Header$, "RECPT.PRN", 1, RetCode%, 5

END SUB

