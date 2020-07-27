DEFINT A-Z
DECLARE SUB PostPayments ()
DECLARE SUB PrintEditList ()
DECLARE SUB EditPayment ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
DECLARE SUB OpenARCustFile (NumOfArRecs, ARFile)
DECLARE SUB EnterPayment ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB DisplayARScrn (ScrnName$)
DECLARE SUB WaitForAction
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


  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

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

   '--Find max menu width
   FOR Cnt = 1 TO UBOUND(Mchoice$)
     TLen = LEN(Mchoice$(Cnt))
     IF TLen > MaxLen THEN
       MaxLen = TLen
     END IF
   NEXT

   '--Center Menu within Screen
   Row = ((25 - (UBOUND(Mchoice$))) \ 2)
   Col = ((80 - MaxLen) \ 2) - 2
   help$ = "Accounts Receivable Payment Menu"
   
   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, "Payment / Transaction Menu ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      'PrintTitle user$
      'printhelp help$
      ShowCursor

      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

      IF Ky$ = CHR$(27) THEN EXIT DO 'choice = 0

      SELECT CASE Choice
          CASE 1
          EnterPayment
          CASE 2
          EditPayment
          CASE 3
          PrintEditList
          CASE 4
          PostPayments
          CASE IS = 5
          END
      END SELECT
   LOOP
   RUN "armenu"

SUB DisplayARScrn (ScrnName$)
 LibFile2Scrn "AR", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB EditPayment
EditMainBody:
  CustomerGrabed = 0

  ARIdxFile = 0
  ARFile = 0
Continue:
  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  help$ = "Edit Customer Balance Entry's"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$


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

  IF AccountRecord = 0 THEN CLOSE : EXIT SUB

  GET AREdFile, AccountRecord, EditPaymentRec(1)
  GET ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
  Form$(2, 0) = Num2Date$(EditPaymentRec(1).TranDate)
  Form$(3, 0) = EditPaymentRec(1).CustNumber
  Form$(4, 0) = ARCustRec(1).CustName
  Form$(6, 0) = ARCustRec(1).ADDRESS1
  Form$(7, 0) = ARCustRec(1).ADDRESS2
  Form$(8, 0) = ARCustRec(1).CITY
  Form$(9, 0) = ARCustRec(1).STATE
  Form$(10, 0) = ARCustRec(1).ZIPCODE
  Form$(11, 0) = EditPaymentRec(1).Desc
  Form$(12, 0) = STR$(EditPaymentRec(1).Amount)
  Form$(13, 0) = EditPaymentRec(1).IssueLic
  Form$(14, 0) = EditPaymentRec(1).SetFee

  Fld(1).Protected = True

  FOR Fld = 3 TO 10
    Fld(Fld).Protected = True
  NEXT Fld

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action


    SELECT CASE Frm(1).KeyCode

    CASE F3Key
       GOSUB DeleteRecord
       IF Deleted THEN
        EditPaymentRec(1).Amount = 0
        PUT AREdFile, AccountRecord, EditPaymentRec(1)
        CLOSE AREdFile
        EXIT SUB
        ELSE
        GOTO Continue
        END IF

    CASE F10Key
       EditPaymentRec(1).TranDate = Date2Num(Form$(2, 0))
       EditPaymentRec(1).Amount = Value(Form$(12, 0), a%)
       EditPaymentRec(1).IssueLic = Form$(13, 0)
       PUT AREdFile, AccountRecord, EditPaymentRec(1)
       CLOSE AREdFile
       help$ = "SAVING YOUR CHANGES"
       'printhelp help$
       PRINT CHR$(7)
       SLEEP 1
       CLOSE
       EXIT SUB
    CASE EscKey
      EXIT SUB

    END SELECT

  LOOP

GetEditRecord:

   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = " Cust #    Customer Business Name"

   '--Center Menu within Screen
   Row = 8
   Col = 15

  AREditRecLen = LEN(EditPaymentRec(1))
  AREdFile = FREEFILE
  OPEN "AREDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS AREdFile LEN = AREditRecLen
  NumOfArRecs = LOF(AREdFile) \ AREditRecLen
  IF NumOfArRecs = 0 THEN CLOSE AREdFile: RETURN

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
        AccountRecord = VAL(RIGHT$(Mchoice$(Choice), 5))
        ExitFlag = True
      END IF

   LOOP UNTIL ExitFlag

  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  LOCATE 1, 1, 0
  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  RETURN

DeleteRecord:
  LibName$ = "AR"
  ScrnName$ = "AREDITOK"
  help$ = "Delete Edit Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
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
  Form$(1, 0) = "Y"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode


    CASE EscKey
      IF Form$(1, 0) = "Y" THEN
       help$ = "Entry Deleted!!!"
       'printhelp help$
       PRINT CHR$(7);
       Deleted = True
       RETURN
      END IF
       Action = 1
       RETURN

    END SELECT

  LOOP

END SUB

SUB EnterPayment
REDIM TempScrn(0)
mainbody:
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0

  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  help$ = "Enter Payments"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$



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
  IF LEN(PrevDate$) = 0 THEN
   Form$(2, 0) = DATE$
   ELSE
   Form$(2, 0) = PrevDate$
  END IF

  
  Fld(1).Protected = True
  Frm(1).FldNo = 2
  Form$(13, 0) = "Y"
  Form$(14, 0) = "Y"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF Frm(1).PrevFld = 4 AND CustomerGrabed = 0 THEN
      GOSUB GetCustomer
      IF AccountRecord = 0 THEN
        CLOSE : Action = 1: GOTO mainbody
      END IF

        REM check for existing transaction
        CLOSE ARFile
        AREditRecLen = LEN(EditPaymentRec(1))
        ARFile = FREEFILE
        OPEN "AREDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = AREditRecLen
        NumOfArRecs = LOF(ARFile) \ AREditRecLen
        IF NumOfArRecs > 0 THEN
         Cnt = 0
         WHILE Cnt < NumOfArRecs
          Cnt = Cnt + 1
          GET ARFile, Cnt, EditPaymentRec(1)
         IF VAL(EditPaymentRec(1).CustNumber) = VAL(Form$(3, 0)) THEN
           SaveScrn TempScrn()
           LOCATE , , 0
           PRINT CHR$(7);
           DisplayARScrn "ARDUPCUS"
           WaitForAction
           RestScrn TempScrn()
           LOCATE , , 1
           CLOSE
           GOTO mainbody
          END IF
           WEND
        END IF

     Action = 1
   END IF

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      IF VAL(Form$(12, 0)) <> 0 THEN
        GOSUB SaveRecord
      END IF
      CLOSE
      Done = True
      GOTO mainbody
    CASE EscKey
      EXIT SUB

    END SELECT

  LOOP


SaveRecord:
    PrevDate$ = Form$(2, 0): REM Keep Default Date Here
    CLOSE
    AREditRecLen = LEN(EditPaymentRec(1))
    ARFile = FREEFILE
    OPEN "AREDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = AREditRecLen
    NumOfArRecs = LOF(ARFile) \ AREditRecLen
     EditPaymentRec(1).TranDate = Date2Num(Form$(2, 0))
     EditPaymentRec(1).CustNumber = Form$(3, 0)
     EditPaymentRec(1).CustName = Form$(4, 0)
     EditPaymentRec(1).Desc = Form$(11, 0)
     EditPaymentRec(1).Amount = Value(Form$(12, 0), a%)
     EditPaymentRec(1).IssueLic = Form$(13, 0)
     EditPaymentRec(1).SetFee = Form$(14, 0)
    PUT ARFile, NumOfArRecs + 1, EditPaymentRec(1)
    CLOSE ARFile
    RETURN


GetCustomer:

CustomerGrabed = 0
AccountRecord = VAL(Form$(3, 0))

REM **************************************************************************

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

   REDIM Mchoice$(1 TO NumOfARIdxRecs)

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfARIdxRecs
    GET ARIdxFile, Cnt, ARCustIdxRec(1)
    IF LEFT$(ARCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
     ChoiceCounter = ChoiceCounter + 1
     Mchoice$(ChoiceCounter) = SPACE$(50)
     LSET Mchoice$(ChoiceCounter) = STR$(ARCustIdxRec(1).IDXRECORD)
     MID$(Mchoice$(ChoiceCounter), 10) = ARCustIdxRec(1).IDXNAME
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
        AccountRecord = VAL(LEFT$(Mchoice$(Choice), 8))
        ExitFlag = True
      END IF

   LOOP UNTIL ExitFlag

  LibName$ = "AR"
  ScrnName$ = "ARTRANEN"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%


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
    Form$(5, 0) = ARCustRec(1).BILLNAME
    Form$(6, 0) = ARCustRec(1).ADDRESS1
    Form$(7, 0) = ARCustRec(1).ADDRESS2
    Form$(8, 0) = ARCustRec(1).CITY
    Form$(9, 0) = ARCustRec(1).STATE
    Form$(10, 0) = ARCustRec(1).ZIPCODE
    Fld(3).Protected = True
    Frm(1).FldNo = 11
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
    'printhelp help$

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
    'printhelp help$

    Done = False
    Action = 1


  DO

   EditForm Form$(), Fld(), Frm(1), Cnf, Action

   SELECT CASE Frm(1).KeyCode
    CASE F10Key
     RETURN
    END SELECT
  LOOP


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

SUB PostPayments
REDIM TempScrn(0)
REDIM ARCatCodeRec(1) AS ARCatCodeRecType
CatCodeRecLen = LEN(ARCatCodeRec(1))

LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode


MainPostBody:
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0

  LibName$ = "AR"
  ScrnName$ = "AROKPOST"
  help$ = "Set Customer Balances"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$

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

  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
  OpenARCustFile NumOfArRecs, ARFile

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
  OpenARCustFile NumOfArRecs, ARFile

  AREditRecLen = LEN(EditPaymentRec(1))
  AREdFile = FREEFILE
  OPEN "AREDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS AREdFile LEN = AREditRecLen
  NumOfArRecs = LOF(AREdFile) \ AREditRecLen

' See if any records to post
  IF NumOfArRecs = 0 THEN
   PRINT CHR$(7);
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

 IF EditPaymentRec(1).Amount <> 0 THEN
   GET ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
   help$ = "Posting: " + LEFT$(ARCustRec(1).BILLNAME, 30)
   'printhelp help$


  'Get Catagory Code Record # Here
   Code$ = QPTrim$(ARCustRec(1).BILLCAT)
   CatFile = FREEFILE
   OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
   NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
   CatRecord = 0
   FOR SCnt! = 1 TO NumOfCatRecs
    GET CatFile, SCnt!, ARCatCodeRec(1)
        ARCode$ = QPTrim$(ARCatCodeRec(1).CATCODE)
    IF ARCode$ = Code$ THEN
     CatRecord% = SCnt!
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
     ARTransRec(1).TransType = 2                                 ' Type 2 = Payment
     ARTransRec(1).TransDesc = "Payment"
     ARTransRec(1).CashAmount = EditPaymentRec(1).Amount
     ARTransRec(1).ChkAmount = 0
     ARTransRec(1).BalanceAfterTrans = NewBalance#
     ARTransRec(1).ExtraRoom = ""
     ARTransRec(1).Posted2GL = "N"
     ARTransRec(1).CatCodeRec = CatRecord%
     ARTransRec(1).NextTrans = 0
     PUT ARTransFile, NextTransRec, ARTransRec(1)

   'Update Customer
     GET ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
     ARCustRec(1).IssueLicense = EditPaymentRec(1).IssueLic
     ARCustRec(1).AcctBal = ARCustRec(1).AcctBal - EditPaymentRec(1).Amount

  IF EditPaymentRec(1).SetFee = "Y" THEN
     ARCustRec(1).FeeAmt = EditPaymentRec(1).Amount
  END IF

     PUT ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)

  IF ARCustRec(1).FirstTrans = 0 THEN
     ARCustRec(1).FirstTrans = NextTransRec
     ARCustRec(1).LastTrans = NextTransRec
     PUT ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
  ELSE
     Prev! = ARCustRec(1).LastTrans
     ARCustRec(1).LastTrans = NextTransRec
     PUT ARFile, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)
     GET ARTransFile, Prev!, ARTransRec(1)
     ARTransRec(1).NextTrans = NextTransRec
     PUT ARTransFile, Prev!, ARTransRec(1)
  END IF
     NextTransRec = NextTransRec + 1
 END IF

 LOOP UNTIL Cnt > NumOfArRecs
 CLOSE
 KILL "AREDPYT.DAT"

 ' Show All Posted
   PRINT CHR$(7);
   SaveScrn TempScrn()
   DisplayARScrn "ARPOSTED"
   WaitForAction
   RestScrn TempScrn()
   LOCATE , , 1
   CLOSE
   RETURN

END SUB

SUB PrintEditList
  SHARED Choice$()
  ReportFile$ = "ARPAYED.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  LPTPort% = 1
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  size = 2500
  Start = 1               'start at array element 1
  Dir = 0                 'sort direction - use anything else for descending
  SSize = 16               'total size of each TYPE element
  MOff = 0                'offset into the TYPE for the key element
  MSize = 16              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

'  REDIM array(1 TO size) AS Struct

  GOSUB SelectOutput


  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen

  AREditRecLen = LEN(EditPaymentRec(1))
  ARFile = FREEFILE
  OPEN "AREDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = AREditRecLen
  NumOfArRecs = LOF(ARFile) \ AREditRecLen

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintRptHeader

  FOR Cnt = 1 TO NumOfArRecs
   GET ARFile, Cnt, EditPaymentRec(1)
    GET TrHandle, VAL(EditPaymentRec(1).CustNumber), ARCustRec(1)

   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   IF EditPaymentRec(1).Amount <> 0 THEN
   PRINT #RptHandle, VAL(ARCustRec(1).CUSTNUMB); TAB(10); ARCustRec(1).BILLNAME; TAB(55); USING "$$#####,#.##"; EditPaymentRec(1).Amount;
    PRINT #RptHandle, TAB(75); EditPaymentRec(1).IssueLic
    TotalCust = TotalCust + 1
    TotalValue# = TotalValue# + EditPaymentRec(1).Amount
    TotalValue# = INT((TotalValue# * 100) + .5) / 100
    LineCnt = LineCnt + 1
   END IF

  NEXT Cnt
    GOSUB PrintRptEnding
    PRINT #RptHandle, CHR$(18); ' oki 320 10 cpi
    CLOSE                       'Close all open files now

  IF DevSpec$ = "S" THEN
       EntryPoint = 2
       ELSE
       EntryPoint = 5
  END IF
  Header$ = "Payment Edit Listing"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintRptHeader:
    page = page + 1
    PRINT #RptHandle, TAB(18); "Business License : Payment Edit Listing"
    PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Cust #"; TAB(10); "Billing Name"; TAB(55); "Payment Amt"; TAB(73); "Prt Lic"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 5
    RETURN

PrintRptEnding:
     PRINT #RptHandle, STRING$(80, "-")
     PRINT #RptHandle, "Number of Entries .. "; USING "####,#"; TotalCust;
     PRINT #RptHandle, TAB(55); USING "$$#####,#.##"; TotalValue#
     PRINT #RptHandle, FF$
     RETURN




SelectOutput:
  LibName$ = "AR"
  ScrnName$ = "WHERPRNT"

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

  REDIM Choice$(2, 0)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "SCREEN"
  Choice$(2, 0) = "PRINTER"


  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$
  Action = 1
  COLOR 14: LOCATE 9, 23: PRINT "Payment Edit List"; : COLOR 10: PRINT " ]";


  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       DevSpec$ = LEFT$(Form$(1, 0), 1)
       RETURN
     CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP
  RETURN





END SUB

