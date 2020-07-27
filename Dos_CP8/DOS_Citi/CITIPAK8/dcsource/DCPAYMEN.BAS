DEFINT A-Z
DECLARE SUB SortDCNameIndex ()
DECLARE SUB ShowNoCodes ()
DECLARE SUB AddCustomer ()
DECLARE SUB PostPayments ()
DECLARE SUB PrintEditList ()
DECLARE SUB EditPayment ()
DECLARE SUB OpenDCCustIdxFile (NumOfDCIdxRecs, DCIdxFile)
DECLARE SUB OpenDCCustFile (NumOfDcRecs, DCFile)
DECLARE SUB EnterPayment ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB SaveScrn (array%())
DECLARE SUB RestScrn (array%())
DECLARE SUB DisplayDCScrn (ScrnName$)
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
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

   CONST False = 0, True = NOT False
   DIM SHARED DCCustRec(1) AS DCCustRecType
   DIM SHARED DCCustIdxRec(1) AS DCCustIDXRecType
   DIM SHARED EditPaymentRec(1) AS DCEditPaymentRecType

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE

 STACK 8000

   '--Dim the choice array to the number of menu items
   REDIM Mchoice$(1 TO 5)

   Mchoice$(1) = " Enter Decal Purchase  "
   Mchoice$(2) = " Edit Purchases "
   Mchoice$(3) = " Print Edit List "
   Mchoice$(4) = " Post Purchases to Accts "
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
   help$ = "Decal Purchases"
   
   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, "Decal Purchases ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      PrintTitle user$
      printhelp help$
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
   RUN "dcmenu"

SUB AddCustomer

mainbody:
  LibName$ = "DC"
  ScrnName$ = "DCCUST"
  help$ = "NEW Decal Customer Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  printhelp help$



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

  REM check for code file

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen
  CLOSE DCCatFile
  IF NumOFDCCatRecs = 0 THEN
   ShowNoCodes
   EXIT SUB
  END IF



  OpenDCCustFile NumOfDcRecs, DCFile




  Fld(1).Protected = True

  Frm(1).FldNo = 2

  LOCATE 5, 28: COLOR 15: PRINT "        PENDING"
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF Frm(1).FldNo = 9 AND LEFT$(Form$(9, 0), 1) = " " THEN
      GOSUB SelectCatagory
      ShowCursor
      LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      printhelp help$
    END IF

    'IF Frm(1).PrevFld = 4 THEN Form$(11, 0) = Form$(4, 0): action = 1

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      GOSUB SaveRecord
      Done = True
      GOTO mainbody
    CASE EscKey
      NeedtoSort = True         ' set to true for testing
      IF NeedtoSort = True THEN
       SortDCNameIndex
      END IF

      EXIT SUB
    END SELECT

  LOOP


SaveRecord:
    IF LEFT$(Form$(9, 0), 1) = " " THEN
     CLOSE DCFile
     CLOSE DCCatFile
     ELSE
    DCCustRec(1).SORTNAME = Form$(2, 0)
    DCCustRec(1).BILLNAME = Form$(3, 0)
    DCCustRec(1).ADDRESS1 = Form$(4, 0)
    DCCustRec(1).ADDRESS2 = Form$(5, 0)
    DCCustRec(1).CITY = Form$(6, 0)
    DCCustRec(1).STATE = Form$(7, 0)
    DCCustRec(1).ZIPCODE = Form$(8, 0)
    DCCustRec(1).BILLCAT = Form$(9, 0)
    DCCustRec(1).SOSEC = Form$(10, 0)
    DCCustRec(1).DRVLIC = Form$(11, 0)
    DCCustRec(1).DATEOPED = Date2Num(Form$(12, 0))
    DCCustRec(1).BILLCMT = Form$(13, 0)
    DCCustRec(1).PAYCMT = Form$(14, 0)
    DCCustRec(1).CASHONLY = Form$(15, 0)
    DCCustRec(1).RESIDENT = Form$(16, 0)
    DCCustRec(1).OWNER = Form$(17, 0)
    DCCustRec(1).HPHONE = Form$(18, 0)
    DCCustRec(1).WPHONE = Form$(19, 0)
    DCCustRec(1).FeeAmt = Value#(Form$(20, 0), a%)  'DOUBLE
    DCCustRec(1).LICENSE = Form$(21, 0)
    DCCustRec(1).VALID = Date2Num(Form$(22, 0))       'INTEGER Date Function
    DCCustRec(1).AcctBal = 0
    DCCustRec(1).FirstTrans = 0
    DCCustRec(1).LastTrans = 0
    DCCustRec(1).Deleted = "N"
    DCCustRec(1).RoomtoGrow = ""
    NextAccount = NumOfDcRecs + 1
    DCCustRec(1).CUSTNUMB = STR$(NextAccount)
    PUT DCFile, NextAccount, DCCustRec(1)
    LOCATE 5, 28: COLOR 15: PRINT NextAccount; "  ASSIGNED": SLEEP 2
    CLOSE DCFile
    CLOSE DCCatFile
    NeedtoSort = True
    END IF
    RETURN

SelectCatagory:

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))

  DCCatFile = FREEFILE

  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen

  REDIM Mchoice$(1 TO NumOFDCCatRecs)
  FOR Cnt = 1 TO NumOFDCCatRecs
    GET DCCatFile, Cnt, DCCatCodeRec(1)
    Mchoice$(Cnt) = SPACE$(50)
    LSET Mchoice$(Cnt) = DCCatCodeRec(1).CATCODE
    MID$(Mchoice$(Cnt), 5) = DCCatCodeRec(1).CODEDESC
  NEXT Cnt

   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = "  Code    Description"

   '--Center Menu within Screen
   Row = 8
   Col = 15

  '--Set upper left corner of menu, turn off the cursor
   LOCATE Row, Col, 0
   QPrintRC TText$, Row - 1, Col, 112
   VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
   GET DCCatFile, Choice, DCCatCodeRec(1)

   Form$(9, 0) = DCCatCodeRec(1).CATCODE
   Form$(20, 0) = STR$(DCCatCodeRec(1).Fee)
   Frm(1).FldNo = 10
RETURN




END SUB

SUB DisplayDCScrn (ScrnName$)
 LibFile2Scrn "DC", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB EditPayment
EditMainBody:
  CustomerGrabed = 0

  DCIdxFile = 0
  DCFile = 0
Continue:
  LibName$ = "DC"
  ScrnName$ = "DCTRANEN"
  help$ = "Edit Customer Balance Entry's"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  printhelp help$


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

  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
  OpenDCCustFile NumOfDcRecs, DCFile

  IF AccountRecord = 0 THEN
   GOSUB GetEditRecord
  END IF

  IF AccountRecord = 0 THEN CLOSE : EXIT SUB

  GET DCEdFile, AccountRecord, EditPaymentRec(1)
  GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
  Form$(1, 0) = Num2Date$(EditPaymentRec(1).TranDate)
  Form$(2, 0) = EditPaymentRec(1).CustNumber
  Form$(3, 0) = DCCustRec(1).BILLNAME
  Form$(4, 0) = DCCustRec(1).ADDRESS1
  Form$(5, 0) = DCCustRec(1).ADDRESS2
  Form$(6, 0) = DCCustRec(1).CITY
  Form$(7, 0) = DCCustRec(1).STATE
  Form$(8, 0) = DCCustRec(1).ZIPCODE
  Form$(9, 0) = EditPaymentRec(1).DecalCat
  Form$(10, 0) = EditPaymentRec(1).Sticker
  Form$(11, 0) = EditPaymentRec(1).Desc
  Form$(12, 0) = EditPaymentRec(1).MakeModel
  Form$(13, 0) = EditPaymentRec(1).StateTag
  Form$(14, 0) = STR$(EditPaymentRec(1).Amount)
  Form$(15, 0) = Num2Date$(EditPaymentRec(1).ExpDate)
  

  

  FOR Fld = 2 TO 9
    Fld(Fld).Protected = True
  NEXT Fld

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action


    SELECT CASE Frm(1).KeyCode

    CASE F3Key
       GOSUB DeleteRecord
       IF Deleted THEN
        EditPaymentRec(1).Amount = 0
        PUT DCEdFile, AccountRecord, EditPaymentRec(1)
        CLOSE DCEdFile
        EXIT SUB
        ELSE
        GOTO Continue
        END IF

    CASE F10Key

      IF Date2Num(Form$(15, 0)) < 0 THEN
       COLOR 12
       PRINT CHR$(7);
       LOCATE 19, 40: PRINT "MUST ENTER EXPIRATION DATE"
       SLEEP 2
       LOCATE 19, 40: PRINT "                          "
       Frm(1).FldNo = 15
       Action = 1
      ELSE
       EditPaymentRec(1).TranDate = Date2Num(Form$(1, 0))
       EditPaymentRec(1).Sticker = Form$(10, 0)
       EditPaymentRec(1).Desc = Form$(11, 0)
       EditPaymentRec(1).MakeModel = Form$(12, 0)
       EditPaymentRec(1).StateTag = Form$(13, 0)
       EditPaymentRec(1).Amount = Value(Form$(14, 0), a%)
       EditPaymentRec(1).ExpDate = Date2Num(Form$(15, 0))
       PUT DCEdFile, AccountRecord, EditPaymentRec(1)
       CLOSE DCEdFile
       help$ = "SAVING YOUR CHANGES"
       printhelp help$
       PRINT CHR$(7)
       SLEEP 1
       CLOSE
       EXIT SUB
       END IF
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
   LSET TText$ = " Cust #    Customer Name                   Trans #"

   '--Center Menu within Screen
   Row = 8
   Col = 15

  DCEditRecLen = LEN(EditPaymentRec(1))
  DCEdFile = FREEFILE
  OPEN "DCEdPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCEdFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCEdFile) \ DCEditRecLen
  IF NumOfDcRecs = 0 THEN CLOSE DCEdFile: RETURN

   REDIM Mchoice$(1 TO NumOfDcRecs)

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfDcRecs
    GET DCEdFile, Cnt, EditPaymentRec(1)
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
      LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
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

  LibName$ = "DC"
  ScrnName$ = "DCTRANEN"
  LOCATE 1, 1, 0
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  RETURN

DeleteRecord:
  LibName$ = "DC"
  ScrnName$ = "DCEditOK"
  help$ = "Delete Edit Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  printhelp help$
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
       printhelp help$
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

PayMainBody:
  CustomerGrabed = 0
  DCIdxFile = 0
  DCFile = 0

  LibName$ = "DC"
  ScrnName$ = "DCTRANEN"
  help$ = "Enter Payments"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  printhelp help$



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

  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
  OpenDCCustFile NumOfDcRecs, DCFile
 IF LEN(PrevDate$) = 0 THEN
   Form$(1, 0) = DATE$
   PrevDate$ = DATE$
   ELSE
   Form$(1, 0) = PrevDate$
  END IF

  
  'Fld(1).Protected = True
  'Frm(1).FldNo = 2

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF Frm(1).PrevFld = 2 AND CustomerGrabed = 0 THEN
       GOSUB GetCustomer
      IF AccountRecord = 0 THEN
       GOSUB ADDNEWCUSTOMER
       CLOSE : Action = 1: GOTO PayMainBody
      END IF
    END IF

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      IF Date2Num(Form$(15, 0)) < 0 THEN
      COLOR 12
      PRINT CHR$(7);
      LOCATE 19, 40: PRINT "MUST ENTER EXPIRATION DATE"
      SLEEP 2
      LOCATE 19, 40: PRINT "                          "
      Frm(1).FldNo = 15

      Action = 1
       ELSE
      IF VAL(Form$(14, 0)) <> 0 THEN
        GOSUB PAYSaveRecord
      END IF
      CLOSE
      Done = True
      GOTO PayMainBody
      END IF
    CASE EscKey
      EXIT SUB

    END SELECT

  LOOP


PAYSaveRecord:
    PrevDate$ = Form$(1, 0): REM Keep Default Date Here
    CLOSE
    DCEditRecLen = LEN(EditPaymentRec(1))
    DCFile = FREEFILE
    OPEN "DCEdPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCEditRecLen
    NumOfDcRecs = LOF(DCFile) \ DCEditRecLen
     EditPaymentRec(1).TranDate = Date2Num(Form$(1, 0))
     EditPaymentRec(1).CustNumber = Form$(2, 0)
     EditPaymentRec(1).CustName = Form$(3, 0)
     EditPaymentRec(1).DecalCat = Form$(9, 0)
     EditPaymentRec(1).Sticker = Form$(10, 0)
     EditPaymentRec(1).Desc = Form$(11, 0)
     EditPaymentRec(1).MakeModel = Form$(12, 0)
     EditPaymentRec(1).StateTag = Form$(13, 0)
     EditPaymentRec(1).Amount = Value(Form$(14, 0), a%)
     EditPaymentRec(1).ExpDate = Date2Num(Form$(15, 0))
    PUT DCFile, NumOfDcRecs + 1, EditPaymentRec(1)
    CLOSE DCFile
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

   REDIM Mchoice$(1 TO NumOfDCIdxRecs + 1)

     ChoiceCounter = 1
     Mchoice$(ChoiceCounter) = SPACE$(50)
     LSET Mchoice$(ChoiceCounter) = STR$(0)
     MID$(Mchoice$(ChoiceCounter), 10) = "NEW"

  FOR Cnt = 1 TO NumOfDCIdxRecs
    GET DCIdxFile, Cnt, DCCustIdxRec(1)
    IF LEFT$(DCCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
     ChoiceCounter = ChoiceCounter + 1
     Mchoice$(ChoiceCounter) = SPACE$(50)
     LSET Mchoice$(ChoiceCounter) = STR$(DCCustIdxRec(1).IDXRECORD)
     MID$(Mchoice$(ChoiceCounter), 10) = DCCustIdxRec(1).IDXNAME
    END IF
  NEXT Cnt

   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
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

  LibName$ = "DC"
  ScrnName$ = "DCTRANEN"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%


   END IF

REM ************************************************************************
IF AccountRecord > 0 AND AccountRecord <= NumOfDcRecs THEN
    GET DCFile, AccountRecord, DCCustRec(1)
    IF DCCustRec(1).Deleted = "Y" THEN
     GOSUB CustomerDeleted
     GOTO PayMainBody
    END IF

    Form$(2, 0) = DCCustRec(1).CUSTNUMB
    Form$(3, 0) = DCCustRec(1).BILLNAME
    Form$(4, 0) = DCCustRec(1).ADDRESS1
    Form$(5, 0) = DCCustRec(1).ADDRESS2
    Form$(6, 0) = DCCustRec(1).CITY
    Form$(7, 0) = DCCustRec(1).STATE
    Form$(8, 0) = DCCustRec(1).ZIPCODE
    Form$(9, 0) = DCCustRec(1).BILLCAT
    Form$(14, 0) = STR$(DCCustRec(1).FeeAmt)
    Fld(2).Protected = True
    Fld(3).Protected = True
    Fld(4).Protected = True
    Fld(5).Protected = True
    Fld(6).Protected = True
    Fld(7).Protected = True
    Fld(8).Protected = True
    Fld(9).Protected = True
    Frm(1).FldNo = 10
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    RETURN

   ELSE
   IF AccountRecord = 0 THEN RETURN
    LibName$ = "DC"
    ScrnName$ = "DCBADCUS"
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
    LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    printhelp help$

    Done = False
    Action = 1

  DO

   EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
   SELECT CASE Frm(1).KeyCode
    
    CASE EscKey
     Done = True
     END SELECT
     IF Done = True THEN GOTO PayMainBody
  LOOP

 END IF

CustomerDeleted:
    LibName$ = "DC"
    ScrnName$ = "DCDELCUS"
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
    LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    printhelp help$

    Done = False
    Action = 1


  DO

   EditForm Form$(), Fld(), Frm(1), Cnf, Action

   SELECT CASE Frm(1).KeyCode
    CASE F10Key
     RETURN
    END SELECT
  LOOP

ADDNEWCUSTOMER:
AddCustomer
RETURN

END SUB

SUB OpenDCCustFile (NumOfDcRecs, DCFile)
  CLOSE DCFile
  DCCustRecLen = LEN(DCCustRec(1))
  DCFile = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDcRecs = LOF(DCFile) \ DCCustRecLen
END SUB

SUB OpenDCCustIdxFile (NumOfDCIdxRecs, DCIdxFile)
  CLOSE DCIdxFile
  DCCustIdxRecLen = LEN(DCCustIdxRec(1))
  DCIdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS DCIdxFile LEN = DCCustIdxRecLen
  NumOfDCIdxRecs = LOF(DCIdxFile) \ DCCustIdxRecLen
END SUB

SUB PostPayments
REDIM TempScrn(0)

LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode

MainPostBody:
  CustomerGrabed = 0
  DCIdxFile = 0
  DCFile = 0

  LibName$ = "DC"
  ScrnName$ = "DCOKPOST"
  help$ = "Set Customer Balances"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  printhelp help$

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

  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
  OpenDCCustFile NumOfDcRecs, DCFile

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
  OpenDCCustFile NumOfDcRecs, DCFile

  DCEditRecLen = LEN(EditPaymentRec(1))
  DCEdFile = FREEFILE
  OPEN "DCEdPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCEdFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCEdFile) \ DCEditRecLen

' See if any records to post
  IF NumOfDcRecs = 0 THEN
   PRINT CHR$(7);
   SaveScrn TempScrn()
   DisplayDCScrn "DCNOTRAN"
   WaitForAction
   RestScrn TempScrn()
   LOCATE , , 1
   CLOSE
   EXIT SUB
  END IF


 REDIM DCTransRec(1) AS DCTransRecType
 DCTransRecLen = LEN(DCTransRec(1))
 DCTransFile = FREEFILE
 OPEN "DCTrans.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCTransFile LEN = DCTransRecLen
 NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
 NextTransRec = NumOfTransRecs + 1

 DO
 Cnt = Cnt + 1
 GET DCEdFile, Cnt, EditPaymentRec(1)

 IF EditPaymentRec(1).Amount <> 0 THEN
   GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
   help$ = "Posting: " + LEFT$(DCCustRec(1).BILLNAME, 30)
   printhelp help$

   ' Post Charge First to Offset Payment of Decal
  DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
  DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
  DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
  DCTransRec(1).TransType = 1               ' Type 1 = Charge
  DCTransRec(1).TransDesc = "Decal Purchase " + EditPaymentRec(1).Sticker
  DCTransRec(1).CashAmount = EditPaymentRec(1).Amount
  DCTransRec(1).ChkAmount = 0
  DCTransRec(1).BalanceAfterTrans = DCTransRec(1).BalanceAfterTrans + EditPaymentRec(1).Amount
  DCTransRec(1).MakeModel = EditPaymentRec(1).MakeModel
  DCTransRec(1).StateTag = EditPaymentRec(1).StateTag
  DCTransRec(1).Sticker = EditPaymentRec(1).Sticker
  DCTransRec(1).ExpireDate = EditPaymentRec(1).ExpDate
  DCTransRec(1).ExtraRoom = ""
  DCTransRec(1).NextTrans = 0
  PUT DCTransFile, NextTransRec, DCTransRec(1)


  GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
   DCCustRec(1).AcctBal = DCCustRec(1).AcctBal + EditPaymentRec(1).Amount
   DCCustRec(1).VALID = EditPaymentRec(1).ExpDate
  PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)

  IF DCCustRec(1).FirstTrans = 0 THEN
   DCCustRec(1).FirstTrans = NextTransRec
   DCCustRec(1).LastTrans = NextTransRec
   PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
  ELSE
   Prev! = DCCustRec(1).LastTrans
   DCCustRec(1).LastTrans = NextTransRec
   PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
   GET DCTransFile, Prev!, DCTransRec(1)
   DCTransRec(1).NextTrans = NextTransRec
   PUT DCTransFile, Prev!, DCTransRec(1)
  END IF
  NextTransRec = NextTransRec + 1

   ' Post Transaction Record First
  DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
  DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
  DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
  DCTransRec(1).TransType = 2               ' Type 2 = Payment
  DCTransRec(1).TransDesc = EditPaymentRec(1).Desc
  DCTransRec(1).CashAmount = EditPaymentRec(1).Amount
  DCTransRec(1).ChkAmount = 0
  DCTransRec(1).BalanceAfterTrans = DCTransRec(1).BalanceAfterTrans - EditPaymentRec(1).Amount
  DCTransRec(1).ExtraRoom = ""
  DCTransRec(1).NextTrans = 0
  PUT DCTransFile, NextTransRec, DCTransRec(1)

  GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
   DCCustRec(1).AcctBal = DCCustRec(1).AcctBal - EditPaymentRec(1).Amount
   DCCustRec(1).LICENSE = EditPaymentRec(1).Sticker
  PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)

  IF DCCustRec(1).FirstTrans = 0 THEN
   DCCustRec(1).FirstTrans = NextTransRec
   DCCustRec(1).LastTrans = NextTransRec
   PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
  ELSE
   Prev! = DCCustRec(1).LastTrans
   DCCustRec(1).LastTrans = NextTransRec
   PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
   GET DCTransFile, Prev!, DCTransRec(1)
   DCTransRec(1).NextTrans = NextTransRec
   PUT DCTransFile, Prev!, DCTransRec(1)
  END IF
  NextTransRec = NextTransRec + 1
 END IF

 LOOP UNTIL Cnt > NumOfDcRecs
 CLOSE
 KILL "DCEdPYT.DAT"
 ' Show All Posted
   PRINT CHR$(7);
   SaveScrn TempScrn()
   DisplayDCScrn "DCPOSTED"
   WaitForAction
   RestScrn TempScrn()
   LOCATE , , 1
   CLOSE

 RETURN

END SUB

SUB PrintEditList
  SHARED Choice$()
  ReportFile$ = "DCPAYED.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 53
  LPTPort% = 1
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


  CustRecLen = LEN(DCCustRec(1))
  TrHandle = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen

  DCEditRecLen = LEN(EditPaymentRec(1))
  DCFile = FREEFILE
  OPEN "DCEDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCFile) \ DCEditRecLen

  Rpthandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #Rpthandle

  GOSUB PrintRptHeader

  FOR Cnt = 1 TO NumOfDcRecs
   GET DCFile, Cnt, EditPaymentRec(1)
    GET TrHandle, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)

   IF LineCnt >= MaxLines THEN
    PRINT #Rpthandle, FF$
    GOSUB PrintRptHeader
   END IF
   IF EditPaymentRec(1).Amount <> 0 THEN
   PRINT #Rpthandle, VAL(DCCustRec(1).CUSTNUMB); TAB(10); DCCustRec(1).BILLNAME; TAB(55); USING "$$#####,#.##"; EditPaymentRec(1).Amount;
    PRINT #Rpthandle, TAB(73); EditPaymentRec(1).DecalCat
   PRINT #Rpthandle, "Veh Desc: "; EditPaymentRec(1).MakeModel; TAB(50); "State Tag# "; RTRIM$(EditPaymentRec(1).StateTag)
   PRINT #Rpthandle, STRING$(80, "-")
    TotalCust = TotalCust + 1
    TotalValue# = TotalValue# + EditPaymentRec(1).Amount
    TotalValue# = INT((TotalValue# * 100) + .5) / 100
    LineCnt = LineCnt + 3
   END IF

  NEXT Cnt
    GOSUB PrintRptEnding
    PRINT #Rpthandle, CHR$(18); ' oki 320 10 cpi
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
    PRINT #Rpthandle, TAB(15); "Decal System : Payment (Decal Purchase's) Edit Listing"
    PRINT #Rpthandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
    PRINT #Rpthandle, ""
    PRINT #Rpthandle, "Cust #"; TAB(10); "Billing Name"; TAB(56); "Payment Amt"; TAB(73); "Type"
    PRINT #Rpthandle, STRING$(80, "=")
    LineCnt = 5
    RETURN

PrintRptEnding:
     PRINT #Rpthandle, "Number of Entries .. "; USING "####,#"; TotalCust;
     PRINT #Rpthandle, TAB(55); USING "$$#####,#.##"; TotalValue#
     PRINT #Rpthandle, FF$
     RETURN




SelectOutput:
  LibName$ = "DC"
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
  printhelp help$
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

SUB ShowNoCodes
  LibName$ = "DC"
  ScrnName$ = "DCNOCODE"
  help$ = "NEW Customer Entry"
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
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE EscKey
     EXIT SUB
    END SELECT

  LOOP




END SUB

SUB SortDCNameIndex
  SHARED Mchoice$


  size = 2500
  Start = 1               'start at array element 1
  Dir = 0                 'sort direction - use anything else for descending
  SSize = 16               'total size of each TYPE element
  MOff = 0                'offset into the TYPE for the key element
  MSize = 7              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

  DIM array(1 TO size)  AS Struct
  help$ = "Sorting Customer Index"
  printhelp help$

  DCCustRecLen = LEN(DCCustRec(1))
  DCFile = FREEFILE
  OPEN "DCCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDcRecs = LOF(DCFile) \ DCCustRecLen

  DCCustIdxRecLen = LEN(DCCustIdxRec(1))
  DCIdxFile = FREEFILE
  OPEN "DCCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS DCIdxFile LEN = DCCustIdxRecLen

 FOR Cnt = 1 TO NumOfDcRecs
    GET DCFile, Cnt, DCCustRec(1)
     array(Cnt).who = DCCustRec(1).SORTNAME + "    "
     array(Cnt).RecNum = Cnt
 NEXT Cnt

 SortT array(Start), NumOfDcRecs, Dir, SSize, MOff, MSize

 FOR Cnt = 1 TO NumOfDcRecs
   DCCustIdxRec(1).IDXNAME = array(Cnt).who
   DCCustIdxRec(1).IDXRECORD = array(Cnt).RecNum
   PUT DCIdxFile, Cnt, DCCustIdxRec(1)
 NEXT Cnt
 CLOSE DCFile
 CLOSE DCIdxFile
END SUB

