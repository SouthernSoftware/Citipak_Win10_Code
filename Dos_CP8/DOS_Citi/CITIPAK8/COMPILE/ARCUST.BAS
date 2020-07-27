DEFINT A-Z
DECLARE SUB CustMailLbl ()
DECLARE SUB SetLicense ()
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE SUB SortARNameIndex ()
DECLARE SUB AddCustomer ()
DECLARE SUB EditCustomer ()
DECLARE SUB PrintCustomer ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE
   
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

   DIM SHARED ARCust(1) AS ARCustRecType
   DIM SHARED ARCustRec(1) AS ARCustRecType
   DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType

   STACK 8000
   CONST False = 0, True = NOT False

   '--Dim the choice array to the number of menu items
   REDIM Mchoice$(1 TO 6)

   Mchoice$(1) = "Add New Customer"
   Mchoice$(2) = "Edit Existing Customer"
   Mchoice$(3) = "Print Customer Listing"
   Mchoice$(4) = "Set Customer License's to Print"
   Mchoice$(5) = "Print Customer Mailing Labels"
   Mchoice$(6) = "Exit to OS"

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
   Col = ((80 - MaxLen) \ 2)
   help$ = "Add/Edit/Print Customers"
   
   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, "Customer Maintenance ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      PrintTitle user$
      PrintHelp help$

      ShowCursor

      VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      
      IF Ky$ = CHR$(27) THEN EXIT DO 'choice = 0

      SELECT CASE Choice
          CASE 1
           AddCustomer
          CASE 2
           EditCustomer
          CASE 3
           PrintCustomer
          CASE 4
           SetLicense
          CASE 5
           CustMailLbl
          CASE 6
          END
      END SELECT
   LOOP
   RUN "armenu"

SUB AddCustomer

mainbody:
  LibName$ = "AR"
  ScrnName$ = "ARCUST"
  help$ = "NEW A/R Customer Entry"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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

  REM check for code file

  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  CLOSE ARCatFile
  IF NumOFARCatRecs = 0 THEN
   ShowNoCodes
   EXIT SUB
  END IF



  OpenARCustFile NumOfArRecs, ARFile
  
 

  Form$(16, 0) = "N"
  Form$(21, 0) = "0"
  Fld(1).Protected = True
  'Fld(24).Protected = True

  Frm(1).FldNo = 2
  
  LOCATE 5, 28: COLOR 15: PRINT "        PENDING"
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF Frm(1).FldNo = 10 AND LEFT$(Form$(10, 0), 1) = " " THEN
      GOSUB SelectCatagory
      ShowCursor
      LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp help$
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
       SortARNameIndex
      END IF

      EXIT SUB
    END SELECT

  LOOP

    
SaveRecord:
    IF LEFT$(Form$(10, 0), 1) = " " THEN
     CLOSE ARFile
     CLOSE ARCatFile
     ELSE
    ARCustRec(1).SORTNAME = Form$(2, 0)
    ARCustRec(1).BILLNAME = Form$(3, 0)
    ARCustRec(1).Address1 = Form$(4, 0)
    ARCustRec(1).Address2 = Form$(5, 0)
    ARCustRec(1).CITY = Form$(6, 0)
    ARCustRec(1).STATE = Form$(7, 0)
    ARCustRec(1).ZipCode = Form$(8, 0)
    ARCustRec(1).CustName = Form$(9, 0)
    ARCustRec(1).BILLCAT = Form$(10, 0)
    ARCustRec(1).SOSEC = Form$(11, 0)
    ARCustRec(1).DRVLIC = Form$(12, 0)
    ARCustRec(1).DATEOPED = Date2Num(Form$(13, 0))
    ARCustRec(1).BILLCMT = Form$(14, 0)
    ARCustRec(1).PAYCMT = Form$(15, 0)
    ARCustRec(1).CASHONLY = Form$(16, 0)
    ARCustRec(1).APPNUMB = CVI(MID$(Form$(0, 0), Fld(17).Fields, 2))   'INTEGER
    ARCustRec(1).BILLFORM = CVI(MID$(Form$(0, 0), Fld(18).Fields, 2))  'INTEGER
    ARCustRec(1).HPHONE = Form$(19, 0)
    ARCustRec(1).WPHONE = Form$(20, 0)
    ARCustRec(1).FeeAmt = CVD(MID$(Form$(0, 0), Fld(21).Fields, 8))    'DOUBLE
    ARCustRec(1).LICENSE = Form$(22, 0)
    ARCustRec(1).VALID = Date2Num(Form$(23, 0))       'INTEGER Date Function
    ARCustRec(1).AcctBal = 0
    ARCustRec(1).FirstTrans = 0
    ARCustRec(1).LastTrans = 0
    ARCustRec(1).Deleted = "N"
    ARCustRec(1).IssueLicense = "N"
    ARCustRec(1).RoomtoGrow = ""
    NextAccount = NumOfArRecs + 1
    ARCustRec(1).CUSTNUMB = STR$(NextAccount)
    PUT ARFile, NextAccount, ARCustRec(1)
    LOCATE 5, 28: COLOR 15: PRINT NextAccount; "  ASSIGNED": SLEEP 2
    CLOSE ARFile
    CLOSE ARCatFile
    NeedtoSort = True
    END IF
    RETURN

SelectCatagory:

  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  
  ARCatFile = FREEFILE
  
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
 
  REDIM Mchoice$(1 TO NumOFARCatRecs)
  FOR Cnt = 1 TO NumOFARCatRecs
    GET ARCatFile, Cnt, ARCatCodeRec(1)
    Mchoice$(Cnt) = SPACE$(50)
    LSET Mchoice$(Cnt) = ARCatCodeRec(1).CATCODE
    MID$(Mchoice$(Cnt), 5) = ARCatCodeRec(1).CODEDESC
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
   GET ARCatFile, Choice, ARCatCodeRec(1)

   Form$(10, 0) = ARCatCodeRec(1).CATCODE
   Form$(17, 0) = STR$(ARCatCodeRec(1).APPNUMB)
   Form$(18, 0) = STR$(ARCatCodeRec(1).BILLCODE)

   Fld(17).Protected = True
   Fld(18).Protected = True
   
   Frm(1).FldNo = 11
RETURN


END SUB

SUB EditCustomer

EditMainBody:
  SHARED Mchoice$
  LibName$ = "AR"
  ScrnName$ = "ARCUST"
  help$ = "Edit A/R Customer Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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

  CLOSE ARFile
  OpenARCustFile NumOfArRecs, ARFile

  OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile

  IF NumOfArRecs = 0 THEN
   LibName$ = "AR"
   ScrnName$ = "ARNOCUST"
   help$ = "Edit A/R Customer Entry"
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

  DO
  Done = False
  EditForm Form$(), Fld(), Frm(1), Cnf, Action

  SELECT CASE Frm(1).KeyCode
   CASE EscKey
    Done = True
  END SELECT
    IF Done = True THEN EXIT SUB
  LOOP

  END IF
  COLOR 15
  LOCATE 5, 40: PRINT "PRESS <enter> FOR CUSTOMER LIST"




  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 2 AND CustomerGrabed = 0 THEN
     GOSUB GetCustomer
     LOCK #ARFile, AccountRecord
    END IF

    IF Frm(1).FldNo = 10 AND LEFT$(Form$(10, 0), 1) = " " THEN
      GOSUB EditSelectCatagory
      ShowCursor
      LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp help$

    END IF

    SELECT CASE Frm(1).KeyCode


    CASE F3Key
    IF AccountRecord > 0 THEN
      GOSUB DeleteRecord
      IF NeedtoSort = True THEN
       SortARNameIndex
      END IF
      CLOSE ARFile
      EXIT SUB
    END IF

    CASE F10Key
      GOSUB EditSaveRecord
      IF NeedtoSort = True THEN
       SortARNameIndex
      END IF
      CLOSE ARFile
      EXIT SUB
    CASE EscKey
      IF NeedtoSort = True THEN
       SortARNameIndex
      END IF
      CLOSE ARFile
      EXIT SUB
    END SELECT

  LOOP

  

EditSaveRecord:
    IF AccountRecord = 0 THEN RETURN
    ARCustRec(1).CUSTNUMB = Form$(1, 0)
    ARCustRec(1).SORTNAME = Form$(2, 0)
    ARCustRec(1).BILLNAME = Form$(3, 0)
    ARCustRec(1).Address1 = Form$(4, 0)
    ARCustRec(1).Address2 = Form$(5, 0)
    ARCustRec(1).CITY = Form$(6, 0)
    ARCustRec(1).STATE = Form$(7, 0)
    ARCustRec(1).ZipCode = Form$(8, 0)
    ARCustRec(1).CustName = Form$(9, 0)
    ARCustRec(1).BILLCAT = Form$(10, 0)
    ARCustRec(1).SOSEC = Form$(11, 0)
    ARCustRec(1).DRVLIC = Form$(12, 0)
    ARCustRec(1).DATEOPED = Date2Num(Form$(13, 0))
    ARCustRec(1).BILLCMT = Form$(14, 0)
    ARCustRec(1).PAYCMT = Form$(15, 0)
    ARCustRec(1).CASHONLY = Form$(16, 0)
    ARCustRec(1).APPNUMB = CVI(MID$(Form$(0, 0), Fld(17).Fields, 2))   'INTEGER
    ARCustRec(1).BILLFORM = CVI(MID$(Form$(0, 0), Fld(18).Fields, 2))                                                       'INTEGER
    ARCustRec(1).HPHONE = Form$(19, 0)
    ARCustRec(1).WPHONE = Form$(20, 0)
    ARCustRec(1).FeeAmt = CVD(MID$(Form$(0, 0), Fld(21).Fields, 8))                                                      'INTEGER
    ARCustRec(1).LICENSE = Form$(22, 0)
    ARCustRec(1).VALID = Date2Num(Form$(23, 0))       'INTEGER Date Function
    ARCustRec(1).IssueLicense = Form$(24, 0)
    PUT ARFile, AccountRecord, ARCustRec(1)
    CLOSE ARFile
    NeedtoSort = True
    RETURN

EditSelectCatagory:

  REDIM ARCatCodeRec(1) AS ARCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen

  REDIM Mchoice$(1 TO NumOFARCatRecs)
  FOR Cnt = 1 TO NumOFARCatRecs
    GET ARCatFile, Cnt, ARCatCodeRec(1)
    Mchoice$(Cnt) = SPACE$(50)
    LSET Mchoice$(Cnt) = ARCatCodeRec(1).CATCODE
    MID$(Mchoice$(Cnt), 5) = ARCatCodeRec(1).CODEDESC
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
   GET ARCatFile, Choice, ARCatCodeRec(1)

   Form$(10, 0) = ARCatCodeRec(1).CATCODE
   Form$(17, 0) = STR$(ARCatCodeRec(1).APPNUMB)
   Form$(18, 0) = STR$(ARCatCodeRec(1).BILLCODE)

   Fld(17).Protected = True
   Fld(18).Protected = True
   Frm(1).FldNo = 11
RETURN





GetCustomer:

CustomerGrabed = 0
AccountRecord = VAL(Form$(1, 0))

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
  ScrnName$ = "ARCUST"
  help$ = "Edit A/R Customer Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%

   END IF

REM ************************************************************************
IF AccountRecord > 0 AND AccountRecord <= NumOfArRecs THEN
    GET ARFile, AccountRecord, ARCustRec(1)
    IF ARCustRec(1).Deleted = "Y" THEN GOTO CustomerDeleted
    Form$(1, 0) = ARCustRec(1).CUSTNUMB
    Form$(2, 0) = ARCustRec(1).SORTNAME
    Form$(3, 0) = ARCustRec(1).BILLNAME
    Form$(4, 0) = ARCustRec(1).Address1
    Form$(5, 0) = ARCustRec(1).Address2
    Form$(6, 0) = ARCustRec(1).CITY
    Form$(7, 0) = ARCustRec(1).STATE
    Form$(8, 0) = ARCustRec(1).ZipCode
    Form$(9, 0) = ARCustRec(1).CustName
    Form$(10, 0) = ARCustRec(1).BILLCAT
    Form$(11, 0) = ARCustRec(1).SOSEC
    Form$(12, 0) = ARCustRec(1).DRVLIC
    Form$(13, 0) = Num2Date(ARCustRec(1).DATEOPED)
    Form$(14, 0) = ARCustRec(1).BILLCMT
    Form$(15, 0) = ARCustRec(1).PAYCMT
    Form$(16, 0) = ARCustRec(1).CASHONLY
    Form$(17, 0) = STR$(ARCustRec(1).APPNUMB)
    Form$(18, 0) = STR$(ARCustRec(1).BILLFORM)
    Form$(19, 0) = ARCustRec(1).HPHONE
    Form$(20, 0) = ARCustRec(1).WPHONE
    Form$(21, 0) = STR$(ARCustRec(1).FeeAmt)
    Form$(22, 0) = ARCustRec(1).LICENSE
    Form$(23, 0) = Num2Date$(ARCustRec(1).VALID)
    Form$(24, 0) = ARCustRec(1).IssueLicense
    Fld(1).Protected = True
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    LOCATE 5, 40: PRINT STRING$(39, 32)
    RETURN

   ELSE

    LibName$ = "AR"
    ScrnName$ = "ARBADCUS"
    help$ = "Edit A/R Customer Entry"
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
     IF Done = True THEN GOTO EditMainBody
  LOOP

 END IF

CustomerDeleted:
    LibName$ = "AR"
    ScrnName$ = "ARDELCUS"
    help$ = "Edit Customer"
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
     GOTO EditMainBody
    END SELECT
  LOOP


DeleteRecord:
  LibName$ = "AR"
  ScrnName$ = "ARCUSDEL"
  help$ = "Delete Customer"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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
  Form$(1, 0) = "Y"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode


    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
       help$ = "Account Deleted!!!"
       PrintHelp help$
       PRINT CHR$(7);
       ARCustRec(1).Deleted = "Y"
       ARCustRec(1).SORTNAME = "DELETED"
       PUT ARFile, AccountRecord, ARCustRec(1)
       CLOSE ARFile
       NeedtoSort = True
       RETURN
      END IF
    CASE EscKey
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
  'FOR x = 1 TO NumOfArRecs
  'GET ARFile, x, ARCust(1)
  'PRINT ARCust(1).Custnumb; TAB(15); ARCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP
   END SUB

SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

SUB PrintCustomer

  SHARED Choice$()
  ReportFile$ = "ARCUST.PRN"  'Report File Name
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
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

  REDIM array(1 TO size) AS Struct

  GOSUB SelectOutput

  'REDIM ARCustRec(1) AS ARCustRecType     ' open customer file
  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen

  'REDIM ARCustIdxRec(1) AS ARCustIdxRecType     ' open customer file
  IdxCustRecLen = LEN(ARCustIdxRec(1))
  IdxTrHandle = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS IdxTrHandle LEN = IdxCustRecLen
  IdxTrNumRecs = LOF(IdxTrHandle) \ IdxCustRecLen

   
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintRptHeader

  FOR Cnt = 1 TO IdxTrNumRecs
   GET IdxTrHandle, Cnt, ARCustIdxRec(1)
    GET TrHandle, ARCustIdxRec(1).IDXRECORD, ARCustRec(1)
IF ARCustRec(1).Deleted <> "Y" THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   PRINT #RptHandle, VAL(ARCustRec(1).CUSTNUMB); TAB(10); ARCustRec(1).BILLNAME; TAB(50); ARCustRec(1).BILLCAT;
    PRINT #RptHandle, TAB(65); USING "$$####,#.##"; ARCustRec(1).FeeAmt
    PRINT #RptHandle, TAB(10); ARCustRec(1).CustName; TAB(50); ARCustRec(1).LICENSE; TAB(65); Num2Date$(ARCustRec(1).VALID)
     PRINT #RptHandle, STRING$(79, "-")
    TotalCust = TotalCust + 1
    LineCnt = LineCnt + 3
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

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintRptHeader:
    page = page + 1
    PRINT #RptHandle, TAB(18); "Business License : Customer 'Quick' Listing"
    PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Cust #"; TAB(10); "Billing Name"; TAB(48); "Catagory"; TAB(65); "Fee Amount"
    PRINT #RptHandle, TAB(10); "Customer Name"; TAB(48); "License #"; TAB(65); "Valid To"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 5
    RETURN

PrintRptEnding:
     PRINT #RptHandle, STRING$(80, "-")
     PRINT #RptHandle, "Number of Customers .. "; USING "####,#"; TotalCust
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
  PrintHelp help$
  Action = 1
  COLOR 14: LOCATE 9, 23: PRINT "Customer Listing"
  

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

SUB SetLicense

  LibName$ = "AR"
  ScrnName$ = "ARCNGLIC"

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
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  help$ = "Change License Print Status"
  PrintHelp help$

  Action = 1



  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       DevSpec$ = LEFT$(Form$(1, 0), 1)
       Done = True
     CASE EscKey
      Canceled$ = "Y"
      Done = True
      END SELECT
     LOOP UNTIL Done
     IF Canceled$ = "Y" THEN EXIT SUB



  'REDIM ARCustRec(1) AS ARCustRecType     ' open customer file
  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
 
  FOR Cnt = 1 TO TrNumRecs
   GET TrHandle, Cnt, ARCustRec(1)
IF ARCustRec(1).Deleted <> "Y" THEN
   IF Num2Date$(ARCustRec(1).VALID) = Form$(1, 0) THEN
   help$ = ARCustRec(1).CustName
   PrintHelp help$
    ARCustRec(1).IssueLicense = "Y"
    PUT TrHandle, Cnt, ARCustRec(1)
   END IF
END IF
  NEXT Cnt
    CLOSE                       'Close all open files now
    EXIT SUB







END SUB

SUB ShowNoCodes
  LibName$ = "AR"
  ScrnName$ = "ARNOCODE"
  help$ = "NEW A/R Customer Entry"
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
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE EscKey
     EXIT SUB
    END SELECT

  LOOP




END SUB

SUB SortARNameIndex
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
  PrintHelp help$

  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen

  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  
 FOR Cnt = 1 TO NumOfArRecs
    GET ARFile, Cnt, ARCustRec(1)
     array(Cnt).who = ARCustRec(1).SORTNAME + "    "
     array(Cnt).RecNum = Cnt
 NEXT Cnt

 SortT array(Start), NumOfArRecs, Dir, SSize, MOff, MSize

 FOR Cnt = 1 TO NumOfArRecs
   ARCustIdxRec(1).IDXNAME = array(Cnt).who
   ARCustIdxRec(1).IDXRECORD = array(Cnt).RecNum
   PUT ARIdxFile, Cnt, ARCustIdxRec(1)
 NEXT Cnt
 CLOSE ARFile
 CLOSE ARIdxFile
END SUB

