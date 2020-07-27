DEFINT A-Z
DECLARE SUB ClearCustomer ()
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenDCCustIdxFile (NumOfDCIdxRecs%, DCIdxFile%)
DECLARE SUB OpenDCCustFile (NumOfDcRecs%, DCFile%)
DECLARE SUB SortDCNameIndex ()
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
DECLARE SUB QPrintRC (T$, r%, C%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()

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
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
   DIM SHARED DCCust(1) AS DCCustRecType
   DIM SHARED DCCustRec(1) AS DCCustRecType
   DIM SHARED DCCustIdxRec(1) AS DCCustIDXRecType

   STACK 8000
   CONST False = 0, True = NOT False

   '--Dim the choice array to the number of menu items
   REDIM Mchoice$(1 TO 5)

   Mchoice$(1) = "Add New Customer"
   Mchoice$(2) = "Edit Existing Customer"
   Mchoice$(3) = "Print Customer Listing"
   Mchoice$(4) = "Clear Customer Decal (New Year)"
   Mchoice$(5) = "Exit to OS"

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
   Help$ = "Add/Edit/Print Customers"
   
   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, "Customer Maintenance ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      PrintTitle user$
      PrintHelp Help$

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
           ClearCustomer
          CASE 5
          END
      END SELECT
   LOOP
   RUN "dcmenu"

SUB AddCustomer

mainbody:
  LibName$ = "DC"
  ScrnName$ = "DCCUST"
  Help$ = "NEW Decal Customer Entry"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp Help$



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
    DCCustRec(1).BillName = Form$(3, 0)
    DCCustRec(1).ADDRESS1 = Form$(4, 0)
    DCCustRec(1).ADDRESS2 = Form$(5, 0)
    DCCustRec(1).CITY = Form$(6, 0)
    DCCustRec(1).STATE = Form$(7, 0)
    DCCustRec(1).ZIPCODE = Form$(8, 0)
    DCCustRec(1).SOSEC = Form$(9, 0)
    DCCustRec(1).DRVLIC = Form$(11, 0)
    DCCustRec(1).DATEOPED = Date2Num(Form$(13, 0))
    DCCustRec(1).BILLCMT = Form$(15, 0)
    DCCustRec(1).PAYCMT = Form$(17, 0)
    DCCustRec(1).CASHONLY = Form$(18, 0)
    DCCustRec(1).Resident = Form$(10, 0)
    DCCustRec(1).Owner = Form$(12, 0)
    DCCustRec(1).HPHONE = Form$(14, 0)
    DCCustRec(1).WPHONE = Form$(16, 0)
    DCCustRec(1).FeeAmt = 0
    DCCustRec(1).License = ""
    DCCustRec(1).Valid = -32767     'INTEGER Date Function
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

SUB ClearCustomer

  SHARED Choice$()

  GOSUB SelectFunction
  IF NOT (ClearIt) THEN EXIT SUB
  'REDIM DCCustRec(1) AS DCCustRecType     ' open customer file
  CustRecLen = LEN(DCCustRec(1))
  TrHandle = FREEFILE
  OPEN "DCCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen

  'REDIM DCCustIdxRec(1) AS DCCustIdxRecType     ' open customer file
  IdxCustRecLen = LEN(DCCustIdxRec(1))
  IdxTrHandle = FREEFILE
  OPEN "DCCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS IdxTrHandle LEN = IdxCustRecLen
  IdxTrNumRecs = LOF(IdxTrHandle) \ IdxCustRecLen


  FOR Cnt = 1 TO IdxTrNumRecs
   GET IdxTrHandle, Cnt, DCCustIdxRec(1)
    GET TrHandle, DCCustIdxRec(1).IdxRecord, DCCustRec(1)
    Help$ = DCCustRec(1).BillName
    PrintHelp Help$
IF DCCustRec(1).Deleted <> "Y" THEN
   DCCustRec(1).License = ""
   PUT TrHandle, DCCustIdxRec(1).IdxRecord, DCCustRec(1)
END IF
  NEXT Cnt
    CLOSE                       'Close all open files now
    Help$ = "ALL ACCOUNTS CLEARED"
    PRINT CHR$(7);
    PrintHelp Help$
    SLEEP 5
  EXIT SUB



SelectFunction:
  LibName$ = "DC"
  ScrnName$ = "DCCLEAR"

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
  PrintHelp Help$
  Action = 1
  


  DO


    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
     CASE F10Key
       ClearIt = True
       RETURN
     CASE EscKey
      ClearIt = False
      Canceled$ = "Y"
      RETURN
    END SELECT
 LOOP
  RETURN



END SUB

SUB EditCustomer

EditMainBody:
  SHARED Mchoice$
  LibName$ = "DC"
  ScrnName$ = "DCCUST"
  Help$ = "Edit Customer Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp Help$



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

  CLOSE DCFile
  OpenDCCustFile NumOfDcRecs, DCFile

  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile

  IF NumOfDcRecs = 0 THEN
   LibName$ = "DC"
   ScrnName$ = "DCNOCUST"
   Help$ = "Edit Customer Entry"
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
  PrintHelp Help$

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
     LOCK #DCFile, AccountRecord
    END IF

    IF Frm(1).FldNo = 9 AND LEFT$(Form$(9, 0), 1) = " " THEN
      GOSUB EditSelectCatagory
      ShowCursor
      LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      PrintHelp Help$

    END IF

    SELECT CASE Frm(1).KeyCode


    CASE F3Key
    IF AccountRecord > 0 THEN
      GOSUB DeleteRecord
      IF NeedtoSort = True THEN
       SortDCNameIndex
      END IF
      CLOSE DCFile
      EXIT SUB
    END IF

    CASE F10Key
      GOSUB EditSaveRecord
      IF NeedtoSort = True THEN
       SortDCNameIndex
      END IF
      CLOSE DCFile
      EXIT SUB
    CASE EscKey
      IF NeedtoSort = True THEN
       SortDCNameIndex
      END IF
      CLOSE DCFile
      EXIT SUB
    END SELECT

  LOOP

  

EditSaveRecord:
    IF AccountRecord = 0 THEN RETURN
    DCCustRec(1).CUSTNUMB = Form$(1, 0)
    DCCustRec(1).SORTNAME = Form$(2, 0)
    DCCustRec(1).BillName = Form$(3, 0)
    DCCustRec(1).ADDRESS1 = Form$(4, 0)
    DCCustRec(1).ADDRESS2 = Form$(5, 0)
    DCCustRec(1).CITY = Form$(6, 0)
    DCCustRec(1).STATE = Form$(7, 0)
    DCCustRec(1).ZIPCODE = Form$(8, 0)
    DCCustRec(1).SOSEC = Form$(9, 0)
    DCCustRec(1).DRVLIC = Form$(11, 0)
    DCCustRec(1).DATEOPED = Date2Num(Form$(13, 0))
    DCCustRec(1).BILLCMT = Form$(15, 0)
    DCCustRec(1).PAYCMT = Form$(17, 0)
    DCCustRec(1).CASHONLY = Form$(18, 0)
    DCCustRec(1).Resident = Form$(10, 0)
    DCCustRec(1).Owner = Form$(12, 0)
    DCCustRec(1).HPHONE = Form$(14, 0)
    DCCustRec(1).WPHONE = Form$(16, 0)
    
    PUT DCFile, AccountRecord, DCCustRec(1)
    CLOSE DCFile
    NeedtoSort = True
    RETURN

EditSelectCatagory:

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

   REDIM Mchoice$(1 TO NumOfDCIdxRecs)

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfDCIdxRecs
    GET DCIdxFile, Cnt, DCCustIdxRec(1)
    IF LEFT$(DCCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
     ChoiceCounter = ChoiceCounter + 1
     Mchoice$(ChoiceCounter) = SPACE$(50)
     LSET Mchoice$(ChoiceCounter) = STR$(DCCustIdxRec(1).IdxRecord)
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
  ScrnName$ = "DCCUST"
  Help$ = "Edit A/R Customer Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%

   END IF

REM ************************************************************************
IF AccountRecord > 0 AND AccountRecord <= NumOfDcRecs THEN
    GET DCFile, AccountRecord, DCCustRec(1)
    IF DCCustRec(1).Deleted = "Y" THEN GOTO CustomerDeleted
    Form$(1, 0) = DCCustRec(1).CUSTNUMB
    Form$(2, 0) = DCCustRec(1).SORTNAME
    Form$(3, 0) = DCCustRec(1).BillName
    Form$(4, 0) = DCCustRec(1).ADDRESS1
    Form$(5, 0) = DCCustRec(1).ADDRESS2
    Form$(6, 0) = DCCustRec(1).CITY
    Form$(7, 0) = DCCustRec(1).STATE
    Form$(8, 0) = DCCustRec(1).ZIPCODE
    Form$(9, 0) = DCCustRec(1).SOSEC
    Form$(11, 0) = DCCustRec(1).DRVLIC
    Form$(13, 0) = Num2Date(DCCustRec(1).DATEOPED)
    Form$(15, 0) = DCCustRec(1).BILLCMT
    Form$(17, 0) = DCCustRec(1).PAYCMT
    Form$(18, 0) = DCCustRec(1).CASHONLY
    Form$(10, 0) = DCCustRec(1).Resident
    Form$(12, 0) = DCCustRec(1).Owner
    Form$(14, 0) = DCCustRec(1).HPHONE
    Form$(16, 0) = DCCustRec(1).WPHONE
    
    Fld(1).Protected = True
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    LOCATE 5, 40: PRINT STRING$(39, 32)
    RETURN

   ELSE

    LibName$ = "DC"
    ScrnName$ = "DCBADCUS"
    Help$ = "Edit A/R Customer Entry"
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
    PrintHelp Help$

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
    LibName$ = "DC"
    ScrnName$ = "ARDELCUS"
    Help$ = "Edit Customer"
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
    PrintHelp Help$

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
  LibName$ = "DC"
  ScrnName$ = "DCCUSDEL"
  Help$ = "Delete Customer"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp Help$
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
       Help$ = "Account Deleted!!!"
       PrintHelp Help$
       PRINT CHR$(7);
       DCCustRec(1).Deleted = "Y"
       DCCustRec(1).SORTNAME = "DELETED"
       PUT DCFile, AccountRecord, DCCustRec(1)
       CLOSE DCFile
       NeedtoSort = True
       RETURN
      END IF
    CASE EscKey
       RETURN
    END SELECT








  LOOP


END SUB

SUB OpenDCCustFile (NumOfDcRecs, DCFile)
  CLOSE DCFile
  
  DCCustRecLen = LEN(DCCustRec(1))
  DCFile = FREEFILE
  OPEN "DCCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDcRecs = LOF(DCFile) \ DCCustRecLen
  'FOR x = 1 TO NumOfDcRecs
  'GET DCFile, x, DCCust(1)
  'PRINT DCCust(1).Custnumb; TAB(15); DCCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP
   END SUB

SUB OpenDCCustIdxFile (NumOfDCIdxRecs, DCIdxFile)
  CLOSE DCIdxFile
  DCCustIdxRecLen = LEN(DCCustIdxRec(1))
  DCIdxFile = FREEFILE
  OPEN "DCCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS DCIdxFile LEN = DCCustIdxRecLen
  NumOfDCIdxRecs = LOF(DCIdxFile) \ DCCustIdxRecLen
END SUB

SUB PrintCustomer

  SHARED Choice$()
  ReportFile$ = "DCCUST.PRN"  'Report File Name
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

  'REDIM DCCustRec(1) AS DCCustRecType     ' open customer file
  CustRecLen = LEN(DCCustRec(1))
  TrHandle = FREEFILE
  OPEN "DCCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen

  'REDIM DCCustIdxRec(1) AS DCCustIdxRecType     ' open customer file
  IdxCustRecLen = LEN(DCCustIdxRec(1))
  IdxTrHandle = FREEFILE
  OPEN "DCCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS IdxTrHandle LEN = IdxCustRecLen
  IdxTrNumRecs = LOF(IdxTrHandle) \ IdxCustRecLen

   
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintRptHeader

  FOR Cnt = 1 TO IdxTrNumRecs
   GET IdxTrHandle, Cnt, DCCustIdxRec(1)
    GET TrHandle, DCCustIdxRec(1).IdxRecord, DCCustRec(1)
IF DCCustRec(1).Deleted <> "Y" THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   PRINT #RptHandle, VAL(DCCustRec(1).CUSTNUMB); TAB(10); DCCustRec(1).BillName;
    PRINT #RptHandle, TAB(50); DCCustRec(1).BILLCAT;
    PRINT #RptHandle, TAB(60); USING "$####.##"; DCCustRec(1).FeeAmt;
      PRINT #RptHandle, TAB(70); Num2Date$(DCCustRec(1).Valid)
     TotalCust = TotalCust + 1
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

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintRptHeader:
    page = page + 1
    PRINT #RptHandle, TAB(18); "Business License : Customer 'Quick' Listing"
    PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"; TAB(48); "Catagory"; TAB(59); "Fee Amount"; TAB(72); "Expires"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 5
    RETURN

PrintRptEnding:
     PRINT #RptHandle, STRING$(80, "-")
     PRINT #RptHandle, "Number of Customers .. "; USING "####,#"; TotalCust
     PRINT #RptHandle, FF$
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
  PrintHelp Help$
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

SUB ShowNoCodes
  LibName$ = "DC"
  ScrnName$ = "DCNOCODE"
  Help$ = "NEW Customer Entry"
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
  Help$ = "Sorting Customer Index"
  PrintHelp Help$

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
   DCCustIdxRec(1).IdxRecord = array(Cnt).RecNum
   PUT DCIdxFile, Cnt, DCCustIdxRec(1)
 NEXT Cnt
 CLOSE DCFile
 CLOSE DCIdxFile
END SUB

