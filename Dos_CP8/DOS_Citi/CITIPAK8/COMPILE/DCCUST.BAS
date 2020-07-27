DEFINT A-Z
DECLARE SUB EditVehicles (CustName$, AccountRecord)
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
   DIM SHARED DCVRec(1) AS DCVehType

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
          CLOSE
          CLS
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
  Form$(11, 0) = DATE$
  Form$(14, 0) = "Y"
  Form$(15, 0) = "N"
  Form$(16, 0) = "Y"
  Form$(17, 0) = "N"

  LOCATE 5, 28: COLOR 15: PRINT "        PENDING"
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action


    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      GOSUB SaveRecord
      SortDCNameIndex
      Done = True
      GOSUB AddCars
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
    DCCustRec(1).SORTNAME = Form$(2, 0)
    DCCustRec(1).BILLNAME = Form$(3, 0)
    DCCustRec(1).ADDRESS1 = Form$(4, 0)
    DCCustRec(1).ADDRESS2 = Form$(5, 0)
    DCCustRec(1).CITY = Form$(6, 0)
    DCCustRec(1).STATE = Form$(7, 0)
    DCCustRec(1).ZIPCODE = Form$(8, 0)
    DCCustRec(1).SOSEC = Form$(9, 0)
    DCCustRec(1).DRVLIC = Form$(10, 0)
    DCCustRec(1).DATEOPED = Date2Num(Form$(11, 0))
    DCCustRec(1).CASHONLY = Form$(14, 0)
    DCCustRec(1).RESIDENT = Form$(15, 0)
    DCCustRec(1).OWNER = Form$(16, 0)
    DCCustRec(1).HPHONE = Form$(12, 0)
    DCCustRec(1).WPHONE = Form$(13, 0)
    DCCustRec(1).LICENSE = Form$(10, 0)
    DCCustRec(1).Valid = -32767     'INTEGER Date Function
    DCCustRec(1).AcctBal = 0
    DCCustRec(1).FirstTrans = 0
    DCCustRec(1).LastTrans = 0
    DCCustRec(1).Deleted = "N"
    DCCustRec(1).FirstCar = 0
    DCCustRec(1).LastCar = 0
    DCCustRec(1).RoomtoGrow = ""
    NextAccount = NumOfDcRecs + 1
    DCCustRec(1).CUSTNUMB = LTRIM$(STR$(NextAccount))
    PUT DCFile, NextAccount, DCCustRec(1)
    LOCATE 5, 28: COLOR 15: PRINT NextAccount; "  ASSIGNED  ": SLEEP 2
    NeedtoSort = True
    RETURN


AddCars:
  LibName$ = "DC"
  ScrnName$ = "DCVEHEN"
  Help$ = "NEW Vehicle Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp Help$
  COLOR 11
  LOCATE 5, 20: PRINT "Customer: "; : COLOR 15: PRINT DCCustRec(1).BILLNAME
 
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

  'Set Active Flag Here
  Form$(5, 0) = "Y"

  Action = 1
  
  PrintHelp Help$
  COLOR 11
  LOCATE 5, 20: PRINT "Customer: "; : COLOR 15: PRINT DCCustRec(1).BILLNAME


  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF Frm(1).FldNo = 1 THEN
     GOSUB Selectcatagory
     LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
     Action = 1
    END IF
  

    SELECT CASE Frm(1).KeyCode
    
    CASE F10Key
      GOSUB SaveVehRecord
      Done = True
      GOTO AddCars
    CASE EscKey
      CLOSE DCFile
      CLOSE DCCatFile
      RETURN
    END SELECT

  LOOP


RETURN

SaveVehRecord:
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen
    DCVRec(1).DecalCat = Form$(1, 0)
    DCVRec(1).Fee = Value#(Form$(2, 0), 0)
    DCVRec(1).MakeModel = Form$(6, 0)
    DCVRec(1).StateTag = Form$(7, 0)
    DCVRec(1).ExpireDate = Date2Num%(Form$(4, 0))
    DCVRec(1).Sticker = Form$(3, 0)
    DCVRec(1).Valid = Form$(5, 0)
    DCVRec(1).Active = "Y"
    DCVRec(1).Desc = Form$(8, 0)
    DCVRec(1).Notes = Form$(9, 0)
    DCVRec(1).MoreRoom = ""
    DCVRec(1).NextRec = 0
    DCVRec(1).MasterRecord = NextAccount
    PUT DCVFile, NumOfVrecs + 1, DCVRec(1)
    IF DCCustRec(1).FirstCar = 0 THEN
     DCCustRec(1).FirstCar = NumOfVrecs + 1
     DCCustRec(1).LastCar = NumOfVrecs + 1
     PUT DCFile, NextAccount, DCCustRec(1)
    ELSE
     PrevRec! = DCCustRec(1).LastCar
     DCCustRec(1).LastCar = NumOfVrecs + 1
     PUT DCFile, NextAccount, DCCustRec(1)

     GET DCVFile, PrevRec!, DCVRec(1)
     DCVRec(1).NextRec = NumOfVrecs + 1
     PUT DCVFile, PrevRec!, DCVRec(1)
    END IF
    CLOSE DCVFile
    RETURN

Selectcatagory:
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
   Form$(1, 0) = DCCatCodeRec(1).CATCODE
   Form$(2, 0) = STR$(DCCatCodeRec(1).Fee)
   Frm(1).FldNo = 3
   CLOSE DCCatFile
   RETURN
 
END SUB

SUB ClearCustomer

  SHARED Choice$()

  GOSUB SelectFunction
  IF NOT (ClearIt) THEN EXIT SUB

  'Open Vehicle File
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen


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
    GET TrHandle, DCCustIdxRec(1).IDXRECORD, DCCustRec(1)
    Help$ = DCCustRec(1).BILLNAME
    PrintHelp Help$

    CarRecord! = DCCustRec(1).FirstCar

    WHILE CarRecord! > 0
     GET DCVFile, CarRecord!, DCVRec(1)
       DCVRec(1).Valid = "N"
     PUT DCVFile, CarRecord!, DCVRec(1)
       CarRecord! = DCVRec(1).NextRec
    WEND

  NEXT Cnt
    CLOSE                       'Close all open files now

    REDIM DCOVRec(1) AS DCOldVehType
    DCOVRecLen = LEN(DCOVRec(1))
    DCOVFile = FREEFILE
    OPEN "DCOLDVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCOVFile LEN = DCOVRecLen
    CLOSE DCOVFile
    KILL "DCOLDVEH.DAT"
    Help$ = "ALL ACCOUNTS CLEARED"
    PRINT CHR$(7);
    PrintHelp Help$
    SLEEP 3
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
  ScrnName$ = "DCCUSTMA"
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

    CASE F5Key
      IF AccountRecord > 0 THEN
      UNLOCK #DCFile, AccountRecord
      CustName$ = DCCustRec(1).BILLNAME
      EditVehicles CustName$, AccountRecord
      END IF
      EXIT SUB


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
    DCCustRec(1).BILLNAME = Form$(3, 0)
    DCCustRec(1).ADDRESS1 = Form$(4, 0)
    DCCustRec(1).ADDRESS2 = Form$(5, 0)
    DCCustRec(1).CITY = Form$(6, 0)
    DCCustRec(1).STATE = Form$(7, 0)
    DCCustRec(1).ZIPCODE = Form$(8, 0)
    DCCustRec(1).SOSEC = Form$(9, 0)
    DCCustRec(1).DRVLIC = Form$(10, 0)
    DCCustRec(1).DATEOPED = Date2Num(Form$(11, 0))
    DCCustRec(1).CASHONLY = Form$(14, 0)
    DCCustRec(1).RESIDENT = Form$(15, 0)
    DCCustRec(1).OWNER = Form$(16, 0)
    DCCustRec(1).HPHONE = Form$(12, 0)
    DCCustRec(1).WPHONE = Form$(13, 0)
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
 
   MaxLen = 23     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = " Cust #    Customer Name"

   '--Center Menu within Screen
   Row = 8
   Col = 20
   PRINT FRE(a$)
   A$=INPUT$(1)
   REDIM Mchoice$(1 TO NumOfDCIdxRecs)

  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfDCIdxRecs
    GET DCIdxFile, Cnt, DCCustIdxRec(1)
    IF LEFT$(DCCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
    GET DCFile, DCCustIdxRec(1).IDXRECORD, DCCustRec(1)
     ChoiceCounter = ChoiceCounter + 1
     Mchoice$(ChoiceCounter) = SPACE$(40)
     LSET Mchoice$(ChoiceCounter) = STR$(DCCustIdxRec(1).IDXRECORD)
     MID$(Mchoice$(ChoiceCounter), 10) = RTRIM$(DCCustRec(1).BILLNAME)
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
  ScrnName$ = "DCCUSTMA"
  Help$ = "Edit A/R Customer Entry"
  LOCATE 1, 1, 0

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%

   END IF

REM ************************************************************************
IF AccountRecord > 0 AND AccountRecord <= NumOfDcRecs THEN
    GET DCFile, AccountRecord, DCCustRec(1)
    IF DCCustRec(1).Deleted = "Y" THEN GOTO CustomerDeleted
    Form$(1, 0) = RTRIM$(DCCustRec(1).CUSTNUMB)
    Form$(2, 0) = RTRIM$(DCCustRec(1).SORTNAME)
    Form$(3, 0) = DCCustRec(1).BILLNAME
    Form$(4, 0) = DCCustRec(1).ADDRESS1
    Form$(5, 0) = DCCustRec(1).ADDRESS2
    Form$(6, 0) = DCCustRec(1).CITY
    Form$(7, 0) = DCCustRec(1).STATE
    Form$(8, 0) = DCCustRec(1).ZIPCODE
    Form$(9, 0) = DCCustRec(1).SOSEC
    Form$(10, 0) = DCCustRec(1).DRVLIC
    Form$(11, 0) = Num2Date(DCCustRec(1).DATEOPED)
    Form$(14, 0) = DCCustRec(1).CASHONLY
    Form$(15, 0) = DCCustRec(1).RESIDENT
    Form$(16, 0) = DCCustRec(1).OWNER
    Form$(12, 0) = DCCustRec(1).HPHONE
    Form$(13, 0) = DCCustRec(1).WPHONE
    IF DCCustRec(1).FirstCar > 0 THEN
     Form$(17, 0) = "Y"
     ELSE
     Form$(17, 0) = "N"
    END IF
    
    Fld(1).Protected = True
    CustomerGrabed = 1
    Action = 1
    COLOR 15
    LOCATE 5, 40: PRINT STRING$(34, 32)
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
       DCCustRec(1).SORTNAME = ""
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

SUB EditVehicles (CustName$, AccountRecord)


EditBody:
  SHARED Mchoice$


  LibName$ = "DC"
  ScrnName$ = "DCVEHED"
  Help$ = "Edit Vehicles"
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

  'Open Vehicle File
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen

  'Open Customer File
  OpenDCCustFile NumOfDcRecs, DCFile
  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
  GET DCFile, AccountRecord, DCCustRec(1)
  LOCK #DCFile, AccountRecord

  VehRecord! = DCCustRec(1).FirstCar
  IF VehRecord! <= 0 THEN
   CLOSE
   EXIT SUB
  END IF

MasterLoop:
  GET DCVFile, VehRecord!, DCVRec(1)
  IF DCVRec(1).Active <> "Y" THEN
   VehRecord! = DCVRec(1).NextRec
   IF VehRecord! <= 0 THEN CLOSE : EXIT SUB
  END IF
  GOSUB GetVehRecord
  IF DCVRec(1).NextRec > 0 THEN LOCATE 19, 24: PRINT "MORE VEHICLES ON FILE, PRESS F10 FOR NEXT VEHICLE";
  IF DCVRec(1).NextRec = 0 THEN LOCATE 19, 24: PRINT "                                                 ";

  Action = 1


  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF AddFlag = 1 AND Frm(1).FldNo = 1 THEN
     GOSUB SelectVCatagory
     LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
     Action = 1
    END IF
    IF AddFlag = 0 AND LEFT$(Form$(1, 0), 1) = " " THEN
     GOSUB SelectVCatagory
     LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
     Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    

    CASE F3Key
    
    CASE F5Key
     IF AddFlag = 0 THEN
      FOR F = 1 TO NumFlds
       LSET Form$(F, 0) = ""
      NEXT F
       VehRecord! = NumOfVrecs + 1
       Action = 1
       AddFlag = 1
     END IF
     
    CASE F10Key
     GOSUB SaveVRecord
     Frm(1).FldNo = 3
     IF AddFlag = 0 THEN
     VehRecord! = DCVRec(1).NextRec
     IF VehRecord! > 0 THEN
          GOTO MasterLoop
     END IF
     ELSE
      FOR F = 1 TO NumFlds
       LSET Form$(F, 0) = ""
      NEXT F
       VehRecord! = NumOfVrecs + 1
       Action = 1
       AddFlag = 1

     END IF
     CLOSE
     EXIT SUB

    CASE EscKey
      CLOSE
      EXIT SUB
    END SELECT
  LOOP




GetVehRecord:
    Form$(1, 0) = DCVRec(1).DecalCat
    Form$(2, 0) = STR$(DCVRec(1).Fee)
    Form$(6, 0) = DCVRec(1).MakeModel
    Form$(7, 0) = DCVRec(1).StateTag
    Form$(4, 0) = Num2Date$(DCVRec(1).ExpireDate)
    Form$(3, 0) = DCVRec(1).Sticker
    Form$(5, 0) = DCVRec(1).Valid
    Form$(8, 0) = DCVRec(1).Desc
    Form$(9, 0) = DCVRec(1).Notes
    RETURN

SaveVRecord:
    IF LEN(RTRIM$(Form$(1, 0))) = 0 THEN RETURN
    DCVRec(1).DecalCat = Form$(1, 0)
    DCVRec(1).Fee = Value#(Form$(2, 0), 0)
    DCVRec(1).MakeModel = Form$(6, 0)
    DCVRec(1).StateTag = Form$(7, 0)
    DCVRec(1).ExpireDate = Date2Num%(Form$(4, 0))
    DCVRec(1).Sticker = Form$(3, 0)
    DCVRec(1).Valid = Form$(5, 0)
    DCVRec(1).Active = "Y"
    DCVRec(1).Desc = Form$(8, 0)
    DCVRec(1).Notes = Form$(9, 0)
    DCVRec(1).MoreRoom = ""
    IF AddFlag = 1 THEN
      DCVRec(1).NextRec = 0
    END IF
    DCVRec(1).MasterRecord = NextAccount
    PUT DCVFile, VehRecord!, DCVRec(1)
    IF AddFlag = 1 THEN
     GOSUB UpdateVendorPointer
    END IF
    RETURN

UpdateVendorPointer:
    IF DCCustRec(1).FirstCar = 0 THEN
     DCCustRec(1).FirstCar = VehRecord!
     DCCustRec(1).LastCar = VehRecord!
     PUT DCFile, AccountRecord, DCCustRec(1)
    ELSE
     PrevRec! = DCCustRec(1).LastCar
     DCCustRec(1).LastCar = VehRecord!
     PUT DCFile, AccountRecord, DCCustRec(1)
     GET DCVFile, PrevRec!, DCVRec(1)
     DCVRec(1).NextRec = VehRecord!
     PUT DCVFile, PrevRec!, DCVRec(1)
    END IF
    RETURN
SelectVCatagory:
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
   Form$(1, 0) = DCCatCodeRec(1).CATCODE
   Form$(2, 0) = STR$(DCCatCodeRec(1).Fee)
   Frm(1).FldNo = 3
   CLOSE DCCatFile
   RETURN

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

  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  


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
    GET TrHandle, DCCustIdxRec(1).IDXRECORD, DCCustRec(1)
IF DCCustRec(1).Deleted <> "Y" THEN
   IF LineCnt >= MaxLines THEN
    PRINT #RptHandle, FF$
    GOSUB PrintRptHeader
   END IF
   PRINT #RptHandle, VAL(DCCustRec(1).CUSTNUMB);
    PRINT #RptHandle, TAB(10); DCCustRec(1).BILLNAME
     
    
     TotalCust = TotalCust + 1
    LineCnt = LineCnt + 1
    CarRecord! = DCCustRec(1).FirstCar

    WHILE CarRecord! > 0
     GET DCVFile, CarRecord!, DCVRec(1)
     IF DCVRec(1).Active <> "N" THEN
      PRINT #RptHandle, DCVRec(1).MakeModel;
      PRINT #RptHandle, TAB(30); LEFT$(DCVRec(1).StateTag, 20);
      PRINT #RptHandle, TAB(51); DCVRec(1).DecalCat;
      PRINT #RptHandle, TAB(60); LTRIM$(LEFT$(DCVRec(1).Sticker, 10));
      PRINT #RptHandle, TAB(70); Num2Date$(DCVRec(1).ExpireDate)
      LineCnt = LineCnt + 1
     END IF
      CarRecord! = DCVRec(1).NextRec
    WEND



   PRINT #RptHandle, STRING$(80, "-"): lc = lc + 1

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
    PRINT #RptHandle, TAB(18); "Va. Decal System : Customer 'Quick' Listing"
    PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Cust #"; TAB(10); "Customer Name"
    PRINT #RptHandle, "Make/Model"; TAB(30); "State Tag#"; TAB(45); "Decal Type"; TAB(58); "Sticker#     EXPIRES"
    PRINT #RptHandle, STRING$(80, "=")
    LineCnt = 6
    RETURN

PrintRptEnding:
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
  GoodRecords = 0

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
    IF DCCustRec(1).Deleted <> "Y" THEN
     GoodRecords = GoodRecords + 1
     array(GoodRecords).who = DCCustRec(1).SORTNAME + "    "
     array(GoodRecords).RecNum = Cnt
    END IF
 NEXT Cnt

 SortT array(Start), GoodRecords, Dir, SSize, MOff, MSize

 FOR Cnt = 1 TO GoodRecords
   DCCustIdxRec(1).IDXNAME = array(Cnt).who
   DCCustIdxRec(1).IDXRECORD = array(Cnt).RecNum
   PUT DCIdxFile, Cnt, DCCustIdxRec(1)
 NEXT Cnt
 CLOSE DCFile
 CLOSE DCIdxFile
END SUB

