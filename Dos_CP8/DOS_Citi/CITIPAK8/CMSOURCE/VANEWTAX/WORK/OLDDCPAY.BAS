DEFINT A-Z
DECLARE SUB DCVertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE FUNCTION DCAskSavePayment% (DefaultFlag$)
DECLARE SUB EnterDCPayment (OperRecNumber, PostDate$)
DECLARE SUB SortDCNameIndex ()
DECLARE SUB ShowNoCodes ()
DECLARE SUB AddCustomer ()
DECLARE SUB PostPayments (Operator)
DECLARE SUB PrintEditList ()
DECLARE SUB EditPayment ()
DECLARE SUB OpenDCCustIdxFile (NumOfDCIdxRecs, DCIdxFile)
DECLARE SUB OpenDCCustFile (NumOfDcRecs, DCFile)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB SaveScrn (array%())
DECLARE SUB RestScrn (array%())
DECLARE SUB DisplayDCScrn (ScrnName$)
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
DECLARE FUNCTION AskSavePayment% (DefaultFlag$)
DECLARE SUB PressButton (BYVAL a, BYVAL b, BYVAL c, BYVAL d)
'$INCLUDE: 'DefCnf.BI'
'DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  '$INCLUDE: 'CMFILES.BI'
  
  CONST false = 0, true = NOT false
  DIM SHARED DCCustRec(1) AS DCCustRecType
  DIM SHARED DCCustIdxRec(1) AS DCCustIDXRecType
  DIM SHARED EditPaymentRec(1) AS DCEditPaymentRecType
  DIM SHARED DCVRec(1) AS DCVehType
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  TYPE FLen3
     V AS STRING * 32
  END TYPE
  
  STACK 8000

SUB AddCustomer
  
mainbody:
  
  LibName$ = "DC"
  ScrnName$ = "DCCUST"
  help$ = "NEW Decal Customer Entry"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
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
  
  
  
  
  Fld(1).Protected = true
  
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
      Done = true
      GOSUB AddCars
      GOTO mainbody
    CASE EscKey
      NeedtoSort = true         ' set to true for testing
      IF NeedtoSort = true THEN
        SortDCNameIndex
      END IF
      
      EXIT SUB
    END SELECT
    
  LOOP
  
  
SaveRecord:
  DCCustRec(1).SORTNAME = Form$(2, 0)
  DCCustRec(1).BillName = Form$(3, 0)
  DCCustRec(1).ADDRESS1 = Form$(4, 0)
  DCCustRec(1).ADDRESS2 = Form$(5, 0)
  DCCustRec(1).CITY = Form$(6, 0)
  DCCustRec(1).STATE = Form$(7, 0)
  DCCustRec(1).ZIPCODE = Form$(8, 0)
  DCCustRec(1).SOSEC = Form$(9, 0)
  DCCustRec(1).DRVLIC = Form$(10, 0)
  DCCustRec(1).DATEOPED = Date2Num(Form$(11, 0))
  DCCustRec(1).CASHONLY = Form$(14, 0)
  DCCustRec(1).Resident = Form$(15, 0)
  DCCustRec(1).Owner = Form$(16, 0)
  DCCustRec(1).HPHONE = Form$(12, 0)
  DCCustRec(1).WPHONE = Form$(13, 0)
  DCCustRec(1).LICENSE = Form$(10, 0)
  DCCustRec(1).Valid = -32767   'INTEGER Date Function
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
  NeedtoSort = true
  RETURN
  
  
AddCars:
  LibName$ = "DC"
  ScrnName$ = "DCVEHEN"
  help$ = "NEW Vehicle Entry"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$
  COLOR 11
  LOCATE 5, 20: PRINT "Customer: "; : COLOR 15: PRINT DCCustRec(1).BillName
  
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
  
  PrintHelp help$
  COLOR 11
  LOCATE 5, 20: PRINT "Customer: "; : COLOR 15: PRINT DCCustRec(1).BillName
  
  
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
      Done = true
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
  NumOfVRecs = LOF(DCVFile) \ DCVehRecLen
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
  PUT DCVFile, NumOfVRecs + 1, DCVRec(1)
  IF DCCustRec(1).FirstCar = 0 THEN
    DCCustRec(1).FirstCar = NumOfVRecs + 1
    DCCustRec(1).LastCar = NumOfVRecs + 1
    PUT DCFile, NextAccount, DCCustRec(1)
  ELSE
    PrevRec! = DCCustRec(1).LastCar
    DCCustRec(1).LastCar = NumOfVRecs + 1
    PUT DCFile, NextAccount, DCCustRec(1)
    
    GET DCVFile, PrevRec!, DCVRec(1)
    DCVRec(1).NextRec = NumOfVRecs + 1
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
  
  REDIM MChoice$(1 TO NumOFDCCatRecs)
  FOR Cnt = 1 TO NumOFDCCatRecs
    GET DCCatFile, Cnt, DCCatCodeRec(1)
    MChoice$(Cnt) = SPACE$(50)
    LSET MChoice$(Cnt) = DCCatCodeRec(1).CATCODE
    MID$(MChoice$(Cnt), 5) = DCCatCodeRec(1).CODEDESC
  NEXT Cnt
  
  MaxLen = 50   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = "  Code    Description"
  
  '--Center Menu within Screen
  Row = 8
  Col = 15
  
  '--Set upper left corner of menu, turn off the cursor
  LOCATE Row, Col, 0
  QPrintRC TText$, Row - 1, Col, 112
  VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
  GET DCCatFile, Choice, DCCatCodeRec(1)
  Form$(1, 0) = DCCatCodeRec(1).CATCODE
  Form$(2, 0) = STR$(DCCatCodeRec(1).Fee)
  Frm(1).FldNo = 3
  CLOSE DCCatFile
  RETURN
  
  
END SUB

FUNCTION DCAskSavePayment% (DefaultFlag$)

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  LibName$ = "UB"
  ScrnName$ = "ASKR2PST"

  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = true
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  DisplayUBScrn ScrnName$

  ShowCursor

  Action = 1
  FirstTime = true

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = NOT FirstTime
      LSET Form$(1, 0) = DefaultFlag$
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      DCAskSavePayment% = false
      EXIT DO
    CASE F10Key
      Receipt$ = Form$(1, 0)
      IF Receipt$ = "Y" THEN
        DCAskSavePayment% = 1
      ELSE
        DCAskSavePayment% = true
      END IF
      EXIT DO
    END SELECT

    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14
        SELECT CASE Frm(1).MCol
        CASE 27 TO 40           'ESC Cancel button
          PressButton 27, 14, 27, 40
        CASE 41 TO 56           'F10 Save Button
          PressButton -68, 14, 41, 56
        END SELECT
      END SELECT
    END IF
  LOOP

  RestScrn TempScrn()

END FUNCTION

SUB DisplayDCScrn (ScrnName$)
  LibFile2Scrn "DC", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB EnterDCPayment (OperRecNumber, PostDate$)
  
  REDIM TempScrn(0)
  
  REDIM CMOperRec(1) AS CMOperRecType
  GOSUB GetDCOperator

  PayTName$ = "DCPYT" + QPTrim$(STR$(Operator)) + ".DAT"

PayMainBody:
  CustomerGrabed = 0
  CarGrabed = 0
  DCIdxFile = 0
  DCFile = 0
  
  LibName$ = "DC"
  ScrnName$ = "DCTRANEC"
  help$ = "Enter Payments"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp help$
  
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

  'Set Default Exprire Date
  CLOSE 15
  OPEN "I", 15, "DCEXPIRE.DAT": WIDTH #15, 255
  INPUT #15, ExpireDate$
  CLOSE 15
  Form$(15, 0) = ExpireDate$
  Form$(19, 0) = "Y"
  Form$(20, 0) = "Y"
  
  Action = 1
  
  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
  OpenDCCustFile NumOfDcRecs, DCFile
  IF LEN(PrevDate$) = 0 THEN
    Form$(1, 0) = DATE$
    PrevDate$ = DATE$
  ELSE
    Form$(1, 0) = PrevDate$
  END IF

  FirstTime = true
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = false
      QPrintRC "[" + SPACE$(29) + "]", 2, 44, -1
      QPrintRC " Operator: " + OperName$, 2, 45, 15
    END IF

    IF Frm(1).PrevFld = 3 AND CustomerGrabed = 0 THEN

      SaveScrn TempScrn()
      GOSUB GetCustomer
      RestScrn TempScrn()
      IF AccountRecord = 0 THEN
        GOSUB ADDNEWCUSTOMER
        CLOSE
        Action = 1
        GOTO PayMainBody
      ELSEIF AccountRecord = -1 THEN
        CLOSE
        AccountRecord = 0
        Frm(1).FldNo = 2
        Action = 1
        GOTO PayMainBody
      END IF

    END IF

    IF Frm(1).PrevFld = 9 AND LEN(LTRIM$(Form$(9, 0))) = 0 THEN
      Frm(1).FldNo = 9
    END IF
    
    IF Frm(1).FldNo = 10 AND VehRecord = -1 THEN
      SaveScrn TempScrn()
      GOSUB SelectVCatagory
      RestScrn TempScrn()
      ScrnName$ = "DCTRANEC"
      LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      Action = 1
    END IF
    
    IF Frm(1).PrevFld = 10 AND CarGrabed = 0 THEN
      GOSUB SelectCar
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      TVal1# = Value#(Form$(16, 0), ECode) + Value#(Form$(17, 0), ECode)
      TVal2# = Value#(Form$(14, 0), ECode) + Value#(Form$(18, 0), ECode)

      IF Date2Num(Form$(15, 0)) < 0 THEN
        COLOR 12
        PRINT CHR$(7);
        LOCATE 22, 18
        PRINT "MUST ENTER EXPIRATION DATE"
        SLEEP 2
        LOCATE 22, 18: PRINT "                          "
        Frm(1).FldNo = 15
        Action = 1
      ELSEIF VAL(Form$(14, 0)) = 0 THEN
        COLOR 12
        PRINT CHR$(7);
        LOCATE 22, 18
        PRINT "MUST ENTER AN AMOUNT"
        SLEEP 2
        LOCATE 22, 18
        PRINT "                          "
        Frm(1).FldNo = 14
        Action = 1
      ELSEIF TVal1# <> TVal2# THEN
        COLOR 12
        PRINT CHR$(7);
        LOCATE 22, 18
        PRINT "AMOUNT OUT OF BALANCE"
        SLEEP 2
        LOCATE 22, 18
        PRINT "                          "
        Frm(1).FldNo = 16
        Action = 1

      ELSE

        SELECT CASE DCAskSavePayment("Y")
        CASE 1
          SaveScrn TempScrn()
          DisplayDCScrn "UPDATDSK"
          GOSUB PAYSaveRecord
          RestScrn TempScrn()
          DisplayDCScrn "UPDATEOK"
          GOSUB PrintReceipt
          Action = 1
          GOTO PayMainBody
        CASE true
          ReceiptFlag = false
          SaveScrn TempScrn()
          DisplayDCScrn "UPDATDSK"
          GOSUB PAYSaveRecord
          RestScrn TempScrn()
          DisplayDCScrn "UPDATEOK"
          WaitForAction
          'GOSUB ClearForm
          'Frm(1).FldNo = 1
          Action = 1
          GOTO PayMainBody
        CASE false
          Action = 2
        END SELECT
      END IF
    CASE EscKey
      EXIT SUB
    END SELECT
  LOOP
  
  
PAYSaveRecord:
  PrevDate$ = Form$(1, 0)
  ' Keep Default Date Here

  IF UpdateVeh = 1 THEN
    'Must Update Account With This Vehicle
    GOSUB AddNewVehicle
  END IF
  CLOSE
  DCEditRecLen = LEN(EditPaymentRec(1))
  DCFile = FREEFILE
  OPEN PayTName$ FOR RANDOM SHARED AS DCFile LEN = DCEditRecLen
  EditPaymentRec(1).TranDate = Date2Num(Form$(1, 0))
  EditPaymentRec(1).CustNumber = Form$(2, 0)
  EditPaymentRec(1).CustName = Form$(3, 0)
  EditPaymentRec(1).Sticker = Form$(9, 0)
  EditPaymentRec(1).DecalCat = Form$(10, 0)

  EditPaymentRec(1).MakeModel = Form$(11, 0)
  EditPaymentRec(1).StateTag = Form$(12, 0)
  EditPaymentRec(1).Desc = Form$(13, 0)
  EditPaymentRec(1).Amount = Value(Form$(14, 0), a%)
  EditPaymentRec(1).ExpDate = Date2Num(Form$(15, 0))
  EditPaymentRec(1).Resident = Form$(20, 0)
  EditPaymentRec(1).Owner = Form$(21, 0)
  EditPaymentRec(1).OldMake = Form$(22, 0)
  EditPaymentRec(1).OldDesc = Form$(23, 0)
  EditPaymentRec(1).CashAmt = Value#(Form$(16, 0), ECode)
  EditPaymentRec(1).CheckAmt = Value#(Form$(17, 0), ECode)
  EditPaymentRec(1).Change = Value#(Form$(18, 0), ECode)
  EditPaymentRec(1).VehRecord = VehRecord

  NumOfDcRecs = (LOF(DCFile) \ DCEditRecLen) + 1
  PUT DCFile, NumOfDcRecs, EditPaymentRec(1)
  CLOSE DCFile
  GOSUB DCMakeCMTrans
  PostPayments Operator
  
RETURN

GetCustomer:
  
  CustomerGrabed = 0
  AccountRecord = VAL(Form$(2, 0))
  
  REM **************************************************************************
  
  IF AccountRecord = 0 THEN
    
    MaxLen = 32 'Set menu width to zero
    BoxBot = 17 'limit the box length to go no lower than line 20
    VAction = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
    TText$ = SPACE$(MaxLen + 4)
    LSET TText$ = " Cust #    Customer Name"
    
    '--Center Menu within Screen
    Row = 8
    Col = 25

    IF INSTR(COMMAND$, "DEBUG") > 0 THEN
      QPrintRC "Free:" + STR$(FRE(-1)), 12, 1, 15
      'WaitForAction
    END IF

    REDIM Items(1 TO NumOfDCIdxRecs + 1) AS FLen3
    IF INSTR(COMMAND$, "DEBUG") > 0 THEN
      QPrintRC "Free:" + STR$(FRE(-1)), 13, 1, 15
      'WaitForAction
    END IF

    'REDIM MChoice$(1 TO NumOfDCIdxRecs + 1)
    Num$ = SPACE$(5)

    ChoiceCounter = 1
    Items(ChoiceCounter).V = SPACE$(26)
    LSET Items(ChoiceCounter).V = " ADD A NEW CUSTOMER"
    'MID$(Items(ChoiceCounter).V, 27) = STR$(0)
    
    FOR Cnt = 1 TO NumOfDCIdxRecs
      GET DCIdxFile, Cnt, DCCustIdxRec(1)
      IF LEFT$(DCCustIdxRec(1).IDXNAME, 7) <> "DELETED" THEN
        GET DCFile, DCCustIdxRec(1).IDXRECORD, DCCustRec(1)
        ChoiceCounter = ChoiceCounter + 1
        'Items(ChoiceCounter).V = SPACE$(26)
        LSET Items(ChoiceCounter).V = LEFT$(DCCustIdxRec(1).IDXNAME, 1) + " " + LEFT$(DCCustRec(1).BillName, 23)
        RSET Num$ = STR$(DCCustIdxRec(1).IDXRECORD)
        MID$(Items(ChoiceCounter).V, 27) = Num$
      END IF
    NEXT Cnt
    
    DO
      
      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      'LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      'ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
'HERE
'      VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf

'11/13/00
      DCVertMenuT Items(), Choice, MaxLen%, BoxBot, Ky$, VAction, Cnf

      IF Ky$ = CHR$(27) THEN
        AccountRecord = -1
        ExitFlag = true
      ELSE
        AccountRecord = VAL(MID$(Items(Choice).V, 27, 5))
        ExitFlag = true
        ERASE Items
      END IF
      
    LOOP UNTIL ExitFlag
    
    IF AccountRecord = -1 THEN
      GOTO GetCustExit
    END IF

    LibName$ = "DC"
    ScrnName$ = "DCTRANEC"
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
    LSET Form$(2, 0) = DCCustRec(1).CUSTNUMB
    LSET Form$(3, 0) = DCCustRec(1).BillName
    LSET Form$(4, 0) = DCCustRec(1).ADDRESS1
    LSET Form$(5, 0) = DCCustRec(1).ADDRESS2
    LSET Form$(6, 0) = DCCustRec(1).CITY
    LSET Form$(7, 0) = DCCustRec(1).STATE
    LSET Form$(8, 0) = DCCustRec(1).ZIPCODE
    Fld(2).Protected = true
    Fld(3).Protected = true
    Fld(4).Protected = true
    Fld(5).Protected = true
    Fld(6).Protected = true
    Fld(7).Protected = true
    Fld(8).Protected = true
    Frm(1).FldNo = 9
    CustomerGrabed = 1
    Action = 1
    COLOR 15
RETURN
    
  ELSE
    IF AccountRecord = 0 THEN RETURN
    LibName$ = "DC"
    ScrnName$ = "DCBADCUS"
    help$ = "Invalid Customer"
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
    'PrintHelp help$
    
    Done = false
    Action = 1
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      SELECT CASE Frm(1).KeyCode
        
      CASE EscKey
        Done = true
      END SELECT
      IF Done = true THEN GOTO PayMainBody
    LOOP
  END IF
  
GetCustExit:
 RETURN


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
  'PrintHelp help$
  
'  Done = False
'  Action = 1
'
'  DO
'
'    EditForm Form$(), Fld(), Frm(1), Cnf, Action
'
'    SELECT CASE Frm(1).KeyCode
'    CASE F10Key
'      RETURN
'    END SELECT
'  LOOP

WaitForAction
      RETURN

ADDNEWCUSTOMER:
  AddCustomer
RETURN
  
SelectCar:
  SaveScrn TempScrn()
  'Show List plus Add New Function Here
  UpdateVeh = 0           'Flag to Indicate whether to add this car
                          'to customers file 1=true 0=false
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM SHARED AS DCVFile LEN = DCVehRecLen
  VRecord = DCCustRec(1).FirstCar
  
  MaxLen = 40   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  VAction = 0   '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = "Vehicle Desc"
  
  '--Center Menu within Screen
  Row = 8
  Col = 20
  
  REDIM MChoice$(1 TO 100)
  ChoiceCounter = 1
  MChoice$(ChoiceCounter) = SPACE$(40)
  LSET MChoice$(ChoiceCounter) = "ADD NEW VEHICLE"
  MID$(MChoice$(ChoiceCounter), 35) = " 0"
  
  WHILE VRecord <> 0
    GET DCVFile, VRecord, DCVRec(1)
    IF DCVRec(1).Active <> "N" THEN
      ChoiceCounter = ChoiceCounter + 1
      MChoice$(ChoiceCounter) = SPACE$(40)
      LSET MChoice$(ChoiceCounter) = DCVRec(1).MakeModel
      MID$(MChoice$(ChoiceCounter), 35) = STR$(VRecord)
    END IF
    VRecord = DCVRec(1).NextRec
  WEND
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    'LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    'ShowCursor
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
    IF Ky$ = CHR$(27) THEN
      VehRecord = 0
      CarGrabed = 0
      UpdateVeh = 0
      Frm(1).FldNo = 9
      ExitFlag = true
      Action = 1

    ELSEIF VAL(RIGHT$(MChoice$(Choice), 5)) = 0 THEN
      VehRecord = -1
      CarGrabed = 1
      UpdateVeh = 1
      ExitFlag = true
    ELSE
      VehRecord = VAL(RIGHT$(MChoice$(Choice), 5))
      GET DCVFile, VehRecord, DCVRec(1)
      CarGrabed = 1
      LSET Form$(10, 0) = RTRIM$(DCVRec(1).DecalCat)
      LSET Form$(11, 0) = DCVRec(1).MakeModel
      LSET Form$(12, 0) = DCVRec(1).StateTag
      LSET Form$(13, 0) = DCVRec(1).Desc
      LSET Form$(14, 0) = STR$(DCVRec(1).Fee)
      Fld(10).Protected = true
      LSET Form$(20, 0) = "Y"
      LSET Form$(21, 0) = "Y"
      ExitFlag = true
    END IF
    
  LOOP UNTIL ExitFlag

  CLOSE DCVFile

  LibName$ = "DC"
  ScrnName$ = "DCTRANEC"
  LOCATE 1, 1, 0
  
  'ShowCursor
  'LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  RestScrn TempScrn()
  Action = 1
RETURN
  
  '************************* NEW VEHICLE ENTRY
OpenVehFile:
  'Open Vehicle File
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVRecs = LOF(DCVFile) \ DCVehRecLen
RETURN
  
AddNewVehicle:
  GOSUB OpenVehFile
  GOSUB SaveVRecord
RETURN
  
SaveVRecord:
  DCVRec(1).DecalCat = Form$(10, 0)
  DCVRec(1).Fee = Value#(Form$(14, 0), 0)
  DCVRec(1).MakeModel = Form$(11, 0)
  DCVRec(1).StateTag = Form$(12, 0)
  DCVRec(1).ExpireDate = Date2Num%(Form$(15, 0))
  DCVRec(1).Sticker = Form$(9, 0)
  DCVRec(1).Valid = "Y"
  DCVRec(1).Active = "Y"
  DCVRec(1).Desc = Form$(13, 0)
  DCVRec(1).Notes = ""
  DCVRec(1).MoreRoom = ""
  DCVRec(1).NextRec = 0
  DCVRec(1).MasterRecord = AccountRecord
  PUT DCVFile, VehRecord, DCVRec(1)
  GOSUB UpdateVendorPointer
RETURN
  
UpdateVendorPointer:
  IF DCCustRec(1).FirstCar = 0 THEN
    DCCustRec(1).FirstCar = VehRecord
    DCCustRec(1).LastCar = VehRecord
    PUT DCFile, AccountRecord, DCCustRec(1)
  ELSE
    PrevRec! = DCCustRec(1).LastCar
    DCCustRec(1).LastCar = VehRecord
    PUT DCFile, AccountRecord, DCCustRec(1)
    GET DCVFile, PrevRec!, DCVRec(1)
    DCVRec(1).NextRec = VehRecord
    PUT DCVFile, PrevRec!, DCVRec(1)
  END IF
RETURN
  
  
SelectVCatagory:
  'SaveScrn TempScrn()

  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  
  OPEN "DCCODE.DAT" FOR RANDOM SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen
  
  REDIM MChoice$(1 TO NumOFDCCatRecs)
  FOR Cnt = 1 TO NumOFDCCatRecs
    GET DCCatFile, Cnt, DCCatCodeRec(1)
    MChoice$(Cnt) = SPACE$(50)
    LSET MChoice$(Cnt) = DCCatCodeRec(1).CATCODE
    MID$(MChoice$(Cnt), 5) = DCCatCodeRec(1).CODEDESC
  NEXT Cnt
  
  MaxLen = 50   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  VAction = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = "  Code    Description"
  
  '--Center Menu within Screen
  Row = 8
  Col = 15
  
  '--Set upper left corner of menu, turn off the cursor
  LOCATE Row, Col, 0
  QPrintRC TText$, Row - 1, Col, 112
  VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf

  IF ASC(Ky$) = 27 THEN
    CarGrabed = 0
    Frm(1).FldNo = 9
  ELSE
    GET DCCatFile, Choice, DCCatCodeRec(1)
    LSET Form$(10, 0) = DCCatCodeRec(1).CATCODE
    LSET Form$(14, 0) = STR$(DCCatCodeRec(1).Fee)
    CatRecord = Choice
    Frm(1).FldNo = 11
    IF VehRecord = -1 THEN
      GOSUB OpenVehFile
      VehRecord = NumOfVRecs + 1
    END IF
  END IF
  Action = 1
  CLOSE DCCatFile
  'RestScrn TempScrn()
RETURN
  
GetDCOperator:

  'Action = 1
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

DCMakeCMTrans:

  PostDate = Date2Num%(PostDate$)
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))

  CMTRRec(1).TransDate = EditPaymentRec(1).TranDate
  CMTRRec(1).TransAmount = EditPaymentRec(1).Amount
  CMTRRec(1).TransCash = EditPaymentRec(1).CashAmt
  CMTRRec(1).TransCheck = EditPaymentRec(1).CheckAmt
  CMTRRec(1).TransAmtOwed = EditPaymentRec(1).Amount
  CMTRRec(1).TransDesc = EditPaymentRec(1).Desc
  CMTRRec(1).TransSource = 51
  CMTRRec(1).TransName = EditPaymentRec(1).CustName
  CMTRRec(1).TransAcctNum = AccountRecord
  CMTRRec(1).TransDetNum = 0
  CMTRRec(1).TransOperNum = Operator
  CMTRRec(1).Trans2GL = "Y"
  CMTRRec(1).TransPad = ""
  CMTRRec(1).TransRevAmt(1) = CatRecord
  CMTRRec(1).TransRevAmt(2) = EditPaymentRec(1).Amount
  FOR Cnt = 3 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = 0
  NEXT Cnt
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CHandle LEN = CMTrRecLen
  PUT CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTRRec(1)
  CLOSE CHandle

RETURN

PrintReceipt:
   OPEN "I", 15, "DCHEADER.DAT": WIDTH #15, 255
  INPUT #15, TownDesc$
  CLOSE 15
  TL = LEN(TownDesc$)
  TabStop = INT((28 - TL) / 2)
  IF TabStop <= 1 THEN TabStop = 1

  OPEN "C:\RECPT.PRN" FOR OUTPUT SHARED AS #15
  WIDTH #15, 255

  PRINT #15, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #15, TAB(TabStop); TownDesc$
  PRINT #15, "      DECAL RECEIPT"
  PRINT #15, ""
  PRINT #15, "    Date: "; Form$(1, 0)
  PRINT #15, " Sticker: "; QPTrim$(EditPaymentRec(1).Sticker)
  PRINT #15, " Vehicle: "; QPTrim$(Form$(11, 0))            'Car Make/Model
  PRINT #15, " Expires: "; Num2Date$(EditPaymentRec(1).ExpDate)
  PRINT #15,
  PRINT #15, "Account:"; EditPaymentRec(1).CustNumber
  PRINT #15, Form$(3, 0)   'CustName
  PRINT #15, Form$(4, 0)   'Addr line 1
  PRINT #15, Form$(13, 0)  'Desc
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; EditPaymentRec(1).Amount
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; EditPaymentRec(1).CashAmt
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; EditPaymentRec(1).CheckAmt
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; EditPaymentRec(1).CashAmt + EditPaymentRec(1).CheckAmt
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; EditPaymentRec(1).Change
  PRINT #15,
  PRINT #15, "Operator: "; Operator
  PRINT #15,
  PRINT #15, TAB(7); "T H A N K   Y O U !"
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,

  CLOSE
  PrintRptFile Header$, "C:\RECPT.PRN", GetDEFPort%, RetCode%, 5
  KILL "C:\RECPT.PRN"

RETURN
'

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

SUB PostPayments (Operator)

  PayTName$ = "DCPYT" + QPTrim$(STR$(Operator)) + ".DAT"

  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
  OpenDCCustFile NumOfDcRecs, DCFile
  DCEditRecLen = LEN(EditPaymentRec(1))
  DCEdFile = FREEFILE

  OPEN PayTName$ FOR RANDOM ACCESS READ WRITE SHARED AS DCEdFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCEdFile) \ DCEditRecLen
  
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
      
      GOSUB OldVehPost
      GOSUB UpdateVehRecord
      
      GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      help$ = "Posting: " + LEFT$(DCCustRec(1).BillName, 30)
      PrintHelp help$
      
      ' Post Charge First to Offset Payment of Decal
      DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
      DCTransRec(1).TransType = 1               ' Type 1 = Charge
      DCTransRec(1).TransDesc = "Decal Purchase " + EditPaymentRec(1).Sticker
      DCTransRec(1).CashAmount = EditPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EditPaymentRec(1).CheckAmt
      DCTransRec(1).BalanceAfterTrans = DCTransRec(1).BalanceAfterTrans + EditPaymentRec(1).Amount
      DCTransRec(1).MakeModel = EditPaymentRec(1).MakeModel
      DCTransRec(1).StateTag = EditPaymentRec(1).StateTag
      DCTransRec(1).Sticker = EditPaymentRec(1).Sticker
      DCTransRec(1).ExpireDate = EditPaymentRec(1).ExpDate
      DCTransRec(1).DecalCat = EditPaymentRec(1).DecalCat
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      PUT DCTransFile, NextTransRec, DCTransRec(1)
      
      
      GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      DCCustRec(1).AcctBal = DCCustRec(1).AcctBal + EditPaymentRec(1).Amount
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
      DCTransRec(1).CashAmount = EditPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EditPaymentRec(1).CheckAmt
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
    
  LOOP UNTIL Cnt = NumOfDcRecs
  CLOSE
  KILL PayTName$
  ' Show All Posted
  PRINT CHR$(7);
'  SaveScrn TempScrn()
'  DisplayDCScrn "DCPOSTED"
'  WaitForAction
'  RestScrn TempScrn()
  LOCATE , , 1
  CLOSE
  
  EXIT SUB

OldVehPost:
  IF EditPaymentRec(1).Owner = "Y" THEN RETURN
  REDIM DCOVRec(1) AS DCOldVehType
  DCOVRecLen = LEN(DCOVRec(1))
  DCOVFile = FREEFILE
  OPEN "DCOLDVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCOVFile LEN = DCOVRecLen
  NumOfOVRecs! = LOF(DCOVFile) \ DCOVRecLen
  NextOVRec! = NumOfOVRecs! + 1
  DCOVRec(1).Make = LTRIM$(EditPaymentRec(1).OldMake)
  DCOVRec(1).Year = LTRIM$(EditPaymentRec(1).OldDesc)
  DCOVRec(1).CustRec = VAL(EditPaymentRec(1).CustNumber)
  DCOVRec(1).MoreRoom = ""
  PUT DCOVFile, NextOVRec!, DCOVRec(1)
  CLOSE DCOVFile
  RETURN
  
UpdateVehRecord:
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVRecs! = LOF(DCVFile) \ DCVehRecLen
  VehRecord! = EditPaymentRec(1).VehRecord
  IF VehRecord! <= 0 OR VehRecord! > NumOfVRecs! THEN CLOSE DCVFile: RETURN
  GET DCVFile, VehRecord!, DCVRec(1)
  DCVRec(1).ExpireDate = EditPaymentRec(1).ExpDate
  DCVRec(1).Sticker = LTRIM$(EditPaymentRec(1).Sticker)
  DCVRec(1).Valid = "Y"
  DCVRec(1).Fee = EditPaymentRec(1).Amount
  PUT DCVFile, VehRecord!, DCVRec(1)
  CLOSE DCVFile
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
  SHARED MChoice$
  
  
  size = 2500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 7     'size of the key element - coded as follows:
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  DIM array(1 TO size)  AS Struct
  help$ = "Sorting Customer Index"
  PrintHelp help$
  
  DCCustRecLen = LEN(DCCustRec(1))
  DCFile = FREEFILE
  OPEN "DCCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDcRecs = LOF(DCFile) \ DCCustRecLen
  
  DCCustIdxRecLen = LEN(DCCustIdxRec(1))
  DCIdxFile = FREEFILE
  OPEN "DCCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS DCIdxFile LEN = DCCustIdxRecLen
  
  GoodRecords = 0
  
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

