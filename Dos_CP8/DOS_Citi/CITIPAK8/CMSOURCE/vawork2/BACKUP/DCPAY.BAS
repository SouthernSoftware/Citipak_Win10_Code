DEFINT A-Z
DECLARE SUB ClearBack ()
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB SetDefaultExpireDate ()
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
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB WaitForAction ()
DECLARE SUB SaveScrn (array%())
DECLARE SUB RestScrn (array%())
DECLARE SUB DisplayDCScrn (ScrnName$)
DECLARE SUB QPrint (X$, Colr%, Page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'DC.bi'                        ' FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  CONST False = 0, True = NOT False
  
  DIM SHARED DCCustRec(1) AS DCCustRecType
  DIM SHARED DCCust(1) AS DCCustRecType
  DIM SHARED DCCustIdxRec(1) AS DCCustIDXRecType
  DIM SHARED EditPaymentRec(1) AS DCEditPaymentRecType
  DIM SHARED DCVRec(1) AS DCVehType
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = " Enter Decal Purchase  "
  MChoice$(2) = " Edit Purchases "
  MChoice$(3) = " Print Edit List "
  MChoice$(4) = " Post Purchases to Accts "
  MChoice$(5) = " Set Default License Expiration Date"
  MChoice$(6) = " Exit to OS "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "Decal Purchases"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    TitleBox 3, Col, MaxLen + 3, "Decal Purchases ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    'PrintTitle user$
    'PrintHelp Help$
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      EnterPayment
    CASE 2
      EditPayment
    CASE 3
      PrintEditList
    CASE 4
      PostPayments
    CASE 5
      SetDefaultExpireDate
    CASE IS = 6
      HideCursor
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
  'PrintHelp Help$
  
  
  
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
    CASE ESCKey
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
  DCCustRec(1).Address1 = Form$(4, 0)
  DCCustRec(1).ADDRESS2 = Form$(5, 0)
  DCCustRec(1).City = Form$(6, 0)
  DCCustRec(1).State = Form$(7, 0)
  DCCustRec(1).ZipCode = Form$(8, 0)
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
  NeedtoSort = True
  RETURN
  
  
AddCars:
  LibName$ = "DC"
  ScrnName$ = "DCVEHEN"
  Help$ = "NEW Vehicle Entry"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp Help$
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
  
  'PrintHelp Help$
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
    CASE ESCKey
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
  DCVRec(1).makemodel = Form$(6, 0)
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

SUB EditPayment
EditMainBody:
  CustomerGrabed = 0
  
  DCIdxFile = 0
  DCFile = 0
Continue:
  LibName$ = "DC"
  ScrnName$ = "DCTRANED"
  Help$ = "Edit Customer Balance Entry's"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp Help$
  
  
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
  
  IF AccountRecord = 0 THEN
    CLOSE
    EXIT SUB
  END IF
  

  FOR Fld = 2 TO 10
    Fld(Fld).Protected = True
  NEXT Fld

  FirstTime = True

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    'smallPause2

    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadPaymentInfo
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
      
    CASE F3Key
      GOSUB DeleteRecord
      IF Deleted THEN
        EditPaymentRec(1).Amount = -1
        PUT DCEdFile, AccountRecord, EditPaymentRec(1)
        CLOSE DCEdFile
        EXIT SUB
      ELSE
        GOTO Continue
      END IF
      
    CASE F10Key
      IF Date2Num(Form$(15, 0)) < 0 THEN
        'COLOR 12
        PRINT CHR$(7);
        LOCATE 19, 40: PRINT "MUST ENTER EXPIRATION DATE"
        SLEEP 2
        LOCATE 19, 40: PRINT "                          "
        Frm(1).FldNo = 15
        Action = 1
      ELSE
        EditPaymentRec(1).TranDate = Date2Num(Form$(1, 0))
        EditPaymentRec(1).makemodel = Form$(11, 0)
        EditPaymentRec(1).StateTag = Form$(12, 0)
        EditPaymentRec(1).Desc = Form$(13, 0)
        EditPaymentRec(1).Amount = Value(Form$(14, 0), a%)
        EditPaymentRec(1).ExpDate = Date2Num(Form$(15, 0))
        EditPaymentRec(1).Resident = Form$(16, 0)
        EditPaymentRec(1).Owner = Form$(17, 0)
        EditPaymentRec(1).OldMake = Form$(18, 0)
        EditPaymentRec(1).OldDesc = Form$(19, 0)
        PUT DCEdFile, AccountRecord, EditPaymentRec(1)
        CLOSE DCEdFile
        Help$ = "SAVING YOUR CHANGES"
        'PrintHelp Help$
        PRINT CHR$(7)
        SLEEP 1
        CLOSE
        EXIT SUB
      END IF
    CASE ESCKey
      EXIT SUB
      
    END SELECT
    
  LOOP
  
GetEditRecord:
  
  MaxLen = 50   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
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
  
  REDIM MChoice$(1 TO NumOfDcRecs)
  
  ChoiceCounter = 0
  FOR Cnt = 1 TO NumOfDcRecs
    GET DCEdFile, Cnt, EditPaymentRec(1)
    IF EditPaymentRec(1).Amount >= 0 AND VAL(EditPaymentRec(1).CustNumber) <> 0 THEN
      ChoiceCounter = ChoiceCounter + 1
      MChoice$(ChoiceCounter) = SPACE$(50)
      LSET MChoice$(ChoiceCounter) = EditPaymentRec(1).CustNumber
      MID$(MChoice$(ChoiceCounter), 10, 30) = EditPaymentRec(1).CustName
      MID$(MChoice$(ChoiceCounter), 45, 5) = STR$(Cnt)
    END IF
  NEXT Cnt
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    ShowCursor
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      AccountRecord = 0
      ExitFlag = True
    ELSE
      AccountRecord = VAL(RIGHT$(MChoice$(Choice), 5))
      ExitFlag = True
    END IF
    
  LOOP UNTIL ExitFlag
  
  LibName$ = "DC"
  ScrnName$ = "DCTRANED"
  LOCATE 1, 1, 0
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
RETURN

LoadPaymentInfo:

  GET DCEdFile, AccountRecord, EditPaymentRec(1)
  GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
  LSET Form$(1, 0) = Num2Date$(EditPaymentRec(1).TranDate)
  LSET Form$(2, 0) = EditPaymentRec(1).CustNumber
  LSET Form$(3, 0) = DCCustRec(1).BILLNAME
  LSET Form$(4, 0) = DCCustRec(1).Address1
  LSET Form$(5, 0) = DCCustRec(1).ADDRESS2
  LSET Form$(6, 0) = DCCustRec(1).City
  LSET Form$(7, 0) = DCCustRec(1).State
  LSET Form$(8, 0) = DCCustRec(1).ZipCode
  LSET Form$(9, 0) = LTRIM$(EditPaymentRec(1).Sticker)
  LSET Form$(10, 0) = LTRIM$(EditPaymentRec(1).DecalCat)
  LSET Form$(11, 0) = EditPaymentRec(1).makemodel
  LSET Form$(12, 0) = EditPaymentRec(1).StateTag
  LSET Form$(13, 0) = EditPaymentRec(1).Desc
  LSET Form$(14, 0) = STR$(EditPaymentRec(1).Amount)
  LSET Form$(15, 0) = Num2Date$(EditPaymentRec(1).ExpDate)
  LSET Form$(16, 0) = EditPaymentRec(1).Owner
  LSET Form$(17, 0) = EditPaymentRec(1).Resident
  LSET Form$(18, 0) = EditPaymentRec(1).OldMake
  LSET Form$(19, 0) = EditPaymentRec(1).OldDesc

RETURN

DeleteRecord:
  LibName$ = "DC"
  ScrnName$ = "DCEDITOK"
  Help$ = "Delete Edit Entry"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp Help$
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
  Frm(1).FldNo = 2
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
      
    CASE F10Key
      
      Help$ = "Entry Deleted!!!"
      'PrintHelp Help$
      FOR CLine = 10 TO 14
        LOCATE CLine, 21: PRINT STRING$(40, 32);
      NEXT CLine
      
      QPrintRC "Transaction Deleted!", 12, 29, 122
      PRINT CHR$(7);
      SLEEP 2
      Deleted = True
      RETURN
    CASE ESCKey
      Deleted = False
      Action = 1
      RETURN
      
    END SELECT
    
  LOOP
  
END SUB

SUB EnterPayment
  REDIM TempScrn(0)
  REDIM DCDate(1) AS DCExpireDate
  
  DEFSNG V
  
PayMainBody:
  CustomerGrabed = 0
  CarGrabed = 0
  DCIdxFile = 0
  DCFile = 0
  
  LibName$ = "DC"
  ScrnName$ = "DCTRANEN"
  Help$ = "Enter Payments"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp Help$

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
  CLOSE

  IF LEN(PrevDate$) = 0 THEN
    Form$(1, 0) = DATE$
    PrevDate$ = DATE$
  ELSE
    Form$(1, 0) = PrevDate$
  END IF
  
  DCEFile = FREEFILE
  OPEN "DCEXPIRE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCEFile LEN = 10
  IF LOF(DCEFile) > 0 THEN
    GET DCEFile, 1, DCDate(1)
    Form$(15, 0) = DCDate(1).ExpireDate
  END IF
  CLOSE DCEFile

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    'smallPause2

    SELECT CASE Frm(1).FldNo
    CASE 22, 23, 24, 25
      Frm(1).FldNo = 2
    END SELECT

    IF Frm(1).PrevFld = 3 AND CustomerGrabed = 0 THEN
      GOSUB GetCustomer
      IF CustomerGrabed = 0 THEN
        CLOSE
        Action = 1
        'GOTO PayMainBody
        Frm(1).FldNo = 2
      END IF
    END IF
    
    IF Frm(1).PrevFld = 9 AND LEN(LTRIM$(Form$(9, 0))) = 0 THEN
      Frm(1).FldNo = 9
    END IF
    
    IF Frm(1).FldNo = 10 AND VehRecord = -1 THEN
      GOSUB SelectVCatagory
      ScrnName$ = "DCTRANEN"
      LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
      Action = 1
    END IF
    
    IF Frm(1).PrevFld = 10 AND CarGrabed = 0 THEN
      GOSUB SelectCar
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F3Key
      IF CustomerGrabed = 0 THEN
        GOSUB ADDNEWCUSTOMER
        CLOSE
        Action = 1
        GOTO PayMainBody
      END IF
    CASE F7Key
      IF CustomerGrabed = 0 THEN
        LookUp RecNo&, "Customer", False, True, False
        Action = 1
        IF RecNo& > 0 THEN
          AccountRecord = RecNo&
          GOSUB GrabCust
          Action = 1
        END IF
      END IF
    CASE F10Key
      IF Date2Num(Form$(15, 0)) < 0 THEN
        ok = MsgBox("DC", "BADDATA")
        ''COLOR 12
        'PRINT CHR$(7);
        'LOCATE 21, 18: PRINT "MUST ENTER EXPIRATION DATE"
        'SLEEP 2
        'LOCATE 21, 18: PRINT "                          "
        Frm(1).FldNo = 2
        Action = 1
        
      ELSE
        Help$ = "SAVING THIS ENTRY NOW!!!"
        'PrintHelp Help$
        SLEEP 1
        GOSUB PAYSaveRecord
        CLOSE
        Done = True
        GOTO PayMainBody
      END IF
    CASE ESCKey
      EXIT SUB
      
    END SELECT
    
  LOOP
  
  
PAYSaveRecord:
  PrevDate$ = Form$(1, 0):      REM Keep Default Date Here
  IF UpdateVeh = 1 THEN
    'Must Update Account With This Vehicle
    GOSUB AddNewVehicle
  END IF
  CLOSE
  DCEditRecLen = LEN(EditPaymentRec(1))
  DCFile = FREEFILE
  OPEN "DCEdPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCFile) \ DCEditRecLen
  EditPaymentRec(1).TranDate = Date2Num(Form$(1, 0))
  EditPaymentRec(1).CustNumber = Form$(2, 0)
  EditPaymentRec(1).CustName = Form$(3, 0)
  EditPaymentRec(1).Sticker = Form$(9, 0)
  EditPaymentRec(1).DecalCat = Form$(10, 0)
  EditPaymentRec(1).makemodel = Form$(11, 0)
  EditPaymentRec(1).StateTag = Form$(12, 0)
  EditPaymentRec(1).Desc = Form$(13, 0)
  EditPaymentRec(1).Amount = Value(Form$(14, 0), a%)
  EditPaymentRec(1).ExpDate = Date2Num(Form$(15, 0))
  EditPaymentRec(1).Resident = Form$(16, 0)
  EditPaymentRec(1).Owner = Form$(17, 0)
  EditPaymentRec(1).OldMake = Form$(18, 0)
  EditPaymentRec(1).OldDesc = Form$(19, 0)
  EditPaymentRec(1).VehRecord = VehRecord
  PUT DCFile, NumOfDcRecs + 1, EditPaymentRec(1)
  CLOSE DCFile
  
  RETURN
  
  
GetCustomer:
  CustomerGrabed = 0
  AccountRecord = VAL(Form$(2, 0))
  IF AccountRecord = 0 THEN
    Action = 1
    RETURN
  END IF
  
  
  REM ************************************************************************
GrabCust:
  IF AccountRecord > 0 AND AccountRecord <= NumOfDcRecs THEN
    OpenDCCustFile NumOfDcRecs, DCFile
    GET DCFile, AccountRecord, DCCustRec(1)
    CLOSE
    IF DCCustRec(1).Deleted = "Y" THEN
      GOSUB CustomerDeleted
      GOTO PayMainBody
    END IF
    
    Form$(2, 0) = DCCustRec(1).CUSTNUMB
    Form$(3, 0) = DCCustRec(1).BILLNAME
    Form$(4, 0) = DCCustRec(1).Address1
    Form$(5, 0) = DCCustRec(1).ADDRESS2
    Form$(6, 0) = DCCustRec(1).City
    Form$(7, 0) = DCCustRec(1).State
    Form$(8, 0) = DCCustRec(1).ZipCode
    Fld(2).Protected = True
    Fld(3).Protected = True
    Fld(4).Protected = True
    Fld(5).Protected = True
    Fld(6).Protected = True
    Fld(7).Protected = True
    Fld(8).Protected = True
    Fld(8).Protected = True
    Frm(1).FldNo = 9
    CustomerGrabed = 1
    Action = 1
    'COLOR 15
    RETURN
    
  ELSE
    IF AccountRecord = 0 THEN RETURN
    LibName$ = "DC"
    ScrnName$ = "DCBADCUS"
    Help$ = "Invalid Customer"
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
    'PrintHelp Help$
    
    Done = False
    Action = 1
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      SELECT CASE Frm(1).KeyCode
        
      CASE ESCKey
        Done = True
      END SELECT
      IF Done = True THEN GOTO PayMainBody
    LOOP
    
  END IF
  
CustomerDeleted:
  LibName$ = "DC"
  ScrnName$ = "DCDELCUS"
  Help$ = "Payment Entry"
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
  'PrintHelp Help$
  
  Done = False
  Action = 1
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE ESCKey
      RETURN
    END SELECT
  LOOP
  
ADDNEWCUSTOMER:
  AddCustomer
  RETURN
  
SelectCar:
  'Show List plus Add New Function Here
  UpdateVeh = 0 'Flag to Indicate whether to add this car to customers file 1=true 0=false
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  VRecord = DCCustRec(1).FirstCar
  
  MaxLen = 40   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
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
      LSET MChoice$(ChoiceCounter) = DCVRec(1).makemodel
      MID$(MChoice$(ChoiceCounter), 35) = STR$(VRecord)
    END IF
    VRecord = DCVRec(1).NextRec
  WEND
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    ShowCursor
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) OR VAL(RIGHT$(MChoice$(Choice), 5)) = 0 THEN
      VehRecord = -1
      CarGrabed = 1
      UpdateVeh = 1
      ExitFlag = True
    ELSE
      VehRecord = VAL(RIGHT$(MChoice$(Choice), 5))
      GET DCVFile, VehRecord, DCVRec(1)
      CarGrabed = 1
      Form$(10, 0) = RTRIM$(DCVRec(1).DecalCat)
      Form$(11, 0) = DCVRec(1).makemodel
      Form$(12, 0) = DCVRec(1).StateTag
      Form$(13, 0) = DCVRec(1).Desc
      Form$(14, 0) = STR$(DCVRec(1).Fee)
      Fld(10).Protected = True
      Form$(16, 0) = "Y": Form$(17, 0) = "Y"
      ExitFlag = True
    END IF
    
  LOOP UNTIL ExitFlag
  
  LibName$ = "DC"
  ScrnName$ = "DCTRANEN"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  RETURN
  
  '************************* NEW VEHICLE ENTRY
OpenVehFile:
  'Open Vehicle File
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVRecs = LOF(DCVFile) \ DCVehRecLen
RETURN
  
AddNewVehicle:
  GOSUB OpenVehFile
  GOSUB SaveVRecord
RETURN
  
SaveVRecord:
  DCVRec(1).DecalCat = Form$(10, 0)
  DCVRec(1).Fee = Value#(Form$(14, 0), 0)
  DCVRec(1).makemodel = Form$(11, 0)
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
    OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
    OpenDCCustFile NumOfDcRecs, DCFile
    PUT DCFile, AccountRecord, DCCustRec(1)
    CLOSE
  ELSE
    PrevRec! = DCCustRec(1).LastCar
    DCCustRec(1).LastCar = VehRecord
    CLOSE DCFile
    OpenDCCustFile NumOfDcRecs, DCFile

    PUT DCFile, AccountRecord, DCCustRec(1)
    GET DCVFile, PrevRec!, DCVRec(1)
    DCVRec(1).NextRec = VehRecord
    PUT DCVFile, PrevRec!, DCVRec(1)
  END IF
RETURN
  
  
SelectVCatagory:
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
  Form$(10, 0) = DCCatCodeRec(1).CATCODE
  Form$(14, 0) = STR$(DCCatCodeRec(1).Fee)
  Frm(1).FldNo = 11
  IF VehRecord = -1 THEN
    GOSUB OpenVehFile
    VehRecord = NumOfVRecs + 1
    CLOSE DCVFile
  END IF
  CLOSE DCCatFile
RETURN
  
END SUB

DEFINT V
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
  Help$ = "Set Customer Balances"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp Help$
  
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
    CASE ESCKey
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
    
    IF EditPaymentRec(1).Amount >= 0 AND VAL(EditPaymentRec(1).CustNumber) > 0 THEN
      
      GOSUB OldVehPost
      GOSUB UpdateVehRecord
      
      GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      Help$ = "Posting: " + LEFT$(DCCustRec(1).BILLNAME, 30)
      'PrintHelp Help$
      
      ' Post Charge First to Offset Payment of Decal
      DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
      DCTransRec(1).TransType = 1               ' Type 1 = Charge
      DCTransRec(1).TransDesc = "Decal Purchase " + EditPaymentRec(1).Sticker
      DCTransRec(1).CashAmount = EditPaymentRec(1).Amount
      DCTransRec(1).ChkAmount = 0
      DCTransRec(1).BalanceAfterTrans = DCTransRec(1).BalanceAfterTrans + EditPaymentRec(1).Amount
      DCTransRec(1).makemodel = EditPaymentRec(1).makemodel
      DCTransRec(1).StateTag = EditPaymentRec(1).StateTag
      DCTransRec(1).Sticker = EditPaymentRec(1).Sticker
      DCTransRec(1).ExpireDate = EditPaymentRec(1).ExpDate
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      DCTransRec(1).GLInterfaced = "Y"
      DCTransRec(1).DecalCat = EditPaymentRec(1).DecalCat       'Dale Need This in His Stuff
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
      DCTransRec(1).CashAmount = EditPaymentRec(1).Amount
      DCTransRec(1).ChkAmount = 0
      DCTransRec(1).BalanceAfterTrans = DCTransRec(1).BalanceAfterTrans - EditPaymentRec(1).Amount
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      DCTransRec(1).GLInterfaced = "N"
      DCTransRec(1).DecalCat = EditPaymentRec(1).DecalCat       'DALE NEED THIS LINE IN HIS STUFF
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
OldVehPost:
  IF EditPaymentRec(1).Owner = "Y" THEN RETURN
  REDIM DCOVRec(1) AS DCOldVehType
  DCOVRecLen = LEN(DCOVRec(1))
  DCOVFile = FREEFILE
  OPEN "DCOLDVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCOVFile LEN = DCOVRecLen
  NumOfOVRecs! = LOF(DCOVFile) \ DCOVRecLen
  NextOVRec! = NumOfOVRecs! + 1
  DCOVRec(1).Make = LTRIM$(EditPaymentRec(1).OldMake)
  DCOVRec(1).year = LTRIM$(EditPaymentRec(1).OldDesc)
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

SUB PrintEditList

  SHARED Choice$()
  DIM Cat$(250), CatAmt#(250)   'Set Maximum Catagories at 250
  ReportFile$ = "DCPAYED.PRN"   'Report File Name
  FF$ = CHR$(12)
  MaxLines = 53
  LPTPort% = 1
  LineCnt = 0
  
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
  'PrintHelp Help$
  Action = 1
  QPrintRC "Payment Edit List ]", 9, 23, -1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      DevSpec$ = LEFT$(Form$(1, 0), 1)
      EXIT DO
    CASE ESCKey
      AbortFlag = True
      EXIT DO
    END SELECT
  LOOP
  
  IF AbortFlag THEN
    GOTO DoneHere1:
  END IF
  
  CustRecLen = LEN(DCCustRec(1))
  TrHandle = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
  
  DCEditRecLen = LEN(EditPaymentRec(1))
  DCFile = FREEFILE
  OPEN "DCEDPYT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCFile) \ DCEditRecLen
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO NumOfDcRecs
    GET DCFile, Cnt, EditPaymentRec(1)
    IF VAL(EditPaymentRec(1).CustNumber) = 0 THEN
      
    ELSE
      GET TrHandle, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      IF EditPaymentRec(1).Amount >= 0 THEN
        PRINT #RptHandle, VAL(DCCustRec(1).CUSTNUMB); TAB(10); DCCustRec(1).BILLNAME; TAB(55); USING "$$#####,#.##"; EditPaymentRec(1).Amount;
        PRINT #RptHandle, TAB(73); EditPaymentRec(1).DecalCat
        PRINT #RptHandle, TAB(10); QPTrim$(DCCustRec(1).Address1)
        PRINT #RptHandle, TAB(10); QPTrim$(DCCustRec(1).City); " "; DCCustRec(1).State; " "; DCCustRec(1).ZipCode
        PRINT #RptHandle, "     Resident: "; EditPaymentRec(1).Resident
        PRINT #RptHandle, "Vehicle Owned: "; EditPaymentRec(1).Owner
        PRINT #RptHandle, " Payment Desc: "; EditPaymentRec(1).Desc
        PRINT #RptHandle, "     Veh Desc: "; EditPaymentRec(1).makemodel; TAB(50); "State Tag# "; RTRIM$(EditPaymentRec(1).StateTag)
        PRINT #RptHandle, " Payment Date: "; Num2Date(EditPaymentRec(1).TranDate)
        PRINT #RptHandle, " Expires Date: "; Num2Date(EditPaymentRec(1).ExpDate)
        PRINT #RptHandle, STRING$(80, "-")
        TotalCust = TotalCust + 1
        TotalValue# = TotalValue# + EditPaymentRec(1).Amount
        TotalValue# = INT((TotalValue# * 100) + .5) / 100
        GOSUB CatagoryTotal
        LineCnt = LineCnt + 10
      END IF
    END IF
  '  smallPause
  NEXT Cnt
  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  Header$ = "Payment Edit Listing"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
DoneHere1:
  EXIT SUB
  
CatagoryTotal:
  IF CatCnt = 0 THEN
    CatCnt = 1
    Cat$(1) = LTRIM$(EditPaymentRec(1).DecalCat)
    CatAmt#(1) = EditPaymentRec(1).Amount
    RETURN
  END IF
  FOR CatLoop = 1 TO CatCnt
    IF Cat$(CatLoop) = LTRIM$(EditPaymentRec(1).DecalCat) THEN
      CatAmt#(CatLoop) = CatAmt#(CatLoop) + EditPaymentRec(1).Amount
      RETURN
    END IF
  NEXT CatLoop
  CatCnt = CatCnt + 1
  Cat$(CatCnt) = LTRIM$(EditPaymentRec(1).DecalCat)
  CatAmt#(CatCnt) = EditPaymentRec(1).Amount
  
  RETURN
  
PrintRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Vehicle Decals Payment EDIT Listing"
  PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; Page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Cust #"; TAB(10); "Billing Name"; TAB(56); "Payment Amt"; TAB(73); "Type"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, "Number of Entries .. "; USING "####,#"; TotalCust;
  PRINT #RptHandle, TAB(55); USING "$$#####,#.##"; TotalValue#
  PRINT #RptHandle, FF$
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Vehicle Decals Payment EDIT Listing"
  PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; Page
  PRINT #RptHandle, "Catagory Totals"
  PRINT #RptHandle, "Catagory"; TAB(20); "       Amount"
  FOR Lp = 1 TO CatCnt
    PRINT #RptHandle, Cat$(Lp); TAB(20); USING "$$######,#.##"; CatAmt#(Lp)
  NEXT Lp
  RETURN
  
END SUB

SUB SetDefaultExpireDate
  
  LibName$ = "DC"
  ScrnName$ = "DCEXPIRE"
  Help$ = "Set Default License Expire Date"
  LOCATE 1, 1, 0
  
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'PrintHelp Help$
  
  REDIM DCDate(1) AS DCExpireDate
  
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
  
  DCFile = FREEFILE
  OPEN "DCEXPIRE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = 10
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      
      DCDate(1).ExpireDate = Form$(1, 0)
      PUT DCFile, 1, DCDate(1)
      CLOSE
      Done = True
    CASE ESCKey
      CLOSE
      EXIT SUB
    END SELECT
    IF Done THEN EXIT SUB
  LOOP
  
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
    CASE ESCKey
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
  Help$ = "Sorting Customer Index"
  'PrintHelp Help$
  
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

