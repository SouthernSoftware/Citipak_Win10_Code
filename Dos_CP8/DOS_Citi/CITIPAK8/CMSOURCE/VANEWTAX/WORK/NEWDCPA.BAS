DECLARE SUB PostPayments (Operator%)
DEFINT A-Z
DECLARE SUB DCLookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE FUNCTION DCAskSavePayment% (DefaultFlag$)
DECLARE SUB ClearBack ()
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB SetDefaultExpireDate ()
DECLARE SUB SortDCNameIndex ()
DECLARE SUB ShowNoCodes ()
DECLARE SUB AddCustomer ()
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
DECLARE SUB WaitForAction ()
DECLARE SUB SaveScrn (array%())
DECLARE SUB RestScrn (array%())
DECLARE SUB DisplayDCScrn (ScrnName$)
DECLARE SUB QPrint (X$, Colr%, Page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
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
  '$INCLUDE: 'CMFILES.BI'

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

FUNCTION DCAskSavePayment% (DefaultFlag$)

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  LibName$ = "DC"
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

  DisplayDCScrn ScrnName$

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
    CASE EscKey
      AskSavePayment% = False
      EXIT DO
    CASE F10Key
      Receipt$ = Form$(1, 0)
      IF Receipt$ = "Y" THEN
        DCAskSavePayment% = 1
      ELSE
        DCAskSavePayment% = True
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

SUB EnterDCPayment (OperRecNumber, PostDate$)

  REDIM CMOperRec(1) AS CMOperRecType
  GOSUB GetDCOperator
  PayTName$ = "DCPYT" + QPTrim$(STR$(Operator)) + ".DAT"
  
  REDIM TempScrn(0)
  REDIM DCDate(1) AS DCExpireDate
  PayRecpName$ = "DCRCP.RPT"

PayMainBody:
  CustomerGrabed = 0
  CarGrabed = 0
  DCIdxFile = 0
  DCFile = 0
  
  LibName$ = "DC"
  ScrnName$ = "DCTRANEN"
  help$ = "Enter Payments"
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
  OPEN "DCEXPIRE.DAT" FOR RANDOM SHARED AS DCEFile LEN = 10
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
      
    'CASE F3Key
    '  IF CustomerGrabed = 0 THEN
    '    GOSUB ADDNEWCUSTOMER
    '    CLOSE
    '    Action = 1
    '    GOTO PayMainBody
    '  END IF
    CASE F7Key
      IF CustomerGrabed = 0 THEN
        DCLookUp RecNo&, "Customer", False, True, False
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
        Frm(1).FldNo = 2
        Action = 1
      ELSE
        SELECT CASE DCAskSavePayment%("Y")
        CASE 1 'save and print
          SaveScrn TempScrn()
          GOSUB PAYSaveRecord
          GOSUB PrintReceipt
          RestScrn TempScrn()
          DisplayDCScrn "PRESSKEY"
          WaitForAction
          RestScrn TempScrn()
          Done = True
          GOTO PayMainBody
        CASE True  'just save no reciept
          SaveScrn TempScrn()
          GOSUB PAYSaveRecord
          RestScrn TempScrn()
          DisplayDCScrn "PRESSKEY"
          WaitForAction
          RestScrn TempScrn()
          Done = True
          GOTO PayMainBody
        CASE False  'opps go back to edit
          Frm(1).FldNo = 2
          Action = 1
        END SELECT
        help$ = "SAVING THIS ENTRY NOW!!!"
        'PrintHelp Help$
        'SLEEP 1
        CLOSE
      END IF
    CASE EscKey
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
  OPEN PayTName$ FOR RANDOM SHARED AS DCFile LEN = DCEditRecLen
  'OPEN "DCEdPYT.DAT" FOR RANDOM SHARED AS DCFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCFile) \ DCEditRecLen
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
  EditPaymentRec(1).Resident = Form$(16, 0)
  EditPaymentRec(1).Owner = Form$(17, 0)
  EditPaymentRec(1).PersBuss = Form$(18, 0)
  EditPaymentRec(1).OldMake = Form$(19, 0)
  EditPaymentRec(1).OldDesc = Form$(20, 0)
  EditPaymentRec(1).VehRecord = VehRecord
  PUT DCFile, NumOfDcRecs + 1, EditPaymentRec(1)
  CLOSE DCFile
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM SHARED AS DCVFile LEN = DCVehRecLen
  GET DCVFile, VehRecord, DCVRec(1)
  DCVRec(1).PBFlag = EditPaymentRec(1).PersBuss
  PUT DCVFile, VehRecord, DCVRec(1)
  CLOSE DCVFile
  GOSUB DCMakeCMTrans
  PostPayments Operator

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
  RHandle = FREEFILE
  OPEN PayRecpName$ FOR OUTPUT AS RHandle
  PRINT #RHandle, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #RHandle, CHR$(7)
  PRINT #RHandle,
  PRINT #RHandle, "DECAL PAYMENT"
  PRINT #RHandle, "Date: "; DATE$
  PRINT #RHandle,
  PRINT #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  PRINT #RHandle, EditPaymentRec(1).CustName
  PRINT #RHandle, EditPaymentRec(1).Desc
  PRINT #RHandle, "     Make: "; EditPaymentRec(1).MakeModel
  PRINT #RHandle, "State Tag: "; EditPaymentRec(1).StateTag
  PRINT #RHandle, " Acct. No: "; EditPaymentRec(1).CustNumber
  PRINT #RHandle, " Decal No: "; EditPaymentRec(1).Sticker
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, "Paid: "; USING "$$####,#.##"; EditPaymentRec(1).Amount
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle, "       T H A N K   Y O U !"
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  PRINT #RHandle,
  CLOSE RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$

  PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  KillFile PayRecpName$
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
    Form$(3, 0) = DCCustRec(1).BillName
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
    'PrintHelp Help$
    
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
  'PrintHelp Help$
  
  Done = False
  Action = 1
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      RETURN
    END SELECT
  LOOP
  
  
SelectCar:
  'Show List plus Add New Function Here
  UpdateVeh = 0 'Flag to Indicate whether to add this car to customers file 1=true 0=false
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM SHARED AS DCVFile LEN = DCVehRecLen
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
      LSET MChoice$(ChoiceCounter) = DCVRec(1).MakeModel
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
      'Fld(10).Protected = True
    ELSE
      VehRecord = VAL(RIGHT$(MChoice$(Choice), 5))
      GET DCVFile, VehRecord, DCVRec(1)
      CarGrabed = 1
      LSET Form$(10, 0) = RTRIM$(DCVRec(1).DecalCat)
      LSET Form$(11, 0) = DCVRec(1).MakeModel
      LSET Form$(12, 0) = DCVRec(1).StateTag
      LSET Form$(13, 0) = DCVRec(1).Desc
      LSET Form$(14, 0) = STR$(DCVRec(1).Fee)
      Fld(10).Protected = True
      LSET Form$(16, 0) = "Y"
      LSET Form$(17, 0) = "Y"
      LSET Form$(18, 0) = DCVRec(1).PBFlag

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
  DCVRec(1).PBFlag = QPTrim$(Form$(18, 0))
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
  Fld(10).Protected = True
  IF VehRecord = -1 THEN
    GOSUB OpenVehFile
    VehRecord = NumOfVRecs + 1
    CLOSE DCVFile
  END IF
  CLOSE DCCatFile
RETURN

GetDCOperator:
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

SUB PrintEditList
  
  SHARED Choice$()
  DIM Cat$(250), CatAmt#(250)   'Set Maximum Catagories at 250
  ReportFile$ = "DCPAYED.PRN"   'Report File Name
  FF$ = CHR$(12)
  MaxLines = 53
  LPTPORT% = 1
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
    CASE EscKey
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
        PRINT #RptHandle, VAL(DCCustRec(1).CUSTNUMB); TAB(10); DCCustRec(1).BillName; TAB(55); USING "$$#####,#.##"; EditPaymentRec(1).Amount;
        PRINT #RptHandle, TAB(73); EditPaymentRec(1).DecalCat
        PRINT #RptHandle, TAB(10); QPTrim$(DCCustRec(1).Address1)
        PRINT #RptHandle, TAB(10); QPTrim$(DCCustRec(1).City); " "; DCCustRec(1).State; " "; DCCustRec(1).ZipCode
        PRINT #RptHandle, "     Resident: "; EditPaymentRec(1).Resident
        PRINT #RptHandle, "Vehicle Owned: "; EditPaymentRec(1).Owner
        PRINT #RptHandle, "Pers/Business: "; EditPaymentRec(1).PersBuss
        PRINT #RptHandle, " Payment Desc: "; EditPaymentRec(1).Desc
        PRINT #RptHandle, "     Veh Desc: "; EditPaymentRec(1).MakeModel; TAB(50); "State Tag# "; RTRIM$(EditPaymentRec(1).StateTag)
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
  PrintRptFile Header$, ReportFile$, LPTPORT%, RetCode%, EntryPoint
  
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
  help$ = "Set Default License Expire Date"
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
    CASE EscKey
      CLOSE
      EXIT SUB
    END SELECT
    IF Done THEN EXIT SUB
  LOOP
  
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

