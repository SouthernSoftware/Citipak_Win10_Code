DEFINT A-Z

DECLARE SUB RelinkTransactions ()
DECLARE SUB ClearBack ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB WaitForAction ()
DECLARE SUB CustMailLbl ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB CursorOff ()
DECLARE SUB ClearBack ()
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
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
DECLARE SUB DisplayARScrn (ScrnName$)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB HideCursor ()
DECLARE SUB LibGetAddresses (LibName$, FrmName$, Handle%, Offset&, EndOffset&, ErrCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB KillFile (FileName$)

DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  DIM SHARED ARCust(1) AS ARCustRecType
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
  
  STACK 8000

  CONST False = 0, True = NOT False
  
  '   LibGetAddresses "AR", "ARCUST", Handle%, Offset&, EndOffset&, ErrCode%
  '   CALL FClose(Handle)
  '   a$ = SPACE$((EndOffset& - Offset&) + 1)
  '   OPEN "ar.qfl" FOR BINARY AS #1
  '   SEEK #1, Offset&
  '   GET #1, , a$
  '   CLOSE #1
  '   OPEN "arcust2.frm" FOR BINARY AS #1
  '   PUT #1, , a$
  '   CLOSE
  '   END

  'ARCustLen = LEN(ARCust(1))
  'OPEN "arcust.dat" FOR RANDOM AS #1 LEN = ARCustLen
  'FOR cnt = 1 TO (LOF(1) / ARCustLen)
  'GET #1, cnt, ARCust(1)
  'ARCust(1).ProRate = 100
  'ARCust(1).CUSTNUMB = QPTrim$(STR$(cnt))
  'PUT #1, cnt, ARCust(1)
  'NEXT
  'CLOSE
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 10)
  
  MChoice$(1) = "Add New Customer"
  MChoice$(2) = "Edit Existing Customer"
  MChoice$(3) = "Print Customer Listing"
  MChoice$(4) = "Set Customer License's to Print "
  MChoice$(5) = "Print Customer Mailing Labels"
  MChoice$(6) = "Resort Customer Name Index"
  MChoice$(7) = "Relink Transaction History"
'  MChoice$(8) = "Remove a Customer Transaction"
  MChoice$(8) = "Adjust a Customer Balance"
  MChoice$(9) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2) - 1
  help$ = "Add/Edit/Print Customers"
  

  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ClearBack
    
    TitleBox 3, Col, MaxLen + 3, "Customer Maintenance ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    PrintTitle User$
    PrintHelp help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
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
      DScrn = True
      SortARNameIndex
      DScrn = False
    CASE 7
      RelinkTransactions
    'CASE 8
    '  RUN "ARWACK"
    CASE 8
      RUN "ARadjust"
    CASE 9
      HideCursor
      CLS
      END
    END SELECT
  LOOP

  RUN "armenu"

SUB AddCustomer

  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  
mainbody:
  LibName$ = "AR"
  ScrnName$ = "ARCUST"
  help$ = "NEW A/R Customer Entry"
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
  
  Action = 1
  
  REM check for code file
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  CLOSE
  IF NumOFARCatRecs = 0 THEN
    ShowNoCodes
    EXIT SUB
  END IF
  
  OpenARCustFile NumOfArRecs, ARFile

'  IF NumOfArRecs > 2 THEN
'    ClearBack
'    ok = MsgBox("AR", "DEMOONLY")
'    EXIT SUB
'  END IF

  Fld(1).Protected = True
  Fld(29).Protected = True
  Frm(1).FldNo = 2
  Form$(14, 0) = CHR$(0)
  Form$(17, 0) = CHR$(0)
  Form$(20, 0) = CHR$(0)
  Form$(23, 0) = CHR$(0)
  
  Form$(26, 0) = "0"
  Form$(27, 0) = "I"
  Form$(32, 0) = "N"

  GOSUB AddInitCodes

  ShowCursor
  DisplayARScrn "BAKCLEAR"
  DisplayARScrn ScrnName$

  'PrintHelp help$
  
  LOCATE 4, 28
  COLOR 15
  PRINT "        PENDING"

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).PrevFld
    CASE 11, 14, 17, 20, 23
      IF Frm(1).FldNo <> Frm(1).PrevFld THEN
        Code$ = QPTrim$(Form$(Frm(1).PrevFld, 0))
        IF Code$ = "NONE" THEN
          LSET Form$(Frm(1).PrevFld, 0) = ""
          Action = 2
        ELSE
          IF LEN(Code$) > 0 THEN
            FOR CCnt = 1 TO NumOFARCatRecs
              TCode$ = QPTrim$(LEFT$(Choice$(CCnt, 1), 5))
              IF Code$ = TCode$ THEN
                LSET Form$(Frm(1).PrevFld + 1, 0) = QPTrim$(MID$(Choice$(CCnt, 1), 6))
                EXIT FOR
              END IF
            NEXT
            Action = 2
          END IF
        END IF
      END IF
    END SELECT
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB SaveRecord
      DONE = True
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
  IF LEFT$(Form$(11, 0), 1) <> " " THEN
    ARCustRec(1).SORTNAME = Form$(2, 0)
    ARCustRec(1).BILLNAME = Form$(3, 0)
    ARCustRec(1).ADDRESS1 = Form$(4, 0)
    ARCustRec(1).ADDRESS2 = Form$(5, 0)
    ARCustRec(1).CITY = Form$(6, 0)
    ARCustRec(1).STATE = Form$(7, 0)
    ARCustRec(1).ZIPCODE = Form$(8, 0)
    ARCustRec(1).CUSTNAME = Form$(9, 0)
    ARCustRec(1).CONTACT = Form$(10, 0)
    ARCustRec(1).BILLCAT1 = Form$(11, 0)
    ARCustRec(1).Desc1 = Form$(12, 0)
    ARCustRec(1).REV1 = Value#(Form$(13, 0), ECode%)
    ARCustRec(1).Fee1 = 0
    ARCustRec(1).BILLCAT2 = Form$(14, 0)
    ARCustRec(1).Desc2 = Form$(15, 0)
    ARCustRec(1).REV2 = Value#(Form$(16, 0), ECode%)
    ARCustRec(1).Fee2 = 0
    ARCustRec(1).BILLCAT3 = Form$(17, 0)
    ARCustRec(1).Desc3 = Form$(18, 0)
    ARCustRec(1).REV3 = Value#(Form$(19, 0), ECode%)
    ARCustRec(1).Fee3 = 0
    ARCustRec(1).BILLCAT4 = Form$(20, 0)
    ARCustRec(1).Desc4 = Form$(21, 0)
    ARCustRec(1).REV4 = Value#(Form$(21, 0), ECode%)
    ARCustRec(1).Fee4 = 0
    ARCustRec(1).BILLCAT5 = Form$(23, 0)
    ARCustRec(1).Desc5 = Form$(24, 0)
    ARCustRec(1).REV5 = Value#(Form$(25, 0), ECode%)
    ARCustRec(1).Fee5 = 0
    ARCustRec(1).WPHONE = Form$(28, 0)
    ARCustRec(1).FeeAmt = 0
    ARCustRec(1).LICENSE = Form$(29, 0)
    ARCustRec(1).VALID = Date2Num(Form$(30, 0)) 'INTEGER Date Function
    ARCustRec(1).AcctBal = 0
    ARCustRec(1).FirstTrans = 0
    ARCustRec(1).LastTrans = 0
    ARCustRec(1).Deleted = "N"
    ARCustRec(1).IssueLicense = "N"
    ARCustRec(1).Inactive = Form$(32, 0)
'    ARCustRec(1).ProRate = Value#(Form$(33, 0), ECode)
    ARCustRec(1).RoomtoGrow = ""
    NextAccount = NumOfArRecs + 1
    ARCustRec(1).CUSTNUMB = STR$(NextAccount)
    PUT ARFile, NextAccount, ARCustRec(1)
    CLOSE ARFile
    CLOSE ARCatFile
    NeedtoSort = True
    LOCATE 4, 28
    COLOR 15
    PRINT NextAccount; "  ASSIGNED"
    SLEEP 2
  END IF
RETURN
  
AddInitCodes:
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  IF NumOFARCatRecs < 1 THEN
    REDIM Choice$(0 TO 1, 0 TO 1)
    Choice$(1, 1) = SPACE$(24)
    LSET Choice$(1, 1) = "NONE"
    MID$(Choice$(1, 1), 7) = "NO CODES ON FILE"
  ELSE
    REDIM Choice$(0 TO NumOFARCatRecs, 0 TO 1)
    FOR Cnt = 1 TO NumOFARCatRecs
      GET ARCatFile, Cnt, ARCatCodeRec(1)
      Choice$(Cnt, 1) = SPACE$(24)
      LSET Choice$(Cnt, 1) = ARCatCodeRec(1).CATCODE
      MID$(Choice$(Cnt, 1), 7) = ARCatCodeRec(1).CODEDESC + ARCatCodeRec(1).CodeType
    NEXT
  END IF
  CLOSE ARCatFile
  Choice$(0, 1) = "11, 14, 17, 20, 23"
RETURN

  
END SUB

SUB EditCustomer
  
  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)

BeginAgain:
  ClearBack
  LookUp RecNo&, "Customer", False, True, False
  IF RecNo& = 0 THEN
    EXIT SUB
  END IF
  
'  RecNo& = RecNo&
  
  CLOSE ARFile
  OpenARCustFile NumOfArRecs, ARFile

EditMainBody:
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  
  'SHARED MChoice$

  LibName$ = "AR"
  ScrnName$ = "ARCUST"

  help$ = "Edit A/R Customer Entry"
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
  
  Action = 1
  FirstTime = True

  GOSUB InitCodes

  LOCK #ARFile, RecNo&

  ShowCursor

  DisplayARScrn "BAKCLEAR"
  DisplayARScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadCustInfo
      Action = 1
    END IF

    SELECT CASE Frm(1).PrevFld
    CASE 11, 14, 17, 20, 23
      IF Frm(1).FldNo <> Frm(1).PrevFld THEN
        Code$ = QPTrim$(Form$(Frm(1).PrevFld, 0))
        IF Code$ = "" THEN
          LSET Form$(Frm(1).PrevFld + 1, 0) = ""
          Action = 2
        ELSE
          IF LEN(Code$) > 0 THEN
            FOR CCnt = 1 TO NumOFARCatRecs
              TCode$ = QPTrim$(LEFT$(Choice$(CCnt, 1), 5))
              IF Code$ = TCode$ THEN
                LSET Form$(Frm(1).PrevFld + 1, 0) = QPTrim$(MID$(Choice$(CCnt, 1), 6))
                EXIT FOR
              END IF
            NEXT
            Action = 2
          END IF
        END IF
      END IF
    END SELECT
    
    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      IF RecNo& > 0 THEN
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
  IF RecNo& = 0 THEN RETURN

  ARCustRec(1).CUSTNUMB = Form$(1, 0)
  ARCustRec(1).SORTNAME = Form$(2, 0)
  ARCustRec(1).BILLNAME = Form$(3, 0)
  ARCustRec(1).ADDRESS1 = Form$(4, 0)
  ARCustRec(1).ADDRESS2 = Form$(5, 0)
  ARCustRec(1).CITY = Form$(6, 0)
  ARCustRec(1).STATE = Form$(7, 0)
  ARCustRec(1).ZIPCODE = Form$(8, 0)
  ARCustRec(1).CUSTNAME = Form$(9, 0)
  ARCustRec(1).CONTACT = Form$(10, 0)
  ARCustRec(1).BILLCAT1 = Form$(11, 0)
  ARCustRec(1).Desc1 = Form$(12, 0)
  ARCustRec(1).REV1 = Value#(Form$(13, 0), ECode%)
  ARCustRec(1).BILLCAT2 = Form$(14, 0)
  ARCustRec(1).Desc2 = Form$(15, 0)
  ARCustRec(1).REV2 = Value#(Form$(16, 0), ECode%)
  ARCustRec(1).BILLCAT3 = Form$(17, 0)
  ARCustRec(1).Desc3 = Form$(18, 0)
  ARCustRec(1).REV3 = Value#(Form$(19, 0), ECode%)
  ARCustRec(1).BILLCAT4 = Form$(20, 0)
  ARCustRec(1).Desc4 = Form$(21, 0)
  ARCustRec(1).REV4 = Value#(Form$(22, 0), ECode%)
  ARCustRec(1).BILLCAT5 = Form$(23, 0)
  ARCustRec(1).Desc5 = Form$(24, 0)
  ARCustRec(1).REV5 = Value#(Form$(25, 0), ECode%)
  ARCustRec(1).IssuanceFee = Value#(Form$(26, 0), ECode%)
  ARCustRec(1).CustLocation = Form$(27, 0)
  ARCustRec(1).WPHONE = Form$(28, 0)
  ARCustRec(1).LICENSE = Form$(29, 0)
  ARCustRec(1).VALID = Date2Num%(Form$(30, 0))
  ARCustRec(1).IssueLicense = Form$(31, 0)
  ARCustRec(1).Inactive = Form$(32, 0)
  ARCustRec(1).ProRate = Value#(Form$(33, 0), ECode)
  PUT ARFile, RecNo&, ARCustRec(1)
  CLOSE ARFile
  NeedtoSort = True
RETURN

InitCodes:
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  IF NumOFARCatRecs < 1 THEN
    REDIM Choice$(0 TO 1, 0 TO 1)
    Choice$(1, 1) = SPACE$(24)
    LSET Choice$(1, 1) = "NONE"
    MID$(Choice$(1, 1), 7) = "NO CODES ON FILE"
  ELSE
    REDIM Choice$(0 TO NumOFARCatRecs, 0 TO 1)
    FOR Cnt = 1 TO NumOFARCatRecs
      GET ARCatFile, Cnt, ARCatCodeRec(1)
      Choice$(Cnt, 1) = SPACE$(24)
      LSET Choice$(Cnt, 1) = ARCatCodeRec(1).CATCODE
      MID$(Choice$(Cnt, 1), 7) = ARCatCodeRec(1).CODEDESC + ARCatCodeRec(1).CodeType
    NEXT
  END IF
  CLOSE ARCatFile
  Choice$(0, 1) = "11, 14, 17, 20, 23"
RETURN


LoadCustInfo:
  IF RecNo& > 0 AND RecNo& <= NumOfArRecs THEN
    GET ARFile, RecNo&, ARCustRec(1)
    IF ARCustRec(1).Deleted = "Y" THEN
      GOTO CustomerDeleted
    END IF

    LSET Form$(1, 0) = ARCustRec(1).CUSTNUMB
    LSET Form$(2, 0) = ARCustRec(1).SORTNAME
    LSET Form$(3, 0) = ARCustRec(1).BILLNAME
    LSET Form$(4, 0) = ARCustRec(1).ADDRESS1
    LSET Form$(5, 0) = ARCustRec(1).ADDRESS2
    LSET Form$(6, 0) = ARCustRec(1).CITY
    LSET Form$(7, 0) = ARCustRec(1).STATE
    LSET Form$(8, 0) = ARCustRec(1).ZIPCODE
    LSET Form$(9, 0) = ARCustRec(1).CUSTNAME
    LSET Form$(10, 0) = ARCustRec(1).CONTACT
    LSET Form$(11, 0) = ARCustRec(1).BILLCAT1
    LSET Form$(12, 0) = ARCustRec(1).Desc1
    LSET Form$(13, 0) = STR$(ARCustRec(1).REV1)
    
    IF LEN(QPTrim$(ARCustRec(1).BILLCAT2)) = 0 THEN
      LSET Form$(14, 0) = CHR$(0)
    ELSE
      LSET Form$(14, 0) = ARCustRec(1).BILLCAT2
    END IF
    
    LSET Form$(15, 0) = ARCustRec(1).Desc2
    LSET Form$(16, 0) = QPTrim$(STR$(ARCustRec(1).REV2))
    
    IF LEN(QPTrim$(ARCustRec(1).BILLCAT3)) = 0 THEN
      LSET Form$(17, 0) = CHR$(0)
    ELSE
      LSET Form$(17, 0) = ARCustRec(1).BILLCAT3
    END IF
    
    LSET Form$(18, 0) = ARCustRec(1).Desc3
    LSET Form$(19, 0) = QPTrim$(STR$(ARCustRec(1).REV3))
    
    IF LEN(QPTrim$(ARCustRec(1).BILLCAT4)) = 0 THEN
      LSET Form$(20, 0) = CHR$(0)
    ELSE
      LSET Form$(20, 0) = ARCustRec(1).BILLCAT4
    END IF
    
    LSET Form$(21, 0) = ARCustRec(1).Desc4
    LSET Form$(22, 0) = QPTrim$(STR$(ARCustRec(1).REV4))
    
    IF LEN(QPTrim$(ARCustRec(1).BILLCAT5)) = 0 THEN
      LSET Form$(23, 0) = CHR$(0)
    ELSE
      LSET Form$(23, 0) = ARCustRec(1).BILLCAT5
    END IF
    
    LSET Form$(24, 0) = ARCustRec(1).Desc5
    LSET Form$(25, 0) = QPTrim$(STR$(ARCustRec(1).REV5))
    LSET Form$(26, 0) = QPTrim$(STR$(ARCustRec(1).IssuanceFee))
    LSET Form$(27, 0) = ARCustRec(1).CustLocation
    LSET Form$(28, 0) = ARCustRec(1).WPHONE
    LSET Form$(29, 0) = ARCustRec(1).LICENSE
    LSET Form$(30, 0) = Num2Date$(ARCustRec(1).VALID)
    LSET Form$(31, 0) = ARCustRec(1).IssueLicense
    LSET Form$(32, 0) = ARCustRec(1).Inactive
    LSET Form$(33, 0) = QPTrim$(STR$(ARCustRec(1).ProRate))

    Fld(1).Protected = True
    CustomerGrabed = 1
    Action = 1
    QPrintRC "Acct:" + STR$(RecNo&), 3, 58, 11

    RETURN
    
  ELSE
    
    LibName$ = "AR"
    ScrnName$ = "ARBADCUS"
    help$ = "Edit A/R Customer Entry"
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
    
    DONE = False
    Action = 1
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        DONE = True
      END SELECT
      IF DONE = True THEN GOTO EditMainBody
    LOOP
    
  END IF
  
CustomerDeleted:
  LibName$ = "AR"
  ScrnName$ = "ARDELCUS"
  help$ = "Edit Customer"
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
  'ShowCursor
  DisplayARScrn ScrnName$
  'PrintHelp help$
  
  DONE = False
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOTO BeginAgain
    END SELECT
  LOOP

DeleteRecord:
  LibName$ = "AR"
  ScrnName$ = "ARCUSDEL"
  help$ = "Delete Customer"
  CursorOff
  ShowCursor
  DisplayARScrn ScrnName$
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
    
    CASE F10Key
      help$ = "Account Deleted!!!"
      'PrintHelp help$
      BEEP
      ARCustRec(1).Deleted = "Y"
      ARCustRec(1).SORTNAME = "DELETED"
      PUT ARFile, RecNo&, ARCustRec(1)
      CLOSE ARFile
      NeedtoSort = True
      RETURN
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

  ReportFile$ = "ARCUST.PRN"    'Report File Name
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  Header$ = "Customer Listing Report"

  ClearBack
  PrintTitle ""
  ShowProcessingScrn Header$

  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
  
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
    'GET TrHandle, Cnt, ARCustRec(1)
    IF ARCustRec(1).Deleted <> "Y" THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      PRINT #RptHandle, ARCustRec(1).CUSTNUMB; TAB(14); ARCustRec(1).BILLNAME;
'      PRINT #RptHandle, TAB(60); USING "$$####,#.##"; ARCustRec(1).FeeAmt;
      PRINT #RptHandle, TAB(67); USING " ###%"; ARCustRec(1).ProRate

      PRINT #RptHandle, TAB(14); ARCustRec(1).CUSTNAME; TAB(50); ARCustRec(1).LICENSE; TAB(65); Num2Date$(ARCustRec(1).VALID)
      PRINT #RptHandle, "Categories: "; ARCustRec(1).BILLCAT1$; " / "; ARCustRec(1).BILLCAT2$; " / "; ARCustRec(1).BILLCAT3$; " / "; ARCustRec(1).BILLCAT4$; " / "; ARCustRec(1).BILLCAT5$
      PRINT #RptHandle, STRING$(79, "-")
      TotalCust = TotalCust + 1
      LineCnt = LineCnt + 4
    END IF
    ShowPctComp Cnt, IdxTrNumRecs
  NEXT Cnt

  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB

PrintRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(18); "Business License : Customer 'Quick' Listing"
  PRINT #RptHandle, TAB(21); "      Report Date: "; DATE$; TAB(68); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Cust #"; TAB(10); "Billing Name"; TAB(65); "ProRate"
  PRINT #RptHandle, TAB(10); "Customer Name"; TAB(48); "License #"; TAB(65); "Valid To"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
RETURN
  
PrintRptEnding:
  PRINT #RptHandle, "Number of Customers .. "; USING "####,#"; TotalCust
  PRINT #RptHandle, FF$
RETURN

END SUB

SUB RelinkTransactions

  ClearBack
  ShowProcessingScrn "Relinking Customer Transactions."
  
  CustRecLen = LEN(ARCust(1))

  REDIM ARTran(1 TO 2) AS ARTransRecType
  TranRecLen = LEN(ARTran(1))

  CustFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = CustRecLen
  NumCRec& = LOF(CustFile) / CustRecLen

  TranFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TranRecLen
  NumTRec& = LOF(TranFile) / TranRecLen

  FOR Cnt& = 1 TO NumCRec&
    GET CustFile, Cnt&, ARCust(1)
    ARCust(1).FirstTrans = 0
    ARCust(1).LastTrans = 0
    PUT CustFile, Cnt&, ARCust(1)
    ShowPctComp Cnt&, NumCRec&
  NEXT
  
  FOR Cnt& = 1 TO NumTRec&
    GET TranFile, Cnt&, ARTran(1)
    CustRec& = VAL(ARTran(1).CustomerNumber)
    IF (CustRec& > 0) AND (CustRec& <= NumCRec&) THEN
      GET CustFile, CustRec&, ARCust(1)
      IF ARCust(1).LastTrans = 0 THEN
        ARCust(1).FirstTrans = Cnt&
        ARCust(1).LastTrans = Cnt&
        PUT CustFile, CustRec&, ARCust(1)
        ARTran(1).NextTrans = 0
        PUT TranFile, Cnt&, ARTran(1)
      ELSE
        GET TranFile, ARCust(1).LastTrans, ARTran(2)  'get old last tr
        ARTran(2).NextTrans = Cnt&                    'point it to next tr
        PUT TranFile, ARCust(1).LastTrans, ARTran(2)  'put it back
        ARCust(1).LastTrans = Cnt&                    'set new cust last TR
        PUT CustFile, CustRec&, ARCust(1)             'put it back
        ARTran(1).NextTrans = 0
        PUT TranFile, Cnt&, ARTran(1)
      END IF
      ShowPctCompL Cnt&, NumTRec&
    ELSE
      'BadTran = BadTran + 1
    END IF
  NEXT

  CLOSE
'LOCATE 5, 1: PRINT "Bad:"; BadCnt
'WaitForAction
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
  DisplayARScrn ScrnName$
  help$ = "Change License Print Status"
  PrintHelp help$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      DevSpec$ = LEFT$(Form$(1, 0), 1)
      DONE = True
    CASE EscKey
      ExitFlag = True
      DONE = True
    END SELECT
  LOOP UNTIL DONE
  IF ExitFlag THEN EXIT SUB
  
  'REDIM ARCustRec(1) AS ARCustRecType     ' open customer file
  CustRecLen = LEN(ARCustRec(1))
  TrHandle = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
  
  FOR Cnt = 1 TO TrNumRecs
    GET TrHandle, Cnt, ARCustRec(1)
    IF (ARCustRec(1).Deleted <> "Y") AND (ARCustRec(1).Inactive <> "Y") THEN
      IF Num2Date$(ARCustRec(1).VALID) = Form$(1, 0) THEN
        help$ = ARCustRec(1).CUSTNAME
        PrintHelp help$
        ARCustRec(1).IssueLicense = "Y"
        PUT TrHandle, Cnt, ARCustRec(1)
      END IF
    END IF
  NEXT Cnt
  CLOSE         'Close all open files now
  
END SUB

SUB ShowNoCodes
  ok = MsgBox("AR", "ARNOCODE")
END SUB

SUB SortARNameIndex

  SHARED MChoice$, DScrn

  IF DScrn THEN
    ClearBack
    ShowProcessingScrn "Rebuilding Name Index."
  END IF

  KillFile "ARCUST.IDX"

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
  
  help$ = "Sorting Customer Index"
  PrintHelp help$

  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen

  REDIM Array(1 TO NumOfArRecs)  AS Struct

  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  
  FOR Cnt = 1 TO NumOfArRecs
    GET ARFile, Cnt, ARCustRec(1)
    Array(Cnt).who = ARCustRec(1).SORTNAME + "    "
    Array(Cnt).RecNum = Cnt
  NEXT Cnt
  
  SortT Array(Start), NumOfArRecs, Dir, SSize, MOff, MSize
  
  FOR Cnt = 1 TO NumOfArRecs
    ARCustIdxRec(1).IDXNAME = Array(Cnt).who
    ARCustIdxRec(1).IDXRECORD = Array(Cnt).RecNum
    ShowPctComp Cnt, NumOfArRecs
    PUT ARIdxFile, Cnt, ARCustIdxRec(1)
  NEXT Cnt
  CLOSE
  ERASE Array
  IF DScrn THEN
    ClearBack
    DisplayARScrn "UPDATEOK"
    WaitForAction
  END IF

END SUB

