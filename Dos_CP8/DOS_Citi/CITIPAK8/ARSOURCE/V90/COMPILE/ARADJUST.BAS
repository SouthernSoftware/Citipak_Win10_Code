DEFINT A-Z
DECLARE SUB PostAdjustments ()
DECLARE SUB EnterAdjustment ()
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB PostBegBalance ()
DECLARE SUB PrintEditList ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB DisplayARScrn (ScrnName$)
DECLARE SUB WaitForAction
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
  
  '$INCLUDE: 'DefCnf.BI'
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARv95.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
  
  TYPE Struct
    Who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  CONST False = 0, True = NOT False
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
  DIM SHARED EditBegBalRec(1) AS AREditBegBalRecType
  
  STACK 8000
  
  DisplayARScrn "CLEARBAK"
  
  EnterAdjustment
  
  RUN "arcust"

SUB DisplayARScrn (ScrnName$)
  LibFile2Scrn "AR", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB EnterAdjustment
  REDIM TempScrn(0)
  
  SHARED Choice$()
  
  REDIM Choice$(3, 0)
  
  Choice$(0, 0) = "8"
  Choice$(1, 0) = "Upward"
  Choice$(2, 0) = "Downward"
  
mainbody:
  CustomerGrabed = 0
  ARIdxFile = 0
  ARFile = 0
  FirstTime = True
  LibName$ = "AR"
  ScrnName$ = "ARADJUST"
  help$ = "Set Customer Balances"
  LOCATE 1, 1, 0
  
  ShowCursor
  
  DisplayARScrn ScrnName$
  
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
  
  IF LEN(PrevDate$) = 0 THEN
    Form$(1, 0) = DATE$
  ELSE
    Form$(1, 0) = PrevDate$
  END IF
  
  Frm(1).FldNo = 1
  
  
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(10, 0) = "0"     'was 9
      LSET Form$(11, 0) = "0"     'was 10
      Action = 2
    END IF
    
    IF Frm(1).PrevFld = 3 AND CustomerGrabed = 0 THEN
      
      GOSUB GetCustomer

      IF AccountRecord = 0 THEN
        CLOSE
        Action = 1
        GOTO mainbody
      END IF
      
      REM check for existing transaction
      
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB CheckAdjustmentInfo
      IF AdjOK THEN
        OK = MsgBox("AR", "SAVEADJ")
        IF OK <> 1 THEN
          GOSUB SaveRecord
          PostAdjustments
          DisplayARScrn "UPDATEOK"
          WaitForAction
          CLOSE
          Done = True
          GOTO mainbody
        ELSE
          Action = 2
        END IF
      ELSE
        OK = MsgBox%(LibName$, "BADDATA")
        Action = 2
      END IF
      
    CASE EscKey
      EXIT SUB
    END SELECT
  LOOP
  
  
SaveRecord:
  PrevDate$ = Form$(1, 0)
  ' Keep Default Date Here
  
  CLOSE
  AREditRecLen = LEN(EditBegBalRec(1))
  ARFile = FREEFILE
  OPEN "ARADJSET.DAT" FOR RANDOM SHARED AS ARFile LEN = AREditRecLen
  NumOfArRecs = LOF(ARFile) \ AREditRecLen
  EditBegBalRec(1).TranDate = Date2Num(Form$(1, 0))
  EditBegBalRec(1).CustNumber = Form$(2, 0)
  EditBegBalRec(1).CUSTNAME = Form$(3, 0)
  
  SELECT CASE LEFT$(QPTrim$(Form$(8, 0)), 1)
  CASE "D"
    EditBegBalRec(1).TRType = 1
  CASE "U"
    EditBegBalRec(1).TRType = 2
  END SELECT
  
  EditBegBalRec(1).LicAMT = Value(Form$(10, 0), a%)
  EditBegBalRec(1).PenAMT = Value(Form$(11, 0), a%)
  
  EditBegBalRec(1).Amount = Value(Form$(12, 0), a%)
  EditBegBalRec(1).TDesc = QPTrim$(Form$(9, 0))
  
  PUT ARFile, NumOfArRecs + 1, EditBegBalRec(1)
  CLOSE ARFile
  
  RETURN
  
CheckAdjustmentInfo:
  AdjOK = True  'assume all is well
  
  AdjType$ = QPTrim$(Form$(8, 0))
  IF LEN(AdjType$) = 0 THEN
    AdjOK = False
    GOTO CheckAdjRet
  END IF
  TotalAmt# = Value(Form$(12, 0), ErrorFlag%)
  IF TotalAmt# <= 0 THEN
    AdjOK = False
    GOTO CheckAdjRet
  END IF
  
  
CheckAdjRet:
  RETURN
  
GetCustomer:
  
  CustomerGrabed = 0
  AccountRecord = VAL(Form$(2, 0))
  
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
    
    'LibName$ = "AR"
    'ScrnName$ = "ARTRANEN"
    LOCATE 1, 1, 0
    
    ShowCursor
    LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    
  END IF
  
  REM ************************************************************************
  IF AccountRecord > 0 AND AccountRecord <= NumOfArRecs THEN
    GET ARFile, AccountRecord, ARCustRec(1)
    IF ARCustRec(1).Deleted = "Y" THEN
      OK = MsgBox("ar", "CSTISDEL")
      Action = 2
      GOTO mainbody
    END IF
    
    Bal$ = SPACE$(10)
    RSET Bal$ = FUsing$(STR$(ARCustRec(1).LicBal), "#####.##")
    QPrintRC Bal$, 13, 48, -1
    RSET Bal$ = FUsing$(STR$(ARCustRec(1).PenBal), "#####.##")
    QPrintRC Bal$, 15, 48, -1
    RSET Bal$ = FUsing$(STR$(ARCustRec(1).PenBal + ARCustRec(1).LicBal), "#####.##")
    QPrintRC Bal$, 20, 48, -1
    
    Form$(2, 0) = ARCustRec(1).CUSTNUMB
    Form$(3, 0) = ARCustRec(1).CUSTNAME
    Form$(4, 0) = ARCustRec(1).ADDRESS1
    Form$(5, 0) = ARCustRec(1).CITY
    Form$(6, 0) = ARCustRec(1).STATE
    Form$(7, 0) = ARCustRec(1).ZIPCODE
    
    Fld(2).Protected = True
    Frm(1).FldNo = 8
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
    
    
    ShowCursor
    LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
    
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

SUB PostAdjustments
  
  'REDIM TempScrn(0)
  CLOSE
  
  OpenARCustFile NumOfArRecs, ARFile
  
  AREditRecLen = LEN(EditBegBalRec(1))
  AREdFile = FREEFILE
  OPEN "ARADJSET.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS AREdFile LEN = AREditRecLen
  NumOfArRecs = LOF(AREdFile) \ AREditRecLen
  
  ' See if any records to post
  ' IF NumOfArRecs = 0 THEN
  '   SaveScrn TempScrn()
  '   DisplayARScrn "ARNOTRAN"
  '   WaitForAction
  '   RestScrn TempScrn()
  '   LOCATE , , 1
  '   CLOSE
  '   EXIT SUB
  ' END IF
  
  REDIM ARTransRec(1) AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))
  ARTransFile = FREEFILE
  OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
  NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
  NextTransRec = NumOfTransRecs + 1
  
  Cnt = Cnt + 1
  GET AREdFile, Cnt, EditBegBalRec(1)
  
  IF EditBegBalRec(1).Amount <> 0 THEN
    CustRec = VAL(EditBegBalRec(1).CustNumber)
    GET ARFile, CustRec, ARCustRec(1)
    help$ = "Posting: " + LEFT$(ARCustRec(1).BILLNAME, 30)
    ' Post Transaction Record First
    ARTransRec(1).CustomerNumber = EditBegBalRec(1).CustNumber
    ARTransRec(1).TransDate = EditBegBalRec(1).TranDate
    ARTransRec(1).TransAmount = EditBegBalRec(1).Amount
    
    SELECT CASE EditBegBalRec(1).TRType
      
    CASE 1      'downward
      ARTransRec(1).TransType = 100
      Adj$ = "DNA "
      ARTransRec(1).LicAMT = -EditBegBalRec(1).LicAMT
      ARTransRec(1).PenAMT = -EditBegBalRec(1).PenAMT
      ARCustRec(1).LicBal = Round#(ARCustRec(1).LicBal - EditBegBalRec(1).LicAMT)
      ARCustRec(1).PenBal = Round#(ARCustRec(1).PenBal - EditBegBalRec(1).PenAMT)
      
    CASE 2      'upward
      ARTransRec(1).TransType = 101
      Adj$ = "UPA "
      ARTransRec(1).LicAMT = EditBegBalRec(1).LicAMT
      ARTransRec(1).PenAMT = EditBegBalRec(1).PenAMT
      ARCustRec(1).LicBal = Round#(ARCustRec(1).LicBal + EditBegBalRec(1).LicAMT)
      ARCustRec(1).PenBal = Round#(ARCustRec(1).PenBal + EditBegBalRec(1).PenAMT)
    END SELECT
    TBal# = Round#(ARCustRec(1).LicBal + ARCustRec(1).PenBal)
    ARCustRec(1).AcctBal = TBal#
    ARTransRec(1).BalanceAfterTrans = TBal#
    
    Adj$ = Adj$ + QPTrim$(EditBegBalRec(1).TDesc)
    ARTransRec(1).TransDesc = Adj$
    
    ARTransRec(1).CashAmount = 0                'EditBegBalRec(1).Amount
    ARTransRec(1).ChkAmount = 0
    ARTransRec(1).ExtraRoom = ""
    ARTransRec(1).NextTrans = 0
    
    PUT ARTransFile, NextTransRec, ARTransRec(1)
    
    '      GET ARFile, VAL(EditBegBalRec(1).CustNumber), ARCustRec(1)
    
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
  END IF
  
  CLOSE
  CALL KILLFile("ARADJSET.DAT")
  ' Show All Posted
  'SaveScrn TempScrn()
  'DisplayARScrn "ARPOSTED"
  'WaitForAction
  'RestScrn TempScrn()
  'LOCATE , , 1
  'CLOSE
  
END SUB

