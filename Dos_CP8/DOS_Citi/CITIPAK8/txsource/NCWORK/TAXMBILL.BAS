DEFINT A-Z

DECLARE SUB PostMBills ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB EditTaxEntry ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB OpenMBEditFile (NumOfEditRecs%, MFile%)
DECLARE SUB POSTEDITLIST ()
DECLARE SUB PrintEditList ()
DECLARE SUB WaitForAction ()
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB TaxEntry (RecNo&)
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB AbtractListing ()
DECLARE SUB BalanceListing ()
DECLARE SUB MortgageCodeList ()
DECLARE SUB MasterValueList ()
DECLARE SUB TransactionJournal ()
DECLARE SUB LateListing ()
DECLARE SUB CustomerInquiry ()
DECLARE SUB SrCitizensList ()
DECLARE SUB Labels ()
DECLARE SUB AdListing ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
  
DECLARE SUB CustomerListing ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (B#)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  DIM SHARED MTaxTrans(1) AS TaxMTransactionType
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 12)
  
  MChoice$(1) = "Enter Tax Bill - Manual"
  MChoice$(2) = "Edit Transaction"
  MChoice$(3) = "Print Transaction Edit Listing"
  MChoice$(4) = "Post Transactions"
  MChoice$(5) = "Exit to DOS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 3
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Tax Billing Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      ClearBack
      LookUp RecNo&, "Customer", False, True, False
      IF RecNo& > 0 THEN
        TaxEntry RecNo&
      END IF
    CASE 2
      EditTaxEntry
    CASE 3
      PrintEditList
    CASE 4
      PostMBills
    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  
  IF INSTR(COMMAND$, "TEST") > 0 THEN
    HideCursor
    ClearScrn
    END
  END IF
  
  IF Exist("Taxmenu" + Ext$) THEN
    RUN "TaxMenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB EditTaxEntry
  
  REDIM ScrnArray(0)
  
  SHARED MChoice$
  SHARED Choice$
  
  REDIM Array(1 TO 1)  AS Struct
  
  MaxLen = 40   'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  ClearBack
  
MainBody:
  
  LibName$ = "TAXNP"
  ScrnName$ = "MTAXBILL"
  
  LOCATE 1, 1, 0
  help$ = "Edit Transactions"
  ShowCursor
  
  'LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
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
  REDIM Choice$(3, 0)
  Choice$(0, 0) = "9"
  Choice$(1, 0) = "Combined"
  Choice$(2, 0) = "Real"
  Choice$(3, 0) = "Personal"
  
  Action = 1
  NoProcess = False
  
  MFile = FREEFILE
  OPEN "TAXMEDIT.DAT" FOR RANDOM AS #MFile LEN = LEN(MTaxTrans(1))
  NumOfCatRecs = LOF(MFile) \ LEN(MTaxTrans(1))

'  GFile = FREEFILE
'  OPEN "TAXMEDIT.DA2" FOR RANDOM AS #GFile LEN = LEN(MTaxTrans(1))
 

  IF NumOfCatRecs = 0 THEN
    GOTO MBillExit
  END IF

  REDIM Array(1 TO NumOfCatRecs)  AS Struct

  IF NumOfCatRecs = 0 THEN
    GOTO MBillExit
  END IF
  
  FOR Cnt = 1 TO NumOfCatRecs
    GET MFile, Cnt, MTaxTrans(1)
    LSET Array(Cnt).who = QPTrim$(MTaxTrans(1).SName)
    Array(Cnt).RecNum = Cnt
    'IF MTaxTrans(1).Deleted <> True THEN
    '  PUT GFile, , MTaxTrans(1)
    'END IF
  NEXT
  'CLOSE : END

  SortT Array(1), NumOfCatRecs, 0, 16, 0, 14
  RCnt = 0
  FOR Cnt = 1 TO NumOfCatRecs
    GET MFile, Array(Cnt).RecNum, MTaxTrans(1)
    'MTaxTrans(1).Deleted = 0
    'PUT MFile, Array(Cnt).RecNum, MTaxTrans(1)
    IF NOT MTaxTrans(1).Deleted = True THEN
      RCnt = RCnt + 1
      REDIM PRESERVE MChoice$(1 TO RCnt)
      MChoice$(RCnt) = SPACE$(45)
      LSET MChoice$(RCnt) = LEFT$(MTaxTrans(1).TName, 18) + " "
      MID$(MChoice$(RCnt), 20) = MTaxTrans(1).Desc
      MID$(MChoice$(RCnt), 43) = MKI$(Array(Cnt).RecNum)
    END IF
  NEXT Cnt
  CLOSE

  IF RCnt = 0 THEN
    GOTO MBillExit
  END IF

  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = " Customer Name      Description  <ESC>-QUIT"
  
  '--Center Menu within Screen
  Row = 8
  
  Col = MaxLen \ 2
  VAction = 0
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    help$ = "Add or Edit Catagories"
    LOCATE Row, Col, 0
    ShowCursor
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
    
    IF Ky$ = CHR$(27) THEN
      RecNo = 0
      ExitFlag = True
    ELSE
      RecNo = CVI(MID$(MChoice$(Choice), 43, 2)) 'Array(Choice).RecNum
      ExitFlag = True
    END IF
    
  LOOP UNTIL ExitFlag
  
  IF RecNo = 0 THEN
    CLOSE
    GOTO MBillExit
    'EXIT SUB
  END IF
  
  ShowCursor
  ClearBack
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  LibFile2Scrn LibName$, "MTAXDELB", MonoCode%, -2, ErrorCode%
  
  Action = 1
  
  FirstTime = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadBillInfo
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB SaveEditTrans
      CLOSE MFile
      GOTO MainBody
    CASE F3Key
      Ok = MsgBox(LibName$, "MBOK2DEL")
      IF Ok = 2 THEN
        GOSUB DeleteMBill
        EXIT SUB
      ELSE
        Action = 1
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
MBillExit:
  EXIT SUB
  
SaveEditTrans:
  SaveScrn ScrnArray()
  DisplayTaxScrn "UPDATDSK"
  MTaxTrans(1).Account = Value#(Form$(1, 0), 0)
  MTaxTrans(1).TransDate = Date2Num%(Form$(5, 0))
  MTaxTrans(1).TaxYear = Value#(Form$(4, 0), 0)
  MTaxTrans(1).Desc = "M Tax Bill #" + QPTrim$(Form$(3, 0))
  MTaxTrans(1).TaxAmount = Value#(Form$(6, 0), 0)
  MTaxTrans(1).IntAmount = Value#(Form$(7, 0), 0)
  MTaxTrans(1).AdColAmount = Value#(Form$(8, 0), 0)
  MTaxTrans(1).BillType = LEFT$(Form$(9, 0), 1)
  MTaxTrans(1).Deleted = 0
  MTaxTrans(1).Padding = ""
  MFile = FREEFILE
  OPEN "TAXMEDIT.DAT" FOR RANDOM AS #MFile LEN = LEN(MTaxTrans(1))
  PUT MFile, RecNo, MTaxTrans(1)
  CLOSE MFile
  RestScrn ScrnArray()
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
RETURN
  
LoadBillInfo:
  MFile = FREEFILE
  OPEN "TAXMEDIT.DAT" FOR RANDOM AS #MFile LEN = LEN(MTaxTrans(1))
  GET MFile, RecNo, MTaxTrans(1)
  CLOSE
  LSET Form$(1, 0) = QPTrim$(STR$(MTaxTrans(1).Account))
  LSET Form$(2, 0) = MTaxTrans(1).TName
  LSET Form$(3, 0) = LTRIM$(MID$(MTaxTrans(1).Desc, 11, 10))
  LSET Form$(4, 0) = QPTrim$(STR$(MTaxTrans(1).TaxYear))
  LSET Form$(5, 0) = Num2Date$(MTaxTrans(1).TransDate)
  LSET Form$(6, 0) = QPTrim$(STR$(MTaxTrans(1).TaxAmount))
  LSET Form$(7, 0) = QPTrim$(STR$(MTaxTrans(1).IntAmount))
  LSET Form$(8, 0) = QPTrim$(STR$(MTaxTrans(1).AdColAmount))
  
  IF MTaxTrans(1).BillType = "C" THEN
    LSET Form$(9, 0) = "COMBINED"
  END IF
  
  IF MTaxTrans(1).BillType = "R" THEN
    LSET Form$(9, 0) = "REAL"
  END IF
  
  IF MTaxTrans(1).BillType = "P" THEN
    LSET Form$(9, 0) = "PERSONAL"
  END IF
RETURN

DeleteMBill:
  MFile = FREEFILE
  OPEN "TAXMEDIT.DAT" FOR RANDOM AS #MFile LEN = LEN(MTaxTrans(1))
  GET MFile, RecNo, MTaxTrans(1)
  MTaxTrans(1).Deleted = True
  PUT MFile, RecNo, MTaxTrans(1)
  CLOSE
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
RETURN

END SUB

SUB OpenMBEditFile (NumOfEditRecs, MFile)
  CLOSE MFile
  MFile = FREEFILE
  OPEN "TAXMEDIT.DAT" FOR RANDOM AS #MFile LEN = LEN(MTaxTrans(1))
  NumOfEditRecs = LOF(MFile) / LEN(MTaxTrans(1))
  
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

SUB PostMBills
  
  REDIM TempScrn(0)
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM Revenue(1) AS RevSourceType
  
  RealRecLen = LEN(RealRec(1))
  PersRecLen = LEN(PersRec(1))
  
  ClearBack
  
  OpenMBEditFile NumOfEditRecs, MFile
  
  IF NumOfEditRecs < 1 THEN
    GOSUB ShowPostError
    GOTO PostExit
  END IF
  
  PostFlag = MsgBox%("TAX", "POSTEDIT")
  
  SELECT CASE PostFlag
  CASE 1
    GOTO Postem
  CASE 2
    CursorOff
    ClearBack
    DisplayTaxScrn "POSTCAN"
    WaitForAction
    GOTO PostExit
  END SELECT
  
Postem:
  ClearBack
  ShowProcessingScrn "Posting Manual Billing Transactions"
  
  'Open All Needed Files Here
  
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  
  CustFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM AS CustFile LEN = LEN(TaxCust(1))
  
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  FOR Cnt& = 1 TO NumOfEditRecs
    GET #MFile, Cnt&, MTaxTrans(1)
    IF NOT MTaxTrans(1).Deleted = True THEN
      LineTotal# = Round#(MTaxTrans(1).TaxAmount + MTaxTrans(1).IntAmount + MTaxTrans(1).AdColAmount)
      IF LineTotal# <> 0 THEN
        'Update the Transaction File First
        TaxTrans(1).TransDate = MTaxTrans(1).TransDate
        TaxTrans(1).TaxYear = MTaxTrans(1).TaxYear
        TaxTrans(1).TranType = 1                '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
        TaxTrans(1).BillType = MTaxTrans(1).BillType            'R=Real P=Personal Property C=Combined (NC/GA)
        TaxTrans(1).Amount = LineTotal#         'Total Transaction Amount
        TaxTrans(1).Revenue.Principle1 = MTaxTrans(1).TaxAmount
        TaxTrans(1).Revenue.Principle2 = 0
        TaxTrans(1).Revenue.Principle3 = 0
        TaxTrans(1).Revenue.Principle4 = 0
        TaxTrans(1).Revenue.Principle5 = 0
        TaxTrans(1).Revenue.interest = MTaxTrans(1).IntAmount
        TaxTrans(1).Revenue.Penalty = 0
        TaxTrans(1).Revenue.Collection = MTaxTrans(1).AdColAmount
        TaxTrans(1).Revenue.Future1 = 0
        TaxTrans(1).Revenue.Future2 = 0
        TaxTrans(1).Revenue.Principle1Pd = 0
        TaxTrans(1).Revenue.Principle2Pd = 0
        TaxTrans(1).Revenue.Principle3Pd = 0
        TaxTrans(1).Revenue.Principle4Pd = 0
        TaxTrans(1).Revenue.Principle5Pd = 0
        TaxTrans(1).Revenue.InterestPd = 0
        TaxTrans(1).Revenue.PenaltyPd = 0
        TaxTrans(1).Revenue.CollectionPd = 0
        TaxTrans(1).Revenue.Future1Pd = 0
        TaxTrans(1).Revenue.Future2Pd = 0
        
        TaxTrans(1).Description = MTaxTrans(1).Desc
        TaxTrans(1).Posted2GL = "Y" ' Do Not Allow Posting 2GL of Manual Entries Probably Already Reflected in General Ledger
        TaxTrans(1).CustomerRec = MTaxTrans(1).Account
        TaxTrans(1).LastTrans = 0
        TaxTrans(1).BelongTo = 0
        TaxTrans(1).Padding = ""
        
        'Increment Transaction File Record Count
        NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
        PUT #TransFile, NextRecord&, TaxTrans(1)
        
        'Update the Customer Pointers Now
        GET #CustFile, MTaxTrans(1).Account, TaxCust(1)
        
        IF TaxCust(1).LastTrans = 0 THEN
          TaxCust(1).LastTrans = NextRecord&
          PUT #CustFile, MTaxTrans(1).Account, TaxCust(1)
        ELSE
          Previous& = TaxCust(1).LastTrans
          TaxCust(1).LastTrans = NextRecord&
          PUT #CustFile, MTaxTrans(1).Account, TaxCust(1)
          
          GET #TransFile, NextRecord&, TaxTrans(1)
          TaxTrans(1).LastTrans = Previous&
          PUT #TransFile, NextRecord&, TaxTrans(1)
        END IF
      END IF    'master loop if balance<>0
    END IF
    ShowPctComp Cnt&, NumOfEditRecs
  NEXT
  
  CLOSE
  
  DisplayTaxScrn "UPDATEOK"
  
  WaitForAction
  
  'Now Delete the Tax Edit File so Duplicate's Cannot Be Reproduced
  S$ = "DEL TAXMEDIT.DAT"
  SHELL S$
  
PostExit:
  EXIT SUB
  
  
ShowPostError:
  ClearBack
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "THERE ARE NO ENTRIES TO POST!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
RETURN
  
END SUB

SUB PrintEditList
  SHARED Choice$()
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  REDIM Cnt(3), Year(99), TaxAmt#(99), IntAmt#(99), AdAmt#(99), RYear(99), RTaxAmt#(99), RIntAmt#(99), RAdAmt#(99), PYear(99), PTaxAmt#(99), PIntAmt#(99), PAdAmt#(99)
  ReportFile$ = "TAXMEDIT.PRN"  'Report File Name
  dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "MBEDITRP"
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 1)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingNIndex = True
      ELSE
        UsingAIndex = True
      END IF
      DevSpec$ = LEFT$(Form$(2, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintBalanceRptHeader
  
  OpenMBEditFile NumOfEditRecs, MFile
  
  IF UsingNIndex AND NumOfEditRecs > 0 THEN
    GOSUB GetNameIndex
  END IF
  IF UsingAIndex AND NumOfEditRecs > 0 THEN
    GOSUB GetAcctIndex
  END IF
  
  
  FOR Cnt& = 1 TO NumOfEditRecs
    CustRecNo = Array(Cnt&).RecNum
    GET MFile, CustRecNo, MTaxTrans(1)
    IF NOT MTaxTrans(1).Deleted = True THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintBalanceRptHeader
      END IF
    
      LineTotal# = Round#(MTaxTrans(1).TaxAmount + MTaxTrans(1).IntAmount + MTaxTrans(1).AdColAmount)
      CustCnt = CustCnt + 1       'Totals # of Bills Entered Manually
      PRINT #RptHandle, USING "######"; MTaxTrans(1).Account;
      PRINT #RptHandle, TAB(10); MTaxTrans(1).TName; TAB(40); MTaxTrans(1).Desc; TAB(72); MTaxTrans(1).TaxYear
      PRINT #RptHandle, MTaxTrans(1).TransDate; TAB(15); USING "######,#.##"; MTaxTrans(1).TaxAmount;
      PRINT #RptHandle, TAB(28); USING "####,#.##"; MTaxTrans(1).IntAmount;
      PRINT #RptHandle, TAB(40); USING "####,#.##"; MTaxTrans(1).AdColAmount;
      PRINT #RptHandle, TAB(55); USING "######,#.##"; LineTotal#;
      PRINT #RptHandle, TAB(79); MTaxTrans(1).BillType
      PRINT #RptHandle, STRING$(79, "-")
      LineCnt = LineCnt + 3
    
      'Summarize the Totals Here
    
      GTTax# = GTTax# + MTaxTrans(1).TaxAmount
      GTInt# = GTInt# + MTaxTrans(1).IntAmount
      GTAd# = GTAd# + MTaxTrans(1).AdColAmount
    
    
      IF MTaxTrans(1).BillType = "C" THEN
        IF Cnt(1) = 0 THEN
          Year(1) = MTaxTrans(1).TaxYear
          TaxAmt#(1) = MTaxTrans(1).TaxAmount
          IntAmt#(1) = MTaxTrans(1).IntAmount
          AdAmt#(1) = MTaxTrans(1).AdColAmount
          Cnt(1) = 1
        ELSE
          FOR LCnt = 1 TO Cnt(1)
            IF Year(LCnt) = MTaxTrans(1).TaxYear THEN
              TaxAmt#(LCnt) = TaxAmt#(LCnt) + MTaxTrans(1).TaxAmount
              IntAmt#(LCnt) = IntAmt#(LCnt) + MTaxTrans(1).IntAmount
              AdAmt#(LCnt) = AdAmt#(LCnt) + MTaxTrans(1).AdColAmount
              GOTO NextLine
            END IF
          NEXT LCnt
          Cnt(1) = Cnt(1) + 1
          Year(Cnt(1)) = MTaxTrans(1).TaxYear
          TaxAmt#(Cnt(1)) = MTaxTrans(1).TaxAmount
          IntAmt#(Cnt(1)) = MTaxTrans(1).IntAmount
          AdAmt#(Cnt(1)) = MTaxTrans(1).AdColAmount
        END IF
      END IF
    
      IF MTaxTrans(1).BillType = "R" THEN
        IF Cnt(2) = 0 THEN
          RYear(1) = MTaxTrans(1).TaxYear
          RTaxAmt#(1) = MTaxTrans(1).TaxAmount
          RIntAmt#(1) = MTaxTrans(1).IntAmount
          RAdAmt#(1) = MTaxTrans(1).AdColAmount
          Cnt(2) = 1
        ELSE
          FOR LCnt = 1 TO Cnt(2)
            IF RYear(LCnt) = MTaxTrans(1).TaxYear THEN
              RTaxAmt#(LCnt) = RTaxAmt#(LCnt) + MTaxTrans(1).TaxAmount
              RIntAmt#(LCnt) = RIntAmt#(LCnt) + MTaxTrans(1).IntAmount
              RAdAmt#(LCnt) = RAdAmt#(LCnt) + MTaxTrans(1).AdColAmount
              GOTO NextLine
            END IF
          NEXT LCnt
          Cnt(2) = Cnt(2) + 1
          RYear(Cnt(2)) = MTaxTrans(1).TaxYear
          RTaxAmt#(Cnt(2)) = MTaxTrans(1).TaxAmount
          RIntAmt#(Cnt(2)) = MTaxTrans(1).IntAmount
          RAdAmt#(Cnt(2)) = MTaxTrans(1).AdColAmount
        END IF
      END IF
      IF MTaxTrans(1).BillType = "P" THEN
        IF Cnt(3) = 0 THEN
          PYear(1) = MTaxTrans(1).TaxYear
          PTaxAmt#(1) = MTaxTrans(1).TaxAmount
          PIntAmt#(1) = MTaxTrans(1).IntAmount
          PAdAmt#(1) = MTaxTrans(1).AdColAmount
          Cnt(3) = 1
        ELSE
          FOR LCnt = 1 TO Cnt(3)
            IF PYear(LCnt) = MTaxTrans(1).TaxYear THEN
              PTaxAmt#(LCnt) = PTaxAmt#(LCnt) + MTaxTrans(1).TaxAmount
              PIntAmt#(LCnt) = PIntAmt#(LCnt) + MTaxTrans(1).IntAmount
              PAdAmt#(LCnt) = PAdAmt#(LCnt) + MTaxTrans(1).AdColAmount
              GOTO NextLine
            END IF
          NEXT LCnt
          Cnt(3) = Cnt(3) + 1
          PYear(Cnt(3)) = MTaxTrans(1).TaxYear
          PTaxAmt#(Cnt(3)) = MTaxTrans(1).TaxAmount
          PIntAmt#(Cnt(3)) = MTaxTrans(1).IntAmount
          PAdAmt#(Cnt(3)) = MTaxTrans(1).AdColAmount
        END IF
      
      END IF
    END IF
NextLine:
  NEXT
  
  GOSUB PrintBalanceRptEnding
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, TaxCustRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintBalanceRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(23); "Manual Tax Bill Entry Edit Report"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  PRINT #RptHandle, "Acct #"; TAB(10); "Name"; TAB(40); "Description"; TAB(72); "Tax Year"
  PRINT #RptHandle, "Trans Date"; TAB(15); "Tax Amount"; TAB(28); "Int't Amt"; TAB(40); "Ad/Col Amt"; TAB(55); "Total Amt"; TAB(78); "BT"
  PRINT #RptHandle, dash80$
  LineCnt = 6
RETURN
  
PrintBalanceRptEnding:
  PRINT #RptHandle, dash80$
  PRINT #RptHandle, "Total Bills Entered: "; USING "#####"; CustCnt
  PRINT #RptHandle, "   Total Tax Amount: "; USING "$$######,#.##"; GTTax#
  PRINT #RptHandle, "   Total Int Amount: "; USING "$$######,#.##"; GTInt#
  PRINT #RptHandle, "Total Ad/Col Amount: "; USING "$$######,#.##"; GTAd#
  PRINT #RptHandle, "   Total A/R Amount: "; USING "$$######,#.##"; Round#(GTTax# + GTInt# + GTAd#)
  PRINT #RptHandle, FF$
  'Calc and print summary by years entered
  'Check Combined Bills First
  IF Cnt(1) > 0 THEN
    PRINT #RptHandle, "Combined Bills Total"
    PRINT #RptHandle, "Tax Year"; TAB(12); "Tax Amount"; TAB(25); "Inter't Amt"; TAB(40); "Ad/Col Amount"; TAB(60); "Total Amount"
    PRINT #RptHandle, dash80$
    FOR LCnt = 1 TO Cnt(1)
      PRINT #RptHandle, TAB(2); USING "####"; Year(LCnt);
      PRINT #RptHandle, TAB(10); USING "######,#.##"; TaxAmt#(LCnt);
      PRINT #RptHandle, TAB(25); USING "#####,#.##"; IntAmt#(LCnt);
      PRINT #RptHandle, TAB(40); USING "#####,#.##"; AdAmt#(LCnt);
      PRINT #RptHandle, TAB(60); USING "######,#.##"; Round#(TaxAmt#(LCnt) + IntAmt#(LCnt) + AdAmt#(LCnt))
      GCTaxAmt# = GCTaxAmt# + TaxAmt#(LCnt)
      GCIntAmt# = GCIntAmt# + IntAmt#(LCnt)
      GCAdAmt# = GCAdAmt# + AdAmt#(LCnt)
    NEXT LCnt
    PRINT #RptHandle, "Totals:";
    PRINT #RptHandle, TAB(10); USING "######,#.##"; GCTaxAmt#;
    PRINT #RptHandle, TAB(25); USING "#####,#.##"; GCIntAmt#;
    PRINT #RptHandle, TAB(40); USING "#####,#.##"; GCAdAmt#;
    PRINT #RptHandle, TAB(60); USING "######,#.##"; Round#(GCTaxAmt# + GCIntAmt# + GCAdAmt#)
    PRINT #RptHandle,
  END IF
  
  'Check Real Type
  IF Cnt(2) > 0 THEN
    PRINT #RptHandle, "Real Estate Bills Total"
    PRINT #RptHandle, "Tax Year"; TAB(12); "Tax Amount"; TAB(25); "Inter't Amt"; TAB(40); "Ad/Col Amount"; TAB(60); "Total Amount"
    PRINT #RptHandle, dash80$
    FOR LCnt = 1 TO Cnt(2)
      PRINT #RptHandle, TAB(2); USING "####"; RYear(LCnt);
      PRINT #RptHandle, TAB(10); USING "######,#.##"; RTaxAmt#(LCnt);
      PRINT #RptHandle, TAB(25); USING "#####,#.##"; RIntAmt#(LCnt);
      PRINT #RptHandle, TAB(40); USING "#####,#.##"; RAdAmt#(LCnt);
      PRINT #RptHandle, TAB(60); USING "######,#.##"; Round#(RTaxAmt#(LCnt) + RIntAmt#(LCnt) + RAdAmt#(LCnt))
      GRTaxAmt# = GRTaxAmt# + RTaxAmt#(LCnt)
      GRIntAmt# = GRIntAmt# + RIntAmt#(LCnt)
      GRAdAmt# = GRAdAmt# + RAdAmt#(LCnt)
    NEXT LCnt
    PRINT #RptHandle, "Totals:";
    PRINT #RptHandle, TAB(10); USING "######,#.##"; GRTaxAmt#;
    PRINT #RptHandle, TAB(25); USING "#####,#.##"; GRIntAmt#;
    PRINT #RptHandle, TAB(40); USING "#####,#.##"; GRAdAmt#;
    PRINT #RptHandle, TAB(60); USING "######,#.##"; Round#(GRTaxAmt# + GRIntAmt# + GRAdAmt#)
    PRINT #RptHandle, ""
    
  END IF
  
  'Check Personal Type
  IF Cnt(3) > 0 THEN
    PRINT #RptHandle, "Personal Bills Total"
    PRINT #RptHandle, "Tax Year"; TAB(12); "Tax Amount"; TAB(25); "Inter't Amt"; TAB(40); "Ad/Col Amount"; TAB(60); "Total Amount"
    PRINT #RptHandle, dash80$
    FOR LCnt = 1 TO Cnt(3)
      PRINT #RptHandle, TAB(2); USING "####"; PYear(LCnt);
      PRINT #RptHandle, TAB(10); USING "######,#.##"; PTaxAmt#(LCnt);
      PRINT #RptHandle, TAB(25); USING "#####,#.##"; PIntAmt#(LCnt);
      PRINT #RptHandle, TAB(40); USING "#####,#.##"; PAdAmt#(LCnt);
      PRINT #RptHandle, TAB(60); USING "######,#.##"; Round#(PTaxAmt#(LCnt) + PIntAmt#(LCnt) + PAdAmt#(LCnt))
      GPTaxAmt# = GPTaxAmt# + PTaxAmt#(LCnt)
      GPIntAmt# = GPIntAmt# + PIntAmt#(LCnt)
      GPAdAmt# = GPAdAmt# + PAdAmt#(LCnt)
    NEXT LCnt
    PRINT #RptHandle, "Totals:";
    PRINT #RptHandle, TAB(10); USING "######,#.##"; GPTaxAmt#;
    PRINT #RptHandle, TAB(25); USING "#####,#.##"; GPIntAmt#;
    PRINT #RptHandle, TAB(40); USING "#####,#.##"; GPAdAmt#;
    PRINT #RptHandle, TAB(60); USING "######,#.##"; Round#(GPTaxAmt# + GPIntAmt# + GPAdAmt#)
    
  END IF
  
RETURN
  
GetNameIndex:
  REDIM Array(1 TO NumOfEditRecs) AS Struct
  FOR Cnt = 1 TO NumOfEditRecs
    GET MFile, Cnt, MTaxTrans(1)
    Array(Cnt).who = LEFT$(UCASE$(MTaxTrans(1).SName), 12)
    Array(Cnt).RecNum = Cnt
  NEXT
  'Sort Them Here
  SortT Array(1), NumOfEditRecs, 0, LEN(Array(1)), 0, 14
RETURN
  
GetAcctIndex:
  REDIM Array(1 TO NumOfEditRecs) AS Struct
  FOR Cnt = 1 TO NumOfEditRecs
    GET MFile, Cnt, MTaxTrans(1)
    Acct$ = LTRIM$(STR$(MTaxTrans(1).Account))
    Acct$ = STRING$(14 - LEN(Acct$), " ") + Acct$
    Array(Cnt).who = Acct$
    Array(Cnt).RecNum = Cnt
  NEXT
  'Sort Them Here
  SortT Array(1), NumOfEditRecs, 0, LEN(Array(1)), 0, 14
RETURN
  
END SUB

SUB TaxEntry (RecNo&)
  SHARED Choice$()
  REDIM ScrnArray(0)
  
  LibName$ = "TAXNP"
  ScrnName$ = "MTAXBILL"
  
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
  REDIM Choice$(3, 0)
  Choice$(0, 0) = "9"
  Choice$(1, 0) = "Combined"
  Choice$(2, 0) = "Real"
  Choice$(3, 0) = "Personal"
  OpenTaxCustFile NumOfTaxRecs%, TaxFile%
  
  GET TaxFile%, RecNo&, TaxCustRec(1)
  Form$(1, 0) = STR$(RecNo&)
  Form$(2, 0) = LTRIM$(RTRIM$(TaxCustRec(1).FName) + " " + RTRIM$(TaxCustRec(1).LName$))
  
  
  Action = 1
  ClearBack
  
  LibFile2Scrn "TAXNP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  ShowCursor
  
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB SaveTrans
      ExitFlag = True
    CASE EscKey
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF ExitFlag THEN EXIT SUB
  
SaveTrans:
  SaveScrn ScrnArray()
  DisplayTaxScrn "UPDATDSK"
  MFile = FREEFILE
  OPEN "TAXMEDIT.DAT" FOR RANDOM SHARED AS #MFile LEN = LEN(MTaxTrans(1))
  NextRecord& = (LOF(MFile) / LEN(MTaxTrans(1))) + 1
  MTaxTrans(1).Account = Value#(Form$(1, 0), 0)
  MTaxTrans(1).TransDate = Date2Num%(Form$(5, 0))
  MTaxTrans(1).TaxYear = Value#(Form$(4, 0), 0)
  MTaxTrans(1).Desc = "M Tax Bill #" + QPTrim$(Form$(3, 0))
  MTaxTrans(1).TaxAmount = Value#(Form$(6, 0), 0)
  MTaxTrans(1).IntAmount = Value#(Form$(7, 0), 0)
  MTaxTrans(1).AdColAmount = Value#(Form$(8, 0), 0)
  MTaxTrans(1).BillType = Form$(9, 0)
  MTaxTrans(1).SName = TaxCustRec(1).LName
  MTaxTrans(1).TName = Form$(2, 0)
  MTaxTrans(1).Padding = ""
  PUT MFile, NextRecord&, MTaxTrans(1)
  CLOSE
  RestScrn ScrnArray()
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
RETURN
  
END SUB

