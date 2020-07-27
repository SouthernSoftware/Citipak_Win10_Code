DEFINT A-Z
DECLARE SUB EDRealTR ()
DECLARE SUB EDPersTR ()
DECLARE SUB ShowCustPersList (CustRec&, TaxType%)
DECLARE SUB OpenPPTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenRETaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB ShowCustPropList (CustRec&, TaxType%)
DECLARE SUB ShowCustHistory (CustRec&, TaxType%)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
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
'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()


  CONST False = 0, True = NOT False

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'

  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType


  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  STACK 5000

  '--Dim the choice array to the number of menu items
  REDIM Mchoice$(1 TO 3)

  Mchoice$(1) = "Edit - Real Estate Trans"
  Mchoice$(2) = "Edit - Personal Property Trans"
  Mchoice$(3) = "Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight

  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(Mchoice$)
    TLen = LEN(Mchoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT

  '--Center Menu within Screen
  Row = ((25 - (UBOUND(Mchoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack

    TitleBox 2, Col, MaxLen + 3, "Tax Transaction Adjuster ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      EDRealTR
    CASE 2
      EDPersTR
    CASE 3
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
  IF Exist("Taxsetup" + Ext$) THEN
    RUN "Taxsetup"
  ELSE
    HideCursor
    ClearScrn
  END IF

  END

SUB EDPersTR

  SHARED Choice$()

  REDIM TaxInq(1) AS TaxCustInqType

  TaxType% = 2 'Personal Property

  ClearBack
  LookUp RecNo&, "Customer Inquiry", False, True, False, TaxType%
  IF RecNo& <= 0 THEN
    GOTO CustInqExit2
  END IF
     CustAcct& = RecNo&
  FirstTime = True

  LibName$ = "TAX"
  ScrnName$ = "PCUSTINQ"

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
  ClearBack

  ShowCursor

  DisplayTaxScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadInqData2
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode

    CASE F4KEY
    IF CustAcct& > 0 THEN
      ShowCustHistory -CustAcct&, 2'DON'T CHANGE THIS
      Action = 1
    END IF
    CASE F7KEY
    IF CustAcct& > 0 THEN
      ShowCustPersList CustAcct&, 1'DON'T CHANGE THIS
      Action = 1

    END IF


    CASE EscKey
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag


CustInqExit2:
EXIT SUB

LoadInqData2:
  RealValue# = 0
  SeniExmp# = 0
  OthrExmp# = 0
  PersValue# = 0
  MOBHValue# = 0
  MERHValue# = 0
  FarmValue# = 0
  MachValue# = 0

  OpenPPTaxCustFile NumOfTaxRecs, TaxFile
  GET TaxFile, RecNo&, TaxCustRec(1)
  CLOSE TaxFile

  REDIM PersRecs(0) AS LONG
  REDIM PropRecs(0) AS LONG

  GetPersRecList PersRecs(), RecNo&
  GetPropRecList PropRecs(), RecNo&

  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  IF PropRecs(0) > 0 THEN
    FOR Cnt& = 1 TO PropRecs(0)
      GET PropTaxFile, PropRecs(Cnt&), PropertyRec(1)
      RealValue# = Round#(RealValue# + PropertyRec(1).PROPVALU)
      SeniExmp# = Round#(SeniExmp# + PropertyRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PropertyRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PropTaxFile

  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  IF PersRecs(0) > 0 THEN
    FOR Cnt& = 1 TO PersRecs(0)
      GET PersTaxFile, PersRecs(Cnt&), PersRec(1)
      PersValue# = Round#(PersValue# + PersRec(1).PERSVAL)
      MOBHValue# = Round#(MOBHValue# + PersRec(1).MHVALUE)
      MERHValue# = Round#(MERHValue# + PersRec(1).MCVALUE)
      FarmValue# = Round#(FarmValue# + PersRec(1).CVALUE)
      MachValue# = Round#(MachValue# + PersRec(1).MTVALUE)
      SeniExmp# = Round#(SeniExmp# + PersRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PersRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PersTaxFile

  TaxInq(1).ACCT = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CUSTNAME = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
  TaxInq(1).HPHONE = TaxCustRec(1).HPHONE
  TaxInq(1).CSSN = TaxCustRec(1).CSSN
  TaxInq(1).WPHONE = TaxCustRec(1).WPHONE
  TaxInq(1).ADDR1 = TaxCustRec(1).ADDR1
  TaxInq(1).ADDR2 = TaxCustRec(1).ADDR2
  TaxInq(1).CITY = TaxCustRec(1).CITY
  TaxInq(1).STATE = TaxCustRec(1).STATE
  TaxInq(1).ZIP = TaxCustRec(1).ZIP
  TaxInq(1).ACTIVE = TaxCustRec(1).ACTIVE
  TaxInq(1).INTEREST = TaxCustRec(1).INTEREST
  TaxInq(1).EXEMPT = TaxCustRec(1).TaxExempt
  TaxInq(1).PENALTY = TaxCustRec(1).PENALTY
  TaxInq(1).SRCITDIS = 0
  TaxInq(1).ODISCOUN = 0
  TaxInq(1).PROPVALU = PersValue#
  TaxInq(1).PERSVAL = MOBHValue#
  TaxInq(1).MHVALUE = MERHValue#
  TaxInq(1).MCVALUE = FarmValue#
  TaxInq(1).CVALUE = MachValue#
  TaxInq(1).MTVALUE = 0

  BCopy VARSEG(TaxInq(1)), VARPTR(TaxInq(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
RETURN

END SUB

SUB EDRealTR

  SHARED Choice$()

  REDIM TaxInq(1) AS RETaxCustInqType
  REDIM TempScrn(0)

  TaxType% = 1
  ClearBack
  LookUp RecNo&, "Customer Inquiry", False, True, False, TaxType%

  CustAcct& = RecNo&

  IF RecNo& <= 0 THEN
    GOTO CustInqExit
  END IF

  FirstTime = True

  LibName$ = "TAX"
  ScrnName$ = "RCUSTINQ"

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
  ClearBack

  ShowCursor

  DisplayTaxScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadInqData
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode

    CASE F4KEY
    IF CustAcct& > 0 THEN
      ShowCustHistory CustAcct&, 1'DON'T CHANGE THIS
      Action = 1
    END IF
    
    CASE EscKey
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag


CustInqExit:
EXIT SUB

LoadInqData:
  RealValue# = 0
  SeniExmp# = 0
  OthrExmp# = 0
  PersValue# = 0
  MOBHValue# = 0
  MERHValue# = 0
  FarmValue# = 0
  MachValue# = 0

  OpenRETaxCustFile NumOfTaxRecs, TaxFile
  GET TaxFile, RecNo&, TaxCustRec(1)
  CLOSE TaxFile

  REDIM PersRecs(0) AS LONG
  REDIM PropRecs(0) AS LONG

  GetPersRecList PersRecs(), RecNo&
  GetPropRecList PropRecs(), RecNo&

  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  IF PropRecs(0) > 0 THEN
    FOR Cnt& = 1 TO PropRecs(0)
      GET PropTaxFile, PropRecs(Cnt&), PropertyRec(1)
      RealValue# = Round#(RealValue# + PropertyRec(1).PROPVALU)
      BldgValue# = Round#(SeniExmp# + PropertyRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PropertyRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PropTaxFile


  TaxInq(1).ACCT = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CUSTNAME = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
  TaxInq(1).HPHONE = TaxCustRec(1).HPHONE
  TaxInq(1).CSSN = TaxCustRec(1).CSSN
  TaxInq(1).WPHONE = TaxCustRec(1).WPHONE
  TaxInq(1).ADDR1 = TaxCustRec(1).ADDR1
  TaxInq(1).ADDR2 = TaxCustRec(1).ADDR2
  TaxInq(1).CITY = TaxCustRec(1).CITY
  TaxInq(1).STATE = TaxCustRec(1).STATE
  TaxInq(1).ZIP = TaxCustRec(1).ZIP
  TaxInq(1).ACTIVE = TaxCustRec(1).ACTIVE
  TaxInq(1).INTEREST = TaxCustRec(1).INTEREST
  TaxInq(1).EXEMPT = TaxCustRec(1).TaxExempt
  TaxInq(1).PENALTY = TaxCustRec(1).PENALTY
  TaxInq(1).ODISCOUN = OthrExmp#
  TaxInq(1).PROPVALU = RealValue#
  TaxInq(1).PERSVAL = BldgValue#

  BCopy VARSEG(TaxInq(1)), VARPTR(TaxInq(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
RETURN
END SUB

SUB OpenPPTaxCustFile (NumOfTaxRecs, TaxFile)

  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

END SUB

SUB OpenRETaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
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

