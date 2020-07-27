DEFINT A-Z
DECLARE SUB EDShowCustHistory (CustRec&, TaxType%)
DECLARE SUB DisplayEDTRScrn (ScrnName$)
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
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION FUsing$ (Numb$, Fmt$)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)


  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'

  ''$INCLUDE: 'TAXCONST.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType

  TYPE FLen2
    V AS STRING * 64
  END TYPE

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

SUB DisplayEDTRScrn (ScrnName$)
  LibFile2Scrn "TAXEDTR", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  IF ErrorCode% <> 0 THEN
    PRINT "Screen Error: "; ScrnName$
    END
  END IF


END SUB

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
      EDShowCustHistory -CustAcct&, 2  'DON'T CHANGE THIS
      Action = 1
    END IF
    CASE F7KEY
    IF CustAcct& > 0 THEN
      ShowCustPersList CustAcct&, 1  'DON'T CHANGE THIS
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
      PersValue# = Round#(PersValue# + PersRec(1).PersVal)
      MOBHValue# = Round#(MOBHValue# + PersRec(1).MHVALUE)
      MERHValue# = Round#(MERHValue# + PersRec(1).MCVALUE)
      FarmValue# = Round#(FarmValue# + PersRec(1).CVALUE)
      MachValue# = Round#(MachValue# + PersRec(1).MTVALUE)
      SeniExmp# = Round#(SeniExmp# + PersRec(1).EXMPSENI)
      OthrExmp# = Round#(OthrExmp# + PersRec(1).EXMPOTHR)
    NEXT
  END IF
  CLOSE PersTaxFile

  TaxInq(1).Acct = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CustName = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
  TaxInq(1).HPHONE = TaxCustRec(1).HPHONE
  TaxInq(1).CSSN = TaxCustRec(1).CSSN
  TaxInq(1).WPHONE = TaxCustRec(1).WPHONE
  TaxInq(1).Addr1 = TaxCustRec(1).Addr1
  TaxInq(1).Addr2 = TaxCustRec(1).Addr2
  TaxInq(1).CITY = TaxCustRec(1).CITY
  TaxInq(1).State = TaxCustRec(1).State
  TaxInq(1).ZIP = TaxCustRec(1).ZIP
  TaxInq(1).Active = TaxCustRec(1).Active
  TaxInq(1).interest = TaxCustRec(1).interest
  TaxInq(1).EXEMPT = TaxCustRec(1).TaxExempt
  TaxInq(1).Penalty = TaxCustRec(1).Penalty
  TaxInq(1).SRCITDIS = 0
  TaxInq(1).ODISCOUN = 0
  TaxInq(1).PROPVALU = PersValue#
  TaxInq(1).PersVal = MOBHValue#
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
      EDShowCustHistory -CustAcct&, 1'DON'T CHANGE THIS
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


  TaxInq(1).Acct = RecNo&
  TaxInq(1).OPENDATE = TaxCustRec(1).OPENDATE
  TaxInq(1).CustName = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
  TaxInq(1).HPHONE = TaxCustRec(1).HPHONE
  TaxInq(1).CSSN = TaxCustRec(1).CSSN
  TaxInq(1).WPHONE = TaxCustRec(1).WPHONE
  TaxInq(1).Addr1 = TaxCustRec(1).Addr1
  TaxInq(1).Addr2 = TaxCustRec(1).Addr2
  TaxInq(1).CITY = TaxCustRec(1).CITY
  TaxInq(1).State = TaxCustRec(1).State
  TaxInq(1).ZIP = TaxCustRec(1).ZIP
  TaxInq(1).Active = TaxCustRec(1).Active
  TaxInq(1).interest = TaxCustRec(1).interest
  TaxInq(1).EXEMPT = TaxCustRec(1).TaxExempt
  TaxInq(1).Penalty = TaxCustRec(1).Penalty
  TaxInq(1).ODISCOUN = OthrExmp#
  TaxInq(1).PROPVALU = RealValue#
  TaxInq(1).PersVal = BldgValue#

  BCopy VARSEG(TaxInq(1)), VARPTR(TaxInq(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  Action = 1
RETURN
END SUB

SUB EDShowCustHistory (CustRec&, TaxType%)
  
  u$ = CHR$(24)
  d$ = CHR$(25)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  DisplayTaxScrn "LOADHIST"
  
  REDIM Mchoice(1 TO 1) AS FLen2

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  'REDIM TaxCustRec(1) AS TaxCustType
  
  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTran(1))
  
  TaxFile = FREEFILE
  IF TaxType% = 1 THEN
    OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  ELSE
    OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  END IF
  GET TaxFile, ABS(CustRec&), TaxCustRec(1)
  CLOSE TaxFile
  
  'CurBal# = TaxCustRec(1).CurrBalance
  'PreBal# = TaxCustRec(1).PrevBalance
  
Top:
  
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  
  PrevTranRec& = TaxCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET TaxTran, PrevTranRec&, TaxTran(1)
      'IF TaxTran(1).TranType = 1 THEN
        DCnt = DCnt + 1
        REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
        LSET Mchoice(DCnt).V = Num2Date(TaxTran(1).TransDate)
        TType$ = TaxTran(1).Description
        MID$(Mchoice(DCnt).V, 13) = TType$
        MID$(Mchoice(DCnt).V, 41) = FUsing(STR$(TaxTran(1).Amount), "#####.##")
        'this will show the actual trans number in the list
        MID$(Mchoice(DCnt).V, 52) = FUsing(STR$(PrevTranRec&), "######")
        'MID$(MChoice(DCnt).V, 52) = FUsing(STR$(TaxTran(1).RunBalance), "#####.##")
        MID$(Mchoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      'END IF
      PrevTranRec& = TaxTran(1).LastTrans
    LOOP
    
    CLOSE TaxTran
    IF DCnt < 1 THEN
      Ok = MsgBox%("Tax.QSL", "NOCTRANS")
      GOTO NOGood
    END IF
    
    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 74, 8
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "   Date            Description              Amount             "
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    WazzWind BoxBot + 2, Col, BoxBot + 5, MaxLen + 3 + Col, 10, 4, True
    QPrintRC "  Use:  " + u$ + "-" + d$ + " to select.", BoxBot + 3, Col + 3, 15
    QPrintRC u$, BoxBot + 3, Col + 11, 14
    QPrintRC d$, BoxBot + 3, Col + 13, 14
    QPrintRC "Total: " + STR$(DCnt), BoxBot + 4, Col + 3, 15
    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, Col + 33, 15
    QPrintRC "        [ENTER] for detail.", BoxBot + 4, Col + 33, 15
    QPrintRC "ESC", BoxBot + 3, Col + 43, 14
    QPrintRC "ENTER", BoxBot + 4, Col + 42, 14
    
    QPrintRC Title$, Row - 1, Col, 112

    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    
    DO
      LOCATE Row, Col, 0
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TempScrn()
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        RestScrn TempScrn2()
        GOTO ShowTransDetail
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE TaxTran
    Ok = MsgBox%("Tax.QSL", "NOCTRANS")
    RestScrn TempScrn()
  END IF
  
NOGood:
  RestScrn TempScrn()
  ERASE Mchoice
  ERASE TempScrn, TaxTran, TaxCustRec
  
EXIT SUB
  
ShowTransDetail:
  CursorOff
  TransRecNum& = CVL(RIGHT$(Mchoice(Choice).V, 4))
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET TaxTran, TransRecNum&, TaxTran(1)
  'GOSUB GetTransType
  
  LibName$ = "TAXEDTR"
  ScrnName$ = "EDTAXTR1"

  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ShowCursor
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True

  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      GOSUB GetTransType:
      Action = 1
      'FOR Cnt = 9 TO 17
      '  Fld(Cnt).Protected = True
      'NEXT
      QPrintRC Num2Date(TaxTran(1).TransDate), 5, 28, -1
      QPrintRC TaxTran(1).Description, 6, 28, -1
      QPrintRC FUsing$(STR$(TaxTran(1).Amount), "#####.##"), 7, 28, -1
      QPrintRC BillType$, 8, 28, -1
      QPrintRC TaxYear$, 9, 60, -1
      QPrintRC Post2GL$, 9, 28, -1
      QPrintRC "TR#:" + STR$(TransRecNum&), 5, 50, 15
      QPrintRC "BT#:" + STR$(TaxTran(1).BelongTo), 6, 50, 15
      QPrintRC "AC#:" + STR$(TaxTran(1).CustomerRec), 7, 50, 15
      QPrintRC "ED#:" + STR$(TaxTran(1).Altered), 8, 50, 15
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE 27
      EXIT DO
    CASE F10Key
      GOSUB CheckTRInfo
      IF TRIsOk THEN
        GOSUB SaveTR
        DisplayTaxScrn "UPDATEOK"
        WaitForAction
        EXIT DO
      END IF
    END SELECT
  LOOP

  RestScrn TempScrn2()
  GOTO HistTop

RETURN
  
CheckTRInfo:
  REDIM ChkTrScrn(0)
  SaveScrn ChkTrScrn()

  TRIsOk = True
  FOR Cnt = 1 TO 8
    ThisFld$ = QPTrim$(Form$(Cnt, 0))
    IF LEN(ThisFld$) = 0 THEN
      DisplayEDTRScrn "BADTRINF"
      WaitForAction
      TRIsOk = False
      GOTO BadTRExit
    END IF
  NEXT

'This calcs the
  TempTotal# = 0
  FOR Cnt = 1 TO 8
    TempTotal# = Round#(TempTotal# + Value#(Form$(Cnt, 0), ECode))
  NEXT
  IF INSTR(COMMAND$, "TEST") = 0 THEN
    IF TempTotal# <> TotalTRAmt# THEN
      DisplayEDTRScrn "BADTRAMT"
      WaitForAction
      'TRIsOk = False
      GOTO BadTRExit
    END IF
  END IF

BadTRExit:
  RestScrn ChkTrScrn()
RETURN

SaveTR:

  Princ1# = Value#(Form$(1, 0), ECode)
  Princ2# = Value#(Form$(2, 0), ECode)
  Princ3# = Value#(Form$(3, 0), ECode)
  Princ4# = Value#(Form$(4, 0), ECode)
  Princ5# = Value#(Form$(5, 0), ECode)
  TInt1# = Value#(Form$(6, 0), ECode)
  TPen1# = Value#(Form$(7, 0), ECode)
  TCol1# = Value#(Form$(8, 0), ECode)
  Princ1Pd# = Value#(Form$(9, 0), ECode)
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET TaxTran, TransRecNum&, TaxTran(1)
  TaxTran(1).Revenue.Principle1 = Princ1#
  TaxTran(1).Revenue.Principle2 = Princ2#
  TaxTran(1).Revenue.Principle3 = Princ3#
  TaxTran(1).Revenue.Principle4 = Princ4#
  TaxTran(1).Revenue.Principle5 = Princ5#
  TaxTran(1).Revenue.interest = TInt1#

  TaxTran(1).Revenue.Penalty = TPen1#
  TaxTran(1).Revenue.Collection = TCol1#
  TaxTran(1).Revenue.Principle1Pd = Princ1Pd#
  TaxTran(1).Altered = True
  TaxTran(1).Revenue.InterestPD = Value#(Form$(14, 0), ECode)
  TaxTran(1).Revenue.PenaltyPD = Value#(Form$(15, 0), ECode)
  TaxTran(1).Amount = Value#(Form$(17, 0), ECode)

  'TaxTran(1).BelongTo = 2334

  PUT TaxTran, TransRecNum&, TaxTran(1)
  CLOSE TaxTran  'NOTE: Close must be after GetTransType

RETURN

GetTransType:

  Principle# = 0
  interest# = 0
  Penalty# = 0
  AdCost# = 0

  TType$ = TaxTran(1).Description
  BillType$ = ""
  TaxYear$ = "N/A"
  Post2GL$ = "N"
  IF TaxTran(1).Posted2GL = "Y" THEN
    Post2GL$ = "Y"
  END IF

  SELECT CASE TaxTran(1).TranType

  CASE 1
    SELECT CASE TaxTran(1).BillType
    CASE "R"
      BillType$ = "Real-Estate"
    CASE "P"
      BillType$ = "Personal Property"
    CASE "C"
      BillType$ = "Combined"
    CASE "M"
      BillType$ = "Manual"
    END SELECT

    TaxYear$ = QPTrim$(STR$(TaxTran(1).TaxYear))
    TaxYearValue = VAL(TaxYear$)
    IYear = TaxYearValue + 1
    IYear$ = LTRIM$(STR$(IYear))
    GOSUB SetForm
    'IF TaxType% = 1 OR TaxType% = 2 THEN
     'Principle# = Round#(TaxTran(1).Revenue.Principle1)

     'UnPaidPrinciple# = Round#(TaxTran(1).Revenue.Principle1 - TaxTran(1).Revenue.Principle1PD)
'    ELSE
'     Principle# = Round#(TaxTran(1).Revenue.Principle1 + TaxTran(1).Revenue.Principle2 + TaxTran(1).Revenue.Principle3)
'     Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4 + TaxTran(1).Revenue.Principle5)
'     UnPaidPrinciple# = Round#(Principle# - (TaxTran(1).Revenue.Principle1Pd + TaxTran(1).Revenue.Principle2Pd + TaxTran(1).Revenue.Principle3Pd + TaxTran(1).Revenue.Principle4Pd + TaxTran(1).Revenue.Principle5Pd))
    'END IF

    'Fiqure Penalty & Interest Here
     'LSET Form$(6, 0) = STR$(TaxTran(1).Revenue.interest#)
     'LSET Form$(7, 0) = STR$(TaxTran(1).Revenue.Penalty#)
     'Penalty# = TaxTran(1).Revenue.Penalty# - TaxTran(1).Revenue.PenaltyPd
     'Interest# = TaxTran(1).Revenue.Interest# - TaxTran(1).Revenue.InterestPd
 
  CASE 2
    BillType$ = "Payment"
   IF TaxType% = 1 THEN
     Principle# = Round#(TaxTran(1).Revenue.Principle1Pd)
     Penalty# = TaxTran(1).Revenue.PenaltyPD
     interest# = TaxTran(1).Revenue.InterestPD
   ELSE
     Principle# = Round#(TaxTran(1).Revenue.Principle1Pd + TaxTran(1).Revenue.Principle2Pd + TaxTran(1).Revenue.Principle3Pd)
     Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4Pd + TaxTran(1).Revenue.Principle5Pd)
     Penalty# = TaxTran(1).Revenue.PenaltyPD
     interest# = TaxTran(1).Revenue.InterestPD
   END IF
   GOSUB SetForm
  CASE 3
    BillType$ = "Release"
     LSET Form$(1, 0) = STR$(TaxTran(1).Revenue.Principle1)
     LSET Form$(2, 0) = STR$(TaxTran(1).Revenue.Principle2)
     LSET Form$(3, 0) = STR$(TaxTran(1).Revenue.Principle3)
     LSET Form$(4, 0) = STR$(TaxTran(1).Revenue.Principle4)
     LSET Form$(5, 0) = STR$(TaxTran(1).Revenue.Principle5)
     LSET Form$(6, 0) = STR$(TaxTran(1).Revenue.interest)
     LSET Form$(7, 0) = STR$(TaxTran(1).Revenue.Penalty)
     LSET Form$(8, 0) = STR$(TaxTran(1).Revenue.Collection)
     LSET Form$(9, 0) = STR$(TaxTran(1).Revenue.Principle1Pd)
     LSET Form$(10, 0) = STR$(TaxTran(1).Revenue.Principle2Pd)
     LSET Form$(11, 0) = STR$(TaxTran(1).Revenue.Principle3Pd)
     LSET Form$(12, 0) = STR$(TaxTran(1).Revenue.Principle4Pd)
     LSET Form$(13, 0) = STR$(TaxTran(1).Revenue.Principle5Pd)
     LSET Form$(14, 0) = STR$(TaxTran(1).Revenue.InterestPD)
     LSET Form$(15, 0) = STR$(TaxTran(1).Revenue.PenaltyPD)
     LSET Form$(16, 0) = STR$(TaxTran(1).Revenue.CollectionPd)
     LSET Form$(17, 0) = STR$(TaxTran(1).Amount)
     TotalTRAmt# = TaxTran(1).Amount

  CASE 4
    BillType$ = "Interest"
    interest# = TaxTran(1).Revenue.interest#
    IF TaxTran(1).BelongTo > 0 THEN
      GET TaxTran, TaxTran(1).BelongTo, TaxTran(2)
      TaxYear$ = QPTrim$(STR$(TaxTran(2).TaxYear))
    END IF
  CASE 5
    BillType$ = "Penalty"
    Penalty# = TaxTran(1).Revenue.Penalty#
  CASE ELSE
    BillType$ = "?????"
  END SELECT

     LSET Form$(1, 0) = STR$(TaxTran(1).Revenue.Principle1)
     LSET Form$(2, 0) = STR$(TaxTran(1).Revenue.Principle2)
     LSET Form$(3, 0) = STR$(TaxTran(1).Revenue.Principle3)
     LSET Form$(4, 0) = STR$(TaxTran(1).Revenue.Principle4)
     LSET Form$(5, 0) = STR$(TaxTran(1).Revenue.Principle5)
     LSET Form$(6, 0) = STR$(TaxTran(1).Revenue.interest)
     LSET Form$(7, 0) = STR$(TaxTran(1).Revenue.Penalty)
     LSET Form$(8, 0) = STR$(TaxTran(1).Revenue.Collection)
     LSET Form$(9, 0) = STR$(TaxTran(1).Revenue.Principle1Pd)
     LSET Form$(10, 0) = STR$(TaxTran(1).Revenue.Principle2Pd)
     LSET Form$(11, 0) = STR$(TaxTran(1).Revenue.Principle3Pd)
     LSET Form$(12, 0) = STR$(TaxTran(1).Revenue.Principle4Pd)
     LSET Form$(13, 0) = STR$(TaxTran(1).Revenue.Principle5Pd)
     LSET Form$(14, 0) = STR$(TaxTran(1).Revenue.InterestPD)
     LSET Form$(15, 0) = STR$(TaxTran(1).Revenue.PenaltyPD)
     LSET Form$(16, 0) = STR$(TaxTran(1).Revenue.CollectionPd)
     LSET Form$(17, 0) = STR$(TaxTran(1).Amount)

CLOSE TaxTran  'NOTE: Close must be after GetTransType

RETURN

SetForm:
     LSET Form$(1, 0) = STR$(TaxTran(1).Revenue.Principle1)
     LSET Form$(2, 0) = STR$(TaxTran(1).Revenue.Principle2)
     LSET Form$(3, 0) = STR$(TaxTran(1).Revenue.Principle3)
     LSET Form$(4, 0) = STR$(TaxTran(1).Revenue.Principle4)
     LSET Form$(5, 0) = STR$(TaxTran(1).Revenue.Principle5)
     LSET Form$(6, 0) = STR$(TaxTran(1).Revenue.interest)
     LSET Form$(7, 0) = STR$(TaxTran(1).Revenue.Penalty)
     LSET Form$(8, 0) = STR$(TaxTran(1).Revenue.Collection)
     LSET Form$(9, 0) = STR$(TaxTran(1).Revenue.Principle1Pd)
     LSET Form$(10, 0) = STR$(TaxTran(1).Revenue.Principle2Pd)
     LSET Form$(11, 0) = STR$(TaxTran(1).Revenue.Principle3Pd)
     LSET Form$(12, 0) = STR$(TaxTran(1).Revenue.Principle4Pd)
     LSET Form$(13, 0) = STR$(TaxTran(1).Revenue.Principle5Pd)
     LSET Form$(14, 0) = STR$(TaxTran(1).Revenue.InterestPD)
     LSET Form$(15, 0) = STR$(TaxTran(1).Revenue.PenaltyPD)
     LSET Form$(16, 0) = STR$(TaxTran(1).Revenue.CollectionPd)
     LSET Form$(17, 0) = STR$(TaxTran(1).Amount)
     TotalTRAmt# = Round#(TaxTran(1).Amount)


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

