  DEFINT A-Z
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB EnterPayment ()
DECLARE SUB EditPayment ()
DECLARE SUB PrintPayment ()
DECLARE SUB PostPayment ()
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE SUB Search4IntAcct (AcctNum&, RecNo&)
DECLARE SUB SmallPause ()
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB IntLookUp (RecNo&, Text$, CLSFlag%)
DECLARE SUB Search4IntCust (Search$, RecNo&, CLSFlag%)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB UpDateTicklerFile ()
DECLARE SUB PostTransactions ()
DECLARE SUB CalcInterest ()
DECLARE SUB EditTrans ()
DECLARE SUB PrintTrans ()
DECLARE SUB PostTaxBills ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (array%())
DECLARE SUB SaveScrn (array%())
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB MakeCustIndex (IdxType)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayCaScrn (ScrnName$)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION GetTaxCustCnt& ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION Rounds! (SglNum!)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB UpDateTicklerFile ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'CACUST.BI'
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = "Enter/Edit Transactions"
  MChoice$(2) = "Print Payment Edit Report "
  MChoice$(3) = "Post Payments to Accounts"
  MChoice$(4) = "Exit to OS"
  
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
  Row = ((24 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  IF Exist("TAXFROMM.DAT") THEN
    'KillFile "TAXFROMM.DAT"
    MainFlag = True
  END IF
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Assessment 'Payment' Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      EnterPayment
    CASE 2
      PrintPayment
    CASE 3
      PostPayment
    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  
  RUN "CAMENU"
  

SUB EnterPayment
  
  REDIM CaCust(1)  AS CACUSTType
  REDIM CaTrans(1) AS CaTransType
  REDIM PayTrans(1) AS CaPayEditType
  
  CaCustLen = LEN(CaCust(1))    'Length of Cust Record Structure
  PayEditLen = LEN(PayTrans(1))
  
  REDIM Fmt$(3)
  REDIM TempScrn(0)
  REDIM ScrnArray(0)
  
  LibName$ = "TAX"
  ScrnName$ = "CAPAYTR1"
  
  '--define the multi-choice fields
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    CustFormLen = CustFormLen + Fld(F).StorLen
  NEXT
  
  Frm(1).PageDisable = True
  Frm(1).StayOnField = True
  Fld(1).Protected = True
  
  ClearBack
  LookUp RecNo&, "Customer", False, True, False
  
  IF RecNo& <= 0 THEN
    CLOSE
    EXIT SUB
  END IF
  
  '--Set screen number to one and display screen
  Action = 1
  ClearBack
  DisplayCaScrn ScrnName$
  
  'ShowCursor
  
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB CheckPayList
      GOSUB LoadPayInfo
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key 'F10
      GOSUB SavePayment
      PayExitFlag = True
      
    CASE EscKey
      'ExitFlag = True
      PayExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 6 TO 14            '--F1 Help
          PressButton F1Key, 22, 6, 14
        CASE 15 TO 25           '--F4
          PressButton F4KEY, 22, 15, 25
        CASE 27 TO 37           '--F6
          PressButton F6KEY, 22, 27, 37
        CASE 38 TO 51           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 38, 51
          END IF
        CASE 52 TO 61           '--F10 Save Button
          PressButton F10Key, 22, 52, 61
        CASE 62 TO 74           '--ESC cancel button
          PressButton 27, 22, 62, 74
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL PayExitFlag
  
AddEditExit:
  
  EXIT SUB
  
LoadPayInfo:
CheckPayList:
  PayRec& = 0
  PayFile = FREEFILE
  OPEN "CAPAYED.DAT" FOR RANDOM SHARED AS PayFile LEN = PayEditLen
  NumOfPayRecs& = LOF(PayFile) \ PayEditLen
  IF NumOfPayRecs& > 0 THEN
    FOR Cnt = 1 TO NumOfPayRecs&
      GET #PayFile, Cnt, PayTrans(1)
      IF PayTrans(1).CustAcct = RecNo& THEN
        PayRec& = Cnt
        EXIT FOR
      END IF
    NEXT
  END IF
  CLOSE PayFile
  
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CaFile LEN = CaCustLen
  NumOfCRecs& = LOF(CaFile) \ CaCustLen
  GET CaFile, RecNo&, CaCust(1)
  CLOSE CaFile
  
  LSET Form$(1, 0) = QPTrim$(STR$(RecNo&))
  Form$(2, 0) = CaCust(1).CustName
  Form$(5, 0) = QPTrim$(STR$(CaCust(1).RPBStreet))
  Form$(6, 0) = QPTrim$(STR$(CaCust(1).RPBSide))
  Form$(7, 0) = QPTrim$(STR$(CaCust(1).RPBDrive))
  Form$(8, 0) = QPTrim$(STR$(CaCust(1).RPBDemo))
  Form$(9, 0) = QPTrim$(STR$(CaCust(1).RPBWater))
  Form$(10, 0) = QPTrim$(STR$(CaCust(1).RPBSewer))
  Form$(12, 0) = QPTrim$(STR$(CaCust(1).RIBStreet))
  Form$(13, 0) = QPTrim$(STR$(CaCust(1).RIBSide))
  Form$(14, 0) = QPTrim$(STR$(CaCust(1).RIBDrive))
  Form$(15, 0) = QPTrim$(STR$(CaCust(1).RIBDemo))
  Form$(16, 0) = QPTrim$(STR$(CaCust(1).RIBWater))
  Form$(17, 0) = QPTrim$(STR$(CaCust(1).RIBSewer))
  
  IF PayRec& > 0 THEN
    LSET Form$(3, 0) = Num2Date$(PayTrans(1).PayDate)
    LSET Form$(4, 0) = QPTrim$(STR$(PayTrans(1).TotalAmtPd))
    LSET Form$(19, 0) = QPTrim$(STR$(PayTrans(1).PStreetAmt))
    LSET Form$(20, 0) = QPTrim$(STR$(PayTrans(1).PSideAmt))
    LSET Form$(21, 0) = QPTrim$(STR$(PayTrans(1).PDriveAmt))
    LSET Form$(22, 0) = QPTrim$(STR$(PayTrans(1).PDemoAmt))
    LSET Form$(23, 0) = QPTrim$(STR$(PayTrans(1).PWaterAmt))
    LSET Form$(24, 0) = QPTrim$(STR$(PayTrans(1).PSewerAmt))
    LSET Form$(26, 0) = QPTrim$(STR$(PayTrans(1).IStreetAmt))
    LSET Form$(27, 0) = QPTrim$(STR$(PayTrans(1).ISideAmt))
    LSET Form$(28, 0) = QPTrim$(STR$(PayTrans(1).IDriveAmt))
    LSET Form$(29, 0) = QPTrim$(STR$(PayTrans(1).IDemoAmt))
    LSET Form$(30, 0) = QPTrim$(STR$(PayTrans(1).IWaterAmt))
    LSET Form$(31, 0) = QPTrim$(STR$(PayTrans(1).ISewerAmt))
    LSET Form$(33, 0) = PayTrans(1).payDesc
  ELSE
    LSET Form$(3, 0) = DATE$
    FOR ZCnt = 19 TO 24
      LSET Form$(ZCnt, 0) = "0"
    NEXT
    FOR ZCnt = 26 TO 31
      LSET Form$(ZCnt, 0) = "0"
    NEXT
    
  END IF
  
  RETURN
  
  '  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CaFile LEN = CaCustLen
  '  NumOfCRecs& = LOF(CaFile) \ CaCustLen
  
  
  'LoadPayInfo:
  
  'GOSUB CheckLoadPayInfo
  
ShowInfoErr1:
  SaveScrn TempScrn()
  DisplayCaScrn "ERRSCRN1"
  QPrintRC "Invalid Information!", 10, 26, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn TempScrn()
  RETURN
  
SavePayment:
  
  PayEditLen = LEN(PayTrans(1))
  PayFile = FREEFILE
  OPEN "CAPAYED.DAT" FOR RANDOM SHARED AS PayFile LEN = PayEditLen
  NumOfPayRecs& = LOF(PayFile) \ PayEditLen
  
  SaveScrn ScrnArray()
  DisplayCaScrn "UPDATDSK"
  
  'Total Amount Paid Calc Here
  TotalPaid# = Value#(Form$(19, 0), ecode) + Value#(Form$(20, 0), ecode) + Value#(Form$(21, 0), ecode)
  TotalPaid# = TotalPaid# + (Value#(Form$(22, 0), ecode) + Value#(Form$(23, 0), ecode) + Value#(Form$(24, 0), ecode))
  TotalPaid# = TotalPaid# + (Value#(Form$(26, 0), ecode) + Value#(Form$(27, 0), ecode) + Value#(Form$(28, 0), ecode))
  TotalPaid# = TotalPaid# + (Value#(Form$(29, 0), ecode) + Value#(Form$(30, 0), ecode) + Value#(Form$(31, 0), ecode))
  TotalPaid# = Round#(TotalPaid#)
  
  PayTrans(1).PayDate = Date2Num(Form$(3, 0))
  PayTrans(1).CustAcct = RecNo&
  PayTrans(1).PStreetAmt = Value#(Form$(19, 0), ecode)
  PayTrans(1).PSideAmt = Value#(Form$(20, 0), ecode)
  PayTrans(1).PDriveAmt = Value#(Form$(21, 0), ecode)
  PayTrans(1).PDemoAmt = Value#(Form$(22, 0), ecode)
  PayTrans(1).PWaterAmt = Value#(Form$(23, 0), ecode)
  PayTrans(1).PSewerAmt = Value#(Form$(24, 0), ecode)
  PayTrans(1).IStreetAmt = Value#(Form$(26, 0), ecode)
  PayTrans(1).ISideAmt = Value#(Form$(27, 0), ecode)
  PayTrans(1).IDriveAmt = Value#(Form$(28, 0), ecode)
  PayTrans(1).IDemoAmt = Value#(Form$(29, 0), ecode)
  PayTrans(1).IWaterAmt = Value#(Form$(30, 0), ecode)
  PayTrans(1).ISewerAmt = Value#(Form$(31, 0), ecode)
  PayTrans(1).TotalAmtPd = TotalPaid#
  PayTrans(1).DelFlag = 0
  PayTrans(1).payDesc = QPTrim$(Form$(33, 0))
  
  IF PayRec& > 0 THEN
    PUT PayFile, PayRec&, PayTrans(1)
  ELSE
    PUT PayFile, NumOfPayRecs& + 1, PayTrans(1)
  END IF
  
  RestScrn ScrnArray()
  DisplayCaScrn "UPDATEOK"
  WaitForAction
  RestScrn ScrnArray()
  Action = 1
  EditedFlag = False
  RETURN
  
  
END SUB

SUB IntLookUp (RecNo&, Text$, CLSFlag%)
  
LookUpTop:
  
  REDIM Hlp$(1 TO 2)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Name here."
  
  SName$ = ""
  AcctNum& = 0
  LScrn = 2
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 2) AS STRING * 17
  
  MScrn = 2
  
  LText(1) = " Account Number:"
  LText(2) = "    Search Name:"
  
  LibName$ = "TAX"
  ScrnName$ = "LUPACCT"
  
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
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  Text$ = Text$ + " Look-Up"
  TextLen = LEN(Text$)
  TCol = ((80 - TextLen) \ 2)
  DisplayCaScrn ScrnName$
  
  QPrintRC Text$, 8, TCol, -1
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      SELECT CASE LScrn
      CASE 1, 2
        LSET Form$(1, 0) = ""
        Fld(1).FType = 2
      END SELECT
      Form$(1, 1) = Hlp$(LScrn)
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13                'F10Key    Proceed with look up
      CursorOff
      SELECT CASE LScrn
      CASE 1    'account lookup
        AcctNum& = QPValL(Form$(1, 0))
        IF AcctNum& < 1 THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")
        ELSE
          Search4IntAcct AcctNum&, RecNo&
          IF RecNo& = 0 THEN
            Ok = MsgBox%("TAX.QSL", "BADACCTN")
          ELSE
            OKFlag = True
          END IF
        END IF
        Action = 1
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4IntCust SName$, RecNo&, CLSFlag
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      END SELECT
    CASE -65    'F7Key
      IF LScrn < MScrn THEN
        LScrn = LScrn + 1
      ELSE
        LScrn = 1
      END IF
      LSET Form$(1, 0) = ""
      Action = 1
      FirstTime = True
      SaveField 0, Form$(), Fld(), BadField
      GOSUB DisplayLookupText
    CASE 27
      RecNo& = 0
      ExitFlag = True
    END SELECT
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16
        SELECT CASE Frm(1).MCol
        CASE 22 TO 33           'ESC Cancel button
          PressButton 27, 16, 22, 33
        CASE 35 TO 45           'F7 Toggle Choice
          PressButton -65, 16, 35, 45
        CASE 47 TO 59           'F10 Save Button
          PressButton -68, 16, 47, 59
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  RestScrn ScrnArray()
  
  ERASE ScrnArray, ScrnArray2
  ERASE Frm, Form$, Fld, LText, Hlp$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText(LScrn), 12, 15, -1
  RETURN
  
  
END SUB

SUB PostPayment
  
  REDIM TempScrn(0)
  
  REDIM CaCust(1)  AS CACUSTType
  REDIM CaTrans(1) AS CaTransType
  REDIM CaPayRec(1) AS CaPayEditType
  
  PayEditLen = LEN(CaPayRec(1))
PostBeginHere:
  ClearBack
  ShowProcessingScrn "Verifying Payment Transactions"
  CLOSE
  PayFile = FREEFILE
  OPEN "CAPAYED.DAT" FOR RANDOM AS PayFile LEN = PayEditLen
  NumOfRec& = LOF(PayFile) / PayEditLen
  
  CustFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM AS CustFile LEN = LEN(CaCust(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #PayFile, Cnt&, CaPayRec(1)
    IF CaPayRec(1).DelFlag = 0 THEN
      DidSome = DidSome + 1
      GET #CustFile, CaPayRec(1).CustAcct, CaCust(1)
      IF Rounds!(CaCust(1).RPBStreet) - Rounds!(CaPayRec(1).PStreetAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RPBSide) - Rounds!(CaPayRec(1).PSideAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RPBDrive) - Rounds!(CaPayRec(1).PDriveAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RPBDemo) - Rounds!(CaPayRec(1).PDemoAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RPBWater) - Rounds!(CaPayRec(1).PWaterAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RPBSewer) - Rounds!(CaPayRec(1).PSewerAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RIBStreet) - Rounds!(CaPayRec(1).IStreetAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RIBSide) - Rounds!(CaPayRec(1).ISideAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RIBDrive) - Rounds!(CaPayRec(1).IDriveAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RIBDemo) - Rounds!(CaPayRec(1).IDemoAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RIBWater) - Rounds!(CaPayRec(1).IWaterAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
      IF Rounds!(CaCust(1).RIBSewer) - Rounds!(CaPayRec(1).ISewerAmt) < 0 THEN
        Acct$ = STR$(CaPayRec(1).CustAcct)
        NegBal = 1
      END IF
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  
  
  CLOSE
  SmallPause
  
  
  IF DidSome < 1 THEN
    GOSUB PayShowPostError1
    GOTO PayPostExit
  END IF
  IF NegBal = 1 THEN
    GOSUB PayShowPostError2
    GOTO PayPostExit
    
  END IF
  
  
  
  SELECT CASE PostFlag
  CASE 1
    GOTO PostPay
  CASE 2
    CursorOff
    ClearBack
    DisplayCaScrn "POSTCAN"
    WaitForAction
    GOTO PayPostExit
  END SELECT
  
PostPay:
  ClearBack
  ShowProcessingScrn "Posting Tax Interest Transactions"
  
  'Open All Needed Files Here
  
  PayFile = FREEFILE
  OPEN "CAPAYED.DAT" FOR RANDOM AS PayFile LEN = PayEditLen
  NumOfRec& = LOF(PayFile) / PayEditLen
  
  CustFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM AS CustFile LEN = LEN(CaCust(1))
  
  TransFile = FREEFILE
  OPEN "CATRANS.DAT" FOR RANDOM AS TransFile LEN = LEN(CaTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #PayFile, Cnt&, CaPayRec(1)
    
    IF CaPayRec(1).DelFlag = 0 AND CaPayRec(1).TotalAmtPd > 0 THEN
      REDIM CaTrans(1) AS CaTransType
      CaTrans(1).TranType = 2
      CaTrans(1).Date = Date2Num%(DATE$)
      CaTrans(1).IStreetAmt = CaPayRec(1).IStreetAmt
      CaTrans(1).ISideAmt = CaPayRec(1).ISideAmt
      CaTrans(1).IDriveAmt = CaPayRec(1).IDriveAmt
      CaTrans(1).IDemoAmt = CaPayRec(1).IDemoAmt
      CaTrans(1).IWaterAmt = CaPayRec(1).IWaterAmt
      CaTrans(1).ISewerAmt = CaPayRec(1).ISewerAmt
      CaTrans(1).PStreetAmt = CaPayRec(1).PStreetAmt
      CaTrans(1).PSideAmt = CaPayRec(1).PSideAmt
      CaTrans(1).PDriveAmt = CaPayRec(1).PDriveAmt
      CaTrans(1).PDemoAmt = CaPayRec(1).PDemoAmt
      CaTrans(1).PWaterAmt = CaPayRec(1).PWaterAmt
      CaTrans(1).PSewerAmt = CaPayRec(1).PSewerAmt
      CaTrans(1).NextTrans = 0
      CaTrans(1).CustRecord = CaPayRec(1).CustAcct
      CaTrans(1).Pad = ""
      
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(CaTrans(1))) + 1
      PUT #TransFile, NextRecord&, CaTrans(1)
      
      'Update the Customer Pointers and Balances Now
      GET #CustFile, CaPayRec(1).CustAcct, CaCust(1)
      CaCust(1).RIBStreet = CaCust(1).RIBStreet - CaPayRec(1).IStreetAmt
      CaCust(1).RIBSide = CaCust(1).RIBSide - CaPayRec(1).ISideAmt
      CaCust(1).RIBDrive = CaCust(1).RIBDrive - CaPayRec(1).IDriveAmt
      CaCust(1).RIBDemo = CaCust(1).RIBDemo - CaPayRec(1).IDemoAmt
      CaCust(1).RIBWater = CaCust(1).RIBWater - CaPayRec(1).IWaterAmt
      CaCust(1).RIBSewer = CaCust(1).RIBSewer - CaPayRec(1).ISewerAmt
      CaCust(1).RPBStreet = CaCust(1).RPBStreet - CaPayRec(1).PStreetAmt
      CaCust(1).RPBSide = CaCust(1).RPBSide - CaPayRec(1).PSideAmt
      CaCust(1).RPBDrive = CaCust(1).RPBDrive - CaPayRec(1).PDriveAmt
      CaCust(1).RPBDemo = CaCust(1).RPBDemo - CaPayRec(1).PDemoAmt
      CaCust(1).RPBWater = CaCust(1).RPBWater - CaPayRec(1).PWaterAmt
      CaCust(1).RPBSewer = CaCust(1).RPBSewer - CaPayRec(1).PSewerAmt
      CaCust(1).LastPayDate = CaPayRec(1).PayDate
      
      IF CaCust(1).LastTrans = 0 THEN
        CaCust(1).LastTrans = NextRecord&
        PUT #CustFile, CaPayRec(1).CustAcct, CaCust(1)
      ELSE
        Previous& = CaCust(1).LastTrans
        CaCust(1).LastTrans = NextRecord&
        PUT #CustFile, CaPayRec(1).CustAcct, CaCust(1)
        GET #TransFile, NextRecord&, CaTrans(1)
        CaTrans(1).NextTrans = Previous&
        PUT #TransFile, NextRecord&, CaTrans(1)
      END IF
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  
  UpDateTicklerFile
  
  ClearBack
  DisplayCaScrn "UPDATEOK"
  WaitForAction
  
  'Now Delete the Pay File so Duplicate's Cannot Be Reproduced
  KillFile "CAPAYED.DAT"
  
PayPostExit:
  EXIT SUB
  
  
PayShowPostError1:
  ClearBack
  DisplayCaScrn "ERRSCRN1"
  QPrintRC "No Active Payment Transactions!", 10, 25, -1
  QPrintRC "Press any key to continue.", 12, 28, -1
  WaitForAction
  RETURN
PayShowPostError2:
  ClearBack
  DisplayCaScrn "ERRSCRN1"
  DJ$ = "Overpayment Noted on Acct " + Acct$
  QPrintRC DJ$, 10, 25, -1
  QPrintRC "Press any key to continue.", 12, 28, -1
  WaitForAction
  RETURN
  
  
END SUB

SUB PostTransactions
  
  REDIM TempScrn(0)
  
  REDIM CaCust(1)  AS CACUSTType
  REDIM CaTrans(1) AS CaTransType
  REDIM CaPayRec(1 TO 1) AS CaIntEditType
  CaPayRecLen = LEN(CaPayRec(1))
  IntRecLen = LEN(CaPayRec(1))
BeginHere:
  ClearBack
  ShowProcessingScrn "Verifying Interest Transactions"
  
  IntFile = FREEFILE
  OPEN "CAPAY.DAT" FOR RANDOM AS IntFile LEN = IntRecLen
  NumOfRec& = LOF(IntFile) / IntRecLen
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #IntFile, Cnt&, CaPayRec(1)
    IF CaPayRec(1).DelFlag = 0 THEN
      DidSome = DidSome + 1
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  SmallPause
  
  LOCATE 15, 22
  PRINT "Enter Posting Date (mmddyy):        ";
  LOCATE 15, 49: PRINT "";
  COLOR 15
  INPUT PostDate$
  IF LEN(PostDate$) <> 6 THEN
    GOTO BeginHere
  END IF
  PostDate$ = LEFT$(PostDate$, 2) + "-" + MID$(PostDate$, 3, 2) + "-20" + RIGHT$(PostDate$, 2)
  
  
  IF DidSome < 1 THEN
    GOSUB ShowPostError
    GOTO PostExit
  END IF
  
  LOCATE 19, 24: COLOR 14: PRINT "Posting Date Set to "; PostDate$
  PostFlag = MsgBox%("TAX", "POSTINTR")
  
  
  SELECT CASE PostFlag
  CASE 1
    GOTO Postem
  CASE 2
    CursorOff
    ClearBack
    DisplayCaScrn "POSTCAN"
    WaitForAction
    GOTO PostExit
  END SELECT
  
Postem:
  ClearBack
  ShowProcessingScrn "Posting Tax Interest Transactions"
  
  'Open All Needed Files Here
  
  IntFile = FREEFILE
  OPEN "CAPAY.DAT" FOR RANDOM AS IntFile LEN = IntRecLen
  NumOfRec& = LOF(IntFile) / IntRecLen
  
  CustFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM AS CustFile LEN = LEN(CaCust(1))
  
  TransFile = FREEFILE
  OPEN "CATRANS.DAT" FOR RANDOM AS TransFile LEN = LEN(CaTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #IntFile, Cnt&, CaPayRec(1)
    IF CaPayRec(1).DelFlag = 0 THEN
      REDIM CaTrans(1) AS CaTransType
      CaTrans(1).TranType = 1
      CaTrans(1).Date = Date2Num%(DATE$)
      CaTrans(1).IStreetAmt = CaPayRec(1).IStreetAmt
      CaTrans(1).ISideAmt = CaPayRec(1).ISideAmt
      CaTrans(1).IDriveAmt = CaPayRec(1).IDriveAmt
      CaTrans(1).IDemoAmt = CaPayRec(1).IDemoAmt
      CaTrans(1).IWaterAmt = CaPayRec(1).IWaterAmt
      CaTrans(1).ISewerAmt = CaPayRec(1).ISewerAmt
      CaTrans(1).PStreetAmt = 0
      CaTrans(1).PSideAmt = 0
      CaTrans(1).PDriveAmt = 0
      CaTrans(1).PDemoAmt = 0
      CaTrans(1).PWaterAmt = 0
      CaTrans(1).PSewerAmt = 0
      CaTrans(1).NextTrans = 0
      CaTrans(1).CustRecord = CaPayRec(1).CustRecord
      CaTrans(1).Pad = ""
      
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(CaTrans(1))) + 1
      PUT #TransFile, NextRecord&, CaTrans(1)
      
      'Update the Customer Pointers and Balances Now
      GET #CustFile, CaPayRec(1).CustRecord, CaCust(1)
      CaCust(1).IStreet = CaCust(1).IStreet + CaPayRec(1).IStreetAmt
      CaCust(1).ISide = CaCust(1).ISide + CaPayRec(1).ISideAmt
      CaCust(1).IDrive = CaCust(1).IDrive + CaPayRec(1).IDriveAmt
      CaCust(1).IDemo = CaCust(1).IDemo + CaPayRec(1).IDemoAmt
      CaCust(1).IWater = CaCust(1).IWater + CaPayRec(1).IWaterAmt
      CaCust(1).ISewer = CaCust(1).ISewer + CaPayRec(1).ISewerAmt
      
      CaCust(1).RIBStreet = CaCust(1).RIBStreet + CaPayRec(1).IStreetAmt
      CaCust(1).RIBSide = CaCust(1).RIBSide + CaPayRec(1).ISideAmt
      CaCust(1).RIBDrive = CaCust(1).RIBDrive + CaPayRec(1).IDriveAmt
      CaCust(1).RIBDemo = CaCust(1).RIBDemo + CaPayRec(1).IDemoAmt
      CaCust(1).RIBWater = CaCust(1).RIBWater + CaPayRec(1).IWaterAmt
      CaCust(1).RIBSewer = CaCust(1).RIBSewer + CaPayRec(1).ISewerAmt
      CaCust(1).LastIntDate = Date2Num%(PostDate$)
      
      IF CaCust(1).LastTrans = 0 THEN
        CaCust(1).LastTrans = NextRecord&
        PUT #CustFile, CaPayRec(1).CustRecord, CaCust(1)
      ELSE
        Previous& = CaCust(1).LastTrans
        CaCust(1).LastTrans = NextRecord&
        PUT #CustFile, CaPayRec(1).CustRecord, CaCust(1)
        GET #TransFile, NextRecord&, CaTrans(1)
        CaTrans(1).NextTrans = Previous&
        PUT #TransFile, NextRecord&, CaTrans(1)
      END IF
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  
  UpDateTicklerFile
  
  ClearBack
  DisplayCaScrn "UPDATEOK"
  WaitForAction
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
  KillFile "CAPAY.DAT"
  
PostExit:
  EXIT SUB
  
  
ShowPostError:
  ClearBack
  DisplayCaScrn "ERRSCRN1"
  QPrintRC "No Active Interest Transactions!", 10, 25, -1
  QPrintRC "Press any key to continue.", 12, 28, -1
  WaitForAction
  RETURN
END SUB

SUB PrintPayment
  
  REDIM CaCust(1)  AS CACUSTType
  REDIM CaTrans(1) AS CaTransType
  REDIM PayTrans(1) AS CaPayEditType
  
  SHARED Choice$()
  
  ReportFile$ = "CaPayEdt.PRN"  'Report File Name
  
  Dash80$ = STRING$(80, "=")
  
  size = 5500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  
  REDIM array(1 TO size) AS Struct
  
  FF$ = CHR$(12)
  
  MaxLines = 60
  LineCnt = 0
  CustCnt = 0
  RptHandle = FREEFILE
  
  DevSpec$ = "S"
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintHeader
  
  ' Print Main Body
  
  
  GOSUB GetNameIndex1
  
  FOR Cnt = 1 TO Count
    
    CustRecNo = array(Cnt).RecNum
    
    GET PayFile, CustRecNo, PayTrans(1)
    GET CaFile, PayTrans(1).CustAcct, CaCust(1)
    
    help$ = "Processing Record # " + STR$(Cnt!)
    'Main Print Processing Here
    
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintHeader
    END IF
    
    
    PRINT #RptHandle, PayTrans(1).CustAcct;
    PRINT #RptHandle, TAB(10); LEFT$(RTRIM$(CaCust(1).CustName), 30);
    
    IF PayTrans(1).TotalAmtPd <= 0 THEN
      PRINT #RptHandle, " ** ENTRY WILL NOT POST **"
    ELSE
      PRINT #RptHandle, TAB(42); Num2Date$(PayTrans(1).PayDate);
      PRINT #RptHandle, TAB(54); RTRIM$(PayTrans(1).payDesc);
      PRINT #RptHandle, TAB(69); USING "$$#####,#.##"; PayTrans(1).TotalAmtPd
      PRINT #RptHandle, "ASSESSMENT TYPE"; TAB(23); "Principle Paid"; TAB(44); "Interest Paid"; TAB(62); "Total Paid"
    END IF
    LineCnt = LineCnt + 1
    
    'Streets
    IF PayTrans(1).PStreetAmt + PayTrans(1).IStreetAmt > 0 THEN
      PRINT #RptHandle, "Streets"; TAB(25); USING "$$#####,#.##"; PayTrans(1).PStreetAmt;
      PRINT #RptHandle, TAB(46); USING "$$####,#.##"; PayTrans(1).IStreetAmt;
      PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; PayTrans(1).PStreetAmt + PayTrans(1).IStreetAmt
      TPStreetAmt# = TPStreetAmt# + PayTrans(1).PStreetAmt
      TIStreetAmt# = TIStreetAmt# + PayTrans(1).IStreetAmt
      LineCnt = LineCnt + 1
    END IF
    'SideWalks
    IF PayTrans(1).PSideAmt + PayTrans(1).ISideAmt > 0 THEN
      PRINT #RptHandle, "Sidewalks"; TAB(25); USING "$$#####,#.##"; PayTrans(1).PSideAmt;
      PRINT #RptHandle, TAB(46); USING "$$####,#.##"; PayTrans(1).ISideAmt;
      PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; PayTrans(1).PSideAmt + PayTrans(1).ISideAmt
      TPSideAmt# = TPSideAmt# + PayTrans(1).PSideAmt
      TISideAmt# = TISideAmt# + PayTrans(1).ISideAmt
      LineCnt = LineCnt + 1
    END IF
    
    'DriveWay
    IF PayTrans(1).PDriveAmt + PayTrans(1).IDriveAmt > 0 THEN
      PRINT #RptHandle, "Driveways"; TAB(25); USING "$$#####,#.##"; PayTrans(1).PDriveAmt;
      PRINT #RptHandle, TAB(46); USING "$$####,#.##"; PayTrans(1).IDriveAmt;
      PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; PayTrans(1).PDriveAmt + PayTrans(1).IDriveAmt
      TPDriveAmt# = TPDriveAmt# + PayTrans(1).PDriveAmt
      TIDriveAmt# = TIDriveAmt# + PayTrans(1).IDriveAmt
      LineCnt = LineCnt + 1
    END IF
    
    'Demolition
    IF PayTrans(1).PDemoAmt + PayTrans(1).IDemoAmt > 0 THEN
      PRINT #RptHandle, "Dredging"; TAB(25); USING "$$#####,#.##"; PayTrans(1).PDemoAmt;
      PRINT #RptHandle, TAB(46); USING "$$####,#.##"; PayTrans(1).IDemoAmt;
      PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; PayTrans(1).PDemoAmt + PayTrans(1).IDemoAmt
      TPDemoAmt# = TPDemoAmt# + PayTrans(1).PDemoAmt
      TIDemoAmt# = TIDemoAmt# + PayTrans(1).IDemoAmt
      LineCnt = LineCnt + 1
    END IF
    'Water
    IF PayTrans(1).PWaterAmt + PayTrans(1).IWaterAmt > 0 THEN
      PRINT #RptHandle, "Underground Utilities"; TAB(25); USING "$$#####,#.##"; PayTrans(1).PWaterAmt;
      PRINT #RptHandle, TAB(46); USING "$$####,#.##"; PayTrans(1).IWaterAmt;
      PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; PayTrans(1).PWaterAmt + PayTrans(1).IWaterAmt
      TPWaterAmt# = TPWaterAmt# + PayTrans(1).PWaterAmt
      TIWaterAmt# = TIWaterAmt# + PayTrans(1).IWaterAmt
      LineCnt = LineCnt + 1
    END IF
    
    'Sewer
    IF PayTrans(1).PSewerAmt + PayTrans(1).ISewerAmt > 0 THEN
      PRINT #RptHandle, "Sewer"; TAB(25); USING "$$#####,#.##"; PayTrans(1).PSewerAmt;
      PRINT #RptHandle, TAB(46); USING "$$####,#.##"; PayTrans(1).ISewerAmt;
      PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; PayTrans(1).PSewerAmt + PayTrans(1).ISewerAmt
      TPSewerAmt# = TPSewerAmt# + PayTrans(1).PSewerAmt
      TISewerAmt# = TISewerAmt# + PayTrans(1).ISewerAmt
      LineCnt = LineCnt + 1
    END IF
    PRINT #RptHandle, STRING$(80, "-")
    LineCnt = LineCnt + 1
    CustCnt = CustCnt + 1
    GTotal# = GTotal# + PayTrans(1).TotalAmtPd

CUSTSKIP:
  NEXT Cnt
  
  GOSUB PrintBalanceCustomerRptEnding
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
    '**
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
    '**
  ELSE
    EntryPoint = 1
  END IF
  
  'HideMouse

  HideCursor

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(30); "Assessment Payment Edit Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Acct #"; TAB(10); "Customer Name"; TAB(43); "Pay Date"; TAB(54); "Pay Desc"; TAB(70); "Total Paid"
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN
  
PrintBalanceCustomerRptEnding:
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Totals Page"
  PRINT #RptHandle, "Date: "; DATE$
  PRINT #RptHandle,
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, "       Grand Total Paid: "; USING "$$######,#.##"; GTotal#
  PRINT #RptHandle,
  PRINT #RptHandle, "Assessment Totals Breakdown"
  PRINT #RptHandle, "Type"; TAB(25); "Principle Pd"; TAB(43); "Interest Pd"; TAB(62); "Total Paid"
  PRINT #RptHandle, "Streets "; TAB(25); USING "$$#####,#.##"; TPStreetAmt#;
  PRINT #RptHandle, TAB(42); USING "$$#####,#.##"; TIStreetAmt#;
  PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; (TPStreetAmt# + TIStreetAmt#)
  PRINT #RptHandle, "Sidewalks "; TAB(25); USING "$$#####,#.##"; TPSideAmt#;
  PRINT #RptHandle, TAB(42); USING "$$#####,#.##"; TISideAmt#;
  PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; (TPSideAmt# + TISideAmt#)
  PRINT #RptHandle, "Driveways "; TAB(25); USING "$$#####,#.##"; TPDriveAmt#;
  PRINT #RptHandle, TAB(42); USING "$$#####,#.##"; TIDriveAmt#;
  PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; (TPDriveAmt# + TIDriveAmt#)
  PRINT #RptHandle, "Demolition "; TAB(25); USING "$$#####,#.##"; TPDemoAmt#;
  PRINT #RptHandle, TAB(42); USING "$$#####,#.##"; TIDemoAmt#;
  PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; (TPDemoAmt# + TIDemoAmt#)
  PRINT #RptHandle, "Utility "; TAB(25); USING "$$#####,#.##"; TPWaterAmt#;
  PRINT #RptHandle, TAB(42); USING "$$#####,#.##"; TIWaterAmt#;
  PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; (TPWaterAmt# + TIWaterAmt#)
  PRINT #RptHandle, "Dredging "; TAB(25); USING "$$#####,#.##"; TPSewerAmt#;
  PRINT #RptHandle, TAB(42); USING "$$#####,#.##"; TISewerAmt#;
  PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; (TPSewerAmt# + TISewerAmt#)
  PRINT #RptHandle, STRING$(72, "-")
  PRINT #RptHandle, "Totals"; TAB(25); USING "$$#####,#.##"; (TPStreetAmt# + TPSideAmt# + TPDriveAmt# + TPDemoAmt# + TPWaterAmt# + TPSewerAmt#);
  PRINT #RptHandle, TAB(42); USING "$$#####,#.##"; (TIStreetAmt# + TISideAmt# + TIDriveAmt# + TIDemoAmt# + TIWaterAmt# + TISewerAmt#);
  PRINT #RptHandle, TAB(60); USING "$$#####,#.##"; GTotal#
  
  PRINT #RptHandle, FF$
  RETURN
  
SelectBalanceCustomerOutput:
  LibName$ = "TAX"
  ScrnName$ = "CUSTRPT"
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  
  Action = 1
  ShowCursor
  ClearBack
  DisplayCaScrn ScrnName$
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  LOCATE 8, 33: COLOR 14: PRINT "Balances";
  'printhelp help$
  Form$(2, 0) = "Summary Only"
  Fld(2).Protected = True
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      
      IF LEFT$(Form$(1, 0), 4) = "Name" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      
      Order$ = LEFT$(Form$(1, 0), 1)
      
      IF LEFT$(Form$(2, 0), 1) = "S" THEN
        DetailFlag = False
      ELSE
        DetailFlag = True
      END IF
      Detail$ = Form$(2, 0)
      
      
      ExitFlag = True
      RETURN
    CASE EscKey
      AbortFlag = True
      'Canceled$ = "Y"
      ExitFlag = True           'EXIT DO
      'RETURN
    END SELECT
  LOOP UNTIL ExitFlag
  
  RETURN
  
GetNameIndex1:
  
  
  PayEditLen = LEN(PayTrans(1))
  
  PayFile = FREEFILE
  OPEN "CAPAYED.DAT" FOR RANDOM SHARED AS PayFile LEN = PayEditLen
  NumOfPayRecs& = LOF(PayFile) \ PayEditLen
  
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CaFile LEN = LEN(CaCust(1))
  
  FOR SCnt = 1 TO NumOfPayRecs&
    GET PayFile, SCnt, PayTrans(1)
    
    IF PayTrans(1).CustAcct > 0 THEN
      GET CaFile, PayTrans(1).CustAcct, CaCust(1)
      array(SCnt).who = UCASE$(CaCust(1).CustName)
      array(SCnt).RecNum = SCnt
      Count = Count + 1
    END IF
  NEXT SCnt
  'Sort Them Here
  SortT array(Start), Count, Dir, SSize, MOff, MSize
  RETURN
  
END SUB

SUB Search4IntAcct (AcctNum&, RecNo&)
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
  REDIM ITaxCust(1 TO 1) AS CaIntEditType
  
  ' ICaCustLen = LEN(CaCust(1))
  
  RecNo& = 0
  ' FOpenS CAINT.DAT, CaFile    'open data file
  ' NumOfCust& = FLof&(CaFile) / ICaCustLen
  
  FOR CCnt& = 1 TO NumOfCust&
    '  FGetRTA CaFile, ICaCust(1), CCnt&, ICaCustLen
    'IF ICaCust(1).CustRec = AcctNum& THEN
    '    RecNo& = CCnt&
    EXIT FOR
    ' END IF
    ShowSearchPCT CCnt&, NumOfCust&
  NEXT
  
  FClose CaFile
  RestScrn ScrnArray()
  
  ERASE ScrnArray, ITaxCust
  
END SUB

SUB Search4IntCust (Search$, RecNo&, CLSFlag%)
  
  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(5)
  
  BNum$ = SPACE$(7)
  
  IF LastSEARCH$ <> Search$ THEN
    LastSEARCH$ = Search$
    Choice = 1
  END IF
  
  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
  '091598 Found a bug that caused the lookup to get erroneous records
  '       if this was changed to greater than 32???
  CustBlock = 1
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM ITaxCust(1 TO 1) AS CaIntEditType
  
  'ICaCustLen = LEN(ICaCust(1))
  
  'SearchLen = LEN(Search$)
  'Match = False
  
  ' FOpenS CAINT.DAT, CaFile    'open data file
  ' NumOfCust& = FLof&(CaFile) / ICaCustLen
  
  FOR CCnt& = 1 TO NumOfCust&
    '  FGetRTA CaFile, ICaCust(1), CCnt&, ICaCustLen
    'IF INSTR(ICaCust(1).CustName, Search$) > 0 AND ICaCust(1).DelFlag = 0 THEN
    '  GOSUB CustLoadEM2
    ' END IF
DelSkip3:
    ShowSearchPCT CCnt&, NumOfCust&
  NEXT
  
  FClose CaFile
  
  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2
  ELSE
    
    'SortT MChoice(1), DCnt, Direction%, LEN(MChoice(1).V), 0, 18
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer Name                  Year      Amt    BillNo  Acct"
    '--Find max menu width
    '--Center Menu within Screen
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      ClearBack
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
LoopRestart:
    LOCATE Row, Col, 0
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  END IF
  
ExitSearch2:
  RestScrn ScrnArray()
  
  ERASE ScrnArray, MChoice, ITaxCust
  
  EXIT SUB
  
CustLoadEM2:
  
  DCnt = DCnt + 1
  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
  'LSET MChoice(DCnt).V = LEFT$(QPTrim$(ICaCust(1).CustName), 30)
  
  'MID$(MChoice(DCnt).V, 32) = FUsing$(STR$(ICaCust(1).TAXYEAR), "####")
  'MID$(MChoice(DCnt).V, 38) = FUsing$(STR$(ICaCust(1).Amount), ",####.##")
  'RSET BNum$ = QPTrim$(ICaCust(1).BillNumber)
  'MID$(MChoice(DCnt).V, 48) = BNum$
  'RSET ACCT$ = QPTrim$(STR$(ICaCust(1).CustRec))
  ' MID$(MChoice(DCnt).V, 55, 5) = ACCT$
  'MID$(MChoice(DCnt).V, 61) = MKL$(CCnt&)
  RETURN
  
  
END SUB

