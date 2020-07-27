DEFINT A-Z
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
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB MakeCustIndex (IdxType)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayCaScrn (ScrnName$)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
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
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'CACUST.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Calculate Interest"
  MChoice$(2) = "Edit Interest Transactions"
  MChoice$(3) = "Print Interest Edit Report "
  MChoice$(4) = "Post Interest to Accounts"
  MChoice$(5) = "Exit to OS"
  
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
    
    TitleBox 2, Col, MaxLen + 3, "Assessment 'Interest' Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      CalcInterest
    CASE 2
      EditTrans
    CASE 3
      PrintTrans
    CASE 4
      PostTransactions
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
  
  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  END IF
  
  RUN "CAMENU"

SUB CalcInterest
  
  SHARED Choice$()
  
  REDIM CaCust(1)  AS CACUSTType
  REDIM CaTrans(1) AS CaTransType
  REDIM IntTrans(1) AS CaIntEditType
  
  CaCustLen = LEN(CaCust(1))  'Length of Cust Record Structure
  TIntLen = LEN(IntTrans(1))
  
  REDIM Fmt$(3)
  
  REDIM TempScrn(0)
  
  ClearBack
  
  ShowProcessingScrn "Calculating Interest"
  
  IF Exist%("CAINT.DAT") THEN
    KILL "CAINT.DAT"             'kill any old work file
  END IF
  
  TIntFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM SHARED AS TIntFile LEN = TIntLen
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CaFile LEN = CaCustLen
  NumOfCRecs& = LOF(CaFile) \ CaCustLen
  
  'CalcInt calc start *******************************
  FOR Cnt& = 1 TO NumOfCRecs&
    CustAcct& = Cnt&
    GET #CaFile, CustAcct&, CaCust(1)         'get cust record
   'Set to Zero to Begin for Each Customer
    CustIntFlag = 0                           'flag to id customers charged
    IntTrans(1).IStreetAmt = 0
    IntTrans(1).ISideAmt = 0
    IntTrans(1).IDriveAmt = 0
    IntTrans(1).IDemoAmt = 0
    IntTrans(1).IWaterAmt = 0
    IntTrans(1).ISewerAmt = 0
   'Figure Interest and Post to Edit File if Necessary
   IF CaCust(1).RPBStreet > 0 THEN
     CalcInt# = 0
     CalcInt# = CaCust(1).RPBStreet * (CaCust(1).STRRate / 100)
     CalcInt# = Round#(CalcInt# / 12)
     IntTrans(1).IStreetAmt = CalcInt#
     CustIntFlag = 1
   END IF

   IF CaCust(1).RPBSide > 0 THEN
     CalcInt# = 0
     CalcInt# = CaCust(1).RPBSide * (CaCust(1).SIDERate / 100)
     CalcInt# = Round#(CalcInt# / 12)
     IntTrans(1).ISideAmt = CalcInt#
     CustIntFlag = 1
    END IF

   IF CaCust(1).RPBDrive > 0 THEN
     CalcInt# = 0
     CalcInt# = CaCust(1).RPBDrive * (CaCust(1).DRIVERate / 100)
     CalcInt# = Round#(CalcInt# / 12)
     IntTrans(1).IDriveAmt = CalcInt#
     CustIntFlag = 1
   END IF

   IF CaCust(1).RPBDemo > 0 THEN
     CalcInt# = 0
     CalcInt# = CaCust(1).RPBDemo * (CaCust(1).DEMORate / 100)
     CalcInt# = Round#(CalcInt# / 12)
     IntTrans(1).IDemoAmt = CalcInt#
     CustIntFlag = 1
   END IF

   IF CaCust(1).RPBWater > 0 THEN
     CalcInt# = 0
     CalcInt# = CaCust(1).RPBWater * (CaCust(1).WATERRate / 100)
     CalcInt# = Round#(CalcInt# / 12)
     IntTrans(1).IWaterAmt = CalcInt#
     CustIntFlag = 1
    END IF

   IF CaCust(1).RPBSewer > 0 THEN
     CalcInt# = 0
     CalcInt# = CaCust(1).RPBSewer * (CaCust(1).SEWERRate / 100)
     CalcInt# = Round#(CalcInt# / 12)
     IntTrans(1).ISewerAmt = CalcInt#
     CustIntFlag = 1
   END IF

   IF CustIntFlag = 1 THEN
     IntTrans(1).CustRecord = Cnt&
     IntTrans(1).CustName = CaCust(1).CustName
     IntTrans(1).DelFlag = 0
     IntTrans(1).Pad = ""
     NextIRec& = (LOF(TIntFile) \ TIntLen) + 1
     PUT TIntFile, NextIRec&, IntTrans(1)
   END IF
CalcIntSkip:
    ShowPctComp Cnt&, NumOfCRecs&
  NEXT

  'CalcInt Calc END   *******************************
  CLOSE

  'Print Report Here
   PrintTrans

ExitCalcInt:
   EXIT SUB
  
  
ShowInfoErr1:
  SaveScrn TempScrn()
  DisplayCaScrn "ERRSCRN1"
  QPrintRC "Invalid CalcInting Information!", 10, 25, -1
  QPrintRC "Please correct and try again.", 12, 26, -1
  WaitForAction
  RestScrn TempScrn()
  RETURN
  
  
END SUB

SUB EditTrans
  
  RecDisp$ = SPACE$(14)
  SHARED Choice$()
  
  REDIM EdScrn(0)
  
  REDIM CaIntRec(1 TO 1) AS CaIntEditType
  CaIntRecLen = LEN(CaIntRec(1))
  
  CaFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM AS CaFile LEN = CaIntRecLen      'open data file
  NumIntRec& = LOF(CaFile) / CaIntRecLen
  CLOSE CaFile
  
  CursorOff
  ClearBack
  
  IF NumIntRec& < 1 THEN
    DisplayCaScrn "ERRSCRN1"
    QPrintRC "No Interest Transaction to Edit!", 10, 25, -1
    QPrintRC "Press any key to continue.", 12, 28, -1
    WaitForAction
    GOTO EditIntExit
  ELSE
    GOSUB CheckAnyNotDeleted
    IF NoneToEdit THEN
      DisplayCaScrn "ERRSCRN1"
      QPrintRC "No Active Interest Transactions!", 10, 25, -1
      QPrintRC "Press any key to continue.", 12, 28, -1
      WaitForAction
      GOTO EditIntExit
    END IF
  END IF
  
  LibName$ = "TAX"
  ScrnName$ = "CAINTED"
  
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
  
  DisplayCaScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      Forward = True
      IntRecNo& = 1
      GOSUB GetIntRec
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key ', EnterKey    'F10   Save
      GOSUB UpDateIntRec
    CASE F3Key
      Ok = MsgBox%("TAX", "DELINTRN")
      SELECT CASE Ok
      CASE 2    'if button 2 was pressed
        GOSUB DelIntTrans
        GOSUB GetIntRec
        'FirstTime = True
      CASE ELSE
      END SELECT
      Action = 1
      
    CASE PgUpKey                'previous rec
      Forward = False
      IF IntRecNo& > 1 THEN
        IntRecNo& = IntRecNo& - 1
        GOSUB GetIntRec
      END IF
    CASE PgDnKey                'next rec
      Forward = True
      IF IntRecNo& < NumIntRec& THEN
        IntRecNo& = IntRecNo& + 1
        GOSUB GetIntRec
      END IF

    CASE EscKey
      RecNo& = 0
      ExitFlag = True
    CASE -118  'Ctrl PgDn
      Forward = True
      IF IntRecNo& < NumIntRec& THEN
        IntRecNo& = NumIntRec&
        GOSUB GetIntRec
      END IF

    CASE -132  'Ctrl PgUp
      Forward = False
      IF IntRecNo& > 1 THEN
        IntRecNo& = 1
        GOSUB GetIntRec
      END IF
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 17
        SELECT CASE Frm(1).MCol
        CASE 8 TO 17            'ESC Cancel button
          PressButton EscKey, 17, 8, 17
        CASE 18 TO 28           'F3 Delete
          PressButton F3Key, 17, 18, 28
        CASE 29 TO 39           'F5 Search
          PressButton F5KEY, 17, 29, 39
        CASE 40 TO 49           'F10 Save Button
          PressButton F10Key, 17, 40, 49
        CASE 50 TO 60           'PgUp Prev Rec
          PressButton PgUpKey, 17, 50, 60
        CASE 61 TO 71           'PgDn Next Rec
          PressButton PgDnKey, 17, 61, 71
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  GOTO EditIntExit
  
GetIntRec:
  DidFirst = False
  DidLast = False
  CaFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM AS CaFile LEN = CaIntRecLen      'open data file
  DO
    GET CaFile, IntRecNo&, CaIntRec(1)
    IF CaIntRec(1).DelFlag <> 0 THEN
      SELECT CASE Forward
      CASE True
        IF IntRecNo& < NumIntRec& THEN
          IntRecNo& = IntRecNo& + 1
        ELSE
          Forward = False
          DidLast = True
        END IF
      CASE False
        IF IntRecNo& > 1 THEN
          IntRecNo& = IntRecNo& - 1
        ELSE
          Forward = True
          DidFirst = True
        END IF
      END SELECT
    ELSE
      EXIT DO
    END IF
    IF DidFirst AND DidLast THEN
      CursorOff
      CLOSE CaFile
      DisplayCaScrn "ERRSCRN1"
      QPrintRC "No Active Interest Transactions!", 10, 25, -1
      QPrintRC "Press any key to continue.", 12, 28, -1
      WaitForAction
      GOTO EditIntExit
    END IF
  LOOP
  CLOSE CaFile
  Form$(1, 0) = QPTrim$(STR$(CaIntRec(1).CustRecord))
  Form$(2, 0) = CaIntRec(1).CustName
  Form$(3, 0) = QPTrim$(STR$(CaIntRec(1).IStreetAmt))
  Form$(4, 0) = QPTrim$(STR$(CaIntRec(1).ISideAmt))
  Form$(5, 0) = QPTrim$(STR$(CaIntRec(1).IDriveAmt))
  Form$(6, 0) = QPTrim$(STR$(CaIntRec(1).IDemoAmt))
  Form$(7, 0) = QPTrim$(STR$(CaIntRec(1).IWaterAmt))
  Form$(8, 0) = QPTrim$(STR$(CaIntRec(1).ISewerAmt))
  Action = 1
  GOSUB DisplayRecCnt
RETURN
  
CheckAnyNotDeleted:
  ShowProcessingScrn "Scanning for Active Transactions"
  NoneToEdit = True
  CaFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM AS CaFile LEN = CaIntRecLen      'open data file
  NumIntRec& = LOF(CaFile) / CaIntRecLen
  FOR Cnt& = 1 TO NumIntRec&
    GET CaFile, Cnt&, CaIntRec(1)
    IF CaIntRec(1).DelFlag = 0 THEN
      NoneToEdit = False
    END IF
    ShowPctCompL Cnt&, NumIntRec&
  NEXT
  CLOSE CaFile
  'SmallPause
RETURN
  
UpDateIntRec:
  SaveScrn EdScrn()
  CaFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM AS CaFile LEN = CaIntRecLen      'open data file
  GET CaFile, IntRecNo&, CaIntRec(1)
   CaIntRec(1).IStreetAmt = Value#(Form$(3, 0), ecode)
   CaIntRec(1).ISideAmt = Value#(Form$(4, 0), ecode)
   CaIntRec(1).IDriveAmt = Value#(Form$(5, 0), ecode)
   CaIntRec(1).IDemoAmt = Value#(Form$(6, 0), ecode)
   CaIntRec(1).IWaterAmt = Value#(Form$(7, 0), ecode)
   CaIntRec(1).ISewerAmt = Value#(Form$(8, 0), ecode)
  PUT CaFile, IntRecNo&, CaIntRec(1)
  CLOSE CaFile
  DisplayCaScrn "UPDATEOK"
  WaitForAction
  RestScrn EdScrn()
RETURN
  
DelIntTrans:
  SaveScrn EdScrn()
  CaFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM AS CaFile LEN = CaIntRecLen      'open data file
  GET CaFile, IntRecNo&, CaIntRec(1)
  CaIntRec(1).DelFlag = True
  PUT CaFile, IntRecNo&, CaIntRec(1)
  CLOSE CaFile
  DisplayCaScrn "UPDATEOK"
  WaitForAction
  RestScrn EdScrn()
RETURN
  
DisplayRecCnt:
  LSET RecDisp$ = QPTrim$(STR$(IntRecNo&)) + " of " + QPTrim$(STR$(NumIntRec&))
  QPrintRC RecDisp$, 8, 56, -1
RETURN
  
EditIntExit:
  
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

SUB PostTransactions

  REDIM TempScrn(0)
  
  REDIM CaCust(1)  AS CACUSTType
  REDIM CaTrans(1) AS CaTransType
  REDIM CaIntRec(1 TO 1) AS CaIntEditType
  CaIntRecLen = LEN(CaIntRec(1))
  IntRecLen = LEN(CaIntRec(1))
BeginHere:
  ClearBack
  ShowProcessingScrn "Verifying Interest Transactions"

  IntFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM AS IntFile LEN = IntRecLen
  NumOfRec& = LOF(IntFile) / IntRecLen
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #IntFile, Cnt&, CaIntRec(1)
    IF CaIntRec(1).DelFlag = 0 THEN
      DidSome = DidSome + 1
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  'SmallPause

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
  OPEN "CAINT.DAT" FOR RANDOM AS IntFile LEN = IntRecLen
  NumOfRec& = LOF(IntFile) / IntRecLen
  
  CustFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM AS CustFile LEN = LEN(CaCust(1))
  
  TransFile = FREEFILE
  OPEN "CATRANS.DAT" FOR RANDOM AS TransFile LEN = LEN(CaTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #IntFile, Cnt&, CaIntRec(1)
    IF CaIntRec(1).DelFlag = 0 THEN
      REDIM CaTrans(1) AS CaTransType
     CaTrans(1).TranType = 1
     CaTrans(1).Date = Date2Num%(DATE$)
     CaTrans(1).IStreetAmt = CaIntRec(1).IStreetAmt
     CaTrans(1).ISideAmt = CaIntRec(1).ISideAmt
     CaTrans(1).IDriveAmt = CaIntRec(1).IDriveAmt
     CaTrans(1).IDemoAmt = CaIntRec(1).IDemoAmt
     CaTrans(1).IWaterAmt = CaIntRec(1).IWaterAmt
     CaTrans(1).ISewerAmt = CaIntRec(1).ISewerAmt
     CaTrans(1).PStreetAmt = 0
     CaTrans(1).PSideAmt = 0
     CaTrans(1).PDriveAmt = 0
     CaTrans(1).PDemoAmt = 0
     CaTrans(1).PWaterAmt = 0
     CaTrans(1).PSewerAmt = 0
     CaTrans(1).NextTrans = 0
     CaTrans(1).CustRecord = CaIntRec(1).CustRecord
     CaTrans(1).Pad = ""

     'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(CaTrans(1))) + 1
      PUT #TransFile, NextRecord&, CaTrans(1)

     'Update the Customer Pointers and Balances Now
      GET #CustFile, CaIntRec(1).CustRecord, CaCust(1)
       CaCust(1).IStreet = CaCust(1).IStreet + CaIntRec(1).IStreetAmt
       CaCust(1).ISide = CaCust(1).ISide + CaIntRec(1).ISideAmt
       CaCust(1).IDrive = CaCust(1).IDrive + CaIntRec(1).IDriveAmt
       CaCust(1).IDemo = CaCust(1).IDemo + CaIntRec(1).IDemoAmt
       CaCust(1).IWater = CaCust(1).IWater + CaIntRec(1).IWaterAmt
       CaCust(1).ISewer = CaCust(1).ISewer + CaIntRec(1).ISewerAmt

       CaCust(1).RIBStreet = CaCust(1).RIBStreet + CaIntRec(1).IStreetAmt
       CaCust(1).RIBSide = CaCust(1).RIBSide + CaIntRec(1).ISideAmt
       CaCust(1).RIBDrive = CaCust(1).RIBDrive + CaIntRec(1).IDriveAmt
       CaCust(1).RIBDemo = CaCust(1).RIBDemo + CaIntRec(1).IDemoAmt
       CaCust(1).RIBWater = CaCust(1).RIBWater + CaIntRec(1).IWaterAmt
       CaCust(1).RIBSewer = CaCust(1).RIBSewer + CaIntRec(1).ISewerAmt
       CaCust(1).LastIntDate = Date2Num%(PostDate$)

      IF CaCust(1).LastTrans = 0 THEN
        CaCust(1).LastTrans = NextRecord&
        PUT #CustFile, CaIntRec(1).CustRecord, CaCust(1)
      ELSE
        Previous& = CaCust(1).LastTrans
        CaCust(1).LastTrans = NextRecord&
        PUT #CustFile, CaIntRec(1).CustRecord, CaCust(1)
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
  KillFile "CAINT.DAT"
  
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

DEFSNG A-Z
SUB PrintTrans

  REDIM IntTrans(1) AS CaIntEditType
  TIntLen = LEN(IntTrans(1))
  dash$ = STRING$(79, "-")
  TheDate$ = DATE$
  MaxLines = 60
  LineCnt = 0

  TIntFile = FREEFILE
  OPEN "CAINT.DAT" FOR RANDOM SHARED AS TIntFile LEN = TIntLen
  IntRecord& = LOF(TIntFile) / TIntLen
  IF IntRecord& = 0 THEN
    EXIT SUB
  END IF


  'CalcInt Print Start #############
  InterestReportFile$ = "CAINT.PRN"
  ClearBack
  ShowProcessingScrn "Creating Interest Report"
  RptFile = FREEFILE
  OPEN InterestReportFile$ FOR OUTPUT AS #RptFile

  GOSUB CalcIntHeading1

  FOR Cnt& = 1 TO IntRecord&
    GET #TIntFile, Cnt&, IntTrans(1)
    IF IntTrans(1).DelFlag <> 0 THEN
      GOTO SkipThisIntTrans
    END IF
   TotalInt# = 0
   TotalInt# = IntTrans(1).IStreetAmt + IntTrans(1).ISideAmt + IntTrans(1).IDriveAmt + IntTrans(1).IDemoAmt + IntTrans(1).IWaterAmt + IntTrans(1).ISewerAmt
   TotalInt# = Round#(TotalInt#)

   PRINT #RptFile, USING "######"; IntTrans(1).CustRecord;
    PRINT #RptFile, TAB(10); IntTrans(1).CustName
   PRINT #RptFile, USING "#####.##"; IntTrans(1).IStreetAmt;
    PRINT #RptFile, TAB(12); USING "#####.##"; IntTrans(1).ISideAmt;
    PRINT #RptFile, TAB(24); USING "#####.##"; IntTrans(1).IDriveAmt;
    PRINT #RptFile, TAB(36); USING "#####.##"; IntTrans(1).IDemoAmt;
    PRINT #RptFile, TAB(48); USING "#####.##"; IntTrans(1).IWaterAmt;
    PRINT #RptFile, TAB(60); USING "#####.##"; IntTrans(1).ISewerAmt;
    PRINT #RptFile, TAB(71); USING "#####.##"; TotalInt#
    PRINT #RptFile, dash$
    GTotalInt# = GTotalInt# + TotalInt#
    GTotalInt# = Round#(GTotalInt#)
    LineCnt = LineCnt + 3
   IF LineCnt >= MaxLines THEN
      PRINT #RptFile, CHR$(12);
      GOSUB CalcIntHeading1
    END IF
SkipThisIntTrans:
    ShowPctComp Cnt&, IntRecord&
  NEXT

  
  PRINT #RptFile, "Total Interest Charged .. "; USING "$$######,#.##"; GTotalInt#
  PRINT #RptFile, CHR$(12);
  Page = Page + 1
  CLOSE
  PrintRptFile "Assessment Interest Report.", InterestReportFile$, 1, RetCode%, 0
  'CalcInt Print END   #############

  EXIT SUB

CalcIntHeading1:
  Page = Page + 1
  PRINT #RptFile, TAB(15); "Assessment Billing : Interest Calculation Register"
  PRINT #RptFile,
  PRINT #RptFile, "Date: "; TheDate$; TAB(65); "Page #"; Page
  PRINT #RptFile,
  PRINT #RptFile, "Acct #"; TAB(10); "Customer Name"
  PRINT #RptFile, " Streets"; TAB(12); "SideWalks"; TAB(24); "DriveWays"; TAB(36); "Demolition"; TAB(48); "   Water"; TAB(60); "   Sewer"; TAB(71); "Total Int"
  PRINT #RptFile, dash$
  LineCnt = 7
  RETURN

END SUB

DEFINT A-Z
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
  ACCT$ = SPACE$(5)
  
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

