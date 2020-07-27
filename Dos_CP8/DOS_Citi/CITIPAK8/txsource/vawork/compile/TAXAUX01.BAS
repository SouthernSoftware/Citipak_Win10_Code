DEFINT A-Z
DECLARE SUB Search4Cust4 (SName$, RecNo&, ChkBalFlag%, CLSFlag%, False%, TaxType%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB Search4Cust3 (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
DECLARE SUB Search4Cust2 (SName$, RecNo&, ChkBalFlag%, CLSFlag%, False%, TaxType%)
DECLARE SUB Search4Cust1 (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
DECLARE FUNCTION Round# (N#)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Unique$ (Path$)
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE FUNCTION GetTaxCustCnt& (TaxType%)
DECLARE FUNCTION IsCustDeleted% (AcctNum&, TaxType%)
DECLARE SUB ClearBack ()
DECLARE SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB DelPropAbstract (PropRecs() AS LONG, WhatProp%, CustRec&)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION FUsing$ (Number$, Image$)

DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPStrI$ (Num%)
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
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE FUNCTION DoesCustOwe% (TaxCustRec AS ANY)
DECLARE FUNCTION GetCustName$ (CustRec&, TaxType%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION ParseBillNum$ (Text$)


  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE SortStruct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  DIM SHARED PctC(1) AS STRING * 4
  DIM SHARED TaxInterestCalc AS TaxInterestDateType

SUB ClearBack
  LibFile2Scrn "TAX", "BAKCLEAR", MonoCode%, Attribute%, ErrorCode%
END SUB

SUB ClearScrn STATIC
  WazzWind 1, 1, 25, 80, 7, 0, 0
END SUB

SUB CursorOff STATIC
  LOCATE , , 0
END SUB

SUB DelPersAbstract (PersRecs() AS LONG, WhatPers%, CustRec&)

  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))

  Pers2Free& = PersRecs(WhatPers)
  NumOfPers& = PersRecs(0)

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)

  FirstPers& = TaxCust(1).FirstPersRec

'First free the Personal in question
  GET PersFile, Pers2Free&, PersRec(1)
  PersRec(1).NextRec = 0
  PersRec(1).CustPin = 0
  PersRec(1).Deleted = True
  PUT PersFile, Pers2Free&, PersRec(1)
'Personal has been marked deleted

  IF Pers2Free& = FirstPers& THEN      'if this was the cust's only Pers
    TaxCust(1).FirstPersRec = 0        'set Pers pointer to 0
    PUT TaxFile, CustRec&, TaxCust(1)  'store cust info
    GOTO DonePersDel                    'were finished.
  END IF

  REDIM TPersRecs(0 TO NumOfPers& - 1)

  FOR Cnt& = 1 TO NumOfPers&
    ThisPers& = PersRecs(Cnt&)
    IF ThisPers& <> Pers2Free& THEN
      DidCnt = DidCnt + 1
      TPersRecs(DidCnt) = ThisPers&
    END IF
  NEXT

  FOR Cnt = 1 TO DidCnt
    ThisPers& = TPersRecs(Cnt)
    GET PersFile, ThisPers&, PersRec(1)
    IF Cnt = 1 THEN
      TaxCust(1).FirstPersRec = ThisPers&
      PUT TaxFile, CustRec&, TaxCust(1)
    END IF
    IF Cnt < DidCnt THEN
      NextPers& = TPersRecs(Cnt + 1)
    ELSE
      NextPers& = 0
    END IF
    PersRec(1).NextRec = NextPers&
    PUT PersFile, ThisPers&, PersRec(1)
  NEXT

DonePersDel:
  CLOSE
  ERASE PersRec, TaxCust

  DisplayTaxScrn "UPDATEOK"
  WaitForAction

END SUB

SUB DelPropAbstract (PropRecs() AS LONG, WhatProp%, CustRec&)
                                                            
'PropRecs() holds rec# pointers to all property records
'WhatProp%  Rec# of the property to delete
'CustRec&   Customers rec#

  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))

  Prop2Free& = PropRecs(WhatProp)
  NumOfProp& = PropRecs(0)

  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)

  FirstProp& = TaxCust(1).FirstPropRec

'First free the property in question
  GET PropFile, Prop2Free&, PropRec(1)
  PropRec(1).NextRec = 0
  PropRec(1).CustPin = 0
  PropRec(1).Deleted = True
  PUT PropFile, Prop2Free&, PropRec(1)
'property has been marked deleted

  IF Prop2Free& = FirstProp& THEN      'if this was the cust's only prop
    TaxCust(1).FirstPropRec = 0        'set prop pointer to 0
    PUT TaxFile, CustRec&, TaxCust(1)  'store cust info
    GOTO DoneDelete                    'were finished.
  END IF
  
  REDIM TPropRecs(0 TO NumOfProp& - 1)

  FOR Cnt& = 1 TO NumOfProp&
    ThisProp& = PropRecs(Cnt&)
    IF ThisProp& <> Prop2Free& THEN
      DidCnt = DidCnt + 1
      TPropRecs(DidCnt) = ThisProp&
    END IF
  NEXT

  FOR Cnt = 1 TO DidCnt
    ThisProp& = TPropRecs(Cnt)
    GET PropFile, ThisProp&, PropRec(1)
    IF Cnt = 1 THEN
      TaxCust(1).FirstPropRec = ThisProp&
      PUT TaxFile, CustRec&, TaxCust(1)
    END IF
    IF Cnt < DidCnt THEN
      NextProp& = TPropRecs(Cnt + 1)
    ELSE
      NextProp& = 0
    END IF
    PropRec(1).NextRec = NextProp&
    PUT PropFile, ThisProp&, PropRec(1)
  NEXT

DoneDelete:
  CLOSE
  ERASE PropRec, TaxCust

  DisplayTaxScrn "UPDATEOK"
  WaitForAction

END SUB

SUB DisplayTaxScrn (ScrnName$)
  LibFile2Scrn "TAX", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  IF ErrorCode% <> 0 THEN

    PRINT "Screen Error: "; ScrnName$
    END
  END IF
END SUB

FUNCTION DoesCustOwe% (TaxCustRec AS TaxCustType)
  
  DoesCustOwe% = False          'assume the customer owes nothing
  
  REDIM TaxTrans(1) AS TaxTransactionType
  IF TaxCustRec.LastTrans > 0 THEN
    TransFile = FREEFILE
    OPEN "TaxTrans.dat" FOR RANDOM SHARED AS TransFile LEN = LEN(TaxTrans(1))
    TransRecord& = TaxCustRec.LastTrans
    DO WHILE TransRecord& <> 0
      GET TransFile, TransRecord&, TaxTrans(1)
      IF TaxTrans(1).TranType = 1 THEN
        Balance# = Round#(TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5)
        Balance# = Round#(Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.Penalty + TaxTrans(1).Revenue.Collection)
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd))
        Balance# = Round#(Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd))
        IF Balance# > 0 THEN
          EXIT DO
        END IF
      END IF
      TransRecord& = TaxTrans(1).LastTrans
    LOOP
    CLOSE TransFile
    IF Balance# > 0 THEN
      DoesCustOwe% = True
    END IF
  END IF
  
END FUNCTION

FUNCTION GetCustName$ (CustRec&, TaxType%)

  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure

  TaxFile = FREEFILE
  IF TaxType% = 1 THEN
   OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  ELSE
   OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  END IF
  GET TaxFile, CustRec&, TaxCust(1)
  CLOSE TaxFile
  GetCustName$ = QPTrim$(TaxCust(1).FName) + " " + QPTrim$(TaxCust(1).LNAME)

  ERASE TaxCust

END FUNCTION

SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
  
  'put routine here to create temp file if adding new cust
  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))
  
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  
  REDIM PersRecs(0 TO 0) AS LONG
  
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)
  CLOSE TaxFile
  
  WhatPers& = TaxCust(1).FirstPersRec
  IF WhatPers& > 0 THEN
    DO
      PCnt = PCnt + 1
      REDIM PRESERVE PersRecs(0 TO PCnt) AS LONG
      PersRecs(PCnt) = WhatPers&
      GET PersFile, WhatPers&, PersRec(1)
      WhatPers& = PersRec(1).NextRec
    LOOP WHILE WhatPers& > 0
    PersRecs(0) = PCnt
  ELSE
    PersRecs(0) = 0
  END IF
  
  CLOSE
  
  ERASE PersRec, TaxCust
  
END SUB

SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
  
  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))
  
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  
  REDIM PropRecs(0 TO 0) AS LONG
  
  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)
  CLOSE TaxFile
  
  WhatProp& = TaxCust(1).FirstPropRec
  IF WhatProp& > 0 THEN
    DO
      PCnt = PCnt + 1
      REDIM PRESERVE PropRecs(0 TO PCnt) AS LONG
      PropRecs(PCnt) = WhatProp&
      GET PropFile, WhatProp&, PropRec(1)
      WhatProp& = PropRec(1).NextRec
    LOOP WHILE WhatProp& > 0
    PropRecs(0) = PCnt
  ELSE
    PropRecs(0) = 0
  END IF
  
  CLOSE
  
  ERASE PropRec, TaxCust
  
END SUB

FUNCTION GetTaxCustCnt& (TaxType%)
  
  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  
  TaxFile = FREEFILE
  IF TaxType% = 1 THEN
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  ELSE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  END IF
  
  NumOfRecs& = LOF(TaxFile) \ TaxCustLen
  CLOSE TaxFile
  
  ERASE TaxCust
  
  GetTaxCustCnt& = NumOfRecs&
  
END FUNCTION

FUNCTION IsCustDeleted (AcctNum&, TaxType%)
  IsCustDeleted = False         'assume they aren't deleted
  
  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  SELECT CASE TaxType
  CASE 1    'real
    FileName$ = RETaxCustFile
  CASE 2    'personal
    FileName$ = PPTaxCustFile
  END SELECT

  TaxFile = FREEFILE

  OPEN FileName$ FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  GET TaxFile, AcctNum&, TaxCust(1)
  CLOSE TaxFile
  
  IF TaxCust(1).Deleted <> 0 THEN
    IsCustDeleted = True
  END IF
  
  ERASE TaxCust
  
END FUNCTION

SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)

LookUpTop:

  REDIM Hlp$(1 TO 7)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Search Name here."
  Hlp$(3) = "Enter all or part of the SSN to search for here." + CHR$(13)
  Hlp$(3) = Hlp$(3) + "NOTE: a blank space will match any digit. The" + CHR$(13)
  Hlp$(3) = Hlp$(3) + "Customer and Spouses SSN are searched."
  Hlp$(4) = "Enter Bill Number to Find Here."
  Hlp$(5) = "Enter Street Address Here."
  Hlp$(6) = "Enter G-PIN Number Here."
  Hlp$(7) = "Enter County Account Number Here."
  
  REDIM TaxCust(1) AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))
  
  SName$ = ""
  AcctNum& = 0
  LScrn = 2
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 7) AS STRING * 17
  
  MScrn = 7
  
  LText(1) = " Account Number:"
  LText(2) = "    Search Name:"
  LText(3) = "Social Security:"
  LText(4) = "Tax Bill Number:"
  LText(5) = " Street Address:"
  LText(6) = "   G-Pin Number:"
  LText(7) = "    County Acct:"
  
  LibName$ = "TAX"
  ScrnName$ = "VALUACCT"
  
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
  DisplayTaxScrn ScrnName$
  
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
      CASE 3
        Fld(1).FType = SoSecFld
        LSET Form$(1, 0) = ""
      CASE 4
        LSET Form$(1, 0) = ""
        Fld(1).FType = 2
      CASE 5
        LSET Form$(1, 0) = ""
        Fld(1).FType = 2
      CASE 6
        LSET Form$(1, 0) = ""
        Fld(1).FType = 2
      CASE 7
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
        IF AcctNum& < 1 OR AcctNum& > GetTaxCustCnt&(TaxType%) THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")
        ELSEIF IsCustDeleted(AcctNum&, TaxType%) THEN
          Ok = MsgBox%("TAX.QSL", "DELACCTN")
        ELSEIF ChkBalFlag THEN
          REDIM TaxCust(1) AS TaxCustType
          TaxCustLen = LEN(TaxCust(1))
          IF TaxType% = 1 THEN
          FOpenS RETaxCustFile, TaxFile   'open data file
          ELSE
          FOpenS PPTaxCustFile, TaxFile   'open data file
          END IF
          FGetRTA TaxFile, TaxCust(1), AcctNum&, TaxCustLen
          FClose TaxFile
          IF DoesCustOwe%(TaxCust(1)) THEN
            SaveScrn ScrnArray2()
            CursorOff
            DisplayTaxScrn "ERRSCRN1"
            QPrintRC "This account HAS A BALANCE", 10, 27, -1
            QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
            WaitForAction
            RecNo& = 0
            RestScrn ScrnArray2()
          ELSE
            RecNo& = AcctNum&
            OKFlag = True
          END IF
        ELSE
          RecNo& = AcctNum&
          OKFlag = True
        END IF
        Action = 1
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 3      'ssn
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, True, TaxType
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 4           'tax bill #
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust1 SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 5
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust2 SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 6
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust3 SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 7
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust4 SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType
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
  
  ERASE TaxCust, ScrnArray, ScrnArray2
  ERASE Frm, Form$, Fld, LText, Hlp$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText(LScrn), 12, 15, -1
RETURN

END SUB

SUB MakeCustIndex (IdxType)
  
  ShowProcessingScrn "Creating Customer Name Index"
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  
  CustFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM AS CustFile LEN = TaxRecLen
  NumOfCRecs = LOF(CustFile) / TaxRecLen
  
  QPrintRC "Reading Customer Information", 11, 27, -1
  
  REDIM Array(1 TO NumOfCRecs) AS SortStruct
  FOR Cnt = 1 TO NumOfCRecs
    GET CustFile, Cnt, TaxCust(1)
    SELECT CASE IdxType
    CASE 1
      Array(Cnt).who = (LEFT$(QPTrim$(TaxCust(1).LNAME), 12) + LEFT$(QPTrim$(TaxCust(1).LNAME), 2))
    CASE 2
      Array(Cnt).who = QPTrim$(TaxCust(1).SName)
    CASE 3
      Array(Cnt).who = QPTrim$(TaxCust(1).CSSN)
    END SELECT
    Array(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE
  
  QPrintRC "Sorting Customer Information", 11, 27, -1
  'Sort Them Here
  SortT Array(1), NumOfCRecs, 0, LEN(Array(1)), 0, 14
  'SortT (Element, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  QPrintRC "   Writing Customer Index   ", 11, 27, -1
  
  Idxfile = FREEFILE
  OPEN "TAXTEMP.IDX" FOR RANDOM AS Idxfile LEN = 2
  FOR Cnt = 1 TO NumOfCRecs
    PUT #Idxfile, Cnt, Array(Cnt).RecNum
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE
  
END SUB

SUB MakeCustIndexRE (IdxType)

  ShowProcessingScrn "Creating Customer Name Index"
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))

  CustFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM AS CustFile LEN = TaxRecLen
  NumOfCRecs = LOF(CustFile) / TaxRecLen

  QPrintRC "Reading Customer Information", 11, 27, -1

  REDIM Array(1 TO NumOfCRecs) AS SortStruct
  FOR Cnt = 1 TO NumOfCRecs
    GET CustFile, Cnt, TaxCust(1)
    SELECT CASE IdxType
    CASE 1
      Array(Cnt).who = (LEFT$(QPTrim$(TaxCust(1).LNAME), 12) + LEFT$(QPTrim$(TaxCust(1).LNAME), 2))
    CASE 2
      Array(Cnt).who = QPTrim$(TaxCust(1).SName)
    CASE 3
      Array(Cnt).who = QPTrim$(TaxCust(1).CSSN)
    END SELECT
    Array(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE

  QPrintRC "Sorting Customer Information", 11, 27, -1
  'Sort Them Here
  SortT Array(1), NumOfCRecs, 0, LEN(Array(1)), 0, 14
  'SortT (Element, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)

  QPrintRC "   Writing Customer Index   ", 11, 27, -1

  Idxfile = FREEFILE
  OPEN "TAXTEMP.IDX" FOR RANDOM AS Idxfile LEN = 2
  FOR Cnt = 1 TO NumOfCRecs
    PUT #Idxfile, Cnt, Array(Cnt).RecNum
    ShowPctComp Cnt, NumOfCRecs
  NEXT
  CLOSE


END SUB

SUB MakePersPINFile
  
  ShowProcessingScrn "Creating PIN Search File"
  
  PINFile = FREEFILE
  OPEN TaxPersPINFile FOR OUTPUT AS #PINFile
  CLOSE PINFile
  
  REDIM PersPINS(1) AS PINSearchType
  PersPINSLen = LEN(PersPINS(1))
  
  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))
  
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  NumPersRecs& = LOF(PersFile) / PersRecLen
  
  PPINFile = FREEFILE
  OPEN TaxPersPINFile FOR RANDOM SHARED AS PPINFile LEN = PersPINSLen
  
  FOR Cnt& = 1 TO NumPersRecs&
    GET #PersFile, Cnt&, PersRec(1)
    PersPINS(1).PIN = PersRec(1).PROPPIN
    PersPINS(1).Cust = Cnt&
    PUT #PPINFile, Cnt&, PersPINS(1)
    ShowPctComp Cnt&, NumPersRecs&
  NEXT
  
  CLOSE
  
END SUB

SUB MakeRealPINFile
  
  ShowProcessingScrn "Creating PIN Search File"
  
  PINFile = FREEFILE
  OPEN TaxRealPINFile FOR OUTPUT AS #PINFile
  CLOSE PINFile
  
  REDIM RealPINS(1) AS PINSearchType
  RealPINSLen = LEN(RealPINS(1))
  
  REDIM RealRec(1) AS PropertyRecType
  RealRecLen = LEN(RealRec(1))
  
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  NumRealRecs& = LOF(RealFile) / RealRecLen
  
  RPINFile = FREEFILE
  OPEN TaxRealPINFile FOR RANDOM SHARED AS RPINFile LEN = RealPINSLen
  
  FOR Cnt& = 1 TO NumRealRecs&
    GET #RealFile, Cnt&, RealRec(1)
    RealPINS(1).PIN = RealRec(1).REALPIN
    RealPINS(1).Cust = Cnt&
    PUT #RPINFile, Cnt&, RealPINS(1)
    ShowPctComp Cnt&, NumRealRecs&
  NEXT
  CLOSE
  
END SUB

FUNCTION ParseBillNum$ (Text$)

  BILLNUM$ = QPTrim$(Text$)
  BNumLen = LEN(BILLNUM$)
  IF BNumLen > 0 THEN
    FOR Cnt = BNumLen TO 1 STEP -1
      ThisChar$ = MID$(BILLNUM$, Cnt, 1)
      IF INSTR("0123456789", ThisChar$) <= 0 THEN
        EXIT FOR
      END IF
    NEXT
    GoodPos = Cnt + 1
    BILLNUM$ = MID$(BILLNUM$, GoodPos)
  END IF
  ParseBillNum$ = BILLNUM$
END FUNCTION

FUNCTION PromptSaveData%
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "TAX"
  SaveFlag = 2
  
  FormName$ = "SAVE1ST"
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  DisplayTaxScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      SaveFlag = True
    CASE EscKey
      SaveFlag = 1
    CASE 88, 120                'X Key
      SaveFlag = False
    END SELECT
    
  LOOP WHILE SaveFlag = 2       'proper key not set
  
  PromptSaveData = SaveFlag
  CursorOff
  
  RestScrn TempScrn()
  
  ERASE TempScrn, Form$, Fld, Frm
  
END FUNCTION

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
  
  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(5)

  IF LastSEARCH$ <> Search$ THEN
    LastSEARCH$ = Search$
    Choice = 1
  END IF
  
  IF SSNFlag THEN
    'if searching by the ssn, then strip out the dashes
    TSearch$ = SPACE$(9)
    DashPos = INSTR(Search$, "-")
    DO WHILE DashPos > 0
      Search$ = LEFT$(Search$, DashPos - 1) + MID$(Search$, DashPos + 1)
      DashPos = INSTR(Search$, "-")
    LOOP
    LSET TSearch$ = Search$
    Search$ = TSearch$
  END IF

  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
'091598 Found a bug that caused the lookup to get erroneous records
'       if this was changed to greater than 32???
  CustBlock = 32
  
  REDIM Mchoice(1 TO 1) AS FLen2
  REDIM TaxCust(1 TO CustBlock) AS TaxCustType
  
  TaxCustLen = LEN(TaxCust(1))
  
  SearchLen = LEN(Search$)
  Match = False
  IF TaxType% = 1 THEN
    FOpenS RETaxCustFile, TaxFile   'open data file
  END IF
  IF TaxType% = 2 THEN
    FOpenS PPTaxCustFile, TaxFile   'open data file
  END IF
  
  NumOfCust& = FLof&(TaxFile) / TaxCustLen

  NumChunks& = NumOfCust& \ CustBlock
'      DO NOT CHANGE THIS ^ TO FLOATING POINT DIVISION!!!!!
  
  OddRecs& = NumOfCust& MOD CustBlock
  
  BlockSize& = (0& + TaxCustLen) * CustBlock
  '            ^^^^^ stops an overflow error
  'since TaxCustLen is an integer, basic will try to multiply to
  'an integer result. the above "0&" causes basic to convert to a long
  'then multiply'

  ' Find matching record
  FOR CCnt& = 1 TO NumChunks&
    FGetA TaxFile, TaxCust(1), BlockSize&
    FOR RecCnt = 1 TO CustBlock
      WhatRec& = ((CCnt& - 1) * CustBlock) + RecCnt
      IF SSNFlag THEN
        FOR WhoSSN = 1 TO 2
          SELECT CASE WhoSSN
          CASE 1                'customers
            UBSearchN$ = TaxCust(RecCnt).CSSN
          CASE 2                'spouses
            UBSearchN$ = TaxCust(RecCnt).SSSN
          END SELECT
          SSNOk = True
          FOR DigitCnt = 1 TO 9
            ThisDigit$ = MID$(Search$, DigitCnt, 1)
            IF ThisDigit$ = " " THEN
              'assume a blank in the search$ is any digit
              'and is OK
            ELSE
              SSNDigit$ = MID$(UBSearchN$, DigitCnt, 1)
              IF SSNDigit$ <> ThisDigit$ THEN
                SSNOk = False
                EXIT FOR
              END IF
            END IF
          NEXT
          IF SSNOk THEN
            GOSUB CustLoadEM2
            EXIT FOR
          END IF
        NEXT
      ELSE
        UBSearchN$ = LEFT$(TaxCust(RecCnt).SName, SearchLen)
        IF (Search$ = UBSearchN$) THEN
          GOSUB CustLoadEM2
        END IF
      END IF
      
DelSkip2:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  NEXT
  
  IF OddRecs& > 0 THEN
    NextRec& = (NumChunks& * CustBlock) + 1
    RecCnt = 1
    FOR CCnt& = NextRec& TO NumOfCust&
      FGetRTA TaxFile, TaxCust(1), CCnt&, TaxCustLen
      WhatRec& = CCnt&
      IF SSNFlag THEN
        FOR WhoSSN = 1 TO 2
          SELECT CASE WhoSSN
          CASE 1                'customers
            UBSearchN$ = TaxCust(RecCnt).CSSN
          CASE 2                'spouses
            UBSearchN$ = TaxCust(RecCnt).SSSN
          END SELECT
          SSNOk = True
          FOR DigitCnt = 1 TO 9
            ThisDigit$ = MID$(Search$, DigitCnt, 1)
            IF ThisDigit$ = " " THEN
              'assume a blank in the ssn search is any digit
            ELSE
              SSNDigit$ = MID$(UBSearchN$, DigitCnt, 1)
              IF SSNDigit$ <> ThisDigit$ THEN
                SSNOk = False
                EXIT FOR
              END IF
            END IF
          NEXT
          IF SSNOk THEN
            GOSUB CustLoadEM2
            EXIT FOR
          END IF
        NEXT
      ELSE
        UBSearchN$ = LEFT$(TaxCust(RecCnt).SName, SearchLen)
        IF (Search$ = UBSearchN$) THEN
          GOSUB CustLoadEM2
        END IF
      END IF
DelSkip3:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  END IF
  
  FClose TaxFile
  
  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2
  ELSE
    
    SortT Mchoice(1), DCnt, Direction%, LEN(Mchoice(1).V), 0, 18
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Last           First        City           SSN         Acct"
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
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(Mchoice(Choice).V, 61, 4))
      IF ChkBalFlag THEN
        FOpenS TaxCustFile, TaxFile   'open data file
        FGetRTA TaxFile, TaxCust(1), RecNo&, TaxCustLen
        FClose TaxFile
        IF DoesCustOwe%(TaxCust(1)) THEN
          CursorOff
          ClearBack
          DisplayTaxScrn "ERRSCRN1"
          QPrintRC "This account HAS A BALANCE", 10, 27, -1
          QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
          WaitForAction
          RecNo& = 0
          ClearBack
          GOTO LoopRestart
        END IF
      END IF
    LOOP UNTIL RecNo& > 0
  END IF
  
ExitSearch2:
  RestScrn ScrnArray()
  
  ERASE ScrnArray, Mchoice, TaxCust
  
  EXIT SUB
  
CustLoadEM2:
  
  DCnt = DCnt + 1
  REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
  RSET Acct$ = QPTrim$(STR$(WhatRec&))
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LNAME), 14)
  MID$(Mchoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FName, 10)
  MID$(Mchoice(DCnt).V, 28, 11) = TaxCust(RecCnt).City
  IF NOT SSNFlag THEN
    WhoSSN = 1
  END IF
  IF LEN(QPTrim$(TaxCust(RecCnt).CSSN)) > 0 THEN
    SELECT CASE WhoSSN
    CASE 1
      MID$(Mchoice(DCnt).V, 41, 11) = TaxCust(RecCnt).CSSN
      MID$(Mchoice(DCnt).V, 50, 1) = "c"
    CASE 2
      MID$(Mchoice(DCnt).V, 41, 11) = TaxCust(RecCnt).SSSN
      MID$(Mchoice(DCnt).V, 50, 1) = "s"
    END SELECT
  END IF

  MID$(Mchoice(DCnt).V, 55, 5) = Acct$
  MID$(Mchoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN
  
  
END SUB

SUB Search4Cust1 (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)

  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(5)
  Search$ = QPTrim$(Search$)

  IF LastSEARCH$ <> QPTrim$(Search$) THEN
    LastSEARCH$ = QPTrim$(Search$)
    Choice = 1
  END IF

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  
  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14

  CustBlock = 32

  REDIM Mchoice(1 TO 1) AS FLen2
  REDIM TaxCust(1 TO CustBlock) AS TaxCustType

  TaxCustLen = LEN(TaxCust(1))
  TaxTranRecLen = LEN(TaxTran(1))

  SearchLen = LEN(Search$)
  Match = False
IF TaxType% = 1 THEN
  FOpenS RETaxCustFile, TaxFile   'open data file
END IF
IF TaxType% = 2 THEN
  FOpenS PPTaxCustFile, TaxFile   'open data file
END IF

  TaxTran = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  NumOfCust& = FLof&(TaxFile) / TaxCustLen
  NumChunks& = NumOfCust& \ CustBlock

  OddRecs& = NumOfCust& MOD CustBlock

  BlockSize& = (0& + TaxCustLen) * CustBlock
  '            ^^^^^ stops an overflow error
  'since TaxCustLen is an integer, basic will try to multiply to
  'an integer result. the above "0&" causes basic to convert to a long
  'then multiply'

  ' Find matching record
  FOR CCnt& = 1 TO NumChunks&
    FGetA TaxFile, TaxCust(1), BlockSize&
   FOR RecCnt = 1 TO CustBlock
      WhatRec& = ((CCnt& - 1) * CustBlock) + RecCnt
        UBSearchN$ = LEFT$(TaxCust(RecCnt).SName, SearchLen)
        PrevTranRec& = TaxCust(RecCnt).LastTrans

  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET TaxTran, PrevTranRec&, TaxTran(1)
      FoundTaxBill = INSTR(TaxTran(1).Description, Search$)
      IF TaxTran(1).TranType <> 1 THEN FoundTaxBill = 0
      
      IF FoundTaxBill THEN EXIT DO
      PrevTranRec& = TaxTran(1).LastTrans
    LOOP
  END IF

        IF (FoundTaxBill > 0) THEN
          GOSUB CustLoadEM2a
        END IF

DelSkip2a:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  NEXT

  IF OddRecs& > 0 THEN
    NextRec& = (NumChunks& * CustBlock) + 1
    RecCnt = 1
    FOR CCnt& = NextRec& TO NumOfCust&
      FGetRTA TaxFile, TaxCust(1), CCnt&, TaxCustLen
      WhatRec& = CCnt&

     PrevTranRec& = TaxCust(1).LastTrans

  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET TaxTran, PrevTranRec&, TaxTran(1)
      FoundTaxBill = INSTR(TaxTran(1).Description, Search$)
          IF TaxTran(1).TranType <> 1 THEN FoundTaxBill = 0

      IF FoundTaxBill THEN EXIT DO
      PrevTranRec& = TaxTran(1).LastTrans
    LOOP
  END IF
    
        IF (FoundTaxBill > 0) THEN
          GOSUB CustLoadEM2a
        END IF
      
DelSkip3a:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  END IF

  FClose TaxFile

  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2a
  ELSE

    SortT Mchoice(1), DCnt, Direction%, LEN(Mchoice(1).V), 0, 18

    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Last           First        City           SSN         Acct"
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

LoopRestarta:
    LOCATE Row, Col, 0
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(Mchoice(Choice).V, 61, 4))
      IF ChkBalFlag THEN
        FOpenS TaxCustFile, TaxFile   'open data file
        FGetRTA TaxFile, TaxCust(1), RecNo&, TaxCustLen
        FClose TaxFile
        IF DoesCustOwe%(TaxCust(1)) THEN
          CursorOff
          ClearBack
          DisplayTaxScrn "ERRSCRN1"
          QPrintRC "This account HAS A BALANCE", 10, 27, -1
          QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
          WaitForAction
          RecNo& = 0
          ClearBack
          GOTO LoopRestarta
        END IF
      END IF
    LOOP UNTIL RecNo& > 0
  END IF

ExitSearch2a:
  RestScrn ScrnArray()

  ERASE ScrnArray, Mchoice, TaxCust

  EXIT SUB

CustLoadEM2a:

  DCnt = DCnt + 1
  REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
  RSET Acct$ = QPTrim$(STR$(WhatRec&))
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LNAME), 14)
  MID$(Mchoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FName, 10)
  MID$(Mchoice(DCnt).V, 28, 11) = TaxCust(RecCnt).City
  MID$(Mchoice(DCnt).V, 55, 5) = Acct$
  MID$(Mchoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN

END SUB

SUB Search4Cust2 (SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType)

  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(5)
  Search$ = QPTrim$(SName$)

  IF LastSEARCH$ <> QPTrim$(Search$) THEN
    LastSEARCH$ = QPTrim$(Search$)
    Choice = 1
  END IF

  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14

  CustBlock = 32

  REDIM Mchoice(1 TO 1) AS FLen2
  REDIM TaxCust(1 TO CustBlock) AS TaxCustType

  TaxCustLen = LEN(TaxCust(1))

  SearchLen = LEN(Search$)
  Match = False

  IF TaxType% = 1 THEN
    FOpenS RETaxCustFile, TaxFile   'open data file
  END IF

  IF TaxType% = 2 THEN
    FOpenS PPTaxCustFile, TaxFile   'open data file
  END IF

  NumOfCust& = FLof&(TaxFile) / TaxCustLen
  NumChunks& = NumOfCust& \ CustBlock

  OddRecs& = NumOfCust& MOD CustBlock

  BlockSize& = (0& + TaxCustLen) * CustBlock
  '            ^^^^^ stops an overflow error
  'since TaxCustLen is an integer, basic will try to multiply to
  'an integer result. the above "0&" causes basic to convert to a long
  'then multiply'

  ' Find matching record
  FOR CCnt& = 1 TO NumChunks&
   FGetA TaxFile, TaxCust(1), BlockSize&
   FOR RecCnt = 1 TO CustBlock
      WhatRec& = ((CCnt& - 1) * CustBlock) + RecCnt
       FindAddress = INSTR(TaxCust(RecCnt).Addr1, Search$)
       IF FindAddress = 0 THEN
        FindAddress = INSTR(TaxCust(RecCnt).Addr2, Search$)
       END IF
       IF (FindAddress > 0) THEN
         GOSUB CustLoadEM2b
       END IF
DelSkip2b:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  NEXT

  IF OddRecs& > 0 THEN
    NextRec& = (NumChunks& * CustBlock) + 1
    RecCnt = 1
    FOR CCnt& = NextRec& TO NumOfCust&
      FGetRTA TaxFile, TaxCust(1), CCnt&, TaxCustLen
      WhatRec& = CCnt&
      FindAddress = INSTR(TaxCust(1).Addr1, Search$)
      IF FindAddress = 0 THEN
        FindAddress = INSTR(TaxCust(1).Addr2, Search$)
      END IF
      IF (FindAddress > 0) THEN
        GOSUB CustLoadEM2b
      END IF
DelSkip3b:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  END IF

  FClose TaxFile

  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2b
  ELSE

    SortT Mchoice(1), DCnt, Direction%, LEN(Mchoice(1).V), 0, 18

    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Last           First        Address                  Acct"
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

LoopRestartb:
    LOCATE Row, Col, 0
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(Mchoice(Choice).V, 61, 4))
      IF ChkBalFlag THEN
        FOpenS TaxCustFile, TaxFile   'open data file
        FGetRTA TaxFile, TaxCust(1), RecNo&, TaxCustLen
        FClose TaxFile
        IF DoesCustOwe%(TaxCust(1)) THEN
          CursorOff
          ClearBack
          DisplayTaxScrn "ERRSCRN1"
          QPrintRC "This account HAS A BALANCE", 10, 27, -1
          QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
          WaitForAction
          RecNo& = 0
          ClearBack
          GOTO LoopRestartb
        END IF
      END IF
    LOOP UNTIL RecNo& > 0
  END IF

ExitSearch2b:
  RestScrn ScrnArray()

  ERASE ScrnArray, Mchoice, TaxCust

  EXIT SUB

CustLoadEM2b:

  DCnt = DCnt + 1
  REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
  RSET Acct$ = QPTrim$(STR$(WhatRec&))
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LNAME), 14)
  MID$(Mchoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FName, 10)
  IF LEN(QPTrim$(TaxCust(RecCnt).Addr2)) = 0 THEN
  MID$(Mchoice(DCnt).V, 28, 11) = TaxCust(RecCnt).Addr1
  ELSE
  MID$(Mchoice(DCnt).V, 28, 11) = TaxCust(RecCnt).Addr2
  END IF
  MID$(Mchoice(DCnt).V, 55, 5) = Acct$
  MID$(Mchoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN

END SUB

SUB Search4Cust3 (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
  STATIC Choice, LastSEARCH$
  Acct$ = SPACE$(5)
  Search$ = QPTrim$(Search$)

  IF LastSEARCH$ <> QPTrim$(Search$) THEN
    LastSEARCH$ = QPTrim$(Search$)
    Choice = 1
  END IF

  
  REDIM PropertyRec(1)  AS PropertyRecType
  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14

  CustBlock = 32

  REDIM Mchoice(1 TO 1) AS FLen2
  REDIM TaxCust(1 TO CustBlock) AS TaxCustType

  TaxCustLen = LEN(TaxCust(1))
  

  SearchLen = LEN(Search$)
  Match = False
  IF TaxType% = 1 THEN
    FOpenS RETaxCustFile, TaxFile   'open data file
    PropTaxFile = FREEFILE
    OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
    NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
  END IF

  IF TaxType% = 2 THEN
    FOpenS PPTaxCustFile, TaxFile   'open data file
  END IF
  
  NumOfCust& = FLof&(TaxFile) / TaxCustLen
  NumChunks& = NumOfCust& \ CustBlock

  OddRecs& = NumOfCust& MOD CustBlock

  BlockSize& = (0& + TaxCustLen) * CustBlock
  '            ^^^^^ stops an overflow error
  'since TaxCustLen is an integer, basic will try to multiply to
  'an integer result. the above "0&" causes basic to convert to a long
  'then multiply'

  ' Find matching record
  FOR CCnt& = 1 TO NumChunks&
    FGetA TaxFile, TaxCust(1), BlockSize&
   FOR RecCnt = 1 TO CustBlock
      WhatRec& = ((CCnt& - 1) * CustBlock) + RecCnt
        PrevTranRec& = TaxCust(RecCnt).FirstPropRec

  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET PropTaxFile, PrevTranRec&, PropertyRec(1)
      FoundTaxBill = INSTR(PropertyRec(1).REALPIN, Search$)
      IF FoundTaxBill THEN EXIT DO
      PrevTranRec& = PropertyRec(1).NextRec
    LOOP
  END IF

        IF (FoundTaxBill > 0) THEN
          GOSUB CustLoadEM2c
        END IF

DelSkip2c:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  NEXT

  IF OddRecs& > 0 THEN
    NextRec& = (NumChunks& * CustBlock) + 1
    RecCnt = 1
    FOR CCnt& = NextRec& TO NumOfCust&
      FGetRTA TaxFile, TaxCust(1), CCnt&, TaxCustLen
      WhatRec& = CCnt&

     PrevTranRec& = TaxCust(1).FirstPropRec

  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
     GET PropTaxFile, PrevTranRec&, PropertyRec(1)
      FoundTaxBill = INSTR(PropertyRec(1).REALPIN, Search$)
      IF FoundTaxBill THEN EXIT DO
      PrevTranRec& = PropertyRec(1).NextRec
    LOOP
  END IF

        IF (FoundTaxBill > 0) THEN
          GOSUB CustLoadEM2c
        END IF

DelSkip3c:
      ShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  END IF

  FClose TaxFile
   CLOSE PropTaxFile

  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2c
  ELSE

    SortT Mchoice(1), DCnt, Direction%, LEN(Mchoice(1).V), 0, 18

    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Last           First        City           SSN         Acct"
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

LoopRestartC:
    LOCATE Row, Col, 0
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(Mchoice(Choice).V, 61, 4))
      IF ChkBalFlag THEN
        FOpenS TaxCustFile, TaxFile   'open data file
        FGetRTA TaxFile, TaxCust(1), RecNo&, TaxCustLen
        FClose TaxFile
        IF DoesCustOwe%(TaxCust(1)) THEN
          CursorOff
          ClearBack
          DisplayTaxScrn "ERRSCRN1"
          QPrintRC "This account HAS A BALANCE", 10, 27, -1
          QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
          WaitForAction
          RecNo& = 0
          ClearBack
          GOTO LoopRestartC
        END IF
      END IF
    LOOP UNTIL RecNo& > 0
  END IF

ExitSearch2c:
  RestScrn ScrnArray()

  ERASE ScrnArray, Mchoice, TaxCust

  EXIT SUB

CustLoadEM2c:

  DCnt = DCnt + 1
  REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
  RSET Acct$ = QPTrim$(STR$(WhatRec&))
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LNAME), 14)
  MID$(Mchoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FName, 10)
  MID$(Mchoice(DCnt).V, 28, 11) = TaxCust(RecCnt).City
  MID$(Mchoice(DCnt).V, 55, 5) = Acct$
  MID$(Mchoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN

END SUB

SUB Search4Cust4 (SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType)

  STATIC Choice, LastSEARCH$

  Acct$ = SPACE$(5)
  SName$ = QPTrim$(SName$)
  IF LEN(SName$) = 0 THEN
    RecNo& = 0
    EXIT SUB
  ELSE
    SearchAcct& = VAL(SName$)

  END IF


  'IF LastSEARCH$ <> QPTrim$(Search$) THEN
  '  LastSEARCH$ = QPTrim$(Search$)
    Choice = 1
  'END IF

  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14

  CustBlock = 32

  REDIM Mchoice(1 TO 1) AS FLen2
  REDIM TaxCust(1 TO 1) AS TaxCustType

  TaxCustLen = LEN(TaxCust(1))

  SearchLen = LEN(Search$)
  Match = False

  IF TaxType% = 1 THEN
    FOpenS RETaxCustFile, TaxFile   'open data file
  END IF

  IF TaxType% = 2 THEN
    FOpenS PPTaxCustFile, TaxFile   'open data file
  END IF

  NumOfCust& = FLof&(TaxFile) / TaxCustLen

  NextRec& = 1
  RecCnt = 1
  FOR CCnt& = NextRec& TO NumOfCust&
    FGetRTA TaxFile, TaxCust(1), CCnt&, TaxCustLen
    WhatRec& = CCnt&
    IF SearchAcct& = TaxCust(1).CountyAcct THEN
      GOSUB CustLoadEM2D
    END IF
'DelSkip3b:
    ShowSearchPCT CCnt&, NumOfCust&
  NEXT

  FClose TaxFile

  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2D
  ELSE

    'SortT Mchoice(1), DCnt, Direction%, LEN(Mchoice(1).V), 0, 18

    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Last           First        County Acct#             Acct"
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

LoopRestartD:
    LOCATE Row, Col, 0
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(Mchoice(Choice).V, 61, 4))
      IF ChkBalFlag THEN
        FOpenS TaxCustFile, TaxFile   'open data file
        FGetRTA TaxFile, TaxCust(1), RecNo&, TaxCustLen
        FClose TaxFile
        IF DoesCustOwe%(TaxCust(1)) THEN
          CursorOff
          ClearBack
          DisplayTaxScrn "ERRSCRN1"
          QPrintRC "This account HAS A BALANCE", 10, 27, -1
          QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
          WaitForAction
          RecNo& = 0
          ClearBack
          GOTO LoopRestartD
        END IF
      END IF
    LOOP UNTIL RecNo& > 0
  END IF

ExitSearch2D:
  RestScrn ScrnArray()

  ERASE ScrnArray, Mchoice, TaxCust

  EXIT SUB

CustLoadEM2D:

  DCnt = DCnt + 1
  REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
  RSET Acct$ = QPTrim$(STR$(WhatRec&))
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LNAME), 14)
  MID$(Mchoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FName, 10)
'here
  MID$(Mchoice(DCnt).V, 28, 11) = SName$
  MID$(Mchoice(DCnt).V, 55, 5) = Acct$
  MID$(Mchoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN

END SUB

SUB ShowCustHistory (CustRec&, TaxType%)
  
  u$ = CHR$(24)
  d$ = CHR$(25)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  DisplayTaxScrn "LOADHIST"
  
  REDIM Mchoice(1 TO 1) AS FLen2

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  REDIM TaxCustRec(1) AS TaxCustType
  
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
      DCnt = DCnt + 1
      REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
      GET TaxTran, PrevTranRec&, TaxTran(1)
      LSET Mchoice(DCnt).V = Num2Date(TaxTran(1).TransDate)
      GOSUB GetTransType
      MID$(Mchoice(DCnt).V, 13) = TType$
      MID$(Mchoice(DCnt).V, 41) = FUsing(STR$(TaxTran(1).Amount), "#####.##")
      'this will show the actual trans number in the list
      MID$(Mchoice(DCnt).V, 52) = FUsing(STR$(PrevTranRec&), "######")
      'MID$(MChoice(DCnt).V, 52) = FUsing(STR$(TaxTran(1).RunBalance), "#####.##")
      MID$(Mchoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      PrevTranRec& = TaxTran(1).LastTrans
    LOOP
    
    CLOSE TaxTran
    
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
    Action = 1
  END IF
  
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
  GOSUB GetTransType
  
  CLOSE TaxTran  'NOTE: Close must be after GetTransType

  DisplayTaxScrn "TRDETAIL"
  QPrintRC Num2Date(TaxTran(1).TransDate), 7, 28, -1

  QPrintRC TaxTran(1).Description, 8, 28, -1

  QPrintRC FUsing$(STR$(TaxTran(1).Amount), "#####.##"), 9, 28, -1
  QPrintRC BillType$, 10, 28, -1
  QPrintRC TaxYear$, 11, 60, -1
  QPrintRC Post2GL$, 11, 28, -1
  QPrintRC FUsing$(STR$(Principle#), "#######.##"), 13, 28, -1
  QPrintRC FUsing$(STR$(Interest#), "#######.##"), 14, 28, -1
  QPrintRC FUsing$(STR$(Penalty#), "#######.##"), 15, 28, -1
  QPrintRC FUsing$(STR$(AdCost#), "#######.##"), 16, 28, -1

  QPrintRC FUsing$(STR$(PrinciplePD#), "#######.##"), 13, 48, -1
  QPrintRC FUsing$(STR$(InterestPd#), "#######.##"), 14, 48, -1
  QPrintRC FUsing$(STR$(PenaltyPd#), "#######.##"), 15, 48, -1
  QPrintRC FUsing$(STR$(AdCostPD#), "#######.##"), 16, 48, -1
  
  QPrintRC "BT:" + FUsing$(STR$(TaxTran(1).BelongTo), "#####"), 7, 46, -1

  IF TaxTran(1).TranType = 1 THEN
   QPrintRC "Balance: ", 17, 19, -1
   QPrintRC FUsing$(STR$(UnPaidPrinciple#), "#######.##"), 13, 28, -1
   QPrintRC FUsing$(STR$(UnPaidPrinciple# + Interest# + Penalty# + AdCost#), "#######.##"), 17, 28, -1
  END IF


  WaitForAction
  RestScrn TempScrn2()
  GOTO HistTop
  
GetTransType:

  Principle# = 0
  Interest# = 0
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

    IF TaxType% = 1 THEN
     Principle# = Round#(TaxTran(1).Revenue.Principle1)
     UnPaidPrinciple# = Round#(TaxTran(1).Revenue.Principle1 - TaxTran(1).Revenue.Principle1Pd)
    ELSE
     Principle# = Round#(TaxTran(1).Revenue.Principle1 + TaxTran(1).Revenue.Principle2 + TaxTran(1).Revenue.Principle3)
     Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4 + TaxTran(1).Revenue.Principle5)
     UnPaidPrinciple# = Round#(Principle# - (TaxTran(1).Revenue.Principle1Pd + TaxTran(1).Revenue.Principle2Pd + TaxTran(1).Revenue.Principle3Pd + TaxTran(1).Revenue.Principle4Pd + TaxTran(1).Revenue.Principle5Pd))
    END IF

    'Fiqure Penalty & Interest Here
     Penalty# = TaxTran(1).Revenue.Penalty# - TaxTran(1).Revenue.PenaltyPd
     Interest# = TaxTran(1).Revenue.Interest# - TaxTran(1).Revenue.InterestPd
 
  CASE 2
    BillType$ = "Payment"
    IF TaxType% = 1 THEN
     Principle# = Round#(TaxTran(1).Revenue.Principle1Pd)
     Penalty# = TaxTran(1).Revenue.PenaltyPd
     Interest# = TaxTran(1).Revenue.InterestPd
   ELSE
     Principle# = Round#(TaxTran(1).Revenue.Principle1Pd + TaxTran(1).Revenue.Principle2Pd + TaxTran(1).Revenue.Principle3Pd)
     Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4Pd + TaxTran(1).Revenue.Principle5Pd)
     Penalty# = TaxTran(1).Revenue.PenaltyPd
     Interest# = TaxTran(1).Revenue.InterestPd

    END IF

  CASE 3
    BillType$ = "Release"
  CASE 4
    BillType$ = "Interest"
    Interest# = TaxTran(1).Revenue.Interest#
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

  PrinciplePD# = Round#(TaxTran(1).Revenue.Principle1Pd + TaxTran(1).Revenue.Principle2Pd + TaxTran(1).Revenue.Principle3Pd)
  PrinciplePD# = Round#(PrinciplePD# + TaxTran(1).Revenue.Principle4Pd + TaxTran(1).Revenue.Principle5Pd)
  InterestPd# = TaxTran(1).Revenue.InterestPd#
  AdCostPD# = TaxTran(1).Revenue.CollectionPd

'  TaxTran(1).Revenue.Interest# = 57.69
'  PUT TaxTran, TransRecNum&, TaxTran(1)

RETURN
  

END SUB

SUB ShowCustPersList (CustRec&, TaxType%)

  u$ = CHR$(24)
  d$ = CHR$(25)

  REDIM TempScrn(0)
  SaveScrn TempScrn()


  REDIM Mchoice(1 TO 1) AS FLen2


  REDIM TaxCustRec(1) AS TaxCustType
  REDIM PersRec(1) AS PersonalRecType
  REDIM PersonalRec(1) AS PersonalRecType

  TaxCustRecLen = LEN(TaxCustRec(1))


  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  GET TaxFile, ABS(CustRec&), TaxCustRec(1)
  PropRec& = TaxCustRec(1).FirstPersRec
  CLOSE TaxFile
  
Topb:

  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPropRecs = LOF(PersTaxFile) / LEN(PersRec(1))


  IF PropRec& > 0 THEN
    DO WHILE PropRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
      GET PersTaxFile, PropRec&, PersonalRec(1)
      PropDesc$ = QPTrim$(PersonalRec(1).Desc4)
      PropDesc$ = PropDesc$ + " " + QPTrim$(PersonalRec(1).DESC2)
      LSET Mchoice(DCnt).V = QPTrim$(PersonalRec(1).DESC1)
      MID$(Mchoice(DCnt).V, 20) = LEFT$(PropDesc$, 45)
     
      PropRec& = PersonalRec(1).NextRec
    LOOP

    CLOSE PersTaxFile

    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 74, 8
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()

HistTopb:

    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  VIN                 Property Description"

    '--Find max menu width
    '--Center Menu within Screen

    Col = ((80 - 60) \ 2) - 1

    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20

    WazzWind BoxBot + 2, Col, BoxBot + 5, MaxLen + 3 + Col, 10, 4, True
    'QPrintRC u$, BoxBot + 3, Col + 11, 14
    'QPrintRC d$, BoxBot + 3, Col + 13, 14

    QPrintRC "Total: " + STR$(DCnt), BoxBot + 3, Col + 3, 15
    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, Col + 33, 15
    QPrintRC "ESC", BoxBot + 3, Col + 43, 14


    QPrintRC Title$, Row - 1, Col, 112

    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8

    DO
      LOCATE Row, Col, 0
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TempScrn()
        EXIT DO 'choice = 0
      END IF
    LOOP
   END IF

  RestScrn TempScrn()
  ERASE Mchoice
  ERASE TempScrn, TaxCustRec

EXIT SUB



END SUB

SUB ShowCustPropList (CustRec&, TaxType%)

  u$ = CHR$(24)
  d$ = CHR$(25)

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  
  REDIM Mchoice(1 TO 1) AS FLen2

  
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM PropertyRec(1) AS PropertyRecType

  TaxCustRecLen = LEN(TaxCustRec(1))
  

  TaxFile = FREEFILE
  IF TaxType% = 1 THEN
   OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  END IF
  GET TaxFile, ABS(CustRec&), TaxCustRec(1)
  PropRec& = TaxCustRec(1).FirstPropRec
  CLOSE TaxFile


Topa:

  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))

  
  IF PropRec& > 0 THEN
    DO WHILE PropRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE Mchoice(1 TO DCnt) AS FLen2
      GET PropTaxFile, PropRec&, PropertyRec(1)
      PropDesc$ = QPTrim$(PropertyRec(1).PROPNOT1)
      PropDesc$ = PropDesc$ + " " + QPTrim$(PropertyRec(1).PROPNOT2)
      LSET Mchoice(DCnt).V = QPTrim$(PropertyRec(1).REALPIN)
      MID$(Mchoice(DCnt).V, 20) = LEFT$(PropDesc$, 34)
      MID$(Mchoice(DCnt).V, 56) = QPTrim$(PropertyRec(1).MORTCODE)
      PropRec& = PropertyRec(1).NextRec
    LOOP

    CLOSE PropTaxFile

    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 74, 8
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()

HistTopa:

    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something

    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF

    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  G-Pin            Description                       Mort Code        "

    '--Find max menu width
    '--Center Menu within Screen

    Col = ((80 - 60) \ 2) - 1

    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20

    WazzWind BoxBot + 2, Col, BoxBot + 5, MaxLen + 3 + Col, 10, 4, True
    'QPrintRC u$, BoxBot + 3, Col + 11, 14
    'QPrintRC d$, BoxBot + 3, Col + 13, 14

    QPrintRC "Total: " + STR$(DCnt), BoxBot + 3, Col + 3, 15
    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, Col + 33, 15
    QPrintRC "ESC", BoxBot + 3, Col + 43, 14
    

    QPrintRC Title$, Row - 1, Col, 112

    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8

    DO
      LOCATE Row, Col, 0
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TempScrn()
        EXIT DO 'choice = 0
      END IF
    LOOP
   END IF

  RestScrn TempScrn()
  ERASE Mchoice
  ERASE TempScrn, TaxCustRec

EXIT SUB



END SUB

SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  QPrintRC PctC(1), 13, 39, Cnf.HiLite
END SUB

SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrL$(INT((RecNo& / NumOfRecs&) * 100))
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)
  
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  CursorOff
  'BlockClear
  DisplayTaxScrn "PRORPT"
  
  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
  
END SUB

SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo& / NumOfRecs&) * 100))
  HideCursor
  QPrintRC PctC(1), 12, 38, 15
  ShowCursor
END SUB

SUB smallpause
'  St# = TIMER
'  St# = St# + .2
'  DO: LOOP UNTIL TIMER > St#
END SUB

'********** Unique.Bas - provides a unique file name
'Copyright (c) 1989 Ethan Winer
'NOTE: Although the manual shows no arguments to the Unique$ function, we
'have added the capability to specify a path name as an argument.  This lets
'you create a unique file name, and also be sure a file with that name does
'not exist in any given directory.
FUNCTION Unique$ (Path$)
  
  IF LEN(Path$) AND RIGHT$(Path$, 1) <> "\" THEN Path$ = Path$ + "\"
  Seed& = ABS(TIMER)            'use the TIMER as a seed
  DO
    TempName$ = Path$ + MID$(STR$(Seed&), 2)    'make a string out of it
    TempName$ = TempName$ + ".RPT"
    Seed& = Seed& + 1           'increment for next time
  LOOP UNTIL NOT Exist%(TempName$)              'loop and try another name
  Unique$ = TempName$           'this is the function output
  
END FUNCTION

SUB UpDateTicklerFile
  ThisMonth$ = LEFT$(DATE$, 2)
  TickFile = FREEFILE
  OPEN "TAXINTCK.DAT" FOR OUTPUT AS #TickFile
  PRINT #TickFile, ThisMonth$
  CLOSE TickFile
END SUB

