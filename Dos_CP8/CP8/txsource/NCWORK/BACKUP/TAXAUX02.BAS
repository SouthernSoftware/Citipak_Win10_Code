DEFINT A-Z
DECLARE SUB Search4Cust1 (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION DoesCustOwe% (TaxCustRec AS ANY)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION GetCustBalance# (RecNo&)
DECLARE FUNCTION GetCustName$ (CustRec&)
DECLARE FUNCTION GetTaxCustCnt& ()
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Unique$ (Path$)
DECLARE SUB ClearBack ()
DECLARE SUB CursorOff ()
DECLARE SUB DelPropAbstract (PropRecs() AS LONG, WhatProp%, CustRec&)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB HideCursor ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)


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

  Pers2Free& = WhatPers
  NumOfPers& = PersRecs(0)

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)

  FirstPers& = TaxCust(1).FirstPersRec

'First free the Personal in question
  GET PersFile, Pers2Free&, PersRec(1)
  PersRec(1).NextRec = 0
  PersRec(1).CustPin = 0
  PersRec(1).Deleted = True
  PUT PersFile, Pers2Free&, PersRec(1)
'Personal has been marked deleted

  IF NumOfPers& = 1 THEN               'if this was the cust's only Pers
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

  Prop2Free& = WhatProp
  NumOfProp& = PropRecs(0)

  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)

  FirstProp& = TaxCust(1).FirstPropRec

'First free the property in question
  GET PropFile, Prop2Free&, PropRec(1)
  PropRec(1).NextRec = 0
  PropRec(1).CustPin = 0
  PropRec(1).Deleted = True
  PUT PropFile, Prop2Free&, PropRec(1)
'property has been marked deleted

  IF NumOfProp& = 1 THEN               'if this was the cust's only prop
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

FUNCTION GetCustBalance# (RecNo&)

  REDIM TaxTran(1) AS TaxTransactionType
  REDIM TaxCustRec(1) AS TaxCustType

  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTran(1))

  TaxFile = FREEFILE
  OPEN "TaxCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  GET TaxFile, RecNo&, TaxCustRec(1)
  CLOSE TaxFile

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  PrevTranRec& = TaxCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET TaxTran, PrevTranRec&, TaxTran(1)
      SELECT CASE TaxTran(1).TranType
      CASE 1  'bill
        GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
      CASE 2  'payment
        TPaid# = Round#(TPaid# + TaxTran(1).Amount)
        GTPaid# = Round#(GTPaid# + TaxTran(1).Amount)
      CASE 3 'release

      CASE 4 'interest
        GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
      CASE 6  'collect/add cost
        GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
      CASE 7  'adjustment
        GTPaid# = Round#(GTPaid# + TaxTran(1).Amount)
      CASE 8  'misc cost
        GTOwed# = Round#(GTOwed# + TaxTran(1).Amount)
      CASE ELSE
        'BillType$ = "?????"
      END SELECT
      PrevTranRec& = TaxTran(1).LastTrans
    LOOP

    GetCustBalance# = Round#(GTOwed# - GTPaid#)
  ELSE
    GetCustBalance# = 0
  END IF

  CLOSE

END FUNCTION

FUNCTION GetCustName$ (CustRec&)

  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure

  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  GET TaxFile, CustRec&, TaxCust(1)
  CLOSE TaxFile
  GetCustName$ = QPTrim$(TaxCust(1).FNAME) + " " + QPTrim$(TaxCust(1).LName)

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
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
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
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
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

FUNCTION GetTaxCustCnt&
  
  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfRecs& = LOF(TaxFile) \ TaxCustLen
  CLOSE TaxFile
  
  ERASE TaxCust
  
  GetTaxCustCnt& = NumOfRecs&
  
END FUNCTION

FUNCTION IsCustDeleted (AcctNum&)
  IsCustDeleted = False         'assume they aren't deleted
  
  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  GET TaxFile, AcctNum&, TaxCust(1)
  CLOSE TaxFile
  
  IF TaxCust(1).Deleted <> 0 THEN
    IsCustDeleted = True
  END IF
  
  ERASE TaxCust
  
END FUNCTION

SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)

LookUpTop:

  REDIM Hlp$(1 TO 4)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Search Name here."
  Hlp$(3) = "Enter all or part of the SSN to search for here." + CHR$(13)
  Hlp$(3) = Hlp$(3) + "NOTE: a blank space will match any digit. The" + CHR$(13)
  Hlp$(3) = Hlp$(3) + "Customer and Spouses SSN are searched."
  Hlp$(4) = "Enter all or part of the PIN to search for here."
  REDIM TaxCust(1) AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))
  
  SName$ = ""
  AcctNum& = 0
  PIN$ = ""
  LScrn = 2
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 4) AS STRING * 17
  
  MScrn = 4
  
  LText(1) = " Account Number:"
  LText(2) = "    Search Name:"
  LText(3) = "Social Security:"
  LText(4) = "     PIN Number:"
  
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
        Fld(1).FType = 2
        LSET Form$(1, 0) = ""
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
        IF AcctNum& < 1 OR AcctNum& > GetTaxCustCnt& THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")
        ELSEIF IsCustDeleted(AcctNum&) THEN
          Ok = MsgBox%("TAX.QSL", "DELACCTN")
        ELSEIF ChkBalFlag THEN
          REDIM TaxCust(1) AS TaxCustType
          TaxCustLen = LEN(TaxCust(1))
          FOpenS TaxCustFile, TaxFile   'open data file
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
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, False
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 3
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, True
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      CASE 4
        SName$ = Form$(1, 0)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust1 SName$, RecNo&, ChkBalFlag, CLSFlag, True
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
  OPEN TaxCustFile FOR RANDOM AS CustFile LEN = TaxRecLen
  NumOfCRecs = LOF(CustFile) / TaxRecLen
  
  QPrintRC "Reading Customer Information", 11, 27, -1
  
  REDIM Array(1 TO NumOfCRecs) AS SortStruct
  FOR Cnt = 1 TO NumOfCRecs
    GET CustFile, Cnt, TaxCust(1)
    SELECT CASE IdxType
    CASE 1
      Array(Cnt).who = (LEFT$(QPTrim$(TaxCust(1).LName), 12) + LEFT$(QPTrim$(TaxCust(1).LName), 2))
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

SUB Search4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
  
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
  
  FOpenS TaxCustFile, TaxFile   'open data file
  
  NumOfCust& = FLof&(TaxFile) / TaxCustLen
  NumChunks& = NumOfCust& / CustBlock
  
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
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LName), 14)
  MID$(Mchoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FNAME, 10)
  MID$(Mchoice(DCnt).V, 28, 11) = TaxCust(RecCnt).CITY
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

SUB Search4Cust1 (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%)
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

  FOpenS TaxCustFile, TaxFile   'open data file
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))



  NumOfCust& = FLof&(TaxFile) / TaxCustLen
  NumChunks& = NumOfCust& / CustBlock

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

LoopRestartc:
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
          GOTO LoopRestartc
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
  LSET Mchoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LName), 14)
  MID$(Mchoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FNAME, 10)
  MID$(Mchoice(DCnt).V, 28, 11) = TaxCust(RecCnt).CITY
  MID$(Mchoice(DCnt).V, 55, 5) = Acct$
  MID$(Mchoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN


END SUB

SUB ShowCustHistory (CustRec&)
  
  IF CustRec& < 0 THEN
    AdjShadow = True
    CustRec& = ABS(CustRec&)
  ELSE
    AdjShadow = False
  END IF

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
  OPEN "TaxCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  GET TaxFile, CustRec&, TaxCustRec(1)
  CLOSE TaxFile
  

  CurBal# = GetCustBalance#(CustRec&)
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

    IF AdjShadow THEN
      MPaintBox 6, 4, 18, 76, 8
    ELSE
      MPaintBox 3, 5, 22, 74, 8
    END IF
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    Bal$ = Title$
    LSET Title$ = "   Date            Description              Amount             "
      LSET Bal$ = "   Total Balance:" + FUsing$(STR$(CurBal#), "#######.##")
    '--Find max menu width
    '--Center Menu within Screen
    
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    QPrintRC Bal$, Row - 2, Col, 112
    QPrintRC Title$, Row - 1, Col, 112
    
    WazzWind BoxBot + 2, Col, BoxBot + 5, MaxLen + 3 + Col, 10, 4, True
    QPrintRC "  Use:  " + u$ + "-" + d$ + " to select.", BoxBot + 3, Col + 3, 15
    QPrintRC u$, BoxBot + 3, Col + 11, 14
    QPrintRC d$, BoxBot + 3, Col + 13, 14
    
    QPrintRC "Total: " + STR$(DCnt), BoxBot + 4, Col + 3, 15
    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, Col + 33, 15
    QPrintRC "        [ENTER] for detail.", BoxBot + 4, Col + 33, 15
    QPrintRC "ESC", BoxBot + 3, Col + 43, 14
    QPrintRC "ENTER", BoxBot + 4, Col + 42, 14
    
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    
    DO
      LOCATE Row, Col, 0
      VertMenuT2 Mchoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        'RestScrn TempScrn()
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        RestScrn TempScrn()
        GOTO ShowTransDetail
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE TaxTran
    Ok = MsgBox%("Tax.QSL", "NOCTRANS")
    'RestScrn TempScrn()
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
    Principle# = Round#(TaxTran(1).Revenue.Principle1 + TaxTran(1).Revenue.Principle2 + TaxTran(1).Revenue.Principle3)
    Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4 + TaxTran(1).Revenue.Principle5)
  CASE 2
    BillType$ = "Payment"
    Principle# = Round#(TaxTran(1).Revenue.Principle1Pd + TaxTran(1).Revenue.Principle2Pd + TaxTran(1).Revenue.Principle3Pd)
    Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4Pd + TaxTran(1).Revenue.Principle5Pd)
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
    BillType$ = "Collection/Ad Cost"
  CASE 7
    BillType$ = "Adjustment"
    Principle# = Round#(TaxTran(1).Revenue.Principle1Pd + TaxTran(1).Revenue.Principle2Pd + TaxTran(1).Revenue.Principle3Pd)
    Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4Pd + TaxTran(1).Revenue.Principle5Pd)
    Interest# = TaxTran(1).Revenue.InterestPd#
  CASE ELSE
    BillType$ = "?????"
  END SELECT

RETURN
  

END SUB

SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
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

SUB SmallPause
  St# = TIMER + .6
  DO
  LOOP UNTIL TIMER > St#
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

