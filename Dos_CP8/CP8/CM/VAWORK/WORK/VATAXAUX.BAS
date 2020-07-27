DEFINT A-Z
DECLARE FUNCTION VAUnique$ (Path$)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Unique$ (Path$)
DECLARE FUNCTION VADoesCustOwe% (TaxCustRec AS ANY)
DECLARE FUNCTION VAGetTaxCustCnt& (TaxType%)
DECLARE FUNCTION VAIsCustDeleted% (AcctNum&, TaxType%)
DECLARE SUB ClearBack ()
DECLARE SUB CursorOff ()
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
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VADisplayTaxScrn (ScrnName$)
DECLARE SUB VASearch4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
DECLARE SUB VAShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
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

SUB VADisplayTaxScrn (ScrnName$)
  LibFile2Scrn "TAX", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  IF ErrorCode% <> 0 THEN
    PRINT "Screen Error: "; ScrnName$
    END
  END IF
END SUB

FUNCTION VADoesCustOwe% (TaxCustRec AS TaxCustType)
  
  VADoesCustOwe% = False        'assume the customer owes nothing
  
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
      VADoesCustOwe% = True
    END IF
  END IF
  
END FUNCTION

FUNCTION VAGetTaxCustCnt& (TaxType%)
  
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
  
  VAGetTaxCustCnt& = NumOfRecs&
  
END FUNCTION

FUNCTION VAIsCustDeleted (AcctNum&, TaxType%)
  VAIsCustDeleted = False       'assume they aren't deleted
  
  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  GET TaxFile, AcctNum&, TaxCust(1)
  CLOSE TaxFile
  
  IF TaxCust(1).Deleted <> 0 THEN
    VAIsCustDeleted = True
  END IF
  
  ERASE TaxCust
  
END FUNCTION

SUB VALookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
  
LookUpTop:
  
  REDIM Hlp$(1 TO 3)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Search Name here."
  Hlp$(3) = "Enter all or part of the SSN to search for here." + CHR$(13)
  Hlp$(3) = Hlp$(3) + "NOTE: a blank space will match any digit. The" + CHR$(13)
  Hlp$(3) = Hlp$(3) + "Customer and Spouses SSN are searched."
  
  REDIM TaxCust(1) AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))
  
  SName$ = ""
  AcctNum& = 0
  LScrn = 2
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 3) AS STRING * 17
  
  MScrn = 3
  
  LText(1) = " Account Number:"
  LText(2) = "    Search Name:"
  LText(3) = "Social Security:"
  
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
  VADisplayTaxScrn ScrnName$
  
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
        IF AcctNum& < 1 OR AcctNum& > VAGetTaxCustCnt&(TaxType%) THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")
        ELSEIF VAIsCustDeleted(AcctNum&, TaxType%) THEN
          Ok = MsgBox%("TAX.QSL", "DELACCTN")
        ELSEIF ChkBalFlag THEN
          REDIM TaxCust(1) AS TaxCustType
          TaxCustLen = LEN(TaxCust(1))
          IF TaxType% = 1 THEN
            FOpenS RETaxCustFile, TaxFile       'open data file
          ELSE
            FOpenS PPTaxCustFile, TaxFile       'open data file
          END IF
          FGetRTA TaxFile, TaxCust(1), AcctNum&, TaxCustLen
          FClose TaxFile
          IF VADoesCustOwe%(TaxCust(1)) THEN
            SaveScrn ScrnArray2()
            CursorOff
            VADisplayTaxScrn "ERRSCRN1"
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
        VASearch4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, False, TaxType
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
        VASearch4Cust SName$, RecNo&, ChkBalFlag, CLSFlag, True, TaxType
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

SUB VASearch4Cust (Search$, RecNo&, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
  
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
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM TaxCust(1 TO CustBlock) AS TaxCustType
  
  TaxCustLen = LEN(TaxCust(1))
  
  SearchLen = LEN(Search$)
  Match = False
  IF TaxType% = 1 THEN
    FOpenS RETaxCustFile, TaxFile               'open data file
  END IF
  IF TaxType% = 2 THEN
    FOpenS PPTaxCustFile, TaxFile               'open data file
  END IF
  
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
      VAShowSearchPCT WhatRec&, NumOfCust&
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
      VAShowSearchPCT WhatRec&, NumOfCust&
    NEXT
  END IF
  
  FClose TaxFile
  
  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2
  ELSE
    
    SortT MChoice(1), DCnt, Direction%, LEN(MChoice(1).V), 0, 18
    
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
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
      IF ChkBalFlag THEN
        FOpenS TaxCustFile, TaxFile             'open data file
        FGetRTA TaxFile, TaxCust(1), RecNo&, TaxCustLen
        FClose TaxFile
        IF VADoesCustOwe%(TaxCust(1)) THEN
          CursorOff
          ClearBack
          VADisplayTaxScrn "ERRSCRN1"
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
  
  ERASE ScrnArray, MChoice, TaxCust
  
  EXIT SUB
  
CustLoadEM2:
  
  DCnt = DCnt + 1
  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
  RSET Acct$ = QPTrim$(STR$(WhatRec&))
  LSET MChoice(DCnt).V = LEFT$(QPTrim$(TaxCust(RecCnt).LName), 14)
  MID$(MChoice(DCnt).V, 16) = LEFT$(TaxCust(RecCnt).FNAME, 10)
  MID$(MChoice(DCnt).V, 28, 11) = TaxCust(RecCnt).CITY
  IF NOT SSNFlag THEN
    WhoSSN = 1
  END IF
  IF LEN(QPTrim$(TaxCust(RecCnt).CSSN)) > 0 THEN
    SELECT CASE WhoSSN
    CASE 1
      MID$(MChoice(DCnt).V, 41, 11) = TaxCust(RecCnt).CSSN
      MID$(MChoice(DCnt).V, 50, 1) = "c"
    CASE 2
      MID$(MChoice(DCnt).V, 41, 11) = TaxCust(RecCnt).SSSN
      MID$(MChoice(DCnt).V, 50, 1) = "s"
    END SELECT
  END IF
  
  MID$(MChoice(DCnt).V, 55, 5) = Acct$
  MID$(MChoice(DCnt).V, 61) = MKL$(WhatRec&)
RETURN
  
  
END SUB

SUB VAShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo& / NumOfRecs&) * 100))
  HideCursor
  QPrintRC PctC(1), 12, 38, 15
  ShowCursor
END SUB

'********** Unique.Bas - provides a unique file name
'Copyright (c) 1989 Ethan Winer
'NOTE: Although the manual shows no arguments to the Unique$ function, we
'have added the capability to specify a path name as an argument.  This lets
'you create a unique file name, and also be sure a file with that name does
'not exist in any given directory.
FUNCTION VAUnique$ (Path$)
  
  IF LEN(Path$) AND RIGHT$(Path$, 1) <> "\" THEN Path$ = Path$ + "\"
  Seed& = ABS(TIMER)            'use the TIMER as a seed
  DO
    TempName$ = Path$ + MID$(STR$(Seed&), 2)    'make a string out of it
    TempName$ = TempName$ + ".RPT"
    Seed& = Seed& + 1           'increment for next time
  LOOP UNTIL NOT Exist%(TempName$)              'loop and try another name
  VAUnique$ = TempName$           'this is the function output
  
END FUNCTION

