DEFINT A-Z
DECLARE SUB PrintNoticeDates ()
DECLARE SUB MakePostalIndex2 (IndexText$)
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE FUNCTION MakeDay$ (DayNum%)
DECLARE SUB SaveRateTblRec (Form$(), RecNo%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Num2Day% (DayNum%)
DECLARE FUNCTION Date2Num% (TDate$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE FUNCTION Round# (n#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, Kolor%)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MakeTempIndex (Indextype%)
DECLARE SUB MakeZipCodeIndex ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Key2Stuff$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WaitForAction ()
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  ''$INCLUDE: 'ubcust.bi'
  ''$INCLUDE: 'ubloc.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubsetbal.BI'
  
  CONST False = 0, True = NOT False
  
  'PRINT Num2Day(Date2Num("01-16-1997"))
  
  'Date2Num (Form$(1, 0))
  CLS
  PrintNoticeDates
  
ExitPrint:
  
    ClearScrn
    HideCursor
    END

FUNCTION MakeDay$ (DayNum)
  SELECT CASE DayNum
  CASE 1
    MakeDay$ = "Sunday"
  CASE 2
    MakeDay$ = "Monday"
  CASE 3
    MakeDay$ = "Tuesday"
  CASE 4
    MakeDay$ = "Wednesday"
  CASE 5
    MakeDay$ = "Thursday"
  CASE 6
    MakeDay$ = "Friday"
  CASE 7
    MakeDay$ = "Saturday"
  END SELECT
  
END FUNCTION

FUNCTION MakeMonth$ (TDate$)
  Month = VAL(LEFT$(TDate$, 2))
  SELECT CASE Month
  CASE 1
    MakeMonth$ = "January"
  CASE 2
    MakeMonth$ = "February"
  CASE 3
    MakeMonth$ = "March"
  CASE 4
    MakeMonth$ = "April"
  CASE 5
    MakeMonth$ = "May"
  CASE 6
    MakeMonth$ = "June"
  CASE 7
    MakeMonth$ = "July"
  CASE 8
    MakeMonth$ = "August"
  CASE 9
    MakeMonth$ = "September"
  CASE 10
    MakeMonth$ = "October"
  CASE 11
    MakeMonth$ = "November"
  CASE 12
    MakeMonth$ = "December"
  END SELECT
END FUNCTION

SUB MakePostalIndex2 (IndexText$)
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs = GetNumOfCust%
  
  REDIM PostalIndex(1 TO NumCustRecs)  AS UBPostalIndexType
  IndexRecLen = LEN(PostalIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs
    GET CHandle, Cnt, UBCustRec(1)
    PostalIndex(Cnt).ZIPCODE = UBCustRec(1).ZIPCODE
    PostalIndex(Cnt).Route = UBCustRec(1).POSTRTE
    PostalIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT PostalIndex(1), NumCustRecs, 0, 16, 10, 4
  
  SortT PostalIndex(1), NumCustRecs, 0, 16, 0, 14
  
  'SortT (Array(1), NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IHandle = FREEFILE
  
  FCreate TempIndexName
  
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO NumCustRecs
    Prec& = PostalIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, PostalIndex
  
END SUB

SUB PrintNoticeDates

  REDIM Msgs$(5)

  REDIM ScrnArray(0)
  
  LibName$ = "UBNOTPRN"
  ScrnName$ = "UBNOTPRN"
  NumScrns = 1
  
  '--define the multi-choice fields
  SHARED Choice$()

  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  Scr = 1
  BlockClear

  'DisplayUBScrn ScrnName$
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute, ErrorCode

  FirstTime = True
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False

      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      SaveScrn ScrnArray()
      'Section to check for customer modifications
      PrintRptFile " ", "UBLALIGN.DAT", 1, RetCode, 4
      RestScrn ScrnArray()

    CASE -68    'F10Key
      FOR Cnt = 1 TO 5
        Msgs$(Cnt) = Form$(Cnt, 0)
      NEXT
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 21 TO 33           '--cancel button
          PressButton 27, 16, 21, 33
        CASE 34 TO 48           '--Alignment
          PressButton F5KEY, 16, 34, 48
        CASE 49 TO 60           '--Save Button
          PressButton -68, 16, 49, 60
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  IF Frm(1).KeyCode = 27 THEN
    GOTO ExitEdit
  ELSE
    GOTO StartHere:
  END IF


StartHere:

  'load setup file
  'Section to check for customer modifications
  
  IdxTypeText$ = "Zipcode"
  MakePostalIndex2 IdxTypeText$
  IndexName$ = TempIndexName

  OKFlag = True

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  TBooks = 0

  NumOfRecs = FileSize(IndexName$) \ 4
  REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IndexArray(1), 4, NumOfRecs

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBNOTPRN.RPT" FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn "Processing Notices"

  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, IndexArray(Cnt).RecNum, UBCustRec(1)
    IF NOT UBCustRec(1).DelFlag THEN
      Location$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
      Acct$ = QPTrim$(STR$(IndexArray(Cnt).RecNum))
      AcctNo& = IndexArray(Cnt).RecNum
      Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
      Ziplen = LEN(Zip$)
      SELECT CASE Ziplen
      CASE 9, 10
        Zip$ = LEFT$(Zip$, 5) + "-" + RIGHT$(Zip$, 4)
      CASE ELSE
        Zip$ = LEFT$(Zip$, 5)
      END SELECT

      IF UBCustRec(1).STATUS = "A" THEN
        Print1 = Print1 + 1
        GOSUB PrintThemOne
      END IF

      '***********************************
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
SkipEm:
    ShowPctComp Cnt, NumOfRecs
'    IF Print1 > 2 THEN EXIT FOR
  NEXT

  CLOSE

  PrintRptFile "Notices Printing", "UBNOTPRN.RPT", 1, RetCode, 0

  GOTO ExitPreReport

PrintThemOne:
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    'AcctNum = AcctNo&
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)


    PRINT #UBRpt, "~"
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; AcctNo&;
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 19); TAB(50); USING "########"; AcctNo&;
    PRINT #UBRpt,
    PRINT #UBRpt,

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

    PCnt = 0
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(47); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZIPCODE, 5)
    PRINT #UBRpt,
    PRINT #UBRpt, "  ***** NOTICE THIS IS NOT A BILL! *****"
    PRINT #UBRpt,
    PRINT #UBRpt, " "; Msgs$(1)
    PRINT #UBRpt, " "; Msgs$(2)
    PRINT #UBRpt, " "; Msgs$(3)
    PRINT #UBRpt, " "; Msgs$(4)
    PRINT #UBRpt, " "; Msgs$(5)
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

  RETURN

ExitEdit:
  HideCursor
  EXIT SUB

ExitPreReport:

  
END SUB

