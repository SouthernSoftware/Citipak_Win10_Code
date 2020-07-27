DEFINT A-Z

DECLARE SUB SaveNewLocaRec (Form$(), LocatRec&, CustRec&)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SaveNewCustRec (FormC$(), NewRecNum&)
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION Num2Date$ (Daty%)

DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE SUB SaveOldCustRec (Form$(), RecNo&)
DECLARE SUB CustLookUp (RecNo%)
DECLARE SUB SearchGetCust (SEARCH$, RecNo&, CLSFlag%)
DECLARE SUB PrintCustList ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB EditCustomer ()
DECLARE SUB AddCustomer ()
DECLARE SUB LoadMeterRec (Form$())
DECLARE SUB AddNewMeter (Cust$, EdFlag)
DECLARE SUB EditMeter ()
DECLARE SUB SaveMeterRec (Form$())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB AddEditCustomer (RecNo&, LocatFlag%, F5Flag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB SaveCustRec (Form$())
DECLARE SUB AddNewCustomer ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB SaveSetUpRec (Form$())
DECLARE SUB LoadSetUpRec (Form$())
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB EditCust ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB StuffBuf (Ky$)
DECLARE FUNCTION FileSize& (FileName$)
  
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
  
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB WaitForAction ()
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION Round# (DblNumber#)
DECLARE SUB LoadCustRec (Form$(), RecNo&, LocatRec&, TFormLen%, LockedFlag%)
DECLARE SUB BlockClear ()
  
  ''$INCLUDE: 'owbtypes.bi'
  ''$INCLUDE: 'sprucust.bi'
  '''$INCLUDE: 'mcorcust.bi'
  ''$INCLUDE: 'wblong.bi'
  '$INCLUDE: 'WadesUB.bi'
  
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubRate.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  DIM WBCustRec  AS WBLongViewCustRecType
  DIM WBTransRec AS WBTransType
  
  DIM UBBookSeq AS BookSeqRecType
  
  REDIM TUBCustRec(1) AS UBCustRecType
  REDIM TUBLocaRec(1) AS UBLocationRecType
  
  LibName$ = "UB"
  ScrnName$ = "UBCUST"
  NumScrns = 2
  
  '--Initialize the form name array
  REDIM FormName$(1 TO NumScrns)
  FOR Scr = 1 TO NumScrns
    FormName$(Scr) = ScrnName$ + LTRIM$(STR$(Scr))
  NEXT
  
  '--Get the total number of fields from all pages
  NumFlds = -1
  FOR Scr = 1 TO NumScrns
    NumFlds = NumFlds + LibNumberOfFields(LibName$, FormName$(Scr)) + 1
  NEXT
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  FOR Scr = 1 TO NumScrns
    LibGetFldDef LibName$, FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
    StartEl = StartEl + Fld(StartEl).Fields + 1
  NEXT
  
  CustFormLen = Fld(0).Row
  
  LibName$ = "UB"
  ScrnName$ = "UBLOC"
  NumScrns = 4
  
  '--define the multi-choice fields
  
  '--Initialize the form name array
  REDIM FormName$(1 TO NumScrns)
  FOR Scr = 1 TO NumScrns
    FormName$(Scr) = ScrnName$ + LTRIM$(STR$(Scr))
  NEXT
  
  '--Get the total number of fields from all pages
  NumFlds = -1
  FOR Scr = 1 TO NumScrns
    NumFlds = NumFlds + LibNumberOfFields(LibName$, FormName$(Scr)) + 1
  NEXT
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  FOR Scr = 1 TO NumScrns
    LibGetFldDef LibName$, FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
    StartEl = StartEl + Fld(StartEl).Fields + 1
  NEXT
  
  LocatFormLen = Fld(0).Row
  TodayDate = Date2Num(DATE$)

  CLS
  
  REDIM FormC$(0, 0)
  REDIM FormL$(0, 0)
  
  DIM RateTbl(1 TO 9) AS STRING * 4

  RateTbl(1) = "LFR"  'Land Fill"
  RateTbl(2) = "SRI"  'Sewer Inside"
  RateTbl(5) = "SWO"  'Sewer Outside"
  RateTbl(6) = "WRI"  'Water Inside"
  RateTbl(7) = "WRO"  'Water Outside"
  RateTbl(3) = "SRSI" 'Sewer Senior Inside"
  RateTbl(4) = "SRSO" 'Sewer Senior Outside"
  RateTbl(8) = "WRSI" 'Water Senior Inside"
  RateTbl(9) = "WRSO" 'Water Senior Outside"

  
  WBCustRecLen = LEN(WBCustRec)
  TUBCustRecLen = LEN(TUBCustRec(1))
  TUBLocaRecLen = LEN(TUBLocaRec(1))
  WBTransRecLen = LEN(WBTransRec)

'GOTO Skip2Here
  
  CustAcct$ = SPACE$(10)
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = WBCustRecLen
  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  CurDate = Date2Num(DATE$)
  
    FOR Cnt = 1 TO NumOfRecs
      St# = TIMER
      
      REDIM TUBCustRec(1) AS UBCustRecType
      REDIM TUBLocaRec(1) AS UBLocationRecType
      REDIM TUBTransRec(1) AS UBTransRecType
      UBTranRecLen = LEN(TUBTransRec(1))

      BadBook = 0
      
      GET #CustHandle, Cnt, WBCustRec
      
      IF VAL(WBCustRec.Meter) < 1 THEN
        BlankMeter = BlankMeter + 1
        GOTO SkipThisCustomer
      END IF
      
      LSET CustAcct$ = WBCustRec.Meter
      LOCATE 1, 1: PRINT " Acct:"; CustAcct$;
      
      TUBCustRec(1).SEARCH = WBCustRec.LastName
      
      TFirstName$ = LTRIM$(RTRIM$(WBCustRec.FirstName))
      TLastName$ = LTRIM$(RTRIM$(WBCustRec.LastName))
      
      IF LEN(TFirstName$) = 0 THEN
        IF LEN(TLastName$) = 0 THEN
          BlankCnt = BlankCnt + 1
          GOTO SkipThisCustomer
        END IF
      END IF
      
      TransCnt& = 0
      IF WBCustRec.FirstTrans > 0 THEN
        OPEN "WBARTRAN.DAT" FOR RANDOM AS #20 LEN = WBTransRecLen
        ThisTrans& = WBCustRec.FirstTrans
        DO WHILE ThisTrans& > 0
          TransCnt& = TransCnt& + 1
          GET #20, ThisTrans&, WBTransRec
          ThisTrans& = WBTransRec.NextTrans
        LOOP
        CLOSE 20
      END IF
      
      TUBLocaRec(1).SEQNUMB = "000000"
      TUBLocaRec(1).Book = "00"
      
      DashPos = INSTR(WBCustRec.Meter, "-")
      
      IF DashPos > 1 THEN
        Book = QPValI(LEFT$(WBCustRec.Meter, DashPos - 1))
        IF Book < 10 THEN
          TUBLocaRec(1).Book = "0" + LTRIM$(STR$(Book))
        ELSE
          TUBLocaRec(1).Book = LTRIM$(STR$(Book))
        END IF
        
        SEQ$ = MID$(WBCustRec.Meter, DashPos + 1)
        SlashPos = INSTR(SEQ$, "/")
        IF SlashPos > 1 THEN
          SEQ$ = LEFT$(SEQ$, SlashPos - 1)
        ELSE
          SEQ$ = RTRIM$(MID$(WBCustRec.Meter, DashPos + 1))
        END IF
        SeqLen = LEN(SEQ$)
        MID$(TUBLocaRec(1).SEQNUMB, 7 - SeqLen) = SEQ$
      ELSE
        'TUBLocaRec(1).Book = LTRIM$(STR$(Book))
      END IF
      
      TempBook$ = TUBLocaRec(1).Book + TUBLocaRec(1).SEQNUMB
      UBBookSeq.BookSeq = QPValL(TempBook$)
      
      BookHand = FREEFILE
      OPEN "UBOOKSEQ.DAT" FOR RANDOM AS BookHand LEN = 4
      NextBookRec = (LOF(BookHand) / 4) + 1
      PUT BookHand, NextBookRec, UBBookSeq
      CLOSE BookHand
      
      TUBCustRec(1).TAXEXPT = "N"
      TUBCustRec(1).CUSTTYPE = WBCustRec.CUSTTYPE
      
      TUBCustRec(1).SRCIT = "N"
      TUBCustRec(1).EPP = "N"
      TUBCustRec(1).EPPAMT = 0
      TUBCustRec(1).CASHONLY = "N"
      TUBCustRec(1).LATEFEE = "Y"
      TUBCustRec(1).CUTOFFYN = "Y"
      
      TUBCustRec(1).AMTOWED = 0
      TUBCustRec(1).AMTPD = 0
      TUBCustRec(1).PAYMENT = 0
      
      'insert rate selection here
      
      TUBCustRec(1).DRAFTYPE = "N"
      
      FOR MCnt = 1 TO 7
        TUBLocaRec(1).LocMeters(MCnt).CurDate = -32767
        TUBLocaRec(1).LocMeters(MCnt).InsDate = -32767
        TUBLocaRec(1).LocMeters(1).CurRead = 0
        TUBLocaRec(1).LocMeters(1).PrevRead = 0
        TUBLocaRec(1).LocMeters(1).AvgUse = 0
        TUBLocaRec(1).LocMeters(1).UseCnt = 0
      NEXT

      'WBCustRec.Tax
      '^^^  Tipping Exempt
       
      SELECT CASE QPTrim$(WBCustRec.CityLimit)
      CASE "I"
        OutSideFlag = 0
      CASE "O"
        OutSideFlag = -1
      END SELECT


        SELECT CASE WBCustRec.WRate
              'This actually tells what services
        CASE 1  'Water & Sewer
          IF OutSideFlag THEN
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(7) 'Water Outside
            TUBLocaRec(1).Serv(2).RATECODE = RateTbl(5) 'Sewer Outside"
          ELSE
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(6) 'Water Inside"
            TUBLocaRec(1).Serv(2).RATECODE = RateTbl(2) 'Sewer Inside"
          END IF
          TUBLocaRec(1).Serv(1).RMTRTYPE = "C"
          TUBLocaRec(1).Serv(2).RMTRTYPE = "C"

          IF WBCustRec.NumMeters = 1 THEN
            TUBLocaRec(1).LocMeters(1).MTRNUM = STR$(WBCustRec.WMeterNum1)
            TUBLocaRec(1).LocMeters(1).CurRead = WBCustRec.WCurrRead1
            TUBLocaRec(1).LocMeters(1).PrevRead = WBCustRec.WPastRead1
            TUBLocaRec(1).LocMeters(1).MTRType = "C"
            TUBLocaRec(1).LocMeters(1).MTRMulti = 100
            TUBLocaRec(1).LocMeters(1).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(1).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(1).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(1).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 2 THEN
            TUBLocaRec(1).LocMeters(2).MTRNUM = STR$(WBCustRec.WMeterNum2)
            TUBLocaRec(1).LocMeters(2).CurRead = WBCustRec.WCurrRead2
            TUBLocaRec(1).LocMeters(2).PrevRead = WBCustRec.WPastRead2
            TUBLocaRec(1).LocMeters(2).MTRType = "C"
            TUBLocaRec(1).LocMeters(2).MTRMulti = 100
            TUBLocaRec(1).LocMeters(2).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(2).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(2).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(2).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 3 THEN
            TUBLocaRec(1).LocMeters(3).MTRNUM = STR$(WBCustRec.WMeterNum3)
            TUBLocaRec(1).LocMeters(3).CurRead = WBCustRec.WCurrRead3
            TUBLocaRec(1).LocMeters(3).PrevRead = WBCustRec.WPastRead3
            TUBLocaRec(1).LocMeters(3).MTRType = "C"
            TUBLocaRec(1).LocMeters(3).MTRMulti = 100
            TUBLocaRec(1).LocMeters(3).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(3).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(3).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(3).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF

        CASE 2  'Water Only
          IF OutSideFlag THEN
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(7) 'Water Outside
          ELSE
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(6) 'Water Inside"
          END IF
          TUBLocaRec(1).Serv(1).RMTRTYPE = "W"

          IF WBCustRec.NumMeters = 1 THEN
            TUBLocaRec(1).LocMeters(1).MTRNUM = STR$(WBCustRec.WMeterNum1)
            TUBLocaRec(1).LocMeters(1).CurRead = WBCustRec.WCurrRead1
            TUBLocaRec(1).LocMeters(1).PrevRead = WBCustRec.WPastRead1
            TUBLocaRec(1).LocMeters(1).MTRType = "W"
            TUBLocaRec(1).LocMeters(1).MTRMulti = 100
            TUBLocaRec(1).LocMeters(1).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(1).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(1).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(1).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 2 THEN
            TUBLocaRec(1).LocMeters(2).MTRNUM = STR$(WBCustRec.WMeterNum2)
            TUBLocaRec(1).LocMeters(2).CurRead = WBCustRec.WCurrRead2
            TUBLocaRec(1).LocMeters(2).PrevRead = WBCustRec.WPastRead2
            TUBLocaRec(1).LocMeters(2).MTRType = "W"
            TUBLocaRec(1).LocMeters(2).MTRMulti = 100
            TUBLocaRec(1).LocMeters(2).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(2).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(2).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(2).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 3 THEN
            TUBLocaRec(1).LocMeters(3).MTRNUM = STR$(WBCustRec.WMeterNum3)
            TUBLocaRec(1).LocMeters(3).CurRead = WBCustRec.WCurrRead3
            TUBLocaRec(1).LocMeters(3).PrevRead = WBCustRec.WPastRead3
            TUBLocaRec(1).LocMeters(3).MTRType = "W"
            TUBLocaRec(1).LocMeters(3).MTRMulti = 100
            TUBLocaRec(1).LocMeters(3).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(3).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(3).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(3).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF

        CASE 3  'Sewer Only
          IF OutSideFlag THEN
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(5) 'Sewer Outside"
          ELSE
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(2) 'Sewer Inside"
          END IF
          TUBLocaRec(1).Serv(1).RMTRTYPE = "S"
          IF WBCustRec.NumMeters = 1 THEN
            TUBLocaRec(1).LocMeters(1).MTRNUM = STR$(WBCustRec.WMeterNum1)
            TUBLocaRec(1).LocMeters(1).CurRead = WBCustRec.WCurrRead1
            TUBLocaRec(1).LocMeters(1).PrevRead = WBCustRec.WPastRead1
            TUBLocaRec(1).LocMeters(1).MTRType = "W"
            TUBLocaRec(1).LocMeters(1).MTRMulti = 100
            TUBLocaRec(1).LocMeters(1).MTRUnit = "S"
            TUBLocaRec(1).LocMeters(1).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(1).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(1).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 2 THEN
            TUBLocaRec(1).LocMeters(2).MTRNUM = STR$(WBCustRec.WMeterNum2)
            TUBLocaRec(1).LocMeters(2).CurRead = WBCustRec.WCurrRead2
            TUBLocaRec(1).LocMeters(2).PrevRead = WBCustRec.WPastRead2
            TUBLocaRec(1).LocMeters(2).MTRType = "S"
            TUBLocaRec(1).LocMeters(2).MTRMulti = 100
            TUBLocaRec(1).LocMeters(2).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(2).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(2).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(2).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 3 THEN
            TUBLocaRec(1).LocMeters(3).MTRNUM = STR$(WBCustRec.WMeterNum3)
            TUBLocaRec(1).LocMeters(3).CurRead = WBCustRec.WCurrRead3
            TUBLocaRec(1).LocMeters(3).PrevRead = WBCustRec.WPastRead3
            TUBLocaRec(1).LocMeters(3).MTRType = "S"
            TUBLocaRec(1).LocMeters(3).MTRMulti = 100
            TUBLocaRec(1).LocMeters(3).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(3).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(3).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(3).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF

        CASE 4  'Water & Sewer  Senior
          IF OutSideFlag THEN
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(9) 'Water Senior Outside"
            TUBLocaRec(1).Serv(2).RATECODE = RateTbl(4) 'Sewer Senior Outside"
          ELSE
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(8) 'Water Senior Inside"
            TUBLocaRec(1).Serv(2).RATECODE = RateTbl(3) 'Sewer Senior Inside"
          END IF
          TUBLocaRec(1).Serv(1).RMTRTYPE = "C"
          TUBLocaRec(1).Serv(2).RMTRTYPE = "C"
          IF WBCustRec.NumMeters = 1 THEN
            TUBLocaRec(1).LocMeters(1).MTRNUM = STR$(WBCustRec.WMeterNum1)
            TUBLocaRec(1).LocMeters(1).CurRead = WBCustRec.WCurrRead1
            TUBLocaRec(1).LocMeters(1).PrevRead = WBCustRec.WPastRead1
            TUBLocaRec(1).LocMeters(1).MTRType = "C"
            TUBLocaRec(1).LocMeters(1).MTRMulti = 100
            TUBLocaRec(1).LocMeters(1).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(1).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(1).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(1).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 2 THEN
            TUBLocaRec(1).LocMeters(2).MTRNUM = STR$(WBCustRec.WMeterNum2)
            TUBLocaRec(1).LocMeters(2).CurRead = WBCustRec.WCurrRead2
            TUBLocaRec(1).LocMeters(2).PrevRead = WBCustRec.WPastRead2
            TUBLocaRec(1).LocMeters(2).MTRType = "C"
            TUBLocaRec(1).LocMeters(2).MTRMulti = 100
            TUBLocaRec(1).LocMeters(2).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(2).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(2).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(2).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 3 THEN
            TUBLocaRec(1).LocMeters(3).MTRNUM = STR$(WBCustRec.WMeterNum3)
            TUBLocaRec(1).LocMeters(3).CurRead = WBCustRec.WCurrRead3
            TUBLocaRec(1).LocMeters(3).PrevRead = WBCustRec.WPastRead3
            TUBLocaRec(1).LocMeters(3).MTRType = "C"
            TUBLocaRec(1).LocMeters(3).MTRMulti = 100
            TUBLocaRec(1).LocMeters(3).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(3).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(3).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(3).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF

        CASE 5  'No Water-Sewer

        CASE 6  'Water Only  Senoir
          IF OutSideFlag THEN
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(9) 'Water Senior Outside
          ELSE
            TUBLocaRec(1).Serv(1).RATECODE = RateTbl(8) 'Water Senior Inside"
          END IF
          TUBLocaRec(1).Serv(1).RMTRTYPE = "W"
          IF WBCustRec.NumMeters = 1 THEN
            TUBLocaRec(1).LocMeters(1).MTRNUM = STR$(WBCustRec.WMeterNum1)
            TUBLocaRec(1).LocMeters(1).CurRead = WBCustRec.WCurrRead1
            TUBLocaRec(1).LocMeters(1).PrevRead = WBCustRec.WPastRead1
            TUBLocaRec(1).LocMeters(1).MTRType = "W"
            TUBLocaRec(1).LocMeters(1).MTRMulti = 100
            TUBLocaRec(1).LocMeters(1).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(1).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(1).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(1).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 2 THEN
            TUBLocaRec(1).LocMeters(2).MTRNUM = STR$(WBCustRec.WMeterNum2)
            TUBLocaRec(1).LocMeters(2).CurRead = WBCustRec.WCurrRead2
            TUBLocaRec(1).LocMeters(2).PrevRead = WBCustRec.WPastRead2
            TUBLocaRec(1).LocMeters(2).MTRType = "W"
            TUBLocaRec(1).LocMeters(2).MTRMulti = 100
            TUBLocaRec(1).LocMeters(2).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(2).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(2).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(2).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF
          IF WBCustRec.NumMeters = 3 THEN
            TUBLocaRec(1).LocMeters(3).MTRNUM = STR$(WBCustRec.WMeterNum3)
            TUBLocaRec(1).LocMeters(3).CurRead = WBCustRec.WCurrRead3
            TUBLocaRec(1).LocMeters(3).PrevRead = WBCustRec.WPastRead3
            TUBLocaRec(1).LocMeters(3).MTRType = "W"
            TUBLocaRec(1).LocMeters(3).MTRMulti = 100
            TUBLocaRec(1).LocMeters(3).MTRUnit = "G"
            TUBLocaRec(1).LocMeters(3).NumUser = WBCustRec.WMin
            TUBLocaRec(1).LocMeters(3).AvgUse = WBCustRec.WAvgUse
            TUBLocaRec(1).LocMeters(3).UseCnt = TransCnt&           'WBCustRec.WAvgCnt
          END IF


        CASE ELSE
        END SELECT

      LOCATE 3, 1
      PRINT "W Rate:"; WBCustRec.WRate; TAB(22); WaterRate$
      LOCATE 4, 1
      PRINT "Tr Cnt:"; TransCnt&;
      
'Flat Rates---------------------------------------------------
      
      FOR FRCnt = 1 TO 4
        TUBLocaRec(1).FlatRates(1).FRDESC = ""
        TUBLocaRec(1).FlatRates(1).FRAMT = 0
      NEXT
      
'      Desc$ = UCASE$(QPTrim$(WBCustRec.RDesc1))
'      IF LEN(Desc$) > 0 THEN
'        TUBLocaRec(1).FlatRates(1).FRDesc = Desc$
'        TUBLocaRec(1).FlatRates(1).FRAmt = Round(0# + WBCustRec.RAmt1)
'        IF INSTR(Desc$, "TRA") THEN
'          TUBLocaRec(1).FlatRates(1).RevSrc = 9 'WBCustRec.RAcct1
'          TUBLocaRec(1).FlatRates(1).NumMin = 1
'        END IF
'      END IF
      
'      Desc$ = UCASE$(QPTrim$(WBCustRec.RDesc2))
'      IF LEN(Desc$) > 0 THEN
'        TUBLocaRec(1).FlatRates(2).FRDesc = Desc$
'        TUBLocaRec(1).FlatRates(2).FRAmt = Round(0# + WBCustRec.RAmt2)
'        IF INSTR(Desc$, "TRA") THEN
'          TUBLocaRec(1).FlatRates(2).RevSrc = 9 'WBCustRec.RAcct1
'          TUBLocaRec(1).FlatRates(1).NumMin = 1
'        END IF
'      END IF
      
'      Desc$ = UCASE$(QPTrim$(WBCustRec.RDesc3))
'      IF LEN(Desc$) > 0 THEN
'        TUBLocaRec(1).FlatRates(3).FRDesc = Desc$
'        TUBLocaRec(1).FlatRates(3).FRAmt = Round(0# + WBCustRec.RAmt3)
'        IF INSTR(Desc$, "TRA") THEN
'          TUBLocaRec(1).FlatRates(3).RevSrc = 9 'WBCustRec.RAcct1
'          TUBLocaRec(1).FlatRates(1).NumMin = 1
'        END IF
'      END IF

'Processes security as a flat rate item
'      IF WBCustRec.Security1 > 0 THEN
'        TUBLocaRec(1).FlatRates(2).RevSrc = 11
'        TUBLocaRec(1).FlatRates(2).FRDesc = "SECURITY LIGHTS"
'        TUBLocaRec(1).FlatRates(2).FRAmt = 6.9
'        TUBLocaRec(1).FlatRates(2).NumMin = WBCustRec.Security1
'      ELSEIF WBCustRec.Security2 > 0 THEN
'        TUBLocaRec(1).FlatRates(2).RevSrc = 11
'        TUBLocaRec(1).FlatRates(2).FRDesc = "SECURITY LIGHTS"
'        TUBLocaRec(1).FlatRates(2).FRAmt = 12
'        TUBLocaRec(1).FlatRates(2).NumMin = WBCustRec.Security2
'      END IF
      
'--------------------------------------------------------------
      
      IF Book = 0 THEN
        BadBook = -1
        NoBook = NoBook + 1
      END IF
      
      TMeter$ = MID$(WBCustRec.Meter, 4)

      LOCATE 1, 25
      PRINT "New:"; TUBLocaRec(1).Book; "-"; TUBLocaRec(1).SEQNUMB
      
      IF LEN(TFirstName$) = 0 THEN
        TUBCustRec(1).CUSTNAME = TLastName$
      ELSE
        TUBCustRec(1).CUSTNAME = TFirstName$ + " " + LTRIM$(RTRIM$(WBCustRec.LastName))
      END IF
      
      TUBCustRec(1).ADDR1 = WBCustRec.add1
      TUBCustRec(1).ADDR2 = WBCustRec.Add2
      
      TUBCustRec(1).CITY = WBCustRec.CITY
      TUBCustRec(1).STATE = WBCustRec.STATE
      
      TUBCustRec(1).BILLCOPY = 1
      TUBCustRec(1).ZIPCODE = WBCustRec.Zip
      
      RSET TUBCustRec(1).SOSEC = WBCustRec.Phone
      
'      TUBLocaRec(1).MFEE1 = WBCustRec.WaterDep
      TUBLocaRec(1).ZONE = WBCustRec.CityLimit
      
      SeAddr$ = LTRIM$(RTRIM$(WBCustRec.Add2))
      IF LEN(SeAddr$) > 0 THEN
        TUBLocaRec(1).SERVADDR = SeAddr$
      ELSE
        TUBLocaRec(1).SERVADDR = "UNKNOWN"
      END IF
      
      IF UCASE$(WBCustRec.Active) = "Y" THEN
        TUBCustRec(1).STATUS = "A"
        TUBLocaRec(1).LocatStatus = "A"
      ELSE
        TUBCustRec(1).STATUS = "I"
        TUBLocaRec(1).LocatStatus = "I"
      END IF
      
      FormC$(0, 0) = STRING$(CustFormLen, 0)
      FormL$(0, 0) = STRING$(LocatFormLen, 0)
      TUBCustRec(1).OPENDATE = -32767
      'TUBLocaRec(1).LocMeters(1).CurDate = CurDate
      TUBLocaRec(1).BILLWHO = "C"
      
      BCopy VARSEG(TUBCustRec(1)), VARPTR(TUBCustRec(1)), SSEG(FormC$(0, 0)), SADD(FormC$(0, 0)), CustFormLen, 0
      BCopy VARSEG(TUBLocaRec(1)), VARPTR(TUBLocaRec(1)), SSEG(FormL$(0, 0)), SADD(FormL$(0, 0)), LocatFormLen, 0
      
      SaveNewCustRec FormC$(), CustRec&
      'SaveNewLocaRec FormL$(), LocatRec&, CustRec&
      
'Put current balances in the cust file
      OPEN "Ubcust.dat" FOR RANDOM SHARED AS #20 LEN = TUBCustRecLen
      GET #20, CustRec&, TUBCustRec(1)
      TUBCustRec(1).CurrRevAmts(1) = Round(0# + WBCustRec.IRevenue1)
      TUBCustRec(1).CurrRevAmts(2) = Round(0# + WBCustRec.IRevenue2)
      TUBCustRec(1).CurrRevAmts(3) = Round(0# + WBCustRec.IRevenue3)
      TUBCustRec(1).CurrRevAmts(4) = Round(0# + WBCustRec.IRevenue4)
      TUBCustRec(1).CurrRevAmts(5) = Round(0# + WBCustRec.IRevenue5)
      TUBCustRec(1).CurrRevAmts(6) = Round(0# + WBCustRec.IRevenue6)
      TUBCustRec(1).CurrBalance = Round(0# + WBCustRec.CurBal)
      TUBCustRec(1).PrevBalance = Round(0# + WBCustRec.PastDue)
      PUT #20, CustRec&, TUBCustRec(1)
      CLOSE 20


'Processes & convert transactions
      FirstTrDate = TodayDate
      IF WBCustRec.FirstTrans > 0 THEN
        OPEN "WBARTRAN.DAT" FOR RANDOM AS #20 LEN = WBTransRecLen
        ThisTrans& = WBCustRec.FirstTrans
        DO WHILE ThisTrans& > 0
          REDIM TUBTransRec(1) AS UBTransRecType
          GET #20, ThisTrans&, WBTransRec
          TDate$ = QPTrim$(WBTransRec.TransDate)
          IF LEN(TDate$) = 6 THEN
            TDate$ = LEFT$(TDate$, 4) + "19" + RIGHT$(TDate$, 2)
          END IF
          ThisTRDate = Date2Num(TDate$)
          IF ThisTRDate < FirstTrDate THEN
            FirstTrDate = ThisTRDate
          END IF

          TUBTransRec(1).TransDate = ThisTRDate
          TUBTransRec(1).TransDesc = WBTransRec.TransDesc
          TUBTransRec(1).TransAmt = Round(0# + WBTransRec.TCurrAmt)
          TUBTransRec(1).CustLocation = LocatRec& 'Pointer to Location Rec
          TUBTransRec(1).CustAcctNo = CustRec&    'Pointer to Customer Rec

          SELECT CASE WBTransRec.TransType
            CASE 1  'CHARGE  utility bill
               TUBTransRec(1).TransType = 101
               TUBTransRec(1).TransDesc = "UTILITY CHARGE"
               TUBTransRec(1).TransAmt = Round(0# + WBTransRec.TCurrAmt)
               TUBTransRec(1).RunBalance = WBTransRec.TransAmt
            CASE 2   'PAYMENT
               TUBTransRec(1).TransType = 104
               TUBTransRec(1).TransDesc = "UTILITY PAYMENT"
               TUBTransRec(1).TransAmt = Round(0# + WBTransRec.TCurrAmt)
               TUBTransRec(1).RunBalance = WBTransRec.TransAmt
            CASE 3   'CHARGE  Penalty + Reconnect
               TUBTransRec(1).TransType = 102
               TUBTransRec(1).TransDesc = "PENALTY/RECONNECT"
               TUBTransRec(1).TransAmt = Round(0# + WBTransRec.TCurrAmt)
               TUBTransRec(1).RunBalance = WBTransRec.TransAmt
            CASE 4     'Adjustment
               TUBTransRec(1).TransType = 109
               TUBTransRec(1).TransDesc = "UTILITY ADJUSTMENTS"
               TUBTransRec(1).TransAmt = Round(0# + WBTransRec.TCurrAmt)
               TUBTransRec(1).RunBalance = WBTransRec.TransAmt
          END SELECT
          SELECT CASE WBCustRec.WRate
          CASE 1, 4 'Water & Sewer
            TUBTransRec(1).RevAmt(1) = Round(0# + WBTransRec.TCurrAmt)
          CASE 2, 6 'Water Only
            TUBTransRec(1).RevAmt(1) = Round(0# + WBTransRec.TCurrAmt)
          CASE 3  'Sewer Only
            TUBTransRec(1).RevAmt(2) = Round(0# + WBTransRec.TCurrAmt)
          CASE 5  'No Water-Sewer

          CASE ELSE
          END SELECT
          GOSUB PostTransaction
          ThisTrans& = WBTransRec.NextTrans
        LOOP
        CLOSE 20
      END IF

'Make deposit transactions
      TWaterDep# = Round(0# + WBCustRec.WaterDep)
      REDIM TUBTransRec(1) AS UBTransRecType
      IF TWaterDep# > 0 THEN
        TUBTransRec(1).TransDate = FirstTrDate
        TUBTransRec(1).TransType = 107  'Deposit
        TUBTransRec(1).TransDesc = "WATER DEPOSIT"
        TUBTransRec(1).TransAmt = TWaterDep#
        TUBTransRec(1).CustAcctNo = CustRec&
        TUBTransRec(1).CustLocation = LocatRec&
        GOSUB PostTransaction
      END IF
      
      zz$ = INKEY$
      IF LEN(zz$) THEN
        IF zz$ = CHR$(27) THEN
          GOTO ExitCnvrt
        END IF
      END IF

      St2# = TIMER
      Elasped# = St2# - St#
      IF Elasped# > Worst# THEN
        Worst# = Elasped#
      END IF
      
      LOCATE 12, 1
      NumAcct = GetNumOfAcct%
      PRINT LEFT$(TUBCustRec(1).CUSTNAME, 20), NumAcct, Cnt, USING "##.####"; Elasped#;
      IF Cnt > 200 THEN
        LOCATE 8, 1
        PRINT USING "Worst time:  ##.####"; Worst#;
      END IF
      
SkipThisCustomer:
    NEXT
  
ExitCnvrt:
  CLOSE
  LOCATE 19, 1
  PRINT
  PRINT "Blank Meter:"; BlankMeter
  PRINT "Blank Names:"; BlankCnt
  PRINT "Blank Books:"; NoBook
  
  END

PostTransaction:

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  UBCust = FREEFILE
  OPEN "Ubcust.dat" FOR RANDOM SHARED AS UBCust LEN = TUBCustRecLen
  GET UBCust, CustRec&, TUBCustRec(1)

  NumOfTranRecs& = LOF(UBTran) \ UBTranRecLen
  NumOfTranRecs& = NumOfTranRecs& + 1 'point to next trans to write

  PrevLastTrans& = TUBCustRec(1).LastTrans
  TUBTransRec(1).PrevTrans = PrevLastTrans&
  TUBCustRec(1).LastTrans = NumOfTranRecs&

  PUT UBCust, CustRec&, TUBCustRec(1)
  PUT UBTran, NumOfTranRecs&, TUBTransRec(1)

  CLOSE UBTran, UBCust

RETURN
  
  
  

Skip2Here:

        OPEN "WBARTRAN.DAT" FOR RANDOM AS #20 LEN = 50
        NumOfTranRecs& = LOF(20) \ WBTransRecLen
        FOR ZZCnt& = 1 TO NumOfTranRecs&
        GET #20, ZZCnt&, WBTransRec
        TDate$ = QPTrim$(WBTransRec.TransDate)
        IF LEN(TDate$) = 6 THEN
          TDate$ = LEFT$(TDate$, 4) + "19" + RIGHT$(TDate$, 2)
        END IF
        PRINT Num2Date$(Date2Num(TDate$)), , WBTransRec.TransDesc, WBTransRec.TransType
        zz$ = INKEY$
        IF zz$ = CHR$(27) THEN EXIT FOR
        NEXT
        CLOSE
          'TUBTransRec(1).TransDate = ThisTRDate
          'TUBTransRec(1).TransDesc = WBTransRec.TransDesc
          'TUBTransRec(1).TransAmt = WBTransRec.TCurrAmt
          'TUBTransRec(1).CustLocation = LocatRec& 'Pointer to Location Rec
          'TUBTransRec(1).CustAcctNo = CustRec&    'Pointer to Customer Rec

