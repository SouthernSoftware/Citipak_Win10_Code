DEFINT A-Z
DECLARE SUB SaveNewLocaRec (Form$(), LocatRec&, CustRec&)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SaveNewCustRec (FormC$(), NewRecNum&)
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION FileSize& (FileName$)
'DECLARE SUB LookUp (RecNo&, Text$, CLSFlag%)
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
  '$INCLUDE: 'mcorcust.bi'
  
  '$INCLUDE: 'ubcust.bi'
  '$INCLUDE: 'ubLoc.bi'
  '$INCLUDE: 'ubRate.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR

  CONST False = 0, True = NOT False

  DIM WBCustRec AS WBMCCoreCustRecType
  DIM WBTransRec AS WBTransType
  
  DIM UBBookSeq AS BookSeqRecType
  REDIM TUBCustRec(1) AS UBCustRecType
  REDIM TUBLocaRec(1) AS UBLocationRecType
  
  GOSUB ConvertRateCodes
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
  
  
  'WBCustRecLen = LEN(WBCustRec)
  TUBCustRecLen = LEN(TUBCustRec(1))
  TUBLocaRecLen = LEN(TUBLocaRec(1))
  
'GOTO Skip2Here
  
  CustAcct$ = SPACE$(10)

  CurDate = Date2Num(DATE$)
  FirstOne = True

  CustHandle = FREEFILE
  OPEN "OUT-ACCT.TXT" FOR INPUT AS #CustHandle


  DO UNTIL EOF(CustHandle)
      OutSideFlag = False
      St# = TIMER
      REDIM TUBCustRec(1) AS UBCustRecType
      REDIM TUBLocaRec(1) AS UBLocationRecType
      REDIM TUBTransRec(1) AS UBTransRecType
      UBTranRecLen = LEN(TUBTransRec(1))

      BadBook = 0
      
      INPUT #CustHandle, ACCOUNT$, SHORTNAME$, CUSTNAME$, SVCADDR$, BILLADDR$, CITYSTATE$, ZIPCODE$, Status$, BANKNBR$, BANKACCT$, STARTDATE$, DISCDATE$, WATERSRC$, SECLIGHTS$, DEMAND$, MULT1$, MULT2$, MULT3$
      INPUT #CustHandle, RATECODE1$, RATECODE2$, RATECODE3$, RATECODE4$, RATECODE5$, RATECODE6$, RATECODE7$, CONNDATE1$, CONNDATE2$, CONNDATE3$, DEPDATE1$, DEPDATE2$, DEPDATE3$, DEPAMT1$, DEPAMT2$, DEPAMT3$, METERNBR1$, METERNBR2$, METERNBR3$
      INPUT #CustHandle, STATUS1$, STATUS2$, STATUS3$, SVC1BAL$, SVC2BAL$, SVC3BAL$, SVC4BAL$, SVC5BAL$, SVC6BAL$, SVC7BAL$, MISCBAL$, PENBAL$, TAXBAL$, ENERGYCRED$
      INPUT #CustHandle, COMPMTR1$, COMPMTR2$, COMPMTR3

      IF FirstOne THEN
        FirstOne = False
        GOTO SkipThisCustomer
      END IF

      LOCATE 1, 1: PRINT " Acct:               ";
      LOCATE 1, 8: PRINT ACCOUNT$;

      Book = QPValI(LEFT$(ACCOUNT$, 2))
      IF Book < 10 THEN
        TUBLocaRec(1).Book = "0" + LTRIM$(STR$(Book))
      ELSE
        TUBLocaRec(1).Book = LTRIM$(STR$(Book))
      END IF
        
      Seq& = QPValL(MID$(QPTrim$(ACCOUNT$), 3))
      Seq$ = QPTrim$(STR$(Seq&))
      TUBLocaRec(1).SeqNumb = RIGHT$("000000" + Seq$, 6)
      
      TempBook$ = TUBLocaRec(1).Book + TUBLocaRec(1).SeqNumb
      UBBookSeq.BookSeq = QPValL(TempBook$)

      LOCATE 1, 25
      PRINT "New:"; TUBLocaRec(1).Book; "-"; TUBLocaRec(1).SeqNumb
      
      DO
      zz$ = INKEY$
      LOOP WHILE LEN(zz$)

      BookHand = FREEFILE
      OPEN "UBOOKSEQ.DAT" FOR RANDOM AS BookHand LEN = 4
      NextBookRec = (LOF(BookHand) / 4) + 1
      PUT BookHand, NextBookRec, UBBookSeq
      CLOSE BookHand
      
      TUBCustRec(1).SEARCH = QPTrim$(SHORTNAME$)
      TUBCustRec(1).CUSTNAME = QPTrim$(CUSTNAME$)

      SeAddr$ = QPTrim$(SVCADDR$)

      IF LEN(SeAddr$) > 0 THEN
        TUBLocaRec(1).SERVADDR = SeAddr$
      ELSE
        TUBLocaRec(1).SERVADDR = "UNKNOWN"
      END IF
      
      CPos = INSTR(CITYSTATE$, ",")
      IF CPos > 0 THEN
        TUBCustRec(1).City = LEFT$(CITYSTATE$, CPos - 1)
        State$ = MID$(CITYSTATE$, CPos + 1)
        DO
          PPos = INSTR(State$, ".")
          IF PPos > 0 THEN
            State$ = LEFT$(State$, PPos - 1) + MID$(State$, PPos + 1)
          END IF
        LOOP WHILE PPos
        DO
          PPos = INSTR(State$, " ")
          IF PPos > 0 THEN
            State$ = LEFT$(State$, PPos - 1) + MID$(State$, PPos + 1)
          END IF
        LOOP WHILE PPos

        TUBCustRec(1).State = QPTrim$(State$)
      END IF
      TUBCustRec(1).ZIPCODE = ZIPCODE$

      IF LEN(BILLADDR$) = 0 THEN
        TUBCustRec(1).Addr1 = TUBLocaRec(1).SERVADDR
      ELSE
        TUBCustRec(1).Addr1 = BILLADDR$
      END IF

      TUBCustRec(1).Status = Status$
      TUBLocaRec(1).LocatStatus = Status$
      
      DO
        DPos = INSTR(METERLOC$, "-")
        IF DPos > 0 THEN
          METERLOC$ = LEFT$(METERLOC$, DPos - 1) + MID$(METERLOC$, DPos + 1)
        END IF
      LOOP WHILE DPos > 0

      SELECT CASE LEN(METERLOC$)
      CASE 7
        TUBCustRec(1).HPhone = "   " + METERLOC$
      CASE 9
        TUBCustRec(1).HPhone = METERLOC$
      
      END SELECT
      FOR MCnt = 1 TO 7
        TUBLocaRec(1).LocMeters(MCnt).CurDate = -32767
        TUBLocaRec(1).LocMeters(MCnt).InsDate = -32767
        TUBLocaRec(1).LocMeters(1).CurRead = 0
        TUBLocaRec(1).LocMeters(1).PrevRead = 0
        TUBLocaRec(1).LocMeters(1).AvgUse = 0
        TUBLocaRec(1).LocMeters(1).UseCnt = 0
      NEXT

      Combined = False
      IF LEN(WATERCODE$) > 0 AND LEN(SEWERCODE$) > 0 THEN
        Combined = True
      END IF

      IF LEN(WATERCODE$) > 0 THEN
        TUBLocaRec(1).LocMeters(1).MTRNum = "0"
        TUBLocaRec(1).LocMeters(1).CurRead = 0
        TUBLocaRec(1).LocMeters(1).PrevRead = 0
        IF Combined THEN
          TUBLocaRec(1).LocMeters(1).MTRType = "C"
        ELSE
          TUBLocaRec(1).LocMeters(1).MTRType = "W"
        END IF
        TUBLocaRec(1).LocMeters(1).MTRMulti = 1
        TUBLocaRec(1).LocMeters(1).MTRUnit = "G"
        TUBLocaRec(1).LocMeters(1).NumUser = 1
        IF Combined THEN
          TUBLocaRec(1).Serv(1).RMTRTYPE = "C"
        ELSE
          TUBLocaRec(1).Serv(1).RMTRTYPE = "W"
        END IF

      END IF

      IF LEN(SEWERCODE$) > 0 THEN
        IF NOT Combined THEN
          TUBLocaRec(1).LocMeters(2).MTRNum = "0"
          TUBLocaRec(1).LocMeters(2).CurRead = 0
          TUBLocaRec(1).LocMeters(2).PrevRead = 0
          TUBLocaRec(1).LocMeters(2).MTRType = "S"
          TUBLocaRec(1).LocMeters(2).MTRMulti = 1
          TUBLocaRec(1).LocMeters(2).MTRUnit = "G"
          TUBLocaRec(1).LocMeters(2).NumUser = 1
          TUBLocaRec(1).Serv(2).RMTRTYPE = "S"
        ELSE
          TUBLocaRec(1).Serv(2).RMTRTYPE = "C"
        END IF

      END IF

      IF LEN(SVC4CODE$) > 0 THEN

      END IF

      IF LEN(GARBGCODE$) > 0 THEN
      END IF

      TUBCustRec(1).TAXEXPT = "N"
      
      TUBCustRec(1).SRCIT = "N"
      TUBCustRec(1).EPP = "N"
      TUBCustRec(1).EPPAMT = 0
      TUBCustRec(1).CASHONLY = "N"
      TUBCustRec(1).LATEFEE = "Y"
      TUBCustRec(1).CUTOFFYN = "Y"
      
      'TUBCustRec(1).AMTOWED = 0
      'TUBCustRec(1).AMTPD = 0
      'TUBCustRec(1).PAYMENT = 0
      
      'insert rate selection here

'--------------------------------------------------------------
      
      LOCATE 2, 25
      PRINT "New:"; TUBLocaRec(1).Book; "-"; TUBLocaRec(1).SeqNumb
      
      TUBCustRec(1).BILLCOPY = 1
      
'      TUBLocaRec(1).MFEE1 = WBCustRec.WATERDEP
'      TUBLocaRec(1).ZONE = WBCustRec.CityLimit

      FormC$(0, 0) = STRING$(CustFormLen, 0)
      FormL$(0, 0) = STRING$(LocatFormLen, 0)
      TUBCustRec(1).OPENDATE = -32767
      'TUBLocaRec(1).LocMeters(1).CurDate = CurDate
      TUBLocaRec(1).BILLWHO = "C"
      
      BCopy VARSEG(TUBCustRec(1)), VARPTR(TUBCustRec(1)), SSEG(FormC$(0, 0)), SADD(FormC$(0, 0)), CustFormLen, 0
      BCopy VARSEG(TUBLocaRec(1)), VARPTR(TUBLocaRec(1)), SSEG(FormL$(0, 0)), SADD(FormL$(0, 0)), LocatFormLen, 0
      
      SaveNewCustRec FormC$(), CustRec&
      SaveNewLocaRec FormL$(), LocatRec&, CustRec&
       
'Make deposit transactions
      ' WATERDEP$, DEPDATE$
      TWaterDep# = Value(WATERDEP$, E)
      REDIM TUBTransRec(1) AS UBTransRecType
      IF TWaterDep# > 0 THEN
        TUBTransRec(1).TransDate = Date2Num(DEPDATE$)
        TUBTransRec(1).TransType = 107  'Deposit
        TUBTransRec(1).TransDesc = "WATER DEPOSIT"
        TUBTransRec(1).TransAmt = TWaterDep#
        TUBTransRec(1).RevAmt(1) = TWaterDep#
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
      NumAcct = GetNumOfCust%
      PRINT LEFT$(TUBCustRec(1).CUSTNAME, 20), NumAcct, Cnt, USING "##.####"; Elasped#;
      'IF Cnt > 200 THEN
        LOCATE 8, 1
        PRINT USING "Worst time:  ##.####"; Worst#;
      'END IF
      
SkipThisCustomer:
    LOOP
    'NEXT
  
'do meter types and readings

  FirstOne = True
  ReadHandle = FREEFILE
  OPEN "OUT-READ.TXT" FOR INPUT AS #ReadHandle
  DO UNTIL EOF(ReadHandle)

    INPUT #ReadHandle, RACCOUNT$, MType$, MPreRead$, MCurRead$, RDate$
      IF FirstOne THEN
        FirstOne = False
        GOTO SkipThisReading
      END IF
      REDIM TUBLocaRec(1) AS UBLocationRecType
      Book = QPValI(LEFT$(ACCOUNT$, 2))
      IF Book < 10 THEN
        TUBLocaRec(1).Book = "0" + LTRIM$(STR$(Book))
      ELSE
        TUBLocaRec(1).Book = LTRIM$(STR$(Book))
      END IF
      Seq& = QPValL(MID$(QPTrim$(ACCOUNT$), 3))
      Seq$ = QPTrim$(STR$(Seq&))
      TUBLocaRec(1).SeqNumb = RIGHT$("000000" + Seq$, 6)

      'TempBook$ = TUBLocaRec(1).Book + TUBLocaRec(1).SeqNumb

 '     GOSUB SearchForLoc

SkipThisReading:
  LOOP



ExitCnvrt:
  CLOSE
  LOCATE 19, 1
  PRINT
  PRINT "Blank Meter:"; BlankMeter
  PRINT "Blank Names:"; BlankCnt
  PRINT "Blank Books:"; NoBook
  
  END

SearchForLoc:

  REDIM UBLocaRec(1) AS UBLocationRecType
  UBLocaRecLen = LEN(UBLocaRec(1))

  UBFile = FREEFILE
  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBFile LEN = UBLocaRecLen
  NumOfRecs& = LOF(UBFile) \ UBLocaRecLen

  FOR LCnt = 1 TO NumOfRecs&
    GET UBFile, NextRec&, UBLocaRec(1)
    IF UBLocaRec(1).Book = TUBLocaRec(1).Book AND UBLocaRec(1).SeqNumb = TUBLocaRec(1).SeqNumb THEN
      SELECT CASE MType$
      CASE "E"
      CASE "S"
      CASE "W"
      END SELECT
    END IF
  NEXT


  CLOSE UBFile

  UBFile = FREEFILE
  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBFile LEN = UBLocaRecLen
  PUT UBFile, NextRec&, UBLocaRec(1)
  CLOSE UBFile


RETURN

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
  
ConvertRateCodes:

  RETURN

  CLS

  RateHandle = FREEFILE
  OPEN "OUT-RATE.TXT" FOR INPUT AS #RateHandle
  DO UNTIL EOF(RateHandle)
      OutSideFlag = False

      BadBook = 0

      INPUT #RateHandle, RATECODE$, RATEDESC$, RATETYPE$, FIRST$, FIRSTAMT$, NEXT1$, NEXT1AMT$, NEXT2$, NEXT2AMT$, NEXT3, NEXT3AMT, NEXT4, NEXT4AMT$
      INPUT #RateHandle, NEXT5$, NEXT5AMT$, NEXT6$, NEXT6AMT$, NEXT7$, NEXT7AMT$, NEXT8$, NEXT8AMT$, BALAMT$, PCRATE$, FLATRATE$

      PRINT CUSTNAME$;
      PRINT RATECODE$
      PRINT RATEDESC$
      PRINT RATETYPE$
      PRINT FIRST$, FIRSTAMT$, NEXT1$, NEXT1AMT$
      PRINT NEXT2$, NEXT2AMT$, NEXT3, NEXT3AMT
      PRINT NEXT4, NEXT4AMT$, NEXT5$, NEXT5AMT$
      PRINT NEXT6$, NEXT6AMT$, NEXT7$, NEXT7AMT$
      PRINT NEXT8$, NEXT8AMT$
      PRINT BALAMT$, PCRATE$, FLATRATE$
      PRINT
  'St! = TIMER: DO: LOOP UNTIL TIMER > St! + 2!
  DO
    A$ = INKEY$:
    IF LEN(A$) THEN
      IF A$ = CHR$(27) THEN GOTO DONE
      EXIT DO
    END IF
  LOOP

  LOOP
DONE:
  CLOSE

END
RETURN


Skip2Here:
  CustHandle = FREEFILE
  OPEN "OUT-ACCT.TXT" FOR INPUT AS #CustHandle
  DO UNTIL EOF(CustHandle)
      OutSideFlag = False
      St# = TIMER

      REDIM TUBCustRec(1) AS UBCustRecType
      REDIM TUBLocaRec(1) AS UBLocationRecType
      REDIM TUBTransRec(1) AS UBTransRecType
      UBTranRecLen = LEN(TUBTransRec(1))

      BadBook = 0

      'GET #CustHandle, Cnt, WBCustRec

      INPUT #CustHandle, ACCOUNT$, SHORTNAME$, CUSTNAME$, SVCADDR$, BILLADDR$, CITYSTATE$, ZIPCODE$, Status$, BANKNBR$, BANKACCT$, STARTDATE$, DISCDATE$, WATERSRC$, SECLIGHTS$, DEMAND$, MULT1$, MULT2$, MULT3$
      INPUT #CustHandle, RATECODE1$, RATECODE2$, RATECODE3$, RATECODE4$, RATECODE5$, RATECODE6$, RATECODE7$, CONNDATE1$, CONNDATE2$, CONNDATE3$, DEPDATE1$, DEPDATE2$, DEPDATE3$, DEPAMT1$, DEPAMT2$, DEPAMT3$, METERNBR1$, METERNBR2$, METERNBR3$
      INPUT #CustHandle, STATUS1$, STATUS2$, STATUS3$, SVC1BAL$, SVC2BAL$, SVC3BAL$, SVC4BAL$, SVC5BAL$, SVC6BAL$, SVC7BAL$, MISCBAL$, PENBAL$, TAXBAL$, ENERGYCRED$
      INPUT #CustHandle, COMPMTR1$, COMPMTR2$, COMPMTR3$

      PRINT CUSTNAME$;
      PRINT RATECODE1$
      PRINT RATECODE2$
      PRINT RATECODE3$
      PRINT RATECODE4$
      PRINT RATECODE5$
      PRINT RATECODE6$
      PRINT RATECODE7$
      PRINT RATECODE8$

  St! = TIMER: DO: LOOP UNTIL TIMER > St! + .2
  A$ = INKEY$: IF LEN(A$) THEN EXIT DO

  LOOP
  CLOSE
  
  

