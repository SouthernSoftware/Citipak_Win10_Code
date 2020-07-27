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
  CLS
  'REDIM WadesRec(1) AS WadesDataType
  '
  'WadesRecLen = LEN(WadesRec(1))
  '
  'PRINT "O"; oCnt, "W"; wCnt, "-"; dCnt
  'PRINT "X"; xCnt, "S"; sCnt, "I"; iCnt
  'PRINT "E"; eCnt
  'PRINT
  'PRINT "Total:"; (oCnt + wCnt + dCnt + xCnt + sCnt + iCnt + eCnt)

'END


  LibName$ = "UB"
  ScrnName$ = "UBCUST"
  NumScrns = 4
  
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
  
  TodayDate = Date2Num(DATE$)

  CLS
  
  REDIM FormC$(0, 0)
  
'GOTO Skip2Here
  
  CustAcct$ = SPACE$(10)
  CurDate = Date2Num(DATE$)


  FirstTime = -1

  WadesHandle = FREEFILE
  OPEN "DUH.TXT" FOR INPUT AS WadesHandle

  DO
    INPUT #WadesHandle, CName$, add1$, Add2$, City$, State$, Zip$, SerAdd$, AcctNo$, BookNo$, Rate$, AODate$, BCycle$, Zone$, PRoute$, SSN$, DLicNo$, HPhone$, MSerNo$, MInDate$, CurRead$, CurDate$, Multi$
                      '   1       2      3      4      5      6       7       8         9      10       11       12      13     14      15      16      17       18        19        20        21       22
      REDIM UBCustRec(1) AS NewUBCustRecType

      UBCustRec(1).STATUS = "A"

      IF FirstTime THEN
        LastBook = VAL(BookNo$)
        FirstTime = 0
      END IF
      
      CName$ = QPTrim$(UCASE$(CName$))
      NameLen = LEN(CName$)

      SEARCH$ = ""

      FOR Cnt = NameLen TO 1 STEP -1
        Char$ = MID$(CName$, Cnt, 1)
        IF Char$ = " " THEN
          SEARCH$ = MID$(CName$, Cnt + 1)
          EXIT FOR
        END IF
      NEXT

      UBCustRec(1).SEARCH = SEARCH$
      UBCustRec(1).CustName = CName$
      
      UBCustRec(1).ADDR1 = add1$
      UBCustRec(1).ADDR2 = Add2$
      UBCustRec(1).City = City$
      UBCustRec(1).State = State$
      UBCustRec(1).ZIPCODE = Zip$
      UBCustRec(1).SERVADDR = SerAdd$
      
      ThisBook = VAL(BookNo$)

      IF ThisBook <> LastBook THEN
        LastBook = ThisBook
        SeqNum = 10
      ELSE
        SeqNum = SeqNum + 10
      END IF
      
      UBCustRec(1).SeqNumb = "000000"
      UBCustRec(1).Book = "00"
      
      IF ThisBook < 10 THEN
        UBCustRec(1).Book = "0" + LTRIM$(STR$(ThisBook))
      ELSE
        UBCustRec(1).Book = LTRIM$(STR$(ThisBook))
      END IF
        
      SEQ$ = QPTrim$(STR$(SeqNum))

      SeqLen = LEN(SEQ$)
      MID$(UBCustRec(1).SeqNumb, 7 - SeqLen) = SEQ$
      
      TempBook$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb

      REDIM UBBookSeq(1) AS BookSeqRecType
      UBBookSeq(1).BookSeq = QPValL(TempBook$)
      BookHand = FREEFILE
      OPEN "UBOOKSEQ.DAT" FOR RANDOM AS BookHand LEN = 4
      NextBookRec = (LOF(BookHand) / 4) + 1
      PUT BookHand, NextBookRec, UBBookSeq(1)
      CLOSE BookHand
      
      UBCustRec(1).TAXEXPT = "N"
      UBCustRec(1).SRCIT = "N"
      UBCustRec(1).EPP = "N"
      UBCustRec(1).EPPAMT = 0
      UBCustRec(1).CASHONLY = "N"
      UBCustRec(1).LATEFEE = "Y"
      UBCustRec(1).CUTOFFYN = "Y"
      
      UBCustRec(1).OpenDate = Date2Num(AODate$)

      Month = VAL(LEFT$(CurDate$, 2))
      IF Month < 10 THEN
        CurDate$ = "0" + CurDate$
      END IF
      Day = VAL(MID$(CurDate$, 4, 2))
      IF Day < 10 THEN
        CurDate$ = LEFT$(CurDate$, 3) + "0" + MID$(CurDate$, 4)
      END IF

      SELECT CASE Rate$
      CASE "O"
        UBCustRec(1).CUSTTYPE = "O"
        UBCustRec(1).Serv(1).RateCode = "WO"
        UBCustRec(1).Serv(1).RMTRTYPE = "C"
        UBCustRec(1).Serv(2).RateCode = "SO"
        UBCustRec(1).Serv(2).RMTRTYPE = "C"
      CASE "W"
      CASE "-"
      CASE "X"
      CASE " "
      CASE "I"
        UBCustRec(1).CUSTTYPE = "I"
        UBCustRec(1).Serv(1).RateCode = "WI"
        UBCustRec(1).Serv(1).RMTRTYPE = "C"
        UBCustRec(1).Serv(2).RateCode = "SI"
        UBCustRec(1).Serv(2).RMTRTYPE = "C"
        UBCustRec(1).Serv(3).RateCode = "G1"
      CASE ELSE
      END SELECT

      'UBCustRec(1).CUSTTYPE = WBCustRec.CUSTTYPE
      UBCustRec(1).Zone = Zone$
      UBCustRec(1).BILLCOPY = 1

      UBCustRec(1).BILLCYCL = VAL(BCycle$)
      UBCustRec(1).POSTRTE = PRoute$
      UBCustRec(1).SOSEC = SSN$
      UBCustRec(1).DRVLIC = DLicNo$

      UBCustRec(1).HPhone = HPhone$
      
      FOR MCnt = 1 TO 7
        UBCustRec(1).LocMeters(MCnt).CurDate = -32767
        UBCustRec(1).LocMeters(MCnt).InsDate = -32767
        UBCustRec(1).LocMeters(MCnt).CurRead = 0
        UBCustRec(1).LocMeters(MCnt).PrevRead = 0
        UBCustRec(1).LocMeters(MCnt).AvgUse = 0
        UBCustRec(1).LocMeters(MCnt).UseCnt = 0
      NEXT

      UBCustRec(1).LocMeters(1).MTRNUM = MSerNo$

      Month = VAL(LEFT$(MInDate$, 2))
      IF Month < 10 THEN
        MInDate$ = "0" + MInDate$
      END IF
      Day = VAL(MID$(MInDate$, 4, 2))
      IF Day < 10 THEN
        MInDate$ = LEFT$(MInDate$, 3) + "0" + MID$(MInDate$, 4)
      END IF

      UBCustRec(1).LocMeters(1).InsDate = Date2Num(MInDate$)
      UBCustRec(1).LocMeters(1).CurDate = Date2Num(CurDate$)
      UBCustRec(1).LocMeters(1).PastDate = UBCustRec(1).LocMeters(1).CurDate
      UBCustRec(1).LocMeters(1).CurRead = QPValL(CurRead$)
      UBCustRec(1).LocMeters(1).PrevRead = UBCustRec(1).LocMeters(1).CurRead
      UBCustRec(1).LocMeters(1).MTRMulti = QPValI(Multi$)
      UBCustRec(1).LocMeters(1).MTRType = "C"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"
      UBCustRec(1).LocMeters(1).NumUser = 1

      FOR FRCnt = 1 TO 4
        UBCustRec(1).FlatRates(1).FRDESC = ""
        UBCustRec(1).FlatRates(1).FRAMT = 0
      NEXT

      FormC$(0, 0) = STRING$(CustFormLen, 0)
      BCopy VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), SSEG(FormC$(0, 0)), SADD(FormC$(0, 0)), CustFormLen, 0
      SaveNewCustRec FormC$(), CustRec&
      LOCATE 5, 1
      PRINT UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb, CustRec&;
      a$ = INKEY$
      IF a$ = CHR$(27) THEN EXIT DO
  LOOP UNTIL EOF(WadesHandle)

  CLOSE


END

      
      
      

