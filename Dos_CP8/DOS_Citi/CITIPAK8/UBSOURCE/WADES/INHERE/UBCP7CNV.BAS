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
  
  '$INCLUDE: 'CP71UTIL.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  DIM WBCustRec  AS Cust71RecType
  DIM WBCustRec1  AS Cust71RecType1
  DIM WBCustRec2  AS Cust71RecType2

  DIM WBTransRec AS WBTransRecType
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTransRec(1) AS UBTransRecType

  CLS
  
  
  WBCustRecLen = LEN(WBCustRec)
  WBCust1RecLen = LEN(WBCustRec1)
  WBCust2RecLen = LEN(WBCustRec2)

  UBCustRecLen = LEN(UBCustRec(1))

  WBTransRecLen = LEN(WBTransRec)
  UBTransRecLen = LEN(UBTransRec(1))

'GOTO Skip2Here
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = WBCustRecLen
  CustHandle1 = FREEFILE
  OPEN "wbcust1.dat" FOR RANDOM AS CustHandle1 LEN = WBCust1RecLen
  CustHandle2 = FREEFILE
  OPEN "wbcust2.dat" FOR RANDOM AS CustHandle2 LEN = WBCust2RecLen


  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  
    FOR Cnt = 1 TO NumOfRecs
      
      GET #CustHandle, Cnt, WBCustRec
      GET #CustHandle1, Cnt, WBCustRec1
      GET #CustHandle2, Cnt, WBCustRec2
    

      LOCATE 1, 1: PRINT "Converting Customer Record "; Cnt; " Of "; NumOfRecs
      SEQNUMB$ = RTRIM$(MID$(WBCustRec.meter, 4, 6))
      kk = INSTR(SEQNUMB$, "/")
      IF kk > 0 THEN
       SEQNUMB$ = LEFT$(SEQNUMB$, kk - 1)
      END IF
      IF LEN(SEQNUMB$) < 6 THEN SEQNUMB$ = STRING$(6 - LEN(SEQNUMB$), "0") + SEQNUMB$
      IF VAL(SEQNUMB$) = 0 THEN SEQNUMB$ = "      "
      IF LEN(QPTrim$(WBCustRec.FIRSTNAME)) = 0 THEN
       CUSTNAME$ = QPTrim$(WBCustRec.LNAME)
      ELSE
       CUSTNAME$ = QPTrim$(WBCustRec.FIRSTNAME) + " " + WBCustRec.LNAME
      END IF
    UBCustRec(1).BOOK = LEFT$(WBCustRec.meter, 2)
    UBCustRec(1).SEQNUMB = SEQNUMB$
    IF VAL(LEFT$(WBCustRec.meter, 2)) = 0 AND VAL(SEQNUMB$) = 0 THEN
     UBCustRec(1).STATUS = "I"
     ELSE
     UBCustRec(1).STATUS = WBCustRec.cycle
    END IF

    UBCustRec(1).OPENDATE = Date2Num%(MINSTALL$)
    UBCustRec(1).SEARCH = WBCustRec.LNAME
    UBCustRec(1).CUSTNAME = CUSTNAME$
    UBCustRec(1).ADDR1 = WBCustRec.add1
    UBCustRec(1).ADDR2 = WBCustRec.Add3
    UBCustRec(1).SERVADDR = WBCustRec.add2
    UBCustRec(1).CITY = WBCustRec.CITY
    UBCustRec(1).STATE = WBCustRec.st
    UBCustRec(1).ZIPCODE = WBCustRec.zip
    UBCustRec(1).HPHONE = WBCustRec.phone
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = WBCustRec.CUSTTYPE
    UBCustRec(1).Addr911 = WBCustRec.MeterNum
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = ""
    UBCustRec(1).SEQ = 0
    UBCustRec(1).CASHONLY = "N"
    UBCustRec(1).LATEFEE = "Y"
    UBCustRec(1).CUTOFFYN = "Y"
    UBCustRec(1).TAXEXPT = "N"
    UBCustRec(1).SRCIT = "N"
    UBCustRec(1).EPP = "N"
    UBCustRec(1).EPPAMT = 0
    UBCustRec(1).USEDRAFT = "N"

    UBCustRec(1).BANKNAME = ""
    UBCustRec(1).BANKLOC = ""
    UBCustRec(1).TRANSIT = ""
    UBCustRec(1).BANKACCT = ""

    UBCustRec(1).BILLCMNT = ""
    UBCustRec(1).PAYCMNT = ""
    UBCustRec(1).PUMPCODE = ""
    UBCustRec(1).USERCODE1 = ""
    UBCustRec(1).USERCODE2 = ""
    UBCustRec(1).HHMSG1 = ""
    UBCustRec(1).HHMSG2 = ""
    UBCustRec(1).HHMSG3 = ""

    UBCustRec(1).Serv(1).RateCode = WBCustRec.Rate
    UBCustRec(1).Serv(2).RateCode = WBCustRec.SewerRate
    UBCustRec(1).Serv(3).RateCode = WBCustRec.SprinklerRate
    UBCustRec(1).Serv(4).RateCode = ""
    UBCustRec(1).Serv(5).RateCode = ""
    UBCustRec(1).Serv(6).RateCode = ""
    UBCustRec(1).Serv(7).RateCode = ""
    UBCustRec(1).Serv(8).RateCode = ""
    UBCustRec(1).Serv(1).RMTRTYPE = "C"
    UBCustRec(1).Serv(2).RMTRTYPE = ""
    UBCustRec(1).Serv(3).RMTRTYPE = ""
    UBCustRec(1).Serv(4).RMTRTYPE = ""
    UBCustRec(1).Serv(5).RMTRTYPE = ""
    UBCustRec(1).Serv(6).RMTRTYPE = ""
    UBCustRec(1).Serv(7).RMTRTYPE = ""
    UBCustRec(1).Serv(8).RMTRTYPE = ""
    FOR ll = 9 TO 15
     UBCustRec(1).Serv(ll).RateCode = ""
     UBCustRec(1).Serv(ll).RMTRTYPE = ""
    NEXT ll


    FOR ll = 1 TO 4
     UBCustRec(1).FlatRates(ll).FRDESC = ""
     UBCustRec(1).FlatRates(ll).FRAMT = 0
     UBCustRec(1).FlatRates(ll).FRFREQ = ""
     UBCustRec(1).FlatRates(ll).REVSRC = 0
     UBCustRec(1).FlatRates(ll).NumMin = 0
    NEXT ll

    UBCustRec(1).Monthly(1).AMTOWED = 0
    UBCustRec(1).Monthly(1).TotAmtPD = 0
    UBCustRec(1).Monthly(1).PayAmt = 0
    UBCustRec(1).Monthly(1).RevSource = 0

    UBCustRec(1).Monthly(2).AMTOWED = 0
    UBCustRec(1).Monthly(2).TotAmtPD = 0
    UBCustRec(1).Monthly(2).PayAmt = 0
    UBCustRec(1).Monthly(2).RevSource = 0
    

    UBCustRec(1).MFEE1 = 0
    UBCustRec(1).MFEE2 = 0
    
    FOR ll = 1 TO 7
    UBCustRec(1).LocMeters(ll).MTRNUM = ""
    UBCustRec(1).LocMeters(ll).MTRMulti = 0
    UBCustRec(1).LocMeters(ll).MTRType = ""
    UBCustRec(1).LocMeters(ll).MTRUnit = ""
    UBCustRec(1).LocMeters(ll).NumUser = 1
    UBCustRec(1).LocMeters(ll).InsDate = -32767
    UBCustRec(1).LocMeters(ll).CurRead = 0
    UBCustRec(1).LocMeters(ll).PrevRead = 0
    UBCustRec(1).LocMeters(ll).CurDate = -32767
    UBCustRec(1).LocMeters(ll).PastDate = -32767
    UBCustRec(1).LocMeters(ll).ReadFlag = "N"
    UBCustRec(1).LocMeters(ll).AvgUse = 0
    UBCustRec(1).LocMeters(ll).UseCnt = 1
    NEXT ll
  
'END OF Quick Screen Form
    UBCustRec(1).CustPIN = Cnt
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = WBCustRec.curbal
    UBCustRec(1).PrevBalance = WBCustRec.pastdue
    UBCustRec(1).CurrRevAmts(1) = WBCustRec.Irev1
    UBCustRec(1).CurrRevAmts(2) = WBCustRec.Irev2
    UBCustRec(1).CurrRevAmts(3) = WBCustRec.Irev3
    UBCustRec(1).CurrRevAmts(4) = WBCustRec.Irev4
    UBCustRec(1).CurrRevAmts(5) = WBCustRec.Irev5
    UBCustRec(1).CurrRevAmts(6) = WBCustRec.Irev6
    UBCustRec(1).CurrRevAmts(7) = WBCustRec.Irev7
    UBCustRec(1).CurrRevAmts(8) = WBCustRec.Irev8

    FOR ll = 9 TO 15
     UBCustRec(1).CurrRevAmts(ll) = 0
    NEXT ll

    FOR ll = 1 TO 15
     UBCustRec(1).PrevRevAmts(ll) = 0
    NEXT ll

    UBCustRec(1).DepositAmt = WBCustRec.waterdep
    UBCustRec(1).DelFlag = 0
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
    PUT UBFile, Cnt, UBCustRec(1)
    NEXT Cnt
    CLOSE

  TransHandle = FREEFILE
  OPEN "wbartran.dat" FOR RANDOM AS TransHandle LEN = 128

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = LEN(UBTransRec(1))

  FOR Cnt = 1 TO LOF(TransHandle) \ 128
  LOCATE 2, 1: PRINT "Converting Transaction Record "; Cnt; " Of "; LOF(TransHandle) \ 128
   GET TransHandle, Cnt, WBTransRec

   UBTransRec(1).TransDate = Date2Num%(WBTransRec.TDATE)
   IF WBTransRec.TTYPE = 3 THEN
     UBTransRec(1).TransType = 104           ' Make provision for payment
    ELSE
     UBTransRec(1).TransType = WBTransRec.TTYPE + 100
   END IF
   UBTransRec(1).TransDesc = WBTransRec.TDESC
   UBTransRec(1).TransAmt = WBTransRec.curamt
   UBTransRec(1).RevAmt(1) = WBTransRec.wbrev1
   UBTransRec(1).RevAmt(2) = WBTransRec.wbrev2
   UBTransRec(1).RevAmt(3) = WBTransRec.wbrev3
   UBTransRec(1).RevAmt(4) = WBTransRec.wbrev4
   UBTransRec(1).RevAmt(5) = WBTransRec.wbrev5
   UBTransRec(1).RevAmt(6) = WBTransRec.wbrev6
   UBTransRec(1).RevAmt(7) = WBTransRec.wbrev7
   UBTransRec(1).RevAmt(8) = WBTransRec.wbrev8

   FOR ll = 9 TO 15
    UBTransRec(1).RevAmt(ll) = 0
   NEXT ll

   FOR ll = 1 TO 15
    UBTransRec(1).TaxAmt(ll) = 0
   NEXT ll

  FOR ll = 1 TO 7
   UBTransRec(1).MtrTypes(ll) = 0
  NEXT ll
  UBTransRec(1).CurRead(1) = 0
  UBTransRec(1).PrevRead(1) = 0
  
  FOR ll = 2 TO 7
   UBTransRec(1).CurRead(ll) = 0
   UBTransRec(1).PrevRead(ll) = 0
  NEXT ll
  FOR ll = 1 TO 7
   UBTransRec(1).EstRead(ll) = "N"
  NEXT ll
  UBTransRec(1).BillNumber = 0
  UBTransRec(1).ReadDate = -32767
  UBTransRec(1).BillDate = -32767
  UBTransRec(1).PastDueDate = -32767
  UBTransRec(1).DraftDate = -32767
  UBTransRec(1).CustLocation = WBTransRec.tmrec
  UBTransRec(1).CustStatus = ""
  UBTransRec(1).CustCode = WBTransRec.TCUSTTYPE
  UBTransRec(1).PostRoute = ""
  UBTransRec(1).BillCycle = ""
  UBTransRec(1).PayTypeCode = 1
  UBTransRec(1).OperatorNumber = 0
  UBTransRec(1).CustAcctNo = WBTransRec.tmrec
  UBTransRec(1).PrevTrans = 0
  UBTransRec(1).NextTrans = 0
  UBTransRec(1).ActiveFlag = 0
  UBTransRec(1).RunBalance = WBTransRec.tamt
  UBTransRec(1).CheckAmount = 0
  UBTransRec(1).CashAmount = 0
  UBTransRec(1).BillMsg = ""
  UBTransRec(1).ApplyDepFlag = "N"
  UBTransRec(1).Posted2GL = "N"
  UBTransRec(1).PrevDate = -32767
  UBTransRec(1).Pad = ""

  PUT UBFile, Cnt, UBTransRec(1)
  NEXT Cnt
  CLOSE
  END
    

