
DEFINT A-Z

DECLARE SUB SaveNewLocaRec (Form$(), LocatRec&, CustRec&)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SaveNewCustRec (FormC$(), NewRecNum&)
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION Num2Date$ (Daty%)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()

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
  
  '$INCLUDE: 'CP72UTIL.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  DIM WBCUSTREC  AS Cust72RecType
  DIM WBCustRec1  AS Cust72RecType1
  DIM WBHH AS WBCUSTHHRECTYPE

  DIM WBTransRec AS WBTransRecType
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTransRec(1) AS UBTransRecType

  CLS
  
  
  WBCustRecLen = 512
  WBCust1RecLen = 256
  

  UBCustRecLen = LEN(UBCustRec(1))

  WBTransRecLen = LEN(WBTransRec)
  UBTransRecLen = LEN(UBTransRec(1))

'GOTO Skip2Here
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = WBCustRecLen
  CustHandle1 = FREEFILE
  OPEN "wbcust1.dat" FOR RANDOM AS CustHandle1 LEN = WBCust1RecLen
  custhandle2 = FREEFILE
  OPEN "wbcusthh.dat" FOR RANDOM AS custhandle2 LEN = 60


  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  
    FOR cnt = 1 TO NumOfRecs
      
  
      GET #CustHandle, cnt, WBCUSTREC
      GET #CustHandle1, cnt, WBCustRec1
      GET #custhandle2, cnt, WBHH
      
      LOCATE 1, 1: PRINT "Converting Customer Record "; cnt; " Of "; NumOfRecs

      UBCustRec(1).DelFlag = 0

      SEQNUMB$ = RTRIM$(MID$(WBCUSTREC.meter, 4, 6))
      kk = INSTR(SEQNUMB$, "/")
      IF kk > 0 THEN
       SEQNUMB$ = LEFT$(SEQNUMB$, kk - 1)
      END IF
      IF LEN(SEQNUMB$) < 6 THEN SEQNUMB$ = STRING$(6 - LEN(SEQNUMB$), "0") + SEQNUMB$
      IF VAL(SEQNUMB$) = 0 THEN SEQNUMB$ = "      "
      IF LEN(QPTrim$(WBCUSTREC.FIRSTNAME)) = 0 THEN
       CUSTNAME$ = QPTrim$(WBCUSTREC.LNAME)
      ELSE
       CUSTNAME$ = QPTrim$(WBCUSTREC.FIRSTNAME) + " " + WBCUSTREC.LNAME
      END IF

    UBCustRec(1).BOOK = LEFT$(WBCUSTREC.meter, 2)
    UBCustRec(1).SEQNUMB = SEQNUMB$

    IF VAL(LEFT$(WBCUSTREC.meter, 2)) = 0 AND VAL(SEQNUMB$) = 0 THEN
     UBCustRec(1).STATUS = "I"
     UBCustRec(1).DelFlag = -1
     UBCustRec(1).BOOK = ""
     UBCustRec(1).SEQNUMB = ""
     ELSE
     UBCustRec(1).STATUS = WBCUSTREC.cycle
    END IF

    IF UBCustRec(1).STATUS = "I" THEN
    IF LEN(QPTrim$(rate1$)) = 0 AND LEN(QPTrim$(rate2$)) = 0 THEN
     UBCustRec(1).DelFlag = -1
     UBCustRec(1).BOOK = ""
     UBCustRec(1).SEQNUMB = ""
    END IF
    END IF

    
    UBCustRec(1).OPENDATE = Date2Num(WBCUSTREC.minstall)
    UBCustRec(1).SEARCH = WBCUSTREC.LNAME
    UBCustRec(1).CUSTNAME = CUSTNAME$
    UBCustRec(1).ADDR1 = WBCUSTREC.add1
    UBCustRec(1).ADDR2 = WBCUSTREC.Add3
    UBCustRec(1).SERVADDR = WBCUSTREC.add2
    UBCustRec(1).CITY = WBCUSTREC.CITY
    UBCustRec(1).STATE = WBCUSTREC.st
    UBCustRec(1).ZIPCODE = WBCUSTREC.zip
    UBCustRec(1).HPHONE = WBCUSTREC.PHONE
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = WBCUSTREC.CUSTTYPE
    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = WBCUSTREC.won
    UBCustRec(1).SEQ = WBCUSTREC.SeqNum
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

    UBCustRec(1).Serv(1).RateCode = WBCUSTREC.rate1
    UBCustRec(1).Serv(2).RateCode = WBCUSTREC.rate2
    UBCustRec(1).Serv(3).RateCode = ""
    UBCustRec(1).Serv(4).RateCode = ""
    UBCustRec(1).Serv(5).RateCode = ""
    UBCustRec(1).Serv(6).RateCode = ""
    UBCustRec(1).Serv(7).RateCode = ""
    UBCustRec(1).Serv(8).RateCode = ""

    UBCustRec(1).Serv(1).RMTRTYPE = "W"
    UBCustRec(1).Serv(2).RMTRTYPE = ""
    UBCustRec(1).Serv(3).RMTRTYPE = ""
    UBCustRec(1).Serv(4).RMTRTYPE = ""
    UBCustRec(1).Serv(5).RMTRTYPE = ""
    UBCustRec(1).Serv(6).RMTRTYPE = ""
    UBCustRec(1).Serv(7).RMTRTYPE = ""
    UBCustRec(1).Serv(8).RMTRTYPE = ""
    FOR LL = 9 TO 15
     UBCustRec(1).Serv(LL).RateCode = ""
     UBCustRec(1).Serv(LL).RMTRTYPE = ""
    NEXT LL

     UBCustRec(1).FlatRates(1).FRDESC = ""
     UBCustRec(1).FlatRates(1).FRAMT = 0
     UBCustRec(1).FlatRates(1).FRFREQ = ""
     UBCustRec(1).FlatRates(1).REVSRC = 0
     UBCustRec(1).FlatRates(1).NumMin = 0

    FOR LL = 2 TO 4
     UBCustRec(1).FlatRates(LL).FRDESC = ""
     UBCustRec(1).FlatRates(LL).FRAMT = 0
     UBCustRec(1).FlatRates(LL).FRFREQ = ""
     UBCustRec(1).FlatRates(LL).REVSRC = 0
     UBCustRec(1).FlatRates(LL).NumMin = 0
    NEXT LL

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
    LL = 1
    UBCustRec(1).LocMeters(LL).MTRNUM = WBCUSTREC.MeterNum
    UBCustRec(1).LocMeters(LL).MTRMulti = 1
    UBCustRec(1).LocMeters(LL).MTRType = "W"
    UBCustRec(1).LocMeters(LL).MTRUnit = "G"
    UBCustRec(1).LocMeters(LL).NumUser = WBCUSTREC.minimums
    UBCustRec(1).LocMeters(LL).InsDate = -32767
    UBCustRec(1).LocMeters(LL).CurRead = WBCUSTREC.CurRead
    UBCustRec(1).LocMeters(LL).PrevRead = WBCUSTREC.PastRead
    IF VAL(WBCUSTREC.MeterReadDate) > 0 THEN
     ReadDate$ = LEFT$(WBCUSTREC.MeterReadDate, 4) + "97"
     ELSE
     ReadDate$ = ""
    END IF
    UBCustRec(1).LocMeters(LL).CurDate = Date2Num(ReadDate$)
    UBCustRec(1).LocMeters(LL).PastDate = -32767
    UBCustRec(1).LocMeters(LL).ReadFlag = "N"
    UBCustRec(1).LocMeters(LL).AvgUse = WBCUSTREC.AvgUse
    UBCustRec(1).LocMeters(LL).UseCnt = WBCUSTREC.AVGCT
    
    FOR LL = 2 TO 7
    UBCustRec(1).LocMeters(LL).MTRNUM = ""
    UBCustRec(1).LocMeters(LL).MTRMulti = 0
    UBCustRec(1).LocMeters(LL).MTRType = ""
    UBCustRec(1).LocMeters(LL).MTRUnit = ""
    UBCustRec(1).LocMeters(LL).NumUser = 1
    UBCustRec(1).LocMeters(LL).InsDate = -32767
    UBCustRec(1).LocMeters(LL).CurRead = 0
    UBCustRec(1).LocMeters(LL).PrevRead = 0
    UBCustRec(1).LocMeters(LL).CurDate = -32767
    UBCustRec(1).LocMeters(LL).PastDate = -32767
    UBCustRec(1).LocMeters(LL).ReadFlag = "N"
    UBCustRec(1).LocMeters(LL).AvgUse = 0
    UBCustRec(1).LocMeters(LL).UseCnt = 1
    NEXT LL
  
'END OF Quick Screen Form
    UBCustRec(1).CustPIN = cnt
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = WBCUSTREC.CURBAL
    UBCustRec(1).PrevBalance = WBCUSTREC.Pastdue
    UBCustRec(1).CurrRevAmts(1) = WBCUSTREC.REV1
    UBCustRec(1).CurrRevAmts(2) = WBCUSTREC.REV2
    UBCustRec(1).CurrRevAmts(3) = WBCUSTREC.REV3
    UBCustRec(1).CurrRevAmts(4) = WBCUSTREC.REV4
    UBCustRec(1).CurrRevAmts(5) = WBCUSTREC.REV5
    UBCustRec(1).CurrRevAmts(6) = WBCUSTREC.REV6
    UBCustRec(1).CurrRevAmts(7) = WBCUSTREC.rev7
    UBCustRec(1).CurrRevAmts(8) = WBCUSTREC.rev8

    FOR LL = 9 TO 15
     UBCustRec(1).CurrRevAmts(LL) = 0
    NEXT LL

    FOR LL = 1 TO 15
     UBCustRec(1).PrevRevAmts(LL) = 0
    NEXT LL

    UBCustRec(1).DepositAmt = WBCUSTREC.WATERDEP
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
endrecord:
    PUT UBFile, cnt, UBCustRec(1)
    NEXT cnt
    CLOSE

  TransHandle = FREEFILE
  OPEN "wbartran.dat" FOR RANDOM AS TransHandle LEN = 128

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = LEN(UBTransRec(1))

  FOR cnt! = 1 TO LOF(TransHandle) \ 128
  LOCATE 2, 1: PRINT "Converting Transaction Record "; cnt!; " Of "; LOF(TransHandle) \ 128
   GET TransHandle, cnt!, WBTransRec

   UBTransRec(1).TransDate = Date2Num%(WBTransRec.TDATE)
   IF WBTransRec.TTYPE = 3 THEN
     UBTransRec(1).TransType = 104           ' Make provision for payment
    ELSE
     UBTransRec(1).TransType = WBTransRec.TTYPE + 100
   END IF

   IF WBTransRec.TTYPE = 5 THEN
    IF WBTransRec.tcuramt < 0 THEN UBTransRec(1).TransType = 111
    IF WBTransRec.tcuramt > 0 THEN UBTransRec(1).TransType = 112
   END IF


   UBTransRec(1).TransDesc = WBTransRec.TDESC
   UBTransRec(1).TransAmt = WBTransRec.tcuramt
   UBTransRec(1).RevAmt(1) = WBTransRec.wbrev1
   UBTransRec(1).RevAmt(2) = WBTransRec.wbrev2
   UBTransRec(1).RevAmt(3) = WBTransRec.wbrev3
   UBTransRec(1).RevAmt(4) = WBTransRec.wbrev4
   UBTransRec(1).RevAmt(5) = WBTransRec.wbrev5
   UBTransRec(1).RevAmt(6) = WBTransRec.wbrev6
   UBTransRec(1).RevAmt(7) = WBTransRec.wbrev7
   UBTransRec(1).RevAmt(8) = WBTransRec.wbrev8

   FOR LL = 9 TO 15
    UBTransRec(1).RevAmt(LL) = 0
   NEXT LL

   FOR LL = 1 TO 15
    UBTransRec(1).TaxAmt(LL) = 0
   NEXT LL

  FOR LL = 1 TO 7
   UBTransRec(1).MtrTypes(LL) = 3
  NEXT LL
  UBTransRec(1).CurRead(1) = 0
  UBTransRec(1).PrevRead(1) = 0
  
  FOR LL = 2 TO 7
   UBTransRec(1).CurRead(LL) = 0
   UBTransRec(1).PrevRead(LL) = 0
  NEXT LL
  FOR LL = 1 TO 7
   UBTransRec(1).EstRead(LL) = "N"
  NEXT LL
  UBTransRec(1).BillNumber = 0
  UBTransRec(1).ReadDate = -32767
  UBTransRec(1).BillDate = -32767
  UBTransRec(1).PastDueDate = -32767
  UBTransRec(1).DraftDate = -32767
  UBTransRec(1).CustLocation = WBTransRec.tmrec
  UBTransRec(1).CustStatus = ""
  UBTransRec(1).CustCode = WBTransRec.tcusttype
  UBTransRec(1).PostRoute = ""
  UBTransRec(1).BillCycle = ""
  UBTransRec(1).PayTypeCode = 1
  UBTransRec(1).OperatorNumber = 0
  UBTransRec(1).CustAcctNo = WBTransRec.tmrec
  UBTransRec(1).PrevTrans = 0
  UBTransRec(1).NextTrans = 0
  UBTransRec(1).ActiveFlag = 0
  UBTransRec(1).RunBalance = WBTransRec.TAMT
  UBTransRec(1).CheckAmount = 0
  UBTransRec(1).CashAmount = 0
  UBTransRec(1).BillMsg = ""
  UBTransRec(1).ApplyDepFlag = "N"
  UBTransRec(1).Posted2GL = "N"
  UBTransRec(1).PrevDate = -32767
  UBTransRec(1).Pad = ""

  PUT UBFile, cnt!, UBTransRec(1)
  NEXT cnt!
  CLOSE
  END
    

