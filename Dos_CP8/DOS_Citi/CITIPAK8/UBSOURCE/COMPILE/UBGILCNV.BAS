DEFINT A-Z

DECLARE SUB SaveNewLocaRec (Form$(), LocatRec&, CustRec&)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SaveNewCustRec (FormC$(), NewRecNum&)
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION Num2Date$ (Daty%)
DECLARE FUNCTION Round# (N#)

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
  DIM WBCUSTREC1  AS WBCust1RecType
  DIM wbcustrec2  AS WBCust2RecType
  DIM WBOwner AS OwnerRecType


  DIM WBTransRec AS WBTransRecType
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTransRec(1) AS UBTransRecType

  CLS
  
  
  WBCustRecLen = LEN(WBCustRec)
  WBCust1RecLen = LEN(WBCUSTREC1)
  WBCust2RecLen = LEN(wbcustrec2)

  UBCustRecLen = LEN(UBCustRec(1))

  WBTransRecLen = LEN(WBTransRec)
  UBTransRecLen = LEN(UBTransRec(1))

'GOTO Skip2Here
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = 256
  CustHandle1 = FREEFILE
  OPEN "wbcust1.dat" FOR RANDOM AS CustHandle1 LEN = 256
  CustHandle2 = FREEFILE
  OPEN "wbcust2.dat" FOR RANDOM AS CustHandle2 LEN = 64



  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ 256
  
  
    FOR Cnt! = 1 TO NumOfRecs
      LOCATE 3, 1: PRINT "Customer # "; Cnt!; " of "; NumOfRecs
      GET #CustHandle, Cnt!, WBCustRec
      GET #CustHandle1, Cnt!, WBCUSTREC1
      GET #CustHandle2, Cnt!, wbcustrec2

    'For Gilbert book is 1st 2 digits & sequence is the 6 digits starting in location 5
    UBCustRec(1).Book = LEFT$(WBCustRec.meter, 2)
    UBCustRec(1).SeqNumb = MID$(WBCustRec.meter, 5, 6)
    kk = INSTR(UBCustRec(1).SeqNumb, "/")
    IF kk > 1 THEN UBCustRec(1).SeqNumb = LEFT$(UBCustRec(1).SeqNumb, kk - 1)

    IF LEN(RTRIM$(UBCustRec(1).SeqNumb)) < 6 THEN
     UBCustRec(1).SeqNumb = STRING$(6 - LEN(RTRIM$(UBCustRec(1).SeqNumb)), "0") + UBCustRec(1).SeqNumb
    END IF


    UBCustRec(1).SEARCH = UCASE$(WBCustRec.LNAME)
    UBCustRec(1).STATUS = WBCustRec.cycle

    CustomerName$ = UCASE$(RTRIM$(LTRIM$(WBCustRec.FIRSTNAME)))
    IF LEN(CustomerName$) = 0 THEN
     CustomerName$ = UCASE$(WBCustRec.LNAME)
     ELSE
     CustomerName$ = CustomerName$ + " " + UCASE$(WBCustRec.LNAME)
    END IF
    
    UBCustRec(1).CUSTNAME = CustomerName$
    UBCustRec(1).ADDR1 = UCASE$(WBCustRec.add1)
    UBCustRec(1).ADDR2 = ""
    UBCustRec(1).SERVADDR = UCASE$(WBCustRec.add2)
    UBCustRec(1).CITY = UCASE$(WBCustRec.CITY)
    UBCustRec(1).STATE = UCASE$(WBCustRec.st)
    UBCustRec(1).ZIPCODE = WBCustRec.zip
    PHONENUMBER$ = QPTrim$(WBCustRec.PHONE)
    IF LEN(PHONENUMBER$) < 10 THEN PHONENUMBER$ = STRING$(10 - LEN(PHONENUMBER$), " ") + PHONENUMBER$

    UBCustRec(1).HPHONE = PHONENUMBER$
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    IF WBCUSTREC1.CUSTTYPE = 1 THEN
     UBCustRec(1).CUSTTYPE = "R"
    ELSE
     UBCustRec(1).CUSTTYPE = "C"
    END IF

    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = WBCustRec.WON
    UBCustRec(1).SEQ = 0
    UBCustRec(1).CASHONLY = "N"
    UBCustRec(1).LATEFEE = "Y"
    UBCustRec(1).CUTOFFYN = "Y"
    UBCustRec(1).TAXEXPT = "N"
    UBCustRec(1).SRCIT = "N"
    UBCustRec(1).EPP = "N"
    UBCustRec(1).EPPAMT = 0
    UBCustRec(1).USEDRAFT = ""
    UBCustRec(1).BANKNAME = ""
    UBCustRec(1).BANKLOC = ""
    UBCustRec(1).TRANSIT = ""
    UBCustRec(1).BANKACCT = ""
    UBCustRec(1).BILLCMNT = ""
    UBCustRec(1).PAYCMNT = ""
    UBCustRec(1).PUMPCODE = ""
    UBCustRec(1).USERCODE1 = ""
    UBCustRec(1).USERCODE2 = ""
    UBCustRec(1).HHMSG1 = WBCUSTREC1.HHMess
    UBCustRec(1).HHMSG2 = ""
    UBCustRec(1).HHMSG3 = ""

    FOR LL = 1 TO 15
     UBCustRec(1).Serv(LL).RateCode = ""
     UBCustRec(1).Serv(LL).RMTRTYPE = ""
    NEXT LL


    IF WBCustRec.Rate = 1 THEN
      UBCustRec(1).Serv(1).RateCode = "WR" + LTRIM$(STR$(WBCustRec.Rate))
      UBCustRec(1).Serv(1).RMTRTYPE = "T"
    END IF
    
     LL = 1
     IF WBCUSTREC1.ramt1 > 0 THEN
     UBCustRec(1).FlatRates(LL).FRDESC = WBCUSTREC1.RDesc1
     UBCustRec(1).FlatRates(LL).FRAMT = WBCUSTREC1.ramt1
     UBCustRec(1).FlatRates(LL).FRFREQ = "M"
     UBCustRec(1).FlatRates(LL).REVSRC = 2
     UBCustRec(1).FlatRates(LL).NumMin = 1
     ELSE
     UBCustRec(1).FlatRates(LL).FRDESC = ""
     UBCustRec(1).FlatRates(LL).FRAMT = 0
     UBCustRec(1).FlatRates(LL).FRFREQ = ""
     UBCustRec(1).FlatRates(LL).REVSRC = 0
     UBCustRec(1).FlatRates(LL).NumMin = 0
     END IF
    FOR LL = 2 TO 4
     UBCustRec(1).FlatRates(LL).FRDESC = ""
     UBCustRec(1).FlatRates(LL).FRAMT = 0
     UBCustRec(1).FlatRates(LL).FRFREQ = ""
     UBCustRec(1).FlatRates(LL).REVSRC = 0
     UBCustRec(1).FlatRates(LL).NumMin = 0
    NEXT LL

    FOR LL = 1 TO 2
     UBCustRec(1).Monthly(LL).AMTOWED = 0
     UBCustRec(1).Monthly(LL).TotAmtPD = 0
     UBCustRec(1).Monthly(LL).PayAmt = 0
     UBCustRec(1).Monthly(LL).RevSource = 0
    NEXT LL
    
    UBCustRec(1).MFEE1 = 0
    UBCustRec(1).MFEE2 = 0

    LL = 1
    UBCustRec(1).LocMeters(LL).MTRNUM = WBCUSTREC1.MeterNum
    UBCustRec(1).LocMeters(LL).MTRMulti = 1000
    UBCustRec(1).LocMeters(LL).MTRType = "T"
    UBCustRec(1).LocMeters(LL).MTRUnit = "G"
    UBCustRec(1).LocMeters(LL).NumUser = 1
    UBCustRec(1).LocMeters(LL).InsDate = Date2Num%(WBCustRec.MInstall)
    UBCustRec(1).LocMeters(LL).CurRead = WBCustRec.CurRead
    UBCustRec(1).LocMeters(LL).PrevRead = WBCustRec.PastRead
    UBCustRec(1).LocMeters(LL).CurDate = -32767
    UBCustRec(1).LocMeters(LL).PastDate = -32767
    UBCustRec(1).LocMeters(LL).ReadFlag = "N"
    UBCustRec(1).LocMeters(LL).AvgUse = 0
    UBCustRec(1).LocMeters(LL).UseCnt = 0

    FOR LL = 2 TO 7
    UBCustRec(1).LocMeters(LL).MTRNUM = ""
    UBCustRec(1).LocMeters(LL).MTRMulti = 1
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
    UBCustRec(1).LocMeters(LL).UseCnt = 0
    NEXT LL

    UBCustRec(1).CustPIN = 0
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = WBCustRec.CurBal
    UBCustRec(1).PrevBalance = WBCustRec.Pastdue

    AcctBal# = WBCustRec.CurBal + WBCustRec.Pastdue
    AcctBal# = INT((AcctBal# * 100) + .5) / 100
    IF AcctBal# < 0 THEN
    UBCustRec(1).CurrRevAmts(1) = AcctBal#
    UBCustRec(1).CurrRevAmts(2) = 0
    UBCustRec(1).CurrRevAmts(3) = 0
    UBCustRec(1).CurrRevAmts(4) = 0
    UBCustRec(1).CurrRevAmts(5) = 0
    UBCustRec(1).CurrRevAmts(6) = 0
    ELSE
    UBCustRec(1).CurrRevAmts(1) = wbcustrec2.REV1
    UBCustRec(1).CurrRevAmts(2) = wbcustrec2.REV2 + wbcustrec2.REV4
    UBCustRec(1).CurrRevAmts(3) = wbcustrec2.REV3
    UBCustRec(1).CurrRevAmts(4) = 0
    UBCustRec(1).CurrRevAmts(5) = wbcustrec2.REV5
    UBCustRec(1).CurrRevAmts(6) = wbcustrec2.REV6
    END IF

    FOR LL = 1 TO 6
    UBCustRec(1).PrevRevAmts(LL) = 0
    NEXT LL
    FOR LL = 7 TO 15
     UBCustRec(1).CurrRevAmts(LL) = 0
     UBCustRec(1).PrevRevAmts(LL) = 0
    NEXT LL

    UBCustRec(1).DepositAmt = WBCustRec.WATERDEP
   
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""

    IF VAL(LEFT$(WBCustRec.meter, 2)) <= 0 THEN
     UBCustRec(1).DelFlag = -1
     UBCustRec(1).STATUS = "I"
     ELSE
     UBCustRec(1).DelFlag = 0
    END IF
    IF UCASE$(WBCustRec.cycle) = "I" THEN
     UBCustRec(1).Book = ""
     UBCustRec(1).SeqNumb = ""
    END IF

PUT UBFile, Cnt!, UBCustRec(1)
    NEXT Cnt!
    CLOSE




    TransHandle = FREEFILE
    OPEN "wbartran.dat" FOR RANDOM AS TransHandle LEN = 50
    UBFile = FREEFILE
    OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = LEN(UBTransRec(1))

  FOR Cnt! = 1 TO LOF(TransHandle) \ 50
   LOCATE 2, 1: PRINT "Converting Transaction Record "; Cnt!; " Of "; LOF(TransHandle) \ 50
   GET TransHandle, Cnt!, WBTransRec
   GET UBFile, Cnt!, UBTransRec(1)

   TransRecord# = WBTransRec.TMREC
   IF TransRecord# < 1 OR TransRecord# > 99999 THEN GOTO NextTranaction
   IF WBTransRec.TTYPE = 1 THEN
     UBTransRec(1).TransType = 101
   END IF
   IF WBTransRec.TTYPE = 2 THEN
     UBTransRec(1).TransType = 102           ' Make provision for payment
   END IF

   IF WBTransRec.TTYPE = 3 THEN
    UBTransRec(1).TransType = 104
   END IF
   IF WBTransRec.TTYPE = 4 THEN
    UBTransRec(1).TransType = 104
   END IF
   IF WBTransRec.TTYPE = 5 AND WBTransRec.tcuramt > 0 THEN
    UBTransRec(1).TransType = 112
    ELSE
    UBTransRec(1).TransType = 111
   END IF
   IF WBTransRec.TTYPE = 6 THEN
    UBTransRec(1).TransType = 107
   END IF
   IF WBTransRec.TTYPE = 7 THEN
    UBTransRec(1).TransType = 104
   END IF
   UBTransRec(1).TransDate = Date2Num%(WBTransRec.TDATE)
   UBTransRec(1).TransDesc = WBTransRec.TDESC
   UBTransRec(1).TransAmt = WBTransRec.tcuramt
   UBTransRec(1).RevAmt(1) = 0
   UBTransRec(1).RevAmt(2) = 0
   UBTransRec(1).RevAmt(3) = 0
   UBTransRec(1).RevAmt(4) = 0
   UBTransRec(1).RevAmt(5) = 0
   UBTransRec(1).RevAmt(6) = 0
   UBTransRec(1).RevAmt(7) = 0
   UBTransRec(1).RevAmt(8) = 0
   FOR LL = 9 TO 15
    UBTransRec(1).RevAmt(LL) = 0
   NEXT LL

   FOR LL = 1 TO 15
    UBTransRec(1).TaxAmt(LL) = 0
   NEXT LL
   FOR LL = 1 TO 7
    UBTransRec(1).MtrTypes(LL) = 0
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
  UBTransRec(1).CustLocation = WBTransRec.TMREC
  UBTransRec(1).CustStatus = ""
  UBTransRec(1).CustCode = ""
  UBTransRec(1).PostRoute = ""
  UBTransRec(1).BillCycle = ""
  UBTransRec(1).PayTypeCode = 1
  UBTransRec(1).OperatorNumber = 0
  UBTransRec(1).CustAcctNo = WBTransRec.TMREC
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
  PUT UBFile, Cnt!, UBTransRec(1)
NextTranaction:
  NEXT Cnt!
  CLOSE
  END

