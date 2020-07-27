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
  
  DIM WBCUSTREC  AS Cust71RecType
  DIM WBCUSTREC1  AS WBCust1RecType
  'DIM WBCustRec2  AS Cust71RecType2

  
  REDIM UBCustRec(1) AS NewUBCustRecType
  'REDIM UBTransRec(1) AS UBTransRecType

  CLS
  
  
  WBCustRecLen = LEN(WBCUSTREC)
  WBCust1RecLen = LEN(WBCUSTREC1)

  WBCust2RecLen = LEN(WBCustRec2)

  UBCustRecLen = LEN(UBCustRec(1))


'GOTO Skip2Here
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = WBCustRecLen
  CustHandle1 = FREEFILE
  OPEN "wbcust1.dat" FOR RANDOM AS CustHandle1 LEN = 256
  'CustHandle2 = FREEFILE
  'OPEN "wbcust2.dat" FOR RANDOM AS CustHandle2 LEN = WBCust2RecLen


  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
    FOR Cnt = 1 TO NumOfRecs
      REDIM UBCustRec(1) AS NewUBCustRecType

      GET #CustHandle, Cnt, WBCUSTREC
      GET #CustHandle1, Cnt, WBCUSTREC1
      'GET #CustHandle2, Cnt, WBCustRec2
    

      LOCATE 1, 1: PRINT "Converting Customer Record "; Cnt; " Of "; NumOfRecs

    'UBCustRec(1).BOOK = "0" + LEFT$(WBCUSTREC.meter, 1)
    UBCustRec(1).SEQNUMB = MID$(WBCUSTREC.meter, 4, 6)
    kk = INSTR(UBCustRec(1).SEQNUMB, "/")
    IF kk > 1 THEN UBCustRec(1).SEQNUMB = LEFT$(UBCustRec(1).SEQNUMB, kk - 1)

    IF LEN(RTRIM$(UBCustRec(1).SEQNUMB)) < 6 THEN
     UBCustRec(1).SEQNUMB = STRING$(6 - LEN(RTRIM$(UBCustRec(1).SEQNUMB)), "0") + UBCustRec(1).SEQNUMB
    END IF


   '   SEQNUMB$ = RTRIM$(MID$(WBCUSTREC.meter, 4, 6))
   '   kk = INSTR(SEQNUMB$, "/")
   '   IF kk > 0 THEN
   '    SEQNUMB$ = LEFT$(SEQNUMB$, kk - 1)
   '   END IF
   '   IF LEN(SEQNUMB$) < 6 THEN SEQNUMB$ = STRING$(6 - LEN(SEQNUMB$), "0") + SEQNUMB$
   '   IF VAL(SEQNUMB$) = 0 THEN SEQNUMB$ = "      "
   '
    UBCustRec(1).BOOK = LEFT$(WBCUSTREC.meter, 2)
'    IF LEN(LTRIM$(RTRIM$(UBCustRec(1).BOOK))) = 0 THEN
'      UBCustRec(1).SEQNUMB = ""
'    ELSE
'      SEQ& = SEQ& + 100
'      TSeq$ = "000000" + LTRIM$(STR$(SEQ&))
'      UBCustRec(1).SEQNUMB = RIGHT$(TSeq$, 6)
'    END IF

      IF LEN(QPTrim$(WBCUSTREC.FIRSTNAME)) = 0 THEN
       CUSTNAME$ = QPTrim$(WBCUSTREC.LNAME)
      ELSE
       CUSTNAME$ = QPTrim$(WBCUSTREC.FIRSTNAME) + " " + WBCUSTREC.LNAME
      END IF

    IF VAL(LEFT$(WBCUSTREC.meter, 2)) = 0 AND VAL(SEQNUMB$) = 0 THEN
      UBCustRec(1).STATUS = "I"
      UBCustRec(1).SEQNUMB = ""
      UBCustRec(1).BOOK = ""
    ELSE
      UBCustRec(1).STATUS = WBCUSTREC.cycle
      IF UBCustRec(1).STATUS = "I" THEN
        UBCustRec(1).SEQNUMB = ""
        UBCustRec(1).BOOK = ""
      END IF
    END IF

    UBCustRec(1).OPENDATE = Date2Num%(MInstall$)
    UBCustRec(1).SEARCH = WBCUSTREC.LNAME
    UBCustRec(1).CUSTNAME = CUSTNAME$
    UBCustRec(1).ADDR1 = WBCUSTREC.add1
    UBCustRec(1).ADDR2 = WBCUSTREC1.add3
    UBCustRec(1).SERVADDR = WBCUSTREC.add2
    UBCustRec(1).CITY = WBCUSTREC.CITY
    UBCustRec(1).STATE = WBCUSTREC.st
    UBCustRec(1).ZIPCODE = WBCUSTREC.zip
    UBCustRec(1).HPHONE = WBCUSTREC.PHONE
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    IF WBCUSTREC1.CUSTTYPE = 1 THEN
      ResCom$ = "R"
    ELSE
      ResCom$ = "C"
    END IF
    UBCustRec(1).CUSTTYPE = ResCom$ + LTRIM$(WBCUSTREC.WON)
    'UBCustRec(1).Addr911 = WBCustRec.MeterNum
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


    FOR zzcnt = 1 TO 7
      UBCustRec(1).LocMeters(zzcnt).NumUser = -32767
      UBCustRec(1).LocMeters(zzcnt).InsDate = -32767
      UBCustRec(1).LocMeters(zzcnt).CurRead = -2147483647
      UBCustRec(1).LocMeters(zzcnt).PrevRead = -2147483647
      UBCustRec(1).LocMeters(zzcnt).CurDate = -32767
      UBCustRec(1).LocMeters(zzcnt).PastDate = -32767
      UBCustRec(1).LocMeters(zzcnt).ReadFlag = ""
      UBCustRec(1).LocMeters(zzcnt).AvgUse = 0
      UBCustRec(1).LocMeters(zzcnt).UseCnt = 0
    NEXT
    
    SELECT CASE WBCUSTREC.RATE
    CASE 1     'water & sewer
      UBCustRec(1).Serv(1).RateCode = "WAT1"
      UBCustRec(1).Serv(1).RMTRTYPE = "C"
      UBCustRec(1).Serv(2).RateCode = "SEW1"
      UBCustRec(1).Serv(2).RMTRTYPE = "C"
      UBCustRec(1).LocMeters(1).NumUser = 1
      UBCustRec(1).LocMeters(1).MTRNUM = WBCUSTREC1.meterNum
      UBCustRec(1).LocMeters(1).MTRMulti = WBCUSTREC1.minimum
      UBCustRec(1).LocMeters(1).CurRead = WBCUSTREC.CurRead
      UBCustRec(1).LocMeters(1).PrevRead = WBCUSTREC.PastRead
      UBCustRec(1).LocMeters(1).MTRType = "C"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"

    CASE 2     'water only
      UBCustRec(1).Serv(1).RateCode = "WAT2"
      UBCustRec(1).Serv(1).RMTRTYPE = "W"
      UBCustRec(1).LocMeters(1).NumUser = 1
      UBCustRec(1).LocMeters(1).MTRNUM = WBCUSTREC1.meterNum
      UBCustRec(1).LocMeters(1).MTRMulti = WBCUSTREC1.minimum
      UBCustRec(1).LocMeters(1).CurRead = WBCUSTREC.CurRead
      UBCustRec(1).LocMeters(1).PrevRead = WBCUSTREC.PastRead
      UBCustRec(1).LocMeters(1).MTRType = "W"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"

    CASE 3     'water & sewer discount
      UBCustRec(1).Serv(1).RateCode = "WAT3"
      UBCustRec(1).Serv(1).RMTRTYPE = "C"
      UBCustRec(1).Serv(2).RateCode = "SEW3"
      UBCustRec(1).Serv(2).RMTRTYPE = "C"
      UBCustRec(1).LocMeters(1).NumUser = 1
      UBCustRec(1).LocMeters(1).MTRNUM = WBCUSTREC1.meterNum
      UBCustRec(1).LocMeters(1).MTRMulti = WBCUSTREC1.minimum
      UBCustRec(1).LocMeters(1).CurRead = WBCUSTREC.CurRead
      UBCustRec(1).LocMeters(1).PrevRead = WBCUSTREC.PastRead
      UBCustRec(1).LocMeters(1).MTRType = "C"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"

    CASE ELSE  '???

    END SELECT
    

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
    
  
'END OF Quick Screen Form
    UBCustRec(1).CustPIN = Cnt
    UBCustRec(1).LastTrans = 0
    'UBCustRec(1).CurrBalance = WBCustRec.CURBAL
    'UBCustRec(1).PrevBalance = WBCustRec.Pastdue

    FOR ll = 9 TO 15
     UBCustRec(1).CurrRevAmts(ll) = 0
    NEXT ll

    FOR ll = 1 TO 15
      UBCustRec(1).PrevRevAmts(ll) = 0
    NEXT ll

    UBCustRec(1).DepositAmt = 0
    UBCustRec(1).DelFlag = 0
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
    PUT UBFile, Cnt, UBCustRec(1)
    NEXT Cnt
    CLOSE

  END
    

