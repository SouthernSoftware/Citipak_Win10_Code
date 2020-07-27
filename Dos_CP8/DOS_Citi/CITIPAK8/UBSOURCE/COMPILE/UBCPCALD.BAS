DEFINT A-Z

DECLARE SUB SaveNewLocaRec (Form$(), LocatRec&, CustRec&)
DECLARE SUB SaveNewCustRec (FormC$(), NewRecNum&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()
DECLARE SUB SaveOldCustRec (Form$(), RecNo&)
DECLARE SUB CustLookUp (RecNo%)
DECLARE SUB SearchGetCust (SEARCH$, RecNo&, CLSFlag%)
DECLARE SUB PrintCustList ()
DECLARE SUB DisplayUBScrn (ScrnName$)
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
DECLARE SUB SaveSetUpRec (Form$())
DECLARE SUB LoadSetUpRec (Form$())
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB EditCust ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB ShowCursor ()
DECLARE SUB WaitForAction ()
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LoadCustRec (Form$(), RecNo&, LocatRec&, TFormLen%, LockedFlag%)
DECLARE SUB BlockClear ()
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION Num2Date$ (Daty%)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION Round# (DblNumber#)
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
  DIM WBCustRec1 AS Cust71RecType1
  DIM WBHHRec AS WBNoteType

  REDIM UBCustRec(1) AS NewUBCustRecType

  HiRate = 0
  CLS
  
  
  WBCustRecLen = 512
  UBCustRecLen = LEN(UBCustRec(1))
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = WBCustRecLen
  CustHandle1 = FREEFILE
  OPEN "wbcust1.dat" FOR RANDOM AS CustHandle1 LEN = 256

  OPEN "WBHHMESS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = 64

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  
    FOR Cnt = 1 TO NumOfRecs
      REDIM UBCustRec(1) AS NewUBCustRecType
      GET #CustHandle, Cnt, WBCustRec
      GET #CustHandle1, Cnt, WBCustRec1
      GET 13, Cnt, WBHHRec
    

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

    CUSTNAME$ = UCASE$(CUSTNAME$)

 
    UBCustRec(1).Status = UCASE$((QPTrim$(WBCustRec.cycle)))
    UBCustRec(1).BOOK = LEFT$(WBCustRec.meter, 2)
    UBCustRec(1).SEQNUMB = SEQNUMB$
    
    IF UBCustRec(1).Status = "I" OR UBCustRec(1).Status = "F" OR UBCustRec(1).Status = "B" THEN
      UBCustRec(1).BOOK = ""
      UBCustRec(1).SEQNUMB = ""
    END IF

    UBCustRec(1).OPENDATE = -32767
    UBCustRec(1).SEARCH = WBCustRec.LNAME
    UBCustRec(1).CUSTNAME = CUSTNAME$
    UBCustRec(1).ADDR1 = WBCustRec.add1
    UBCustRec(1).ADDR2 = WBCustRec.add3
    UBCustRec(1).SERVADDR = WBCustRec.add2
    UBCustRec(1).CITY = WBCustRec.CITY
    UBCustRec(1).STATE = WBCustRec.ST
    UBCustRec(1).ZIPCODE = WBCustRec.ZP
    UBCustRec(1).HPHONE = WBCustRec.phone
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = WBCustRec.CUSTTYPE
    UBCustRec(1).Addr911 = ""
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
    UBCustRec(1).USERCODE1 = WBCustRec.OwnerCode
    UBCustRec(1).USERCODE2 = WBCustRec.MOrder
    UBCustRec(1).HHMSG1 = WBHHRec.note1
    UBCustRec(1).HHMSG2 = WBHHRec.note2
    UBCustRec(1).HHMSG3 = WBHHRec.note3
    
    FOR ll = 1 TO 15
     UBCustRec(1).Serv(ll).RateCode = ""
     UBCustRec(1).Serv(ll).RMTRTYPE = ""
    NEXT ll

    UBCustRec(1).Serv(1).RateCode = WBCustRec.Rate1
    UBCustRec(1).Serv(1).RMTRTYPE = "C"

    Rate2$ = WBCustRec.Rate2
    Rate2$ = RTRIM$(Rate2$)
    Rate2$ = LTRIM$(Rate2$)
    IF LEN(Rate2$) > 0 THEN
     UBCustRec(1).Serv(2).RateCode = WBCustRec.Rate2
     UBCustRec(1).Serv(2).RMTRTYPE = "C"
    END IF


    FOR ll = 1 TO 7
    UBCustRec(1).LocMeters(ll).MTRNUM = ""
    UBCustRec(1).LocMeters(ll).MTRMulti = 1
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
    
    ReadDate$ = LEFT$(MeterReadDate$, 2) + "/" + MID$(MeterReadDate$, 3, 2) + "/" + RIGHT$(MeterReadDate$, 2)
    ReadDate$ = LTRIM$(ReadDate$)
    IF LEFT$(ReadDate$, 1) = "/" THEN
     ReadDate$ = ""
    END IF
    ll = 1  'Meter 1
    UBCustRec(1).LocMeters(ll).MTRNUM = WBCustRec.METERNUM
    UBCustRec(1).LocMeters(ll).MTRMulti = 1
    UBCustRec(1).LocMeters(ll).MTRType = "C"
    UBCustRec(1).LocMeters(ll).MTRUnit = "G"
    UBCustRec(1).LocMeters(ll).NumUser = WBCustRec.minimums
    UBCustRec(1).LocMeters(ll).InsDate = -32767
    UBCustRec(1).LocMeters(ll).CurRead = WBCustRec.CurRead
    UBCustRec(1).LocMeters(ll).PrevRead = WBCustRec.PASTREAD
    IF LEN(ReadDate$) > 0 THEN
    UBCustRec(1).LocMeters(ll).CurDate = Date2Num%(ReadDate$)
    ELSE
    UBCustRec(1).LocMeters(ll).CurDate = -32767
    END IF
    UBCustRec(1).LocMeters(ll).PastDate = -32767
    UBCustRec(1).LocMeters(ll).ReadFlag = "N"
    UBCustRec(1).LocMeters(ll).AvgUse = 0
    UBCustRec(1).LocMeters(ll).UseCnt = 1



    FOR ll = 1 TO 4
     UBCustRec(1).FlatRates(ll).FRDESC = ""
     UBCustRec(1).FlatRates(ll).FRAMT = 0
     UBCustRec(1).FlatRates(ll).FRFREQ = ""
     UBCustRec(1).FlatRates(ll).REVSRC = 0
     UBCustRec(1).FlatRates(ll).NumMin = 0
    NEXT ll

    IF WBCustRec1.RAMT1 > 0 THEN
      ll = 1
      UBCustRec(1).FlatRates(ll).FRDESC = WBCustRec1.RDESC1
      UBCustRec(1).FlatRates(ll).FRAMT = WBCustRec1.RAMT1
      UBCustRec(1).FlatRates(ll).FRFREQ = "R"
      UBCustRec(1).FlatRates(ll).REVSRC = VAL(LEFT$(WBCustRec1.RACCT1, 1))
      UBCustRec(1).FlatRates(ll).NumMin = 1
    END IF
    IF WBCustRec1.RAMT2 > 0 THEN
      ll = 2
      UBCustRec(1).FlatRates(ll).FRDESC = WBCustRec1.RDESC2
      UBCustRec(1).FlatRates(ll).FRAMT = WBCustRec1.RAMT2
      UBCustRec(1).FlatRates(ll).FRFREQ = "R"
      UBCustRec(1).FlatRates(ll).REVSRC = VAL(LEFT$(WBCustRec1.RACCT2, 1))
      UBCustRec(1).FlatRates(ll).NumMin = 1
    END IF
    IF WBCustRec1.RAMT3 > 0 THEN
      ll = 3
      UBCustRec(1).FlatRates(ll).FRDESC = WBCustRec1.RDESC3
      UBCustRec(1).FlatRates(ll).FRAMT = WBCustRec1.RAMT3
      UBCustRec(1).FlatRates(ll).FRFREQ = "R"
      UBCustRec(1).FlatRates(ll).REVSRC = VAL(LEFT$(WBCustRec1.RACCT3, 1))
      UBCustRec(1).FlatRates(ll).NumMin = 1
    END IF
   
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
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    UBCustRec(1).CurrRevAmts(1) = 0
    UBCustRec(1).CurrRevAmts(2) = 0
    UBCustRec(1).CurrRevAmts(3) = 0
    UBCustRec(1).CurrRevAmts(4) = 0
    UBCustRec(1).CurrRevAmts(5) = 0
    UBCustRec(1).CurrRevAmts(6) = 0
    UBCustRec(1).CurrRevAmts(7) = 0
    UBCustRec(1).CurrRevAmts(8) = 0

    FOR ll = 9 TO 15
     UBCustRec(1).CurrRevAmts(ll) = 0
    NEXT ll

    FOR ll = 1 TO 15
     UBCustRec(1).PrevRevAmts(ll) = 0
    NEXT ll

    UBCustRec(1).DepositAmt = WBCustRec.waterdep
    IF VAL(WBCustRec.meter) = 0 THEN
     UBCustRec(1).DelFlag = -1
     ELSE
     UBCustRec(1).DelFlag = 0
     END IF
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
    PUT UBFile, Cnt, UBCustRec(1)
    NEXT Cnt
    CLOSE
  END

