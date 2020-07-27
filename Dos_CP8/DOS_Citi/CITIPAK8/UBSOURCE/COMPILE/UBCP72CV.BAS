
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

DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION Num2Date$ (Daty%)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION round# (DblNumber#)
  
  '$INCLUDE: 'CP72UTIL.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  DIM WBCustRec  AS CustRecType1
  DIM WBCustRec1  AS CustRecType2
  'DIM WBhh AS WBCUSTHHRECTYPE
 
  REDIM UBCustRec(1) AS NewUBCustRecType
 

  CLS
  
  
  WBCustRecLen = 512
  WBCust1RecLen = 256
  

  UBCustRecLen = LEN(UBCustRec(1))


'GOTO Skip2Here
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = WBCustRecLen
 CustHandle1 = FREEFILE
 OPEN "wbcust1.dat" FOR RANDOM AS CustHandle1 LEN = WBCust1RecLen
 ' custhandle2 = FREEFILE
 ' OPEN "WBHHMESS.DAT" FOR RANDOM AS custhandle2 LEN = 64
     

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  

    FOR cnt = 1 TO NumOfRecs
      
  
      GET #CustHandle, cnt, WBCustRec
      GET #CustHandle1, cnt, WBCustRec1
      'GET #custhandle2, cnt, WBhh

      

      LOCATE 1, 1: PRINT "Converting Customer Record "; cnt; " Of "; NumOfRecs

      UBCustRec(1).DelFlag = 0

      SEQNUMB$ = RTRIM$(MID$(WBCustRec.meter, 4, 6))
      IF VAL(SEQNUMB$) = 0 THEN
        UBCustRec(1).Status = "I"
        UBCustRec(1).DelFlag = -1
        UBCustRec(1).BOOK = ""
        UBCustRec(1).SEQNUMB = ""
        PUT UBFile, cnt, UBCustRec(1)
      ELSE
      
      kk = INSTR(SEQNUMB$, "/")
      IF kk > 0 THEN
       SEQNUMB$ = LEFT$(SEQNUMB$, kk - 1)
      END IF
      IF LEN(SEQNUMB$) < 6 THEN SEQNUMB$ = STRING$(6 - LEN(SEQNUMB$), "0") + SEQNUMB$
      IF VAL(SEQNUMB$) = 0 THEN SEQNUMB$ = "      "

      IF LEN(QPTrim$(WBCustRec.firstname)) = 0 THEN
       CUSTNAME$ = QPTrim$(WBCustRec.lname)
      ELSE
       CUSTNAME$ = QPTrim$(WBCustRec.firstname) + " " + WBCustRec.lname
      END IF

    UBCustRec(1).BOOK = LEFT$(WBCustRec.meter, 2)
    UBCustRec(1).SEQNUMB = SEQNUMB$

    UBCustRec(1).Status = WBCustRec.cycle
    
    UBCustRec(1).OPENDATE = Date2Num(WBCustRec.MINSTALL)
    UBCustRec(1).SEARCH = WBCustRec.lname
    UBCustRec(1).CUSTNAME = CUSTNAME$
    UBCustRec(1).ADDR1 = WBCustRec.add1
    UBCustRec(1).ADDR2 = WBCustRec.add3
    UBCustRec(1).SERVADDR = WBCustRec.add2
    UBCustRec(1).CITY = WBCustRec.CITY
    UBCustRec(1).STATE = WBCustRec.ST
    UBCustRec(1).ZIPCODE = WBCustRec.Zip


    UBCustRec(1).HPHONE = ""
    phone$ = QPTrim$(WBCustRec.phone)
    
    IF LEN(phone$) = 7 THEN
     UBCustRec(1).HPHONE = "   " + LEFT$(phone$, 3) + RIGHT$(phone$, 4)
    END IF
    IF LEN(phone$) = 10 THEN
     UBCustRec(1).HPHONE = LEFT$(phone$, 3) + MID$(phone$, 4, 3) + RIGHT$(phone$, 4)
    END IF

    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = WBCustRec.CUSTTYPE
    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = LEFT$(WBCustRec.won, 1)
    UBCustRec(1).SEQ = VAL(WBCustRec.CUSTID)
    UBCustRec(1).CASHONLY = "N"
    UBCustRec(1).LATEFEE = "Y"
    UBCustRec(1).CUTOFFYN = "Y"
    UBCustRec(1).TAXEXPT = "N"
    UBCustRec(1).SRCIT = "N"
    'UBCustRec(1).EPP = "N"
    'UBCustRec(1).EPPAMT = 0
    UBCustRec(1).USEDRAFT = "N"

    UBCustRec(1).BANKNAME = ""
    UBCustRec(1).BANKLOC = ""
    UBCustRec(1).TRANSIT = ""
    UBCustRec(1).BANKACCT = ""

    UBCustRec(1).BILLCMNT = ""
    UBCustRec(1).PAYCMNT = ""
    UBCustRec(1).PUMPCODE = ""
    UBCustRec(1).USERCODE1 = "" 'WBCustRec.ReadOrder
    UBCustRec(1).USERCODE2 = "" 'WBCustRec.WON
    UBCustRec(1).HHMSG1 = "" ' WBhh.Line1
    UBCustRec(1).HHMSG2 = "" ' WBhh.Line2
    UBCustRec(1).HHMSG3 = "" ' WBhh.Line3

    UBCustRec(1).Serv(1).RateCode = WBCustRec.rate1
    UBCustRec(1).Serv(2).RateCode = WBCustRec.rate2
    UBCustRec(1).Serv(3).RateCode = WBCustRec.RATE4
    UBCustRec(1).Serv(4).RateCode = ""
    UBCustRec(1).Serv(5).RateCode = ""
    UBCustRec(1).Serv(6).RateCode = ""
    UBCustRec(1).Serv(7).RateCode = ""
    UBCustRec(1).Serv(8).RateCode = ""

    UBCustRec(1).Serv(1).RMtrType = ""
    IF LEN(QPTrim$(WBCustRec.rate1)) > 0 THEN
     UBCustRec(1).Serv(1).RMtrType = "C"
    END IF
    UBCustRec(1).Serv(2).RMtrType = ""
    IF LEN(QPTrim$(WBCustRec.rate2)) > 0 THEN
     UBCustRec(1).Serv(2).RMtrType = "C"
    END IF

    UBCustRec(1).Serv(3).RMtrType = ""
    UBCustRec(1).Serv(4).RMtrType = ""
    UBCustRec(1).Serv(5).RMtrType = ""
    UBCustRec(1).Serv(6).RMtrType = ""
    UBCustRec(1).Serv(7).RMtrType = ""
    UBCustRec(1).Serv(8).RMtrType = ""
    FOR LL = 9 TO 15
     UBCustRec(1).Serv(LL).RateCode = ""
     UBCustRec(1).Serv(LL).RMtrType = ""
    NEXT LL

     FOR LL = 1 TO 4
      UBCustRec(1).FlatRates(LL).FRDESC = ""
      UBCustRec(1).FlatRates(LL).FRAMT = 0
      UBCustRec(1).FlatRates(LL).FRFREQ = ""
      UBCustRec(1).FlatRates(LL).REVSRC = 0
      UBCustRec(1).FlatRates(LL).NumMin = 0
     NEXT LL

     IF WBCustRec1.ramt1 > 0 THEN
      LL = 1
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec1.RDESC1
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec1.ramt1
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(WBCustRec1.racct1)
      UBCustRec(1).FlatRates(LL).NumMin = 1
     END IF
     IF WBCustRec1.ramt2 > 0 THEN
      LL = 2
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec1.RDESC2
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec1.ramt2
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(WBCustRec1.racct2)
      UBCustRec(1).FlatRates(LL).NumMin = 1
     END IF
     IF WBCustRec1.ramt3 > 0 THEN
      LL = 3
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec1.RDESC3
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec1.ramt3
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(WBCustRec1.racct3)
      UBCustRec(1).FlatRates(LL).NumMin = 1
     END IF

    UBCustRec(1).Monthly(1).AMTOWED = 0         ' WBCustRec.AMTOWED
    UBCustRec(1).Monthly(1).TotAmtPD = 0        '  WBCustRec.AmtPd
    UBCustRec(1).Monthly(1).PayAmt = 0          ' WBCustRec.Payment
    UBCustRec(1).Monthly(1).RevSource = 0       '  WBCustRec.RevSource

    UBCustRec(1).Monthly(2).AMTOWED = 0
    UBCustRec(1).Monthly(2).TotAmtPD = 0
    UBCustRec(1).Monthly(2).PayAmt = 0
    UBCustRec(1).Monthly(2).RevSource = 0
    

    UBCustRec(1).MFEE1 = 0
    UBCustRec(1).MFEE2 = 0
    

    IF WBCustRec.CurRead <> 0 OR WBCustRec.pastread <> 0 THEN
    LL = 1

    
    UBCustRec(1).LocMeters(LL).MTRNUM = WBCustRec.METERNUM
    UBCustRec(1).LocMeters(LL).MTRMulti = 100   'Pembroke 100 Multiplier
    UBCustRec(1).LocMeters(LL).MTRType = "C"
    UBCustRec(1).LocMeters(LL).MTRUnit = "G"
    UBCustRec(1).LocMeters(LL).NumUser = WBCustRec.minimums
    UBCustRec(1).LocMeters(LL).InsDate = -32767
    UBCustRec(1).LocMeters(LL).CurRead = WBCustRec.CurRead
    UBCustRec(1).LocMeters(LL).PrevRead = WBCustRec.pastread

    ReadDate$ = LEFT$(WBCustRec.meterreaddate, 2) + "/" + MID$(WBCustRec.meterreaddate, 3, 2) + "/" + MID$(WBCustRec.meterreaddate, 5, 2)
    
    UBCustRec(1).LocMeters(LL).CurDate = Date2Num(ReadDate$)
    UBCustRec(1).LocMeters(LL).PastDate = -32767
    UBCustRec(1).LocMeters(LL).ReadFlag = "N"
    UBCustRec(1).LocMeters(LL).AvgUse = WBCustRec.AvgUse
    UBCustRec(1).LocMeters(LL).UseCnt = WBCustRec.AVGCT

    END IF
    
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
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0

    FOR LL = 1 TO 15
     UBCustRec(1).CurrRevAmts(LL) = 0
    NEXT LL

    FOR LL = 1 TO 15
     UBCustRec(1).PrevRevAmts(LL) = 0
    NEXT LL

    UBCustRec(1).DepositAmt = 0
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
    IF WBCustRec.cycle = "C" THEN UBCustRec(1).Status = "B"
    IF WBCustRec.cycle = "D" OR WBCustRec.cycle = "I" OR WBCustRec.cycle = "C" THEN
      UBCustRec(1).BOOK = ""
      UBCustRec(1).SEQNUMB = ""
    END IF


    IF VAL(LEFT$(WBCustRec.meter, 2)) = 0 THEN
     UBCustRec(1).Status = "I"
     UBCustRec(1).DelFlag = -1
     UBCustRec(1).BOOK = ""
     UBCustRec(1).SEQNUMB = ""
    END IF



    PUT UBFile, cnt, UBCustRec(1)
   END IF
    NEXT cnt
    CLOSE

  END
    

