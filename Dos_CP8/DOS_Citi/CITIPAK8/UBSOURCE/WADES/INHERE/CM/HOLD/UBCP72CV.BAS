
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
DECLARE FUNCTION Round# (DblNumber#)
  
  '$INCLUDE: 'CP72UTIL.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  DIM WBCustRec  AS Cust72RecType
  DIM WBCustRec1  AS Cust72RecType1
'  DIM WBHH AS WBCUSTHHRECTYPE
 
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
'  custhandle2 = FREEFILE
'  OPEN "wbcusthh.dat" FOR RANDOM AS custhandle2 LEN = 60


  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  

    FOR cnt = 1 TO NumOfRecs
      
  
      GET #CustHandle, cnt, WBCustRec
      GET #CustHandle1, cnt, WBCustRec1
     ' GET #custhandle2, cnt, WBHH
      PRINT WBCustRec.firstname
      

      LOCATE 1, 1: PRINT "Converting Customer Record "; cnt; " Of "; NumOfRecs

      UBCustRec(1).DelFlag = 0

      SEQNUMB$ = RTRIM$(MID$(WBCustRec.meter, 4, 6))
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

    UBCustRec(1).STATUS = WBCustRec.cycle
    
    UBCustRec(1).OPENDATE = Date2Num(WBCustRec.MINSTALL)
    UBCustRec(1).SEARCH = WBCustRec.lname
    UBCustRec(1).CUSTNAME = CUSTNAME$
    UBCustRec(1).ADDR1 = WBCustRec.add1
    UBCustRec(1).ADDR2 = WBCustRec.add3
    UBCustRec(1).SERVADDR = WBCustRec.add2
    UBCustRec(1).City = WBCustRec.City
    UBCustRec(1).STATE = WBCustRec.ST
    UBCustRec(1).ZipCode = WBCustRec.Zip
    UBCustRec(1).HPHONE = WBCustRec.phone
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = WBCustRec.CUSTTYPE
    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = WBCustRec.WON
    UBCustRec(1).SEQ = WBCustRec.SeqNum
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
    UBCustRec(1).HHMSG1 = WBCustRec.HHComment
    UBCustRec(1).HHMSG2 = MID$(WBCustRec.HHComment, 21, 10)
    UBCustRec(1).HHMSG3 = ""

    UBCustRec(1).Serv(1).RateCode = WBCustRec.RATE1
    UBCustRec(1).Serv(2).RateCode = WBCustRec.RATE2
    UBCustRec(1).Serv(3).RateCode = WBCustRec.RATE3
    UBCustRec(1).Serv(4).RateCode = ""
    UBCustRec(1).Serv(5).RateCode = ""
    UBCustRec(1).Serv(6).RateCode = WBCustRec.RATE6
    UBCustRec(1).Serv(7).RateCode = ""
    UBCustRec(1).Serv(8).RateCode = ""

    UBCustRec(1).Serv(1).RMTRTYPE = ""
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

    IF WBCustRec.CurRead <> 0 OR WBCustRec.pastread <> 0 THEN
    LL = 1
    UBCustRec(1).Serv(1).RMTRTYPE = "C"
    UBCustRec(1).Serv(2).RMTRTYPE = "C"
    UBCustRec(1).LocMeters(LL).MTRNUM = WBCustRec.METERNUM
    UBCustRec(1).LocMeters(LL).MTRMulti = 1
    UBCustRec(1).LocMeters(LL).MTRType = "C"
    UBCustRec(1).LocMeters(LL).MTRUnit = "G"
    UBCustRec(1).LocMeters(LL).NumUser = WBCustRec.mINIMUMS
    UBCustRec(1).LocMeters(LL).InsDate = -32767
    UBCustRec(1).LocMeters(LL).CurRead = WBCustRec.CurRead
    UBCustRec(1).LocMeters(LL).PrevRead = WBCustRec.pastread

    ReadDate$ = LEFT$(WBCustRec.MeterReadDate, 6)
    
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
    IF WBCustRec.cycle = "F" OR WBCustRec.cycle = "B" WBCustRec.cycle = "I" THEN
      UBCustRec(1).BOOK = ""
      UBCustRec(1).SEQNUMB = ""
    END IF


    IF VAL(LEFT$(WBCustRec.meter, 2)) = 0 THEN
     UBCustRec(1).STATUS = "I"
     UBCustRec(1).DelFlag = -1
     UBCustRec(1).BOOK = ""
     UBCustRec(1).SEQNUMB = ""
    END IF



    PUT UBFile, cnt, UBCustRec(1)

    NEXT cnt
    CLOSE

  END
    

