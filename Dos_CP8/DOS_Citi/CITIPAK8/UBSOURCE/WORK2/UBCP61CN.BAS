DEFINT A-Z

DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION Num2Date$ (Daty%)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()
DECLARE SUB SaveNewCustRec (FormC$(), NewRecNum&)
DECLARE SUB SaveNewLocaRec (Form$(), LocatRec&, CustRec&)
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
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
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
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  DIM WBCustRec  AS WBCustRecType
  DIM WBCustRec1  AS WBCustRecType1
  'DIM WBCustRec2  AS WBCustRecType2
  'DIM WBDraft AS WBDraftType
  'DIM WBCustRec3 AS WBCustRecType3
  DIM WBOwner AS OwnerRecType

  REDIM UBCustRec(1) AS NewUBCustRecType

  CLS
  
  
  WBCustRecLen = 256
  WBCust1RecLen = 256
  WBCust2RecLen = 128
  WBDraftRecLen = 256

  UBCustRecLen = LEN(UBCustRec(1))

'GOTO Skip2Here
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = 256
  CustHandle1 = FREEFILE
  OPEN "wbcust1.dat" FOR RANDOM AS CustHandle1 LEN = 256
  'CustHandle2 = FREEFILE
  'OPEN "wbcust3.dat" FOR RANDOM AS CustHandle2 LEN = 128
  CustHandle3 = FREEFILE
  OPEN "wbowner.dat" FOR RANDOM AS CustHandle3 LEN = 125
 
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
 
  NumOfRecs = LOF(CustHandle) \ 256

  FOR cnt! = 1 TO NumOfRecs
      LOCATE 3, 1: PRINT "Customer # "; cnt!; " of "; NumOfRecs
      GET #CustHandle, cnt!, WBCustRec
      GET #CustHandle1, cnt!, WBCustRec1
      GET #CustHandle3, cnt!, WBOwner


    UBCustRec(1).BOOK = LEFT$(WBCustRec.meter, 2)
    UBCustRec(1).SEQNUMB = MID$(WBCustRec.meter, 4, 6)

    
    kk = INSTR(UBCustRec(1).SEQNUMB, "/")
    IF kk > 1 THEN UBCustRec(1).SEQNUMB = LEFT$(UBCustRec(1).SEQNUMB, kk - 1)

    IF LEN(RTRIM$(UBCustRec(1).SEQNUMB)) < 6 THEN
     UBCustRec(1).SEQNUMB = STRING$(6 - LEN(RTRIM$(UBCustRec(1).SEQNUMB)), "0") + UBCustRec(1).SEQNUMB
    END IF


    UBCustRec(1).SEARCH = UCASE$(WBCustRec.lname)

   'PRESET ALL ACCOUNTS TO ACTIVE STATUS
    UBCustRec(1).Status = WBCustRec.Status


    CustomerName$ = UCASE$(RTRIM$(LTRIM$(WBCustRec.firstname)))
    IF LEN(CustomerName$) = 0 THEN
     CustomerName$ = UCASE$(WBCustRec.lname)
     ELSE
     CustomerName$ = CustomerName$ + " " + UCASE$(WBCustRec.lname)
    END IF
    
    UBCustRec(1).CUSTNAME = CustomerName$
    UBCustRec(1).ADDR1 = UCASE$(WBCustRec.add1)
    UBCustRec(1).ADDR2 = WBCustRec1.add3
    UBCustRec(1).STATE = UCASE$(WBCustRec.ST)
    

    UBCustRec(1).SERVADDR = UCASE$(WBCustRec.add2)
    UBCustRec(1).CITY = UCASE$(WBCustRec.CITY)
    
    UBCustRec(1).ZIPCODE = WBCustRec.Zip
    PHONENUMBER$ = QPTrim$(WBCustRec.phone)
    IF LEN(PHONENUMBER$) < 10 THEN PHONENUMBER$ = STRING$(10 - LEN(PHONENUMBER$), " ") + PHONENUMBER$

    UBCustRec(1).HPHONE = PHONENUMBER$
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    IF WBCustRec1.CUSTTYPE = 1 THEN
     UBCustRec(1).CUSTTYPE = "R"
    ELSE
     UBCustRec(1).CUSTTYPE = "C"
    END IF

    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = LEFT$(WBCustRec.won, 1)
    UBCustRec(1).SEQ = 0
    UBCustRec(1).CASHONLY = "N"
    UBCustRec(1).LATEFEE = "Y"
    UBCustRec(1).CUTOFFYN = "Y"
    UBCustRec(1).TAXEXPT = "N"
    UBCustRec(1).SRCIT = "N"
    UBCustRec(1).EPP = "N"
    UBCustRec(1).EPPAMT = 0
    'Set Draft Information to Null Here

    UBCustRec(1).PreNoteFlag = 0
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
    UBCustRec(1).HHMSG1 = "": 'WBCustRec3.HHMess1
    UBCustRec(1).HHMSG2 = "": 'WBCustRec3.HHMess2
    UBCustRec(1).HHMSG3 = "": ' WBCustRec3.HHMess3
    UBCustRec(1).OPENDATE = Date2Num%(WBCustRec.MINSTALL)

    FOR LL = 1 TO 15
     UBCustRec(1).Serv(LL).RateCode = ""
     UBCustRec(1).Serv(LL).RMTRTYPE = ""
    NEXT LL

     WaterRate = WBCustRec.rate
     Where$ = UCASE$(LEFT$(WBCustRec.won, 1))

     IF WaterRate = 1 AND Where$ = "I" THEN
       UBCustRec(1).Serv(1).RateCode = "WI1"
       UBCustRec(1).Serv(1).RMTRTYPE = "C"
       UBCustRec(1).Serv(2).RateCode = "SI1"
       UBCustRec(1).Serv(2).RMTRTYPE = "C"
     END IF
     IF WaterRate = 1 AND Where$ = "O" THEN
       UBCustRec(1).Serv(1).RateCode = "WO1"
       UBCustRec(1).Serv(1).RMTRTYPE = "C"
       UBCustRec(1).Serv(2).RateCode = "SO1"
       UBCustRec(1).Serv(2).RMTRTYPE = "C"
     END IF

     IF WaterRate = 2 AND Where$ = "I" THEN
       UBCustRec(1).Serv(1).RateCode = "WI1"
       UBCustRec(1).Serv(1).RMTRTYPE = "C"
     END IF
     IF WaterRate = 2 AND Where$ = "O" THEN
       UBCustRec(1).Serv(1).RateCode = "WO1"
       UBCustRec(1).Serv(1).RMTRTYPE = "C"
     END IF
     IF WaterRate = 3 AND Where$ = "I" THEN
       UBCustRec(1).Serv(2).RateCode = "SI1"
       UBCustRec(1).Serv(2).RMTRTYPE = "C"
     END IF
     IF WaterRate = 3 AND Where$ = "O" THEN
       UBCustRec(1).Serv(2).RateCode = "SO1"
       UBCustRec(1).Serv(2).RMTRTYPE = "C"
     END IF

     IF WaterRate = 4 THEN
       UBCustRec(1).Serv(1).RateCode = "W4"
       UBCustRec(1).Serv(1).RMTRTYPE = "C"
     END IF
     IF WaterRate = 5 THEN
       UBCustRec(1).Serv(2).RateCode = "S5"
       UBCustRec(1).Serv(2).RMTRTYPE = "C"
     END IF

     IF WaterRate = 6 THEN
       UBCustRec(1).Serv(1).RateCode = "WIND"
       UBCustRec(1).Serv(1).RMTRTYPE = "C"
       UBCustRec(1).Serv(2).RateCode = "SIND"
       UBCustRec(1).Serv(2).RMTRTYPE = "C"
     END IF
     IF WaterRate = 7 THEN
       UBCustRec(1).Serv(2).RateCode = "S7"
       UBCustRec(1).Serv(2).RMTRTYPE = "C"
     END IF
     IF WBCustRec.GRate > 0 AND WBCustRec.GRate < 8 THEN
       GRate$ = STR$(WBCustRec.GRate)
       GRate$ = LTRIM$(GRate$)
       UBCustRec(1).Serv(3).RateCode = "GB" + GRate$
     END IF



     FOR LL = 1 TO 4
      UBCustRec(1).FlatRates(LL).FRDESC = ""
      UBCustRec(1).FlatRates(LL).FRAMT = 0
      UBCustRec(1).FlatRates(LL).FRFREQ = ""
      UBCustRec(1).FlatRates(LL).REVSRC = 0
      UBCustRec(1).FlatRates(LL).NumMin = 0
     NEXT LL
     
     IF WBCustRec1.RAMT1 > 0 THEN
      LL = 1
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec1.RDESC1
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec1.RAMT1
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(WBCustRec1.RACCT1)
      UBCustRec(1).FlatRates(LL).NumMin = 1
     END IF
     IF WBCustRec1.RAMT2 > 0 THEN
      LL = 2
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec1.RDESC2
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec1.RAMT2
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(WBCustRec1.RACCT2)
      UBCustRec(1).FlatRates(LL).NumMin = 1
     END IF
     IF WBCustRec1.RAMT3 > 0 THEN
      LL = 3
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec1.RDESC3
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec1.RAMT3
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(WBCustRec1.RACCT3)
      UBCustRec(1).FlatRates(LL).NumMin = 1
     END IF


    FOR LL = 1 TO 2
     UBCustRec(1).Monthly(LL).AMTOWED = 0
     UBCustRec(1).Monthly(LL).TotAmtPD = 0
     UBCustRec(1).Monthly(LL).PayAmt = 0
     UBCustRec(1).Monthly(LL).RevSource = 0
    NEXT LL
    
    UBCustRec(1).MFEE1 = 0
    UBCustRec(1).MFEE2 = 0

    'Initialize All Meters
    FOR LL = 1 TO 7
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


    'First Meter is Water/Sewer
    LL = 1
    minimums = WBCustRec1.Minimum
    IF minimums < 1 THEN minimums = 1
    UBCustRec(1).LocMeters(LL).MTRNUM = WBCustRec1.METERNUM
    UBCustRec(1).LocMeters(LL).MTRMulti = 100                'Timmsonv = 100
    UBCustRec(1).LocMeters(LL).MTRType = "C"
    UBCustRec(1).LocMeters(LL).MTRUnit = "G"
    UBCustRec(1).LocMeters(LL).NumUser = minimums       'Dont Let Min < 1
    UBCustRec(1).LocMeters(LL).InsDate = Date2Num%(WBCustRec.MINSTALL)
    UBCustRec(1).LocMeters(LL).CurRead = WBCustRec.CurRead
    UBCustRec(1).LocMeters(LL).PrevRead = WBCustRec.pastread
    UBCustRec(1).LocMeters(LL).CurDate = -32767
    UBCustRec(1).LocMeters(LL).PastDate = -32767
    UBCustRec(1).LocMeters(LL).ReadFlag = "N"
    UBCustRec(1).LocMeters(LL).AvgUse = 0
    UBCustRec(1).LocMeters(LL).UseCnt = 0
    
    UBCustRec(1).CustPIN = 0
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0

    FOR LL = 1 TO 15
     UBCustRec(1).CurrRevAmts(LL) = 0
     UBCustRec(1).PrevRevAmts(LL) = 0
    NEXT LL

    UBCustRec(1).DepositAmt = 0
   
    
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
    UBCustRec(1).DelFlag = 0


    IF UBCustRec(1).Status = "F" OR UBCustRec(1).Status = "B" OR UBCustRec(1).Status = "I" THEN
     UBCustRec(1).BOOK = ""
     UBCustRec(1).SEQNUMB = ""
    END IF
    IF LEFT$(UBCustRec(1).CUSTNAME, 6) = "VACANT" THEN
     UBCustRec(1).Status = "I"
     UBCustRec(1).BOOK = ""
     UBCustRec(1).SEQNUMB = ""
    END IF

    IF VAL(LEFT$(WBCustRec.meter, 2)) = 0 THEN
     UBCustRec(1).Status = "I"
     UBCustRec(1).BOOK = ""
     UBCustRec(1).SEQNUMB = ""
     UBCustRec(1).DelFlag = -1
    END IF

    
    PUT UBFile, cnt!, UBCustRec(1)
    NEXT cnt!
    CLOSE
  END

