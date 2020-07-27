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
  
  
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  

  CLS
  
  

  UBCustRecLen = LEN(UBCustRec(1))

'GOTO Skip2Here
  


  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  OPEN "UTCUS.TXT" FOR INPUT AS #10

  WHILE NOT EOF(10)
  LINE INPUT #10, a$
  PRINT a$


  IF LEFT$(a$, 1) = "D" THEN

  ELSE
     Book$ = "0" + MID$(a$, 893, 1)
     SeqNumb$ = LTRIM$(RTRIM$(MID$(a$, 6, 8)))
     IF LEN(SeqNumb$) < 6 THEN SeqNumb$ = STRING$(6 - LEN(SeqNumb$), "0") + SeqNumb$

     
    UBCustRec(1).Book = Book$
    UBCustRec(1).SeqNumb = SeqNumb$
    UBCustRec(1).STATUS = UCASE$(LEFT$(a$, 1))
    OpenDate$ = MID$(a$, 304, 4) + MID$(a$, 302, 2)
    UBCustRec(1).OpenDate = Date2Num%(OpenDate$)
    UBCustRec(1).SEARCH = MID$(a$, 22, 10)
    UBCustRec(1).CUSTNAME = MID$(a$, 32, 30)
    UBCustRec(1).ADDR1 = MID$(a$, 62, 30)
    UBCustRec(1).ADDR2 = MID$(a$, 92, 30)
    UBCustRec(1).SERVADDR = MID$(a$, 960, 23)
    UBCustRec(1).City = MID$(a$, 122, 20)
    kk = INSTR(UBCustRec(1).City, "VA")
    IF kk > 0 THEN UBCustRec(1).City = LEFT$(UBCustRec(1).City, kk - 1)
    UBCustRec(1).STATE = "VA"
    UBCustRec(1).ZIPCODE = MID$(a$, 143, 5)
    UBCustRec(1).HPHONE = MID$(a$, 171, 10)
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = MID$(a$, 181, 1)
    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = MID$(a$, 901, 1)
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




    UBCustRec(1).Serv(1).RateCode = MID$(a$, 903, 3)
    UBCustRec(1).Serv(2).RateCode = MID$(a$, 907, 3)
    UBCustRec(1).Serv(3).RateCode = MID$(a$, 911, 3)
    UBCustRec(1).Serv(4).RateCode = ""
    UBCustRec(1).Serv(5).RateCode = ""
    UBCustRec(1).Serv(6).RateCode = ""
    UBCustRec(1).Serv(7).RateCode = ""
    UBCustRec(1).Serv(8).RateCode = ""
    UBCustRec(1).Serv(1).RMTRType = ""
    IF MID$(a$, 893, 1) = "6" THEN
    UBCustRec(1).Serv(2).RMTRType = "W"
    ELSE
    UBCustRec(1).Serv(2).RMTRType = ""
    END IF
    UBCustRec(1).Serv(3).RMTRType = ""
    UBCustRec(1).Serv(4).RMTRType = ""
    UBCustRec(1).Serv(5).RMTRType = ""
    UBCustRec(1).Serv(6).RMTRType = ""
    UBCustRec(1).Serv(7).RMTRType = ""
    UBCustRec(1).Serv(8).RMTRType = ""
    FOR ll = 9 TO 15
     UBCustRec(1).Serv(ll).RateCode = ""
     UBCustRec(1).Serv(ll).RMTRType = ""
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
    IF MID$(a$, 893, 1) = "6" THEN
    ll = 1
    UBCustRec(1).LocMeters(ll).MTRNUM = ""
    UBCustRec(1).LocMeters(ll).MTRMulti = 1
    UBCustRec(1).LocMeters(ll).MTRType = "W"
    UBCustRec(1).LocMeters(ll).MTRUnit = "G"
    UBCustRec(1).LocMeters(ll).NumUser = 1
    UBCustRec(1).LocMeters(ll).InsDate = -32767
    UBCustRec(1).LocMeters(ll).CurRead = 0
    UBCustRec(1).LocMeters(ll).PrevRead = 0
    UBCustRec(1).LocMeters(ll).CurDate = -32767
    UBCustRec(1).LocMeters(ll).PastDate = -32767
    UBCustRec(1).LocMeters(ll).ReadFlag = "N"
    UBCustRec(1).LocMeters(ll).AvgUse = 0
    UBCustRec(1).LocMeters(ll).UseCnt = 1
    
    FOR ll = 2 TO 7
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
    ELSE
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

    END IF
  
'END OF Quick Screen Form
    UBCustRec(1).CustPIN = Cnt!
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    UBCustRec(1).CurrRevAmts(1) = 0
    UBCustRec(1).CurrRevAmts(2) = 0
    UBCustRec(1).CurrRevAmts(3) = 0
   
    FOR ll = 4 TO 15
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
    Cnt! = Cnt! + 1
    PUT UBFile, Cnt!, UBCustRec(1)
    END IF
    WEND
    CLOSE
  END
    

