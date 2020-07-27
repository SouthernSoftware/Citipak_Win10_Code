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

  REDIM NewPort(1) AS NewPort
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTransRec(1) AS UBTransRecType

  CLS
  
  
  WBCustRecLen = 728
  UBCustRecLen = LEN(UBCustRec(1))


'GOTO Skip2Here

  FOR MLoop = 1 TO 10
  
  EXT$ = STR$(MLoop)
  EXT$ = LTRIM$(EXT$)
  IF MLoop <= 9 THEN EXT$ = "0" + EXT$

  CustHandle = FREEFILE
  OPEN "ACTMAS" + EXT$ + ".DAT" FOR RANDOM AS CustHandle LEN = WBCustRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  
    FOR Cnt = 1 TO NumOfRecs
      
      GET #CustHandle, Cnt, NewPort(1)
    

      LOCATE 1, 1: PRINT "Converting Customer Record "; Cnt; " Of "; NumOfRecs
      SEQNUMB$ = LEFT$(NewPort(1).Acct, 3) + MID$(NewPort(1).Acct, 5, 2)
      IF LEN(SEQNUMB$) < 6 THEN SEQNUMB$ = STRING$(6 - LEN(SEQNUMB$), "0") + SEQNUMB$
      IF VAL(SEQNUMB$) = 0 THEN SEQNUMB$ = "      "
      
      CustName$ = QPTrim$(NewPort(1).Name)
      kk = INSTR(CustName$, "^")
      IF kk > 1 THEN CustName$ = LEFT$(CustName$, kk - 1)
    UBCustRec(1).BOOK = "01"
    UBCustRec(1).SEQNUMB = SEQNUMB$
    UBCustRec(1).STATUS = "A"

    UBCustRec(1).OPENDATE = Date2Num%(MINSTALL$)
    UBCustRec(1).SEARCH = CustName$
    UBCustRec(1).CustName = CustName$
    UBCustRec(1).ADDR1 = NewPort(1).Address
    UBCustRec(1).ADDR2 = ""
    UBCustRec(1).SERVADDR = NewPort(1).ServAddress
    UBCustRec(1).City = NewPort(1).City
    UBCustRec(1).STATE = NewPort(1).St
    UBCustRec(1).ZipCode = NewPort(1).Zip
    UBCustRec(1).HPHONE = ""
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = NewPort(1).CTYPE
    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""
    UBCustRec(1).BILLCYCL = 1
    UBCustRec(1).ZONE = NewPort(1).CLoc
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

    UBCustRec(1).Serv(1).RateCode = ""
    UBCustRec(1).Serv(2).RateCode = ""
    UBCustRec(1).Serv(3).RateCode = ""
    UBCustRec(1).Serv(4).RateCode = ""
    UBCustRec(1).Serv(5).RateCode = ""
    UBCustRec(1).Serv(6).RateCode = ""
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

    UBCustRec(1).DepositAmt = 0
    UBCustRec(1).DelFlag = 0
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
    NextRecord = (LOF(UBFile) / LEN(UBCustRec(1))) + 1
    PUT UBFile, NextRecord, UBCustRec(1)
    NEXT Cnt
    CLOSE
    NEXT MLoop

  CLOSE
  END
    

