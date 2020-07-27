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
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'table1.bi'

  CLEAR
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM TblRec(1) AS BladenRecType

  CLS

  IF COMMAND$ = "1" THEN Cycle1Flag = -1
  
  UBCustRecLen = LEN(UBCustRec(1))
  TblRecLen = LEN(TblRec(1))

'GOTO Skip2Here
  
  OPEN "account.txt" FOR RANDOM AS #1 LEN = TblRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  'SeqNum&=
  NumOfRecs = LOF(1) / TblRecLen
  
    FOR Cnt = 1 TO NumOfRecs
      REDIM UBCustRec(1) AS NewUBCustRecType

      GET #1, Cnt, TblRec(1)

      LOCATE 1, 1: PRINT "Converting Customer Record "; Cnt; " Of "; NumOfRecs
    First$ = QPTrim$(TblRec(1).CFirstName)
    FLen = LEN(First$)
    Midd$ = QPTrim$(TblRec(1).CMidI)
    MLen = LEN(Midd$)
    Last$ = QPTrim$(TblRec(1).CLastName)
    LLen = LEN(Last$)

    IF (FLen = 0) AND (MLen = 0) AND (LLen = 0) THEN
      CustName$ = TblRec(1).BusiName
      UBCustRec(1).SEARCH = RIGHT$(CustName$, 10)
    ELSEIF LEN(Midd$) > 0 THEN
      CustName$ = First$ + " " + Midd$ + " " + Last$
      UBCustRec(1).SEARCH = Last$
    ELSE
      CustName$ = First$ + " " + Last$
      UBCustRec(1).SEARCH = Last$
    END IF

    UBCustRec(1).CustName = CustName$
    IF LEN(QPTrim$(UBCustRec(1).SEARCH)) = 0 THEN
      TSearch$ = QPTrim$(UBCustRec(1).CustName)
      UBCustRec(1).SEARCH = QPTrim$(RIGHT$(TSearch$, 6))
    END IF

    UBCustRec(1).Status = TblRec(1).Status
    IF UBCustRec(1).Status = "A" THEN
      'SeqNum& = SeqNum& + 50
      TempSeq$ = "000000" + QPTrim$(TblRec(1).Location)
      UBCustRec(1).SEQNUMB = RIGHT$(TempSeq$, 6)
      TBook$ = "00" + QPTrim$(TblRec(1).Book)
      UBCustRec(1).Book = RIGHT$(TBook$, 2)
    END IF

    IF UBCustRec(1).Status = "F" THEN
      UBCustRec(1).Status = "I"
    END IF

    UBCustRec(1).OPENDATE = Date2Num%(TblRec(1).AcctStartDate)
    UBCustRec(1).ADDR1 = UCASE$(TblRec(1).BillAddr1)
    UBCustRec(1).ADDR2 = UCASE$(TblRec(1).BillAddr2)
    UBCustRec(1).SERVADDR = UCASE$(TblRec(1).MeterLoc)
    UBCustRec(1).CITY = UCASE$(TblRec(1).BillCITY)
    UBCustRec(1).STATE = UCASE$(TblRec(1).BillST)
    UBCustRec(1).ZIPCODE = TblRec(1).BillZip

    Phone$ = QPTrim$(TblRec(1).Contact)

    DO
      DashPos = INSTR(Phone$, "-")
      IF DashPos > 0 THEN
        Phone$ = LEFT$(Phone$, DashPos - 1) + MID$(Phone$, DashPos + 1)
      END IF
    LOOP WHILE DashPos > 0


    UBCustRec(1).HPHONE = Phone$
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""

    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""

    Cycle$ = TblRec(1).Cycle
    BCycle = VAL(RIGHT$(Cycle$, 1))

    IF BCycle = 0 THEN
      BCycle = 1
    END IF

    'IF BCycle <> 1 THEN STOP

    UBCustRec(1).BILLCYCL = BCycle

    IF Cycle1Flag THEN
      UBCustRec(1).BILLCYCL = 1
    END IF

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
    UBCustRec(1).HHMSG1 = UCASE$(TblRec(1).ReadMsg)
    'UBCustRec(1).HHMSG2 =
    'UBCustRec(1).HHMSG3 =


    FOR zzcnt = 1 TO 7
      UBCustRec(1).LocMeters(zzcnt).NumUser = -32767
      UBCustRec(1).LocMeters(zzcnt).InsDate = -32767
      UBCustRec(1).LocMeters(zzcnt).CurRead = -2147483647
      UBCustRec(1).LocMeters(zzcnt).PrevRead = -2147483647
      UBCustRec(1).LocMeters(zzcnt).CurDate = -32767
      UBCustRec(1).LocMeters(zzcnt).PastDate = -32767
      UBCustRec(1).LocMeters(zzcnt).ReadFlag = ""
      UBCustRec(1).LocMeters(zzcnt).avguse = 0
      UBCustRec(1).LocMeters(zzcnt).UseCnt = 0
    NEXT
    
    CCode$ = QPTrim$(TblRec(1).CatCode1)
    SELECT CASE VAL(RIGHT$(CCode$, 1))
    CASE 1     'water & sewer
      UBCustRec(1).Serv(1).RateCode = "WAT1"
      UBCustRec(1).Serv(1).RMTRTYPE = "W"
      UBCustRec(1).LocMeters(1).NumUser = 1
      UBCustRec(1).LocMeters(1).MTRNUM = TblRec(1).SerialNo
      UBCustRec(1).LocMeters(1).MTRMulti = 100
      UBCustRec(1).LocMeters(1).CurDate = Date2Num(TblRec(1).CurDate)
      UBCustRec(1).LocMeters(1).PastDate = Date2Num(TblRec(1).PrevDate)
      UBCustRec(1).LocMeters(1).CurRead = VAL(QPTrim$(TblRec(1).CurRead)) \ 100
      UBCustRec(1).LocMeters(1).PrevRead = VAL(QPTrim$(TblRec(1).PrevRead)) \ 100
      UBCustRec(1).LocMeters(1).MTRType = "W"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"

    CASE 2     'water only
      UBCustRec(1).Serv(1).RateCode = "WAT2"
      UBCustRec(1).Serv(1).RMTRTYPE = "W"
      UBCustRec(1).LocMeters(1).NumUser = 1
      UBCustRec(1).LocMeters(1).MTRNUM = TblRec(1).SerialNo
      UBCustRec(1).LocMeters(1).MTRMulti = 100
      UBCustRec(1).LocMeters(1).CurDate = Date2Num(TblRec(1).CurDate)
      UBCustRec(1).LocMeters(1).PastDate = Date2Num(TblRec(1).PrevDate)
      UBCustRec(1).LocMeters(1).CurRead = VAL(QPTrim$(TblRec(1).CurRead)) \ 100
      UBCustRec(1).LocMeters(1).PrevRead = VAL(QPTrim$(TblRec(1).PrevRead)) \ 100
      UBCustRec(1).LocMeters(1).MTRType = "W"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"

    CASE 3     'water & sewer discount
      UBCustRec(1).Serv(1).RateCode = "WAT3"
      UBCustRec(1).Serv(1).RMTRTYPE = "W"
      UBCustRec(1).LocMeters(1).NumUser = 1
      UBCustRec(1).LocMeters(1).MTRNUM = TblRec(1).SerialNo
      UBCustRec(1).LocMeters(1).MTRMulti = 100
      UBCustRec(1).LocMeters(1).CurDate = Date2Num(TblRec(1).CurDate)
      UBCustRec(1).LocMeters(1).PastDate = Date2Num(TblRec(1).PrevDate)
      UBCustRec(1).LocMeters(1).CurRead = VAL(QPTrim$(TblRec(1).CurRead)) \ 100
      UBCustRec(1).LocMeters(1).PrevRead = VAL(QPTrim$(TblRec(1).PrevRead)) \ 100
      UBCustRec(1).LocMeters(1).MTRType = "W"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"


    CASE 4
      UBCustRec(1).Serv(1).RateCode = "WAT4"
      UBCustRec(1).Serv(1).RMTRTYPE = "W"
      UBCustRec(1).LocMeters(1).NumUser = 1
      UBCustRec(1).LocMeters(1).MTRNUM = TblRec(1).SerialNo
      UBCustRec(1).LocMeters(1).MTRMulti = 100
      UBCustRec(1).LocMeters(1).CurDate = Date2Num(TblRec(1).CurDate)
      UBCustRec(1).LocMeters(1).PastDate = Date2Num(TblRec(1).PrevDate)
      UBCustRec(1).LocMeters(1).CurRead = VAL(QPTrim$(TblRec(1).CurRead)) \ 100
      UBCustRec(1).LocMeters(1).PrevRead = VAL(QPTrim$(TblRec(1).PrevRead)) \ 100
      UBCustRec(1).LocMeters(1).MTRType = "W"
      UBCustRec(1).LocMeters(1).MTRUnit = "G"
    CASE 5
       LPRINT Cnt; TAB(10); RIGHT$(QPTrim$(CustName$), 15); TAB(50); "Code: "; QPTrim$(CCode$)
    '  UBCustRec(1).Serv(1).RateCode = "WAT5"
    '  UBCustRec(1).Serv(1).RMTRTYPE = "C"
    '  UBCustRec(1).Serv(2).RateCode = "SEW5"
    '  UBCustRec(1).Serv(2).RMTRTYPE = "C"
    '  UBCustRec(1).LocMeters(1).NumUser = 1
    '  UBCustRec(1).LocMeters(1).MTRNUM = TblRec(1).SerialNo
    '  UBCustRec(1).LocMeters(1).MTRMulti = 1
    '  UBCustRec(1).LocMeters(1).CurDate = Date2Num(TblRec(1).CurDate)
    '  UBCustRec(1).LocMeters(1).PastDate = Date2Num(TblRec(1).PrevDate)
    '  UBCustRec(1).LocMeters(1).CurRead = VAL(QPTrim$(TblRec(1).CurRead))
    '  UBCustRec(1).LocMeters(1).PrevRead = VAL(QPTrim$(TblRec(1).PrevRead))
    '  UBCustRec(1).LocMeters(1).MTRType = "C"
    '  UBCustRec(1).LocMeters(1).MTRUnit = "G"
    '
    CASE ELSE  '???
      LPRINT Cnt; TAB(10); RIGHT$(QPTrim$(CustName$), 15); TAB(50); "Code: ????"
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
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0

    FOR ll = 1 TO 15
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
    

