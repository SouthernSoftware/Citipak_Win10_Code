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
  DIM WBCustRec1  AS Cust71RecType1
  DIM WBCustRec2  AS Cust71RecType2

  DIM WBTransRec AS WBTransRecType
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTransRec(1) AS UBTransRecType

  CLS
  
  
  WBCustRecLen = LEN(WBCustRec)
  WBCust1RecLen = LEN(WBCustRec1)
  WBCust2RecLen = LEN(WBCustRec2)

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
      GET UBFile, Cnt!, UBCustRec(1)
      UBCustRec(1).DepositAmt = 0
      UBCustRec(1).DepositAmt = INT((WBCustRec.waterdep * 100) + .5) / 100
      IF UBCustRec(1).DepositAmt = 12 THEN STOP
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

   TransRecord# = WBTransRec.tmrec
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
    UBTransRec(1).TransType = 107
   END IF
   IF WBTransRec.TTYPE = 5 THEN
    UBTransRec(1).TransType = 112
   END IF
   IF WBTransRec.TTYPE = 6 THEN
    UBTransRec(1).TransType = 107
   END IF
   IF WBTransRec.TTYPE = 7 THEN
    UBTransRec(1).TransType = 104
   END IF
   IF WBTransRec.TTYPE = 8 THEN
    UBTransRec(1).TransType = 103
   END IF
   PUT UBFile, Cnt!, UBTransRec(1)
NextTranaction:
  NEXT Cnt!
  CLOSE
  END

