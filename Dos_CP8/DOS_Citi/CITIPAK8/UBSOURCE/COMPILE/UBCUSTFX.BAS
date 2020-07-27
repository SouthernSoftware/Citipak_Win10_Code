
DEFINT A-Z
DECLARE SUB PrintCustByRate ()
DECLARE SUB UBLog (Text$)
DECLARE SUB DeleteCust ()
DECLARE SUB ReUseCust ()
DECLARE SUB SetCustFinal ()
DECLARE SUB PrintLocaList ()
DECLARE SUB AddCustDeposit (DepAmt#)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB ShowWrkOrdHistory (CustRec&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE SUB SaveOldCustRec (Form$(), RecNo&)
DECLARE SUB CustLookUp (RecNo%)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
DECLARE SUB PrintCustList ()
DECLARE SUB SaveNewCustRec (Form$(), RecNo&)
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
'DECLARE SUB LoadSetUpRec (Form$(), OKFlag)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB EditCust ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB StuffBuf (Ky$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB MScrnSave (UlRow%, UlCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (UlRow%, UlCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB CursorOff ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB ShowCursor ()
DECLARE SUB WaitForAction ()
  
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB LoadCustRec (Form$(), RecNo&, LocatRec&, TFormLen%, LockedFlag%)
DECLARE SUB BlockClear ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num (WhatDate$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE FUNCTION GetNumRateRecs% ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubTrans.BI'
  
  CONST False = 0, True = NOT False
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE

  CrLf$ = CHR$(13) + CHR$(10)

PrintCustList

SUB PrintCustList
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBCULIST.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Customer Listing. (Name Order)"
  
  GOSUB DoCustRptHeader
  
  FOR Cnt = 1 TO IdxNumOfRecs
   GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)

   IF UBCustRec(1).LocMeters(1).ReadFlag = "Y" THEN
   IF UBCustRec(1).LocMeters(1).CurDate > 6423 THEN
      UBCustRec(1).LocMeters(1).CurRead = UBCustRec(1).LocMeters(1).CurRead / 1000
   END IF


   END IF

   PUT UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
  NEXT Cnt

  CLOSE UBCust, UBLoca, UBRpt
  ERASE IdxBuff, UBCustRec
EXIT SUB
  
  
DoCustRptHeader:
RETURN
  
DoCustRptTotals:
RETURN
  
END SUB

