DEFINT A-Z
DECLARE FUNCTION GetBillCycle% ()
DECLARE SUB RefundMenu ()
DECLARE SUB RefundDeposit ()
DECLARE FUNCTION GetCustMeterType% (UBLocatRec() AS ANY, ThisMeter%)
DECLARE SUB UtilAdjustments ()
DECLARE SUB RePrintUtilBills ()
DECLARE SUB PostBillTrans ()
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION IsDeleted% (CustAcctNo&)
DECLARE SUB ClearScrn ()
DECLARE SUB UBBillMenu ()
DECLARE SUB BillPrintMenu ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE FUNCTION AskAbandonPrint% ()
  
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
'DECLARE SUB MakeBillFile (AbortFlag, FuelAdjAmt#, WhatCycle)
DECLARE FUNCTION Round# (N#)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB MakeZipCodeIndex ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB MakeTempIndex (Indextype%)
DECLARE SUB PreBillReport ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION AskSavePayment
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, Kolor%)
DECLARE SUB StuffBuf (Key2Stuff$)
DECLARE SUB ShowCustHistory (CUSTACCT&)
DECLARE SUB CursorOff ()
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB LookUp (CUSTACCT&, Text$, Parm1, Parm2, ActiveOnly%)
DECLARE FUNCTION PromptSaveData ()
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubpaymnt.BI'
  
  CONST False = 0, True = NOT False
  
  TYPE DraftRptType
    TRANSIT  AS STRING * 9
    BankName AS STRING * 34
    CUSTACCT AS STRING * 8
    CustName AS STRING * 35
    BillAmt  AS STRING * 10
    BANKACCT AS STRING * 20
  END TYPE
  
  TYPE BDRptType
    BankName  AS STRING * 14
    CustRec   AS INTEGER
  END TYPE

SUB BankDraftRegister
  
  Dash80$ = STRING$(125, "-")
  Temp1$ = SPACE$(10)
  Temp2$ = SPACE$(12)
  
  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources

  REDIM PayRec(1) AS UBPaymentRecType
  PayRecLen = LEN(PayRec(1))
  PayFileName$ = "UBPAY99.DAT"

  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TownName$ = UBSetUpRec(1).UTILNAME
  
  MaxLines = 60
  
  REDIM MChoice$(1 TO 6)
  
  PageNo = 0
  
  IndexName$ = NameIndexFile
  OKFlag = True
  
  REDIM DFTRec(1) AS DraftRptType
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))

  PayFile = FREEFILE
  OPEN PayFileName$ FOR OUTPUT AS #PayFile
  CLOSE PayFile
  
  UBBill = FREEFILE
  OPEN "UBFBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  NumOfRecs = LOF(UBBill) \ UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBANKDFT.RPT" FOR OUTPUT AS UBRpt

  ListFile = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS ListFile LEN = PayRecLen
  
  BlockClear
  ShowProcessingScrn "Processing Bank Draft Report"
  
  GOSUB PrintBankDFTHeader
  
  FOR Cnt = 1 TO NumOfRecs
    GET UBBill, Cnt, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag <> 0 THEN
      GET UBCust, Cnt, UBCustRec(1)
      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB PrintBankDFTHeader
      END IF
      IF UBCustRec(1).UseDraft = "Y" THEN
        CstCnt = CstCnt + 1
        REDIM PRESERVE BDCust(1 TO CstCnt) AS BDRptType
        BDCust(CstCnt).BankName = QPTrim$(UBCustRec(1).BankName)
        BDCust(CstCnt).CustRec = Cnt
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
DFTskipem:
  NEXT
  
  IF CstCnt = 0 THEN GOTO ExitReport
  SortT BDCust(1), CstCnt, 0, 16, 0, 14
  
  Split = CstCnt
  PrevBank$ = BDCust(1).BankName
  FOR Cnt = 1 TO CstCnt
    ThisBank$ = BDCust(Cnt).BankName
    IF ThisBank$ <> PrevBank$ THEN
      Split = Cnt - 1
      EXIT FOR
    END IF
  NEXT
  
  IF Split < CstCnt THEN
    SortT BDCust(1), Split, 0, 16, 14, -1
    SortT BDCust(Split + 1), CstCnt - Split, 0, 16, 14, -1
  END IF
  '(SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  GET UBCust, BDCust(1).CustRec, UBCustRec(1)
  PrevBank$ = BDCust(1).BankName
  FOR Cnt = 1 TO CstCnt
'    GET UBCust, BDCust(Cnt).CustRec, UBCustRec(1)
'    GET UBBill, BDCust(Cnt).CustRec, UBBillRec(1)
'
'    PayRec(1).OPERNUM = 99
'    PayRec(1).PAYDATE = UBBillRec(1).BillDate
'    PayRec(1).CUSTACCT = UBBillRec(1).CustAcctNo
'    PayRec(1).CustName = UBCustRec(1).CustName
'    PayRec(1).CUSTADDR = UBCustRec(1).Addr1
'    'PayRec(1).CUSTCMNT= UBBillRec(1).
'    PayRec(1).AMTOWED = UBBillRec(1).TransAmt
'    PayRec(1).TENDERTY = "BANK DRAFT"
'    PayRec(1).CASHAMT = 0
'    PayRec(1).CHKAMT = 0
'    PayRec(1).AMTRECD = UBBillRec(1).TransAmt
'    PayRec(1).CHANGE = 0
'    PayRec(1).DESC = "DRAFT PAYMENT TRANS"
'
'    FOR ZZCnt = 1 TO 15
'      PayRec(1).PaidOwed(ZZCnt).AMTOWE1 = UBBillRec(1).RevAmt(ZZCnt)
'      PayRec(1).PaidOwed(ZZCnt).AMTPD1 = UBBillRec(1).RevAmt(ZZCnt)
'    NEXT
'
'    PayRec(1).TOTOWED = UBBillRec(1).TransAmt
'    PayRec(1).AMTPAID = UBBillRec(1).TransAmt
'
'    PUT #ListFile, , PayRec(1)
    
    ThisBank$ = BDCust(Cnt).BankName
    IF ThisBank$ <> PrevBank$ THEN
      PageNo = 0
      PRINT #UBRpt, Dash80$
      PRINT #UBRpt, CHR$(12)
      GOSUB PrintBankDFTHeader
      PrevBank$ = ThisBank$
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB PrintBankDFTHeader
    END IF
    
    LSET DFTRec(1).TRANSIT = QPTrim$(UBCustRec(1).TRANSIT)
    LSET DFTRec(1).BankName = QPTrim$(UBCustRec(1).BankName)
    RSET DFTRec(1).CUSTACCT = QPTrim$(STR$(BDCust(Cnt).CustRec))
    LSET DFTRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
    LSET DFTRec(1).BANKACCT = QPTrim$(UBCustRec(1).BANKACCT)
    LSET DFTRec(1).BillAmt = FUsing$(STR$(UBBillRec(1).TransAmt), "#####.##")
    PRINT #UBRpt, DFTRec(1).TRANSIT; "  "; DFTRec(1).BankName; "  "; DFTRec(1).CUSTACCT;
    PRINT #UBRpt, "  "; DFTRec(1).CustName; "  "; DFTRec(1).BillAmt; "  ";
    PRINT #UBRpt, ; "  "; DFTRec(1).BANKACCT
    LineCnt = LineCnt + 1
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, CHR$(12)
  CLOSE
  
  ERASE UBSetUpRec, MChoice$, DFTRec, UBCustRec, UBBillRec, PayRec
  
  LPTPort = 1
  IF NOT AbortFlag THEN
    PrintRptFile "Bank Draft Register Report", "UBANKDFT.RPT", LPTPort, RetCode, EntryPoint
  END IF
  
ExitReport:
  EXIT SUB
  
PrintBankDFTHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Billing Bank Draft Register. "; TAB(55); TownName$; TAB(110); "Page: "; PageNo
  PRINT #UBRpt, "Date: "; DATE$
  PRINT #UBRpt, "Bank No.  Bank Name"; TAB(48); "Acct No.   Customer Name                      Draft Amt      Bank Acct No."
  PRINT #UBRpt, Dash80$
  LineCnt = 5
  
  
  RETURN
END SUB

