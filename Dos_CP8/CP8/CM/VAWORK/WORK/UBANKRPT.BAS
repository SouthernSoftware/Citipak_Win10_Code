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
DECLARE SUB MakeBillFile (AbortFlag, FuelAdjAmt#, WhatCycle)
DECLARE FUNCTION Round# (N#)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB MakeZipCodeIndex ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
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
  
'  TYPE DraftRptType
'    TRANSIT  AS STRING * 9
'    BankName AS STRING * 34
'    CUSTACCT AS STRING * 8
'    CUSTNAME AS STRING * 35
'    BillAmt  AS STRING * 10
'    BANKACCT AS STRING * 20
'  END TYPE
  

  TYPE DraftRptType
    TRANSIT  AS STRING * 9
    BankName AS STRING * 14
    'BANKNAME AS STRING * 34
    CUSTACCT AS STRING * 5
    CUSTNAME AS STRING * 25
    BillAmt  AS STRING * 10
    BANKACCT AS STRING * 20
  END TYPE

  TYPE BDRptType
    BankName  AS STRING * 14
    CustRec   AS INTEGER
    TransRec  AS LONG
  END TYPE

  TYPE BankTotalsType
    BankName  AS STRING * 14
    Amount    AS DOUBLE
  END TYPE

SUB BankDraftRegister

  Dash80$ = STRING$(80, "-")

  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM BankTotals(1 TO 1) AS BankTotalsType

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TOWNNAME$ = UBSetUpRec(1).UTILNAME

  MaxLines = 58

  'REDIM MChoice$(1 TO 6)

  PageNo = 0

  IndexName$ = NameIndexFile
  OKFlag = True

  REDIM DFTRec(1) AS DraftRptType
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRecs& = LOF(UBCust) \ UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBANKDFT.RPT" FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn "Processing Bank Draft Report"

  GOSUB PrintBankDFTHeader

  FOR Cnt& = 1 TO NumOfRecs&
    GET UBCust, Cnt&, UBCustRec(1)

    IF UBCustRec(1).STATUS = "A" OR UBCustRec(1).STATUS = "B" THEN
    IF (UBCustRec(1).USEDRAFT = "Y") OR (LEN(QPTrim$(UBCustRec(1).BankName)) > 0) THEN
      IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) > 0 THEN
        IF LineCnt > MaxLines THEN
          PRINT #UBRpt, CHR$(12)
          GOSUB PrintBankDFTHeader
        END IF
        CstCnt = CstCnt + 1
        REDIM PRESERVE BDCust(1 TO CstCnt) AS BDRptType
        BDCust(CstCnt).BankName = QPTrim$(UBCustRec(1).BankName)
        BDCust(CstCnt).CustRec = Cnt&
        BDCust(CstCnt).TransRec = Cnt&
      END IF
    END IF
    END IF
    IF AskAbandonPrint% THEN
      ABExit = True
      GOTO NON2PrintExit:
    END IF
    ShowPctCompL Cnt&, NumOfRecs&
DFTskipem:
  NEXT

  IF CstCnt <= 0 THEN
    PRINT #UBRpt, "No Bills found for: "; Num2Date$(BDate)
    PRINT #UBRpt, Dash80$
    GOTO NON2PrintExit
  END IF

  SortT BDCust(1), CstCnt, 0, 20, 0, 14
  BankCnt = 1
  GET UBCust, BDCust(1).CustRec, UBCustRec(1)

  PrevBank$ = BDCust(1).BankName

  BankTotals(BankCnt).BankName = BDCust(1).BankName

  FOR Cnt = 1 TO CstCnt
    GET UBCust, BDCust(Cnt).CustRec, UBCustRec(1)
    IF PrevBank$ <> BDCust(Cnt).BankName THEN
      BankCnt = BankCnt + 1
      REDIM PRESERVE BankTotals(1 TO BankCnt) AS BankTotalsType
      BankTotals(BankCnt).BankName = BDCust(Cnt).BankName
      PrevBank$ = BDCust(Cnt).BankName
    END IF
    BankTotals(BankCnt).Amount = Round#(BankTotals(BankCnt).Amount + UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB PrintBankDFTHeader
    END IF
    LSET DFTRec(1).TRANSIT = QPTrim$(UBCustRec(1).TRANSIT)
    LSET DFTRec(1).BankName = QPTrim$(UBCustRec(1).BankName)
    RSET DFTRec(1).CUSTACCT = QPTrim$(STR$(BDCust(Cnt).CustRec))
    LSET DFTRec(1).CUSTNAME = QPTrim$(UBCustRec(1).CUSTNAME)
    LSET DFTRec(1).BANKACCT = QPTrim$(UBCustRec(1).BANKACCT)
    LSET DFTRec(1).BillAmt = FUsing$(STR$(Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)), "#####.##")
    PRINT #UBRpt, DFTRec(1).TRANSIT; " "; DFTRec(1).BankName;
    PRINT #UBRpt, DFTRec(1).CUSTACCT; "  "; DFTRec(1).CUSTNAME;
    PRINT #UBRpt, " "; DFTRec(1).BillAmt; "  "; DFTRec(1).BANKACCT
    LineCnt = LineCnt + 1
    IF AskAbandonPrint% THEN
      ABExit = True
      GOTO NON2PrintExit:
    END IF
    ShowPctComp Cnt, CstCnt
  NEXT

  PRINT #UBRpt, CHR$(12)
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Billing Bank Draft Register.                "; QPTrim$(TOWNNAME$)
  PRINT #UBRpt, "Date: "; DATE$; TAB(72); "Page: "; PageNo
  PRINT #UBRpt, "Bank Name                 Bank Total."
  PRINT #UBRpt, Dash80$
  GTotal# = 0
  FOR Cnt = 1 TO BankCnt
    PRINT #UBRpt, BankTotals(Cnt).BankName; TAB(30); FUsing$(STR$(BankTotals(Cnt).Amount), "#####.##")
    GTotal# = Round#(GTotal# + BankTotals(Cnt).Amount)
  NEXT
  PRINT #UBRpt,
  PRINT #UBRpt, "Draft Total:"; TAB(30); FUsing$(STR$(GTotal#), "#####.##")
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, CHR$(12)

NON2PrintExit:
  CLOSE

  IF ABExit THEN
    GOTO AbortExit
  END IF

  ERASE UBSetUpRec, DFTRec, UBCustRec

  LPTPort = 1
  IF NOT AbortFlag THEN
    PrintRptFile "Bank Draft Register Report", "UBANKDFT.RPT", LPTPort, RetCode, EntryPoint
  END IF

AbortExit:
  EXIT SUB

PrintBankDFTHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Billing Bank Draft Register.                "; QPTrim$(TOWNNAME$)
  PRINT #UBRpt, "Date: "; DATE$; TAB(72); "Page: "; PageNo
  PRINT #UBRpt, "Bank No.  Bank Name  Acct No.  Customer Name            Draft Amt    Bank Acct."

  PRINT #UBRpt, Dash80$
  LineCnt = 5

RETURN


END SUB

