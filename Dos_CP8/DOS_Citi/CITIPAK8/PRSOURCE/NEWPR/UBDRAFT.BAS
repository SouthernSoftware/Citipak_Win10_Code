DEFINT A-Z

DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION FUsing$ (number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBAcctsToDraft ()
DECLARE SUB UBBuildTransmitFile ()
DECLARE SUB UBDraftListing ()
DECLARE SUB UBDraftTest ()
DECLARE SUB UBPrenote ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'newcust.bi'
  
  TYPE DraftRptType
    TRANSIT  AS STRING * 9
    BANKNAME AS STRING * 14
    'BANKNAME AS STRING * 34
    CUSTACCT AS STRING * 5
    CUSTNAME AS STRING * 25
    BillAmt  AS STRING * 10
    BANKACCT AS STRING * 20
  END TYPE
  
  TYPE BDRptType
    BANKNAME  AS STRING * 14
    CustRec   AS INTEGER
    TransRec  AS LONG
  END TYPE
  
  TYPE BankTotalsType
    BANKNAME  AS STRING * 14
    Amount    AS DOUBLE
  END TYPE
  
  CONST False = 0, True = NOT False

SUB UBAcctsToDraft
  
  Dash80$ = STRING$(80, "-")
  
  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM BankTotals(1 TO 1) AS BankTotalsType

  'LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  'TOWNNAME$ = UBSetUpRec(1).UTILNAME
  
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
    IF (UBCustRec(1).USEDRAFT = "Y") OR (LEN(QPTrim$(UBCustRec(1).BANKNAME)) > 0) THEN
      IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) > 0 THEN
        IF LineCnt > MaxLines THEN
          PRINT #UBRpt, CHR$(12)
          GOSUB PrintBankDFTHeader
        END IF
        CstCnt = CstCnt + 1
        REDIM PRESERVE BDCust(1 TO CstCnt) AS BDRptType
        BDCust(CstCnt).BANKNAME = QPTrim$(UBCustRec(1).BANKNAME)
        BDCust(CstCnt).CustRec = Cnt&
        BDCust(CstCnt).TransRec = Cnt&
      END IF
    END IF
    END IF
    'IF AskAbandonPrint% THEN
    '  ABExit = True
    '  GOTO NON2PrintExit:
    'END IF
    'ShowPctCompL Cnt&, NumOfRecs&
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

  PrevBank$ = BDCust(1).BANKNAME

  BankTotals(BankCnt).BANKNAME = BDCust(1).BANKNAME

  FOR Cnt = 1 TO CstCnt
    GET UBCust, BDCust(Cnt).CustRec, UBCustRec(1)
    IF PrevBank$ <> BDCust(Cnt).BANKNAME THEN
      BankCnt = BankCnt + 1
      REDIM PRESERVE BankTotals(1 TO BankCnt) AS BankTotalsType
      BankTotals(BankCnt).BANKNAME = BDCust(Cnt).BANKNAME
      PrevBank$ = BDCust(Cnt).BANKNAME
    END IF
    BankTotals(BankCnt).Amount = Round#(BankTotals(BankCnt).Amount + UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB PrintBankDFTHeader
    END IF
    LSET DFTRec(1).TRANSIT = QPTrim$(UBCustRec(1).TRANSIT)
    LSET DFTRec(1).BANKNAME = QPTrim$(UBCustRec(1).BANKNAME)
    RSET DFTRec(1).CUSTACCT = QPTrim$(STR$(BDCust(Cnt).CustRec))
    LSET DFTRec(1).CUSTNAME = QPTrim$(UBCustRec(1).CUSTNAME)
    LSET DFTRec(1).BANKACCT = QPTrim$(UBCustRec(1).BANKACCT)
    LSET DFTRec(1).BillAmt = FUsing$(STR$(Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)), "#####.##")
    PRINT #UBRpt, DFTRec(1).TRANSIT; " "; DFTRec(1).BANKNAME;
    PRINT #UBRpt, DFTRec(1).CUSTACCT; "  "; DFTRec(1).CUSTNAME;
    PRINT #UBRpt, " "; DFTRec(1).BillAmt; "  "; DFTRec(1).BANKACCT
    LineCnt = LineCnt + 1
    'IF AskAbandonPrint% THEN
    '  ABExit = True
    '  GOTO NON2PrintExit:
    'END IF
    'ShowPctComp Cnt, CstCnt
  NEXT

  PRINT #UBRpt, CHR$(12)
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Billing Bank Draft Register.                "; QPTrim$(TOWNNAME$)
  PRINT #UBRpt, "Date: "; DATE$; TAB(72); "Page: "; PageNo
  PRINT #UBRpt, "Bank Name                 Bank Total."
  PRINT #UBRpt, Dash80$
  GTotal# = 0
  FOR Cnt = 1 TO BankCnt
    PRINT #UBRpt, BankTotals(Cnt).BANKNAME; TAB(30); FUsing$(STR$(BankTotals(Cnt).Amount), "#####.##")
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

