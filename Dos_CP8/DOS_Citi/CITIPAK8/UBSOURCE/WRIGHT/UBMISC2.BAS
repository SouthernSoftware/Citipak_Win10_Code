DEFINT A-Z

DECLARE SUB ReLinkHistory ()
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ReIndexSystem ()
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB WaitForAction ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB KillFile (File2Kill$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB QPrintRC (Text$, Row%, Col%, Kolor)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum)
  
'$INCLUDE: 'DefCnf.BI'
  
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION GetNumRateRecs ()
DECLARE SUB ClearScrn ()
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'newcust.bi'
  
TYPE UBTransRecTypeS
   TransDate              AS INTEGER      '
   TransType              AS INTEGER      '
   TransDesc              AS STRING * 21  'may change
   TransAmt               AS DOUBLE       'total revenue amount
   RevAmt(1 TO 15)        AS DOUBLE       'Revenue amounts
   TaxAmt(1 TO 15)        AS SINGLE       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 TO 7)       AS INTEGER
'*******************
   CurRead(1 TO 7)        AS LONG         'Last/Current meter readings
   PrevRead(1 TO 7)       AS LONG         'Previous readings
   EstRead(1 TO 7)        AS STRING * 1   'Y/N Flags for meter est's
   BillNumber             AS LONG         'Number on the bill that Printed
   ReadDate               AS INTEGER
   BillDate               AS INTEGER
   PastDueDate            AS INTEGER
   DraftDate              AS INTEGER
   CustLocation           AS LONG         'Pointer to Location RecNo
   CustStatus             AS STRING * 1   'Customer Status at Time of Transaction
   CustCode               AS STRING * 2   'Customer Code at Time of Transaction
   PostRoute              AS STRING * 4   'Postal Route at Time of Transaction
   BillCycle              AS STRING * 2   'Customer Billing Cycle at Time of Transaction
   PayTypeCode            AS INTEGER      'Payment Type:  1=Cash, 2=Check
   OperatorNumber         AS INTEGER      '
   CustAcctNo             AS LONG         'Pointer to RecNo in ubcust.dat
END TYPE
'   PrevTrans              AS LONG
'   NextTrans              AS LONG

TYPE NewUBCustRecTypeS
    BOOK     AS STRING * 2
    SEQNUMB  AS STRING * 6
    STATUS  AS STRING * 1
    OPENDATE AS INTEGER
    SEARCH   AS STRING * 10
    CUSTNAME AS STRING * 35
    ADDR1    AS STRING * 35
    ADDR2    AS STRING * 35
    SERVADDR AS STRING * 35
    City     AS STRING * 18
    STATE    AS STRING * 2
    ZipCode  AS STRING * 10
    HPHONE   AS STRING * 14
    WPHONE   AS STRING * 14
    SOSEC    AS STRING * 11
    DRVLIC   AS STRING * 16
    CUSTTYPE AS STRING * 3
    Addr911  AS STRING * 15
    BILLCOPY AS INTEGER
    POSTRTE  AS STRING * 4
    BILLCYCL AS INTEGER
    ZONE     AS STRING * 3
    SEQ      AS LONG
    CASHONLY AS STRING * 1
    LATEFEE  AS STRING * 1
    CUTOFFYN AS STRING * 1
    TAXEXPT  AS STRING * 1
    SRCIT    AS STRING * 1
    EPP      AS STRING * 1
    EPPAMT   AS DOUBLE
    USEDRAFT AS STRING * 1
    BANKNAME AS STRING * 34
    BANKLOC  AS STRING * 30
    TRANSIT  AS STRING * 9
    BANKACCT AS STRING * 20
    BILLCMNT AS STRING * 25
    PAYCMNT  AS STRING * 25
    PUMPCODE AS STRING * 4
    USERCODE1   AS STRING * 4
    USERCODE2   AS STRING * 4
    HHMSG1   AS STRING * 20
    HHMSG2   AS STRING * 20
    HHMSG3   AS STRING * 20
    Serv(1 TO 15)        AS ServicesType
    FlatRates(1 TO 4)    AS FlatRateType
    Monthly(1 TO 2)      AS MonthlyPayType
    MFEE1    AS DOUBLE
    MFEE2    AS DOUBLE
    LocMeters(1 TO 7)    AS LocMeterType
    CustPIN              AS LONG
END TYPE
'    LastTrans            AS LONG



DIM TRec AS UBTransRecType
DIM CRec AS NewUBCustRecType

   TRecLen = LEN(TRec)
   CRecLen = LEN(CRec)


  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 8)
  
  MChoice$(1) = "Relink Transactions History"
  MChoice$(2) = "Reindex Utility Files "
  MChoice$(3) = "Exit to OS"
  

  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Miscellaneous Functions ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO   'choice = 0
    END IF
    
    SELECT CASE Choice
    CASE 1
      ReLinkHistory
    CASE 2
      ReIndexSystem
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag
  
'  IF WEnvTest AND INSTR(COMMAND$, "TEST") THEN
'    HideCursor
'    ClearScrn
'    END
'  ELSE
'    RUN "ubmenu"
'  END IF
  
'StartHere:

SUB ReIndexSystem
  
'  UBLog " IN: Reindex Utility Files"
  BlockClear
  Ok = MsgBox%("UB", "MUSTEXIT")
  
  SELECT CASE Ok
  CASE 2
    GOTO ExitReindex
  END SELECT
  
  'BlockClear
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure
  
  ShowProcessingScrn "Reading Customer Names"
  'UBLog "BEGIN: Customer Name Reindex"
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs&) AS UBCustReIndexRecType
  
  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
    IdxBuff(Cnt).First = QPTrim$(UBCustRec(1).CUSTNAME)
    IdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  
  CLOSE UBFile
  
  QPrintRC " Sorting Customer Names", 11, 29, -1
  
  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 10, 2
  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 10
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
  
  GOSUB ClearBlock
  QPrintRC "Writing Customer Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  KillFile "UBCUSTNM.IDX"
  
  UBFile = FREEFILE
  OPEN "UBCUSTNM.IDX" FOR RANDOM SHARED AS UBFile LEN = 4
  
  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, IdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile
  ERASE IdxBuff
  'UBLog "FINISH: Customer Name Reindex"
  GOSUB ClearBlock

  QPrintRC "Reading Location Information", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1
  'UBLog "BEGIN: Book\Sequence Reindex"
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  REDIM LIdxBuff(1 TO NumOfRecs&) AS UBLocaReIndexRecType
  
  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    LIdxBuff(Cnt).BOOK = UBCustRec(1).BOOK
    LIdxBuff(Cnt).SEQNUMB = UBCustRec(1).SEQNUMB
    LIdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  
  CLOSE UBFile
  
  QPrintRC " Sorting Locations Names", 11, 29, -1
  
  SortT LIdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 8
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
  
  GOSUB ClearBlock
  QPrintRC "Writing Location Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  KillFile "UBCUSTBK.IDX"
  
  UBFile = FREEFILE
  OPEN "UBCUSTBK.IDX" FOR RANDOM SHARED AS UBFile LEN = 4
  
  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, LIdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile
  'UBLog "FINISH: Book\Sequence Reindex"
  REDIM BookSeq(1) AS BookSeqRecType
  
  KillFile "UBOOKSEQ.DAT"
  'UBLog "BEGIN: Rebuild Book\Sequence List"
  BookHand = FREEFILE
  OPEN "UBOOKSEQ.DAT" FOR RANDOM SHARED AS BookHand LEN = 4
  FOR Cnt = 1 TO NumOfRecs&
    BookSeq(1).BookSeq = QPValL(LIdxBuff(Cnt).BOOK + LIdxBuff(Cnt).SEQNUMB)
    PUT BookHand, Cnt, BookSeq(1)
  NEXT
  CLOSE BookHand
  'UBLog "FINISH: Rebuild Book\Sequence List"
  ERASE LIdxBuff, BookSeq
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitReindex:
  'UBLog "OUT: Reindex Utility Files" + CrLf$
  EXIT SUB
  
ClearBlock:
  HideCursor
  Blank$ = SPACE$(40)
  FOR Cnt = 8 TO 15
    QPrintRC Blank$, Cnt, 21, -1
  NEXT
  ShowCursor
  RETURN
  
END SUB

SUB ReLinkHistory

  'UBLog " IN: Relink Transaction History"

  TOffSet = 323    'TransRec: (prev trans pointer) rec offset
  TCOffSet = 319   'TransRec: (cust acct pointer) rec offset
  COffSet = 1053   ' CustRec: (last trans pointer) rec offset
  CLTPos& = 0      ' CustRec: (last trans pointer) position offset
  
  BlockClear
  Ok = MsgBox%("UB", "MUSTEXIT")
  
  SELECT CASE Ok
  CASE 2
    'UBLog "ABORTED: Relink Transaction History"
    GOTO ExitRelink
  END SELECT
  
  'BlockClear
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure
  
  ShowProcessingScrn "Relinking Transaction History"
  
  QPrintRC "        Pass: 1 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR BINARY SHARED AS UBTran
  NumOfTRecs& = LOF(UBTran) \ UBTranRecLen

  UBFile2 = FREEFILE
  OPEN "UBCUST.DAT" FOR BINARY SHARED AS UBFile2
  NumOfCRecs& = LOF(UBFile2) \ UBCustRecLen

  'UBLog "BEGIN: Pass 1 of 2"

  Zero& = 0
  FOR CCnt& = 1 TO NumOfCRecs&
      'calc byte offset in cust rec
    CLTPos& = (((CCnt& - 1) * UBCustRecLen) + COffSet) + 1
      'seek to offset
    SEEK UBFile2, CLTPos&
      'put a "0" in that position
    PUT UBFile2, , Zero&
    ShowPctComp CCnt&, NumOfCRecs&
  NEXT
  'UBLog "FINISH: Pass 1 of 2"

  'GOTO JumpHere
  
  QPrintRC "        Pass: 2 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  'UBLog "BEGIN: Pass 2 of 2"
  FOR TCnt& = 1 TO NumOfTRecs&
    TPos& = (((TCnt& - 1) * UBTranRecLen) + TCOffSet) + 1
    SEEK UBTran, TPos&
    GET UBTran, , TCustAct&

    'GET UBTran, TCnt&, UBTransRec(1)
    '^^ this is to get the customer account from this transaction
    IF (TCustAct& > 0) AND (TCustAct& <= NumOfCRecs&) THEN

      CLTPos& = (((TCustAct& - 1) * UBCustRecLen) + COffSet) + 1
      SEEK UBFile2, CLTPos&
      GET UBFile2, , LastTrans&
      'this is to get the cust last trans pointer
      SEEK UBFile2, CLTPos&
      PUT UBFile2, , TCnt&
      'put new last trans pointer in cust rec
      PUT UBTran, , LastTrans&
      'put new last trans pointer in trans rec

    END IF
    ShowPctCompL TCnt&, NumOfTRecs&
  NEXT

'********
JumpHere:

  CLOSE
  'UBLog "FINISH: Pass 2 of 2"
  'UBLog "RELINK: Transactions Completed."
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitRelink:
  'UBLog "OUT: Relink Transaction History" + CrLf$

EXIT SUB
  
  
END SUB

