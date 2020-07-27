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
    BANKNAME AS STRING * 34
    CUSTACCT AS STRING * 8
    CUSTNAME AS STRING * 35
    BillAmt  AS STRING * 10
    BANKACCT AS STRING * 20
  END TYPE
  
  TYPE BDRptType
    BANKNAME  AS STRING * 14
    CustRec   AS INTEGER
    TransRec  AS LONG
  END TYPE
  
  CONST False = 0, True = NOT False
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  MChoice$(1) = "Accounts To Draft Report"
  MChoice$(2) = "Prepare Draft Transmission File"
  MChoice$(3) = "Prepare Draft Prenote File"
  MChoice$(4) = "Print Draft Customer Listing"
  MChoice$(5) = "Prepare Draft Test File for Bank"
  MChoice$(6) = "Exit to OS "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "ACH - Draft Processing System ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO
    END IF
    
    SELECT CASE Choice
    CASE 1
      UBAcctsToDraft
    CASE 2
      UBBuildTransmitFile
    CASE 3
      UBPrenote
    CASE 4
      UBDraftListing
    CASE 5
      UBDraftTest
    CASE 6
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag
  
  RUN "UBbillin"

SUB UBAcctsToDraft
  
  Dash80$ = STRING$(125, "-")
  
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
  
  '*********************************
  LibName$ = "UB"
  ScrnName$ = "DRFTINFO"
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      BDate = Date2Num(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      ABExit = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 29 TO 40           '--Cancel button
          PressButton 27, 14, 29, 40
        CASE 43 TO 54           '--F10 button
          PressButton 13, 14, 43, 54
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  ERASE Frm, Form$, Fld
  
  IF ABExit THEN
    GOTO AbortExit
  END IF
  
  '*********************************
  
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
  
  Split = CstCnt
  PrevBank$ = BDCust(1).BANKNAME
  FOR Cnt = 1 TO CstCnt
    ThisBank$ = BDCust(Cnt).BANKNAME
    IF ThisBank$ <> PrevBank$ THEN
      Split = Cnt - 1
      EXIT FOR
    END IF
  NEXT
  
  IF Split < CstCnt THEN
    SortT BDCust(1), Split, 0, 20, 14, -1
    SortT BDCust(Split + 1), CstCnt - Split, 0, 20, 14, -1
  END IF
  '(SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  GET UBCust, BDCust(1).CustRec, UBCustRec(1)
  PrevBank$ = BDCust(1).BANKNAME
  
  FOR Cnt = 1 TO CstCnt
    GET UBCust, BDCust(Cnt).CustRec, UBCustRec(1)
    ThisBank$ = BDCust(Cnt).BANKNAME
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
    LSET DFTRec(1).BANKNAME = QPTrim$(UBCustRec(1).BANKNAME)
    RSET DFTRec(1).CUSTACCT = QPTrim$(STR$(BDCust(Cnt).CustRec))
    LSET DFTRec(1).CUSTNAME = QPTrim$(UBCustRec(1).CUSTNAME)
    LSET DFTRec(1).BANKACCT = QPTrim$(UBCustRec(1).BANKACCT)
    LSET DFTRec(1).BillAmt = FUsing$(STR$(Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)), "#####.##")
    PRINT #UBRpt, DFTRec(1).TRANSIT; "  "; DFTRec(1).BANKNAME;
    PRINT #UBRpt, "  "; DFTRec(1).CUSTACCT; "  "; DFTRec(1).CUSTNAME;
    PRINT #UBRpt, "  "; DFTRec(1).BillAmt; "  "; "  "; DFTRec(1).BANKACCT
    LineCnt = LineCnt + 1
    IF AskAbandonPrint% THEN
      ABExit = True
      GOTO NON2PrintExit:
    END IF
    ShowPctComp Cnt, CstCnt
  NEXT
  
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, CHR$(12)
  
NON2PrintExit:
  CLOSE
  
  IF ABExit THEN
    GOTO AbortExit
  END IF
  
  ERASE UBSetUpRec, MChoice$, DFTRec, UBCustRec
  
  LPTPort = 1
  IF NOT AbortFlag THEN
    PrintRptFile "Bank Draft Register Report", "UBANKDFT.RPT", LPTPort, RetCode, EntryPoint
  END IF
  
AbortExit:
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

SUB UBBuildTransmitFile
  
  SHARED Choice$()
  LibName$ = "UB"
  ScrnName$ = "UBDFTFIN"
  
  '*****************
  REDIM DFTRec(1) AS DraftRptType
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  NumOfRecs& = LOF(UBCust) \ UBCustRecLen
  
  '*****************
  CursorOff

  BlockClear
  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      BDate = Date2Num(Form$(1, 0))
      GOSUB ProcessDraft
      Done = True
    CASE ESC
      Done = True
    END SELECT
    
  LOOP UNTIL Done
  EXIT SUB
  
ProcessDraft:
  
  DraftDate$ = Form$(1, 0)
  DraftDate$ = LEFT$(DraftDate$, 2) + MID$(DraftDate$, 4, 2) + RIGHT$(DraftDate$, 2)
  
  IF VAL(DraftDate$) = 0 THEN EXIT SUB
  
  LibName$ = "UB"
  ScrnName$ = "UBDFTDF1"
  
  CursorOff

  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  
  CursorOff
  
  LSET Form$(1, 0) = "Building Record Type 1"
  Action = 1
  
  DO
DraftFormLoop:
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    ' Process Record Type 1
    IF NOT Step1 THEN
      GOSUB DraftProcessStep1
      LSET Form$(1, 0) = "Building Record Type 1 ..Done!"
      LSET Form$(2, 0) = "Building Record Type 5"
      Action = 1
      GOTO DraftFormLoop
    END IF
    
    IF NOT Step2 THEN
      GOSUB DraftProcessStep2
      LSET Form$(2, 0) = "Building Record Type 5 ..Done!"
      LSET Form$(3, 0) = "Building Record Type 6 "
      Action = 1
      GOTO DraftFormLoop
    END IF
    
    IF NOT Step3 THEN
      GOSUB DraftProcessStep3
      LSET Form$(3, 0) = "Building Record Type 6 ..Done!"
      LSET Form$(4, 0) = "Building Record Type 8 "
      Action = 1
      GOTO DraftFormLoop
    END IF
    
    IF NOT Step4 THEN
      GOSUB DraftProcessStep4
      LSET Form$(4, 0) = "Building Record Type 8 ..Done!"
      LSET Form$(5, 0) = "Building Record Type 9"
      Action = 1
      GOTO DraftFormLoop
    END IF
    
    IF NOT Step5 THEN
      GOSUB DraftProcessStep5
      LSET Form$(5, 0) = "Building Record Type 9 ..Done!"
      LSET Form$(6, 0) = "File Name Is: DS" + DraftDate$
      Action = 1
      GOTO DraftFormLoop
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE ESC
      Done = True
    END SELECT
    
  LOOP UNTIL Done
  
  CLOSE
  
  EXIT SUB
  
RETURN

OpenMainDraftInformation:
  REDIM UBDraftRec(1) AS UBDraftRecType
  DraftFile = FREEFILE
  OPEN "UBSDRAFT.dat" FOR RANDOM ACCESS READ SHARED AS #DraftFile LEN = LEN(UBDraftRec(1))
  GET DraftFile, 1, UBDraftRec(1)
RETURN
  
DraftProcessStep1:
  GOSUB OpenMainDraftInformation
  
  REDIM UBDraftRecord1(1) AS UBDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord1(1))
  UBDraftRecord1(1).Field1 = "1"
  UBDraftRecord1(1).Field2 = "01"
  UBDraftRecord1(1).Field3 = " " + UBDraftRec(1).BANKDEST
  UBDraftRecord1(1).Field4 = " " + UBDraftRec(1).BANKORIG
  UBDraftRecord1(1).Field5 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord1(1).Field6 = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  UBDraftRecord1(1).Field7 = "A"
  UBDraftRecord1(1).Field8 = "094"
  UBDraftRecord1(1).Field9 = "10"
  UBDraftRecord1(1).Field10 = "1"
  UBDraftRecord1(1).Field11 = UCASE$(UBDraftRec(1).BANKNAME)
  UBDraftRecord1(1).Field12 = UCASE$(UBDraftRec(1).BANKLOC)
  UBDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  PUT DraftFileNum, 1, UBDraftRecord1(1)
  CLOSE DraftFileNum
  Step1 = True
RETURN
  
DraftProcessStep2:
  REDIM UBDraftRecord5(1) AS UBDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT5.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord5(1))
  UBDraftRecord5(1).Field1 = "5"
  UBDraftRecord5(1).Field2 = "200"
  UBDraftRecord5(1).Field3 = "WADESBORO NC"
  UBDraftRecord5(1).Field4 = "                    "
  UBDraftRecord5(1).Field5 = "0" + UBDraftRec(1).FEDID
  UBDraftRecord5(1).Field6 = "PPD"
  UBDraftRecord5(1).Field7 = "UTIL BILL"
  UBDraftRecord5(1).Field8 = RIGHT$(DraftDate$, 2) + LEFT$(DraftDate$, 2) + MID$(DraftDate$, 3, 2)
  UBDraftRecord5(1).Field9 = RIGHT$(DraftDate$, 2) + LEFT$(DraftDate$, 2) + MID$(DraftDate$, 3, 2)
  UBDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  UBDraftRecord5(1).Field11 = "1"
  UBDraftRecord5(1).Field12 = LEFT$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord5(1).Field13 = "0000001"
  PUT DraftFileNum, 1, UBDraftRecord5(1)
  CLOSE DraftFileNum
  Step2 = True
RETURN
  
DraftProcessStep3:
  COUNTER = 0
  
  REDIM UBDraftRecord6(1) AS UBDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  CLOSE DraftFileNum
  KILL "UBDRAFT6.DAT"
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  
  'GO THRU DATA FILE HERE
  
  FOR Cnt& = 1 TO NumOfRecs&
    
    GET UBCust, Cnt&, UBCustRec(1)
    IF UBCustRec(1).STATUS = "A" OR UBCustRec(1).STATUS = "B" THEN
    IF (UBCustRec(1).USEDRAFT = "Y") OR (LEN(QPTrim$(UBCustRec(1).BANKNAME)) > 0) THEN
      CustBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      IF CustBal# > 0 THEN
        ' Process Customer Here
        BillAmt$ = (STR$(CustBal# * 100))       'Remove the Decimals
        BillAmt$ = LTRIM$(BillAmt$)
        IF LEN(BillAmt$) < 10 THEN BillAmt$ = STRING$(10 - LEN(BillAmt$), "0") + BillAmt$
        'Get Running Billed Amount
        TotalAmount# = TotalAmount# + (CustBal# * 100)
        COUNTER = COUNTER + 1
        AcctNumber$ = STR$(Cnt&)
        AcctNumber$ = RIGHT$(AcctNumber$, LEN(AcctNumber$) - 1)
        IF LEN(AcctNumber$) < 15 THEN
          AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
        END IF
        nme$ = UBCustRec(1).CUSTNAME
        IF LEN(nme$) < 22 THEN
          nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
        ELSE
          nme$ = LEFT$(nme$, 22)
        END IF
        BANKACCT$ = QPTrim$(UBCustRec(1).BANKACCT)
        IF LEN(BANKACCT$) < 17 THEN BANKACCT$ = BANKACCT$ + STRING$(17 - LEN(BANKACCT$), 32)
        Trace = Trace + 1
        Trace$ = STR$(Trace): Trace$ = RIGHT$(Trace$, LEN(Trace$) - 1)
        IF LEN(Trace$) < 7 THEN Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
        
        UBDraftRecord6(1).Field1 = "6"
        UBDraftRecord6(1).Field2 = "27"         ' Designates Debit Trans
        UBDraftRecord6(1).Field3 = LEFT$(UBCustRec(1).TRANSIT, 8)
        UBDraftRecord6(1).Field4 = RIGHT$(UBCustRec(1).TRANSIT, 1)
        UBDraftRecord6(1).Field5 = LEFT$(BANKACCT$, 17)
        UBDraftRecord6(1).Field6 = BillAmt$     ' All zero's for Prenote
        UBDraftRecord6(1).Field7 = AcctNumber$
        UBDraftRecord6(1).Field8 = UCASE$(nme$)
        UBDraftRecord6(1).Field9 = "  "
        UBDraftRecord6(1).Field10 = "0"
        UBDraftRecord6(1).Field11 = LEFT$(UBDraftRec(1).BANKORIG, 8) + Trace$
        PUT DraftFileNum, COUNTER, UBDraftRecord6(1)
        hash# = hash# + VAL(LEFT$(UBCustRec(1).TRANSIT, 8))
        number = number + 1
      END IF
    END IF
    END IF
  NEXT
  
  CLOSE DraftFileNum
  Step3 = True
RETURN
  
DraftProcessStep4:
  hash$ = STR$(hash#)
  hash$ = RIGHT$(hash$, LEN(hash$) - 1)
  
  IF LEN(hash$) < 10 THEN
    hash$ = STRING$(10 - LEN(hash$), "0") + hash$
  END IF
  IF LEN(hash$) > 10 THEN
    hash$ = RIGHT$(hash$, 10)
  END IF
  
  IF LEN(Trace$) > 6 THEN Trace$ = RIGHT$(Trace$, 6)
  TotalAmount$ = STR$(TotalAmount#)
  TotalAmount$ = RIGHT$(TotalAmount$, LEN(TotalAmount$) - 1)
  IF LEN(TotalAmount$) < 12 THEN TotalAmount$ = STRING$(12 - LEN(TotalAmount$), "0") + TotalAmount$
        
  REDIM UBDraftRecord8(1) AS UBDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT8.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord8(1))
  UBDraftRecord8(1).Field1 = "8"
  UBDraftRecord8(1).Field2 = "200"
  UBDraftRecord8(1).Field3 = Trace$
  UBDraftRecord8(1).Field4 = hash$
  UBDraftRecord8(1).Field5 = TotalAmount$       ' zero for prenote
  UBDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field7 = "0" + UBDraftRec(1).FEDID
  UBDraftRecord8(1).Field8 = STRING$(19, 32)    ' Reserved
  UBDraftRecord8(1).Field9 = STRING$(6, 32)     ' Reserved for Federal Reserve Use
  UBDraftRecord8(1).Field10 = LEFT$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord8(1).Field11 = "0000001"
  PUT DraftFileNum, 1, UBDraftRecord8(1)
  CLOSE DraftFileNum
  Step4 = True
RETURN
  
DraftProcessStep5:
  TotSize# = VAL(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records
  
  IF BlockSize! <> INT(BlockSize!) THEN
    BlockSize! = INT(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize - 1)))
  ELSE
    FillSize! = 0
  END IF
  
  BlockSize$ = STR$(BlockSize!)
  BlockSize$ = RIGHT$(BlockSize$, LEN(BlockSize$) - 1)
  IF LEN(BlockSize$) < 6 THEN BlockSize$ = STRING$(6 - LEN(BlockSize$), "0") + BlockSize$
  IF LEN(Trace$) < 8 THEN Trace$ = STRING$(8 - LEN(Trace$), "0") + Trace$
  
  REDIM UBDraftRecord9(1) AS UBDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT9.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord9(1))
  UBDraftRecord9(1).Field1 = "9"
  UBDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  UBDraftRecord9(1).Field3 = BlockSize$
  UBDraftRecord9(1).Field4 = Trace$
  UBDraftRecord9(1).Field5 = hash$
  UBDraftRecord9(1).Field6 = TotalAmount$       ' zero for prenote
  UBDraftRecord9(1).Field7 = "000000000000"
  UBDraftRecord9(1).Field8 = STRING$(39, 32)    ' Reserved
  PUT DraftFileNum, 1, UBDraftRecord9(1)
  CLOSE DraftFileNum
  
  ' Now Put Them Together In File Name UBDFNOTE
  OutFile = FREEFILE
  OPEN "O", OutFile, "DS" + DraftDate$: WIDTH #OutFile, 255
  
  REDIM UBDraftRecord1(1) AS UBDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord1(1))
  GET DraftFileNum, 1, UBDraftRecord1(1)
  PRINT #OutFile, UBDraftRecord1(1).Field1;
  PRINT #OutFile, UBDraftRecord1(1).Field2;
  PRINT #OutFile, UBDraftRecord1(1).Field3;
  PRINT #OutFile, UBDraftRecord1(1).Field4;
  PRINT #OutFile, UBDraftRecord1(1).Field5;
  PRINT #OutFile, UBDraftRecord1(1).Field6;
  PRINT #OutFile, UBDraftRecord1(1).Field7;
  PRINT #OutFile, UBDraftRecord1(1).Field8;
  PRINT #OutFile, UBDraftRecord1(1).Field9;
  PRINT #OutFile, UBDraftRecord1(1).Field10;
  PRINT #OutFile, UBDraftRecord1(1).Field11;
  PRINT #OutFile, UBDraftRecord1(1).Field12;
  PRINT #OutFile, UBDraftRecord1(1).Field13
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord5(1) AS UBDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT5.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord5(1))
  GET DraftFileNum, 1, UBDraftRecord5(1)
  PRINT #OutFile, UBDraftRecord5(1).Field1;
  PRINT #OutFile, UBDraftRecord5(1).Field2;
  PRINT #OutFile, UBDraftRecord5(1).Field3;
  PRINT #OutFile, UBDraftRecord5(1).Field4;
  PRINT #OutFile, UBDraftRecord5(1).Field5;
  PRINT #OutFile, UBDraftRecord5(1).Field6;
  PRINT #OutFile, UBDraftRecord5(1).Field7;
  PRINT #OutFile, UBDraftRecord5(1).Field8;
  PRINT #OutFile, UBDraftRecord5(1).Field9;
  PRINT #OutFile, UBDraftRecord5(1).Field10;
  PRINT #OutFile, UBDraftRecord5(1).Field11;
  PRINT #OutFile, UBDraftRecord5(1).Field12;
  PRINT #OutFile, UBDraftRecord5(1).Field13
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord6(1) AS UBDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94
  
  FOR Cnt = 1 TO NumOfLines
    GET DraftFileNum, Cnt, UBDraftRecord6(1)
    PRINT #OutFile, UBDraftRecord6(1).Field1;
    PRINT #OutFile, UBDraftRecord6(1).Field2;
    PRINT #OutFile, UBDraftRecord6(1).Field3;
    PRINT #OutFile, UBDraftRecord6(1).Field4;
    PRINT #OutFile, UBDraftRecord6(1).Field5;
    PRINT #OutFile, UBDraftRecord6(1).Field6;
    PRINT #OutFile, UBDraftRecord6(1).Field7;
    PRINT #OutFile, UBDraftRecord6(1).Field8;
    PRINT #OutFile, UBDraftRecord6(1).Field9;
    PRINT #OutFile, UBDraftRecord6(1).Field10;
    PRINT #OutFile, UBDraftRecord6(1).Field11
  NEXT Cnt
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord8(1) AS UBDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT8.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord8(1))
  GET DraftFileNum, 1, UBDraftRecord8(1)
  PRINT #OutFile, UBDraftRecord8(1).Field1;
  PRINT #OutFile, UBDraftRecord8(1).Field2;
  PRINT #OutFile, UBDraftRecord8(1).Field3;
  PRINT #OutFile, UBDraftRecord8(1).Field4;
  PRINT #OutFile, UBDraftRecord8(1).Field5;
  PRINT #OutFile, UBDraftRecord8(1).Field6;
  PRINT #OutFile, UBDraftRecord8(1).Field7;
  PRINT #OutFile, UBDraftRecord8(1).Field8;
  PRINT #OutFile, UBDraftRecord8(1).Field9;
  PRINT #OutFile, UBDraftRecord8(1).Field10;
  PRINT #OutFile, UBDraftRecord8(1).Field11
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord9(1) AS UBDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT9.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord9(1))
  GET DraftFileNum, 1, UBDraftRecord9(1)
  PRINT #OutFile, UBDraftRecord9(1).Field1;
  PRINT #OutFile, UBDraftRecord9(1).Field2;
  PRINT #OutFile, UBDraftRecord9(1).Field3;
  PRINT #OutFile, UBDraftRecord9(1).Field4;
  PRINT #OutFile, UBDraftRecord9(1).Field5;
  PRINT #OutFile, UBDraftRecord9(1).Field6;
  PRINT #OutFile, UBDraftRecord9(1).Field7;
  PRINT #OutFile, UBDraftRecord9(1).Field8
  CLOSE DraftFileNum
  FOR Cnt = 1 TO FillSize! / 94
    PRINT #OutFile, STRING$(94, "9")
  NEXT Cnt
  CLOSE
  Step5 = True
RETURN
  
  
END SUB

SUB UBDraftListing
  
  TotalDraftCustomers = 0
  Dash80$ = STRING$(80, "-")
  Temp1$ = SPACE$(10)
  Temp2$ = SPACE$(12)
  
  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  MaxLines = 50
  
  REDIM MChoice$(1 TO 2)
  
  MChoice$(1) = " Customer Name Order   "
  MChoice$(2) = " Account Number Order  "
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
ReStart:
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  PageNo = 0
  
  DO
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Draft Customer Report ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
    END IF
    
    SELECT CASE Choice
    CASE 0
      ExitFlag = True
    CASE 1
      IndexName$ = NameIndexFile
      UsingName = True
      OKFlag = True
    CASE 2
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  IF ExitFlag THEN GOTO ExitPreReport
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IF UsingAcct THEN             'load the index
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBDFTCUS.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Processing Draft Customer Report"
  
  GOSUB PrintPreHeader
  
  FOR Cnt = 1 TO NumOfRecs
    ShowPctComp Cnt, NumOfRecs
    IF UsingAcct THEN
      GET UBCust, Cnt, UBCustRec(1)
      AcctRecord = Cnt
    ELSE
      GET UBCust, IndexArray(Cnt).RecNum, UBCustRec(1)
      AcctRecord = IndexArray(Cnt).RecNum
    END IF
    '  Process Customer Here
    IF UBCustRec(1).STATUS = "A" THEN
      IF LineCnt >= 60 THEN
        PRINT #UBRpt, CHR$(12); : GOSUB PrintPreHeader
      END IF
      IF UBCustRec(1).USEDRAFT = "Y" THEN
        TotalDraftCustomers = TotalDraftCustomers + 1
        PRINT #UBRpt, USING "#####"; AcctRecord;
        PRINT #UBRpt, TAB(10); UBCustRec(1).CUSTNAME;
        PRINT #UBRpt, TAB(48); LEFT$(UBCustRec(1).TRANSIT, 9);
        PRINT #UBRpt, TAB(59); LEFT$(UBCustRec(1).BANKACCT, 20)
        IF UBCustRec(1).PreNoteFlag = 0 THEN
          PRINT #UBRpt, TAB(2); "N";
        ELSE
          PRINT #UBRpt, TAB(2); "Y";
        END IF
        PRINT #UBRpt, TAB(15); QPTrim$(UBCustRec(1).BANKNAME);
        PRINT #UBRpt, TAB(50); QPTrim$(LEFT$(UBCustRec(1).BANKLOC, 20))
        PRINT #UBRpt, STRING$(79, "-")
        LineCnt = LineCnt + 1
      END IF
    END IF
    
  NEXT Cnt
  PRINT #UBRpt, "Total Draft Customers on File: "; USING "####,#"; TotalDraftCustomers
  PRINT #UBRpt, CHR$(12);
  
  CLOSE
  
  SELECT CASE Choice
  CASE 1
    RptText$ = "(Customer Order)"
  CASE 2
    RptText$ = "(Account Order)"
  END SELECT
  
  IF NOT AbortFlag THEN
    PrintRptFile " Draft Customer Report " + RptText$, "UBDFTCUS.RPT", LPTPort, RetCode, EntryPoint
  END IF
  
  CLOSE
  
  '  IF NOT ExitFlag THEN GOTO ReStart
  EXIT SUB
  
PrintPreHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Draft Customer Listing"; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "Date: "; DATE$
  PRINT #UBRpt,
  PRINT #UBRpt, "Acct #"; TAB(10); "Customer Name"; TAB(48); "Transit #"; TAB(60); "Bank Acct #"
  PRINT #UBRpt, "Prenoted??"; TAB(15); "Bank Name & Location"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  
  RETURN
  
ExitPreReport:
  
END SUB

SUB UBDraftTest
  
  SHARED Choice$()
  LibName$ = "UB"
  ScrnName$ = "UBDFTTST"
  
  CursorOff
  
  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB ProcessTest
      Done = True
    CASE ESC
      Done = True
    END SELECT
    
  LOOP UNTIL Done
  EXIT SUB
  
ProcessTest:
  
  LibName$ = "UB"
  ScrnName$ = "UBDFTTS1"
  
  CursorOff
  
  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  
  Form$(1, 0) = "Building Record Type 1"
  Action = 1
  
  DO
FormTestLoop:
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    ' Process Record Type 1
    
    IF NOT Step1 THEN
      GOSUB TestProcessStep1
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5"
      Action = 1
      GOTO FormTestLoop
    END IF
    
    IF NOT Step2 THEN
      GOSUB TestProcessStep2
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 "
      Action = 1
      GOTO FormTestLoop
    END IF
    
    IF NOT Step3 THEN
      GOSUB TestProcessStep3
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 ..Done!"
      Form$(4, 0) = "Building Record Type 8 "
      Action = 1
      GOTO FormTestLoop
    END IF
    
    IF NOT Step4 THEN
      GOSUB TestProcessStep4
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 ..Done!"
      Form$(4, 0) = "Building Record Type 8 ..Done!"
      Form$(5, 0) = "Building Record Type 9"
      Action = 1
      GOTO FormTestLoop
    END IF
    
    IF NOT Step5 THEN
      GOSUB TestProcessStep5
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 ..Done!"
      Form$(4, 0) = "Building Record Type 8 ..Done!"
      Form$(5, 0) = "Building Record Type 9 ..Done!"
      Form$(6, 0) = "File Name Is: UBDFTEST.DAT"
      Action = 1
      GOTO FormTestLoop
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE ESC
      Done = True
    END SELECT
    
  LOOP UNTIL Done
  EXIT SUB
  
  RETURN
  
OpenDraftInfo:
  REDIM UBDraftRec(1) AS UBDraftRecType
  DraftFile = FREEFILE
  OPEN "UBSDRAFT.dat" FOR RANDOM ACCESS READ SHARED AS #DraftFile LEN = LEN(UBDraftRec(1))
  GET DraftFile, 1, UBDraftRec(1)
  RETURN
  
TestProcessStep1:
  GOSUB OpenDraftInfo
  
  REDIM UBDraftRecord1(1) AS UBDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord1(1))
  UBDraftRecord1(1).Field1 = "1"
  UBDraftRecord1(1).Field2 = "01"
  UBDraftRecord1(1).Field3 = " " + UBDraftRec(1).BANKDEST
  UBDraftRecord1(1).Field4 = " " + UBDraftRec(1).BANKORIG
  UBDraftRecord1(1).Field5 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord1(1).Field6 = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  UBDraftRecord1(1).Field7 = "A"
  UBDraftRecord1(1).Field8 = "094"
  UBDraftRecord1(1).Field9 = "10"
  UBDraftRecord1(1).Field10 = "1"
  UBDraftRecord1(1).Field11 = UCASE$(UBDraftRec(1).BANKNAME)
  UBDraftRecord1(1).Field12 = UCASE$(UBDraftRec(1).BANKLOC)
  UBDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  PUT DraftFileNum, 1, UBDraftRecord1(1)
  CLOSE DraftFileNum
  Step1 = True
  RETURN
  
TestProcessStep2:
  REDIM UBDraftRecord5(1) AS UBDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT5.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord5(1))
  UBDraftRecord5(1).Field1 = "5"
  UBDraftRecord5(1).Field2 = "200"
  UBDraftRecord5(1).Field3 = "COM PUBLIC WORKS"
  UBDraftRecord5(1).Field4 = "                    "
  UBDraftRecord5(1).Field5 = "0" + UBDraftRec(1).FEDID
  UBDraftRecord5(1).Field6 = "PPD"
  UBDraftRecord5(1).Field7 = "UTIL BILL"
  UBDraftRecord5(1).Field8 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord5(1).Field9 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  UBDraftRecord5(1).Field11 = "1"
  UBDraftRecord5(1).Field12 = LEFT$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord5(1).Field13 = "0000001"
  PUT DraftFileNum, 1, UBDraftRecord5(1)
  CLOSE DraftFileNum
  Step2 = True
  RETURN
  
TestProcessStep3:
  COUNTER = 0
  
  REDIM UBDraftRecord6(1) AS UBDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  CLOSE DraftFileNum
  KILL "UBDRAFT6.DAT"
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  
  'GO THRU DATA FILE HERE
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, Cnt, UBCustRec(1)
    AcctRecord = Cnt
    '  Process Customer Here
    IF UBCustRec(1).STATUS = "A" AND UBCustRec(1).USEDRAFT = "Y" THEN
      COUNTER = COUNTER + 1
      AcctNumber$ = STR$(AcctRecord)
      AcctNumber$ = RIGHT$(AcctNumber$, LEN(AcctNumber$) - 1)
      IF LEN(AcctNumber$) < 15 THEN
        AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
      END IF
      nme$ = UBCustRec(1).CUSTNAME
      IF LEN(nme$) < 22 THEN
        nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
      ELSE
        nme$ = LEFT$(nme$, 22)
      END IF
      BANKACCT$ = QPTrim$(UBCustRec(1).BANKACCT)
      IF LEN(BANKACCT$) < 17 THEN BANKACCT$ = BANKACCT$ + STRING$(17 - LEN(BANKACCT$), 32)
      Trace = Trace + 1
      Trace$ = STR$(Trace): Trace$ = RIGHT$(Trace$, LEN(Trace$) - 1)
      IF LEN(Trace$) < 7 THEN Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
      
      UBDraftRecord6(1).Field1 = "6"
      UBDraftRecord6(1).Field2 = "28"           ' Designates Prenote Trans
      UBDraftRecord6(1).Field3 = LEFT$(UBCustRec(1).TRANSIT, 8)
      UBDraftRecord6(1).Field4 = RIGHT$(UBCustRec(1).TRANSIT, 1)
      UBDraftRecord6(1).Field5 = LEFT$(BANKACCT$, 17)
      UBDraftRecord6(1).Field6 = "0000000000"   ' All zero's for Prenote
      UBDraftRecord6(1).Field7 = AcctNumber$
      UBDraftRecord6(1).Field8 = UCASE$(nme$)
      UBDraftRecord6(1).Field9 = "  "
      UBDraftRecord6(1).Field10 = "0"
      UBDraftRecord6(1).Field11 = LEFT$(UBCustRec(1).TRANSIT, 8) + Trace$
      PUT DraftFileNum, COUNTER, UBDraftRecord6(1)
      hash# = hash# + VAL(LEFT$(UBCustRec(1).TRANSIT, 8))
      UBCustRec(1).PreNoteFlag = 1
      number = number + 1
    END IF
  NEXT Cnt
  CLOSE DraftFileNum
  Step3 = True
  RETURN
  
TestProcessStep4:
  hash$ = STR$(hash#)
  hash$ = RIGHT$(hash$, LEN(hash$) - 1)
  
  IF LEN(hash$) < 10 THEN
    hash$ = STRING$(10 - LEN(hash$), "0") + hash$
  END IF
  IF LEN(hash$) > 10 THEN
    hash$ = RIGHT$(hash$, 10)
  END IF
  
  IF LEN(Trace$) > 6 THEN Trace$ = RIGHT$(Trace$, 6)
  
  REDIM UBDraftRecord8(1) AS UBDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT8.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord8(1))
  UBDraftRecord8(1).Field1 = "8"
  UBDraftRecord8(1).Field2 = "200"
  UBDraftRecord8(1).Field3 = Trace$
  UBDraftRecord8(1).Field4 = hash$
  UBDraftRecord8(1).Field5 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field7 = "0" + UBDraftRec(1).FEDID
  UBDraftRecord8(1).Field8 = STRING$(19, 32)    ' Reserved
  UBDraftRecord8(1).Field9 = STRING$(6, 32)     ' Reserved for Federal Reserve Use
  UBDraftRecord8(1).Field10 = LEFT$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord8(1).Field11 = "0000001"
  PUT DraftFileNum, 1, UBDraftRecord8(1)
  CLOSE DraftFileNum
  Step4 = True
  RETURN
  
TestProcessStep5:
  TotSize# = VAL(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records
  
  IF BlockSize! <> INT(BlockSize!) THEN
    BlockSize! = INT(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize - 1)))
  ELSE
    FillSize! = 0
  END IF
  
  BlockSize$ = STR$(BlockSize!)
  BlockSize$ = RIGHT$(BlockSize$, LEN(BlockSize$) - 1)
  IF LEN(BlockSize$) < 6 THEN BlockSize$ = STRING$(6 - LEN(BlockSize$), "0") + BlockSize$
  IF LEN(Trace$) < 8 THEN Trace$ = STRING$(8 - LEN(Trace$), "0") + Trace$
  
  REDIM UBDraftRecord9(1) AS UBDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT9.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord9(1))
  UBDraftRecord9(1).Field1 = "9"
  UBDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  UBDraftRecord9(1).Field3 = BlockSize$
  UBDraftRecord9(1).Field4 = Trace$
  UBDraftRecord9(1).Field5 = hash$
  UBDraftRecord9(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord9(1).Field7 = "000000000000"
  UBDraftRecord9(1).Field8 = STRING$(39, 32)    ' Reserved
  PUT DraftFileNum, 1, UBDraftRecord9(1)
  CLOSE DraftFileNum
  
  ' Now Put Them Together In File Name UBDFNOTE
  OutFile = FREEFILE
  OPEN "O", OutFile, "UBDFTEST.DAT": WIDTH #OutFile, 255
  
  REDIM UBDraftRecord1(1) AS UBDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord1(1))
  GET DraftFileNum, 1, UBDraftRecord1(1)
  PRINT #OutFile, UBDraftRecord1(1).Field1;
  PRINT #OutFile, UBDraftRecord1(1).Field2;
  PRINT #OutFile, UBDraftRecord1(1).Field3;
  PRINT #OutFile, UBDraftRecord1(1).Field4;
  PRINT #OutFile, UBDraftRecord1(1).Field5;
  PRINT #OutFile, UBDraftRecord1(1).Field6;
  PRINT #OutFile, UBDraftRecord1(1).Field7;
  PRINT #OutFile, UBDraftRecord1(1).Field8;
  PRINT #OutFile, UBDraftRecord1(1).Field9;
  PRINT #OutFile, UBDraftRecord1(1).Field10;
  PRINT #OutFile, UBDraftRecord1(1).Field11;
  PRINT #OutFile, UBDraftRecord1(1).Field12;
  PRINT #OutFile, UBDraftRecord1(1).Field13
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord5(1) AS UBDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT5.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord5(1))
  GET DraftFileNum, 1, UBDraftRecord5(1)
  PRINT #OutFile, UBDraftRecord5(1).Field1;
  PRINT #OutFile, UBDraftRecord5(1).Field2;
  PRINT #OutFile, UBDraftRecord5(1).Field3;
  PRINT #OutFile, UBDraftRecord5(1).Field4;
  PRINT #OutFile, UBDraftRecord5(1).Field5;
  PRINT #OutFile, UBDraftRecord5(1).Field6;
  PRINT #OutFile, UBDraftRecord5(1).Field7;
  PRINT #OutFile, UBDraftRecord5(1).Field8;
  PRINT #OutFile, UBDraftRecord5(1).Field9;
  PRINT #OutFile, UBDraftRecord5(1).Field10;
  PRINT #OutFile, UBDraftRecord5(1).Field11;
  PRINT #OutFile, UBDraftRecord5(1).Field12;
  PRINT #OutFile, UBDraftRecord5(1).Field13
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord6(1) AS UBDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94
  
  FOR Cnt = 1 TO NumOfLines
    GET DraftFileNum, Cnt, UBDraftRecord6(1)
    PRINT #OutFile, UBDraftRecord6(1).Field1;
    PRINT #OutFile, UBDraftRecord6(1).Field2;
    PRINT #OutFile, UBDraftRecord6(1).Field3;
    PRINT #OutFile, UBDraftRecord6(1).Field4;
    PRINT #OutFile, UBDraftRecord6(1).Field5;
    PRINT #OutFile, UBDraftRecord6(1).Field6;
    PRINT #OutFile, UBDraftRecord6(1).Field7;
    PRINT #OutFile, UBDraftRecord6(1).Field8;
    PRINT #OutFile, UBDraftRecord6(1).Field9;
    PRINT #OutFile, UBDraftRecord6(1).Field10;
    PRINT #OutFile, UBDraftRecord6(1).Field11
  NEXT Cnt
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord8(1) AS UBDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT8.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord8(1))
  GET DraftFileNum, 1, UBDraftRecord8(1)
  PRINT #OutFile, UBDraftRecord8(1).Field1;
  PRINT #OutFile, UBDraftRecord8(1).Field2;
  PRINT #OutFile, UBDraftRecord8(1).Field3;
  PRINT #OutFile, UBDraftRecord8(1).Field4;
  PRINT #OutFile, UBDraftRecord8(1).Field5;
  PRINT #OutFile, UBDraftRecord8(1).Field6;
  PRINT #OutFile, UBDraftRecord8(1).Field7;
  PRINT #OutFile, UBDraftRecord8(1).Field8;
  PRINT #OutFile, UBDraftRecord8(1).Field9;
  PRINT #OutFile, UBDraftRecord8(1).Field10;
  PRINT #OutFile, UBDraftRecord8(1).Field11
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord9(1) AS UBDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT9.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord9(1))
  GET DraftFileNum, 1, UBDraftRecord9(1)
  PRINT #OutFile, UBDraftRecord9(1).Field1;
  PRINT #OutFile, UBDraftRecord9(1).Field2;
  PRINT #OutFile, UBDraftRecord9(1).Field3;
  PRINT #OutFile, UBDraftRecord9(1).Field4;
  PRINT #OutFile, UBDraftRecord9(1).Field5;
  PRINT #OutFile, UBDraftRecord9(1).Field6;
  PRINT #OutFile, UBDraftRecord9(1).Field7;
  PRINT #OutFile, UBDraftRecord9(1).Field8
  CLOSE DraftFileNum
  FOR Cnt = 1 TO FillSize! / 94
    PRINT #OutFile, STRING$(94, "9")
  NEXT Cnt
  CLOSE
  Step5 = True
  RETURN
  
END SUB

SUB UBPrenote
  
  SHARED Choice$()
  LibName$ = "UB"
  ScrnName$ = "UBDFTPRE"
  
  CursorOff
  
  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB ProcessPrenote
      Done = True
    CASE ESC
      Done = True
    END SELECT
    
  LOOP UNTIL Done
  EXIT SUB
  
ProcessPrenote:
  
  LibName$ = "UB"
  ScrnName$ = "UBDFTPR1"
  
  CursorOff
  DisplayUBScrn ScrnName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  
  CursorOff
  
  Form$(1, 0) = "Building Record Type 1"
  Action = 1
  
  DO
FormLoop:
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    ' Process Record Type 1
    IF NOT Step1 THEN
      GOSUB ProcessStep1
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5"
      Action = 1
      GOTO FormLoop
    END IF
    
    IF NOT Step2 THEN
      GOSUB ProcessStep2
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 "
      Action = 1
      GOTO FormLoop
    END IF
    
    IF NOT Step3 THEN
      GOSUB ProcessStep3
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 ..Done!"
      Form$(4, 0) = "Building Record Type 8 "
      Action = 1
      GOTO FormLoop
    END IF
    
    IF NOT Step4 THEN
      GOSUB ProcessStep4
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 ..Done!"
      Form$(4, 0) = "Building Record Type 8 ..Done!"
      Form$(5, 0) = "Building Record Type 9"
      Action = 1
      GOTO FormLoop
    END IF
    
    IF NOT Step5 THEN
      GOSUB ProcessStep5
      Form$(1, 0) = "Building Record Type 1 ..Done!"
      Form$(2, 0) = "Building Record Type 5 ..Done!"
      Form$(3, 0) = "Building Record Type 6 ..Done!"
      Form$(4, 0) = "Building Record Type 8 ..Done!"
      Form$(5, 0) = "Building Record Type 9 ..Done!"
      Form$(6, 0) = "File Name Is: UBDFNOTE.DAT"
      Action = 1
      GOTO FormLoop
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE ESC
      Done = True
    END SELECT
    
  LOOP UNTIL Done
  EXIT SUB
  
  RETURN
  
  
OpenMainDraftInfo:
  REDIM UBDraftRec(1) AS UBDraftRecType
  DraftFile = FREEFILE
  OPEN "UBSDRAFT.dat" FOR RANDOM ACCESS READ SHARED AS #DraftFile LEN = LEN(UBDraftRec(1))
  GET DraftFile, 1, UBDraftRec(1)
  RETURN
  
ProcessStep1:
  GOSUB OpenMainDraftInfo
  
  REDIM UBDraftRecord1(1) AS UBDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord1(1))
  UBDraftRecord1(1).Field1 = "1"
  UBDraftRecord1(1).Field2 = "01"
  UBDraftRecord1(1).Field3 = " " + UBDraftRec(1).BANKDEST
  UBDraftRecord1(1).Field4 = " " + UBDraftRec(1).BANKORIG
  UBDraftRecord1(1).Field5 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord1(1).Field6 = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  UBDraftRecord1(1).Field7 = "A"
  UBDraftRecord1(1).Field8 = "094"
  UBDraftRecord1(1).Field9 = "10"
  UBDraftRecord1(1).Field10 = "1"
  UBDraftRecord1(1).Field11 = UCASE$(UBDraftRec(1).BANKNAME)
  UBDraftRecord1(1).Field12 = UCASE$(UBDraftRec(1).BANKLOC)
  UBDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  PUT DraftFileNum, 1, UBDraftRecord1(1)
  CLOSE DraftFileNum
  Step1 = True
  RETURN
  
ProcessStep2:
  REDIM UBDraftRecord5(1) AS UBDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT5.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord5(1))
  UBDraftRecord5(1).Field1 = "5"
  UBDraftRecord5(1).Field2 = "200"
  UBDraftRecord5(1).Field3 = "WADESBORO NC"
  UBDraftRecord5(1).Field4 = "                    "
  UBDraftRecord5(1).Field5 = "0" + UBDraftRec(1).FEDID
  UBDraftRecord5(1).Field6 = "PPD"
  UBDraftRecord5(1).Field7 = "UTIL BILL"
  UBDraftRecord5(1).Field8 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord5(1).Field9 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  UBDraftRecord5(1).Field11 = "1"
  UBDraftRecord5(1).Field12 = LEFT$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord5(1).Field13 = "0000001"
  PUT DraftFileNum, 1, UBDraftRecord5(1)
  CLOSE DraftFileNum
  Step2 = True
  RETURN
  
ProcessStep3:
  COUNTER = 0
  
  REDIM UBDraftRecord6(1) AS UBDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  CLOSE DraftFileNum
  KILL "UBDRAFT6.DAT"
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  
  'GO THRU DATA FILE HERE
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, Cnt, UBCustRec(1)
    AcctRecord = Cnt
    
    '  Process Customer Here
    IF UBCustRec(1).STATUS = "A" AND UBCustRec(1).USEDRAFT = "Y" AND UBCustRec(1).PreNoteFlag = 0 THEN
      IF LEFT$(UBCustRec(1).TRANSIT, 8) = "05318221" THEN UBCustRec(1).TRANSIT = "053108221": PUT UBCust, Cnt, UBCustRec(1)
      COUNTER = COUNTER + 1
      AcctNumber$ = STR$(AcctRecord)
      AcctNumber$ = RIGHT$(AcctNumber$, LEN(AcctNumber$) - 1)
      IF LEN(AcctNumber$) < 15 THEN
        AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
      END IF
      nme$ = UBCustRec(1).CUSTNAME
      IF LEN(nme$) < 22 THEN
        nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
      ELSE
        nme$ = LEFT$(nme$, 22)
      END IF
      BANKACCT$ = QPTrim$(UBCustRec(1).BANKACCT)
      IF LEN(BANKACCT$) < 17 THEN BANKACCT$ = BANKACCT$ + STRING$(17 - LEN(BANKACCT$), 32)
      Trace = Trace + 1
      Trace$ = STR$(Trace): Trace$ = RIGHT$(Trace$, LEN(Trace$) - 1)
      IF LEN(Trace$) < 7 THEN Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
      
      UBDraftRecord6(1).Field1 = "6"
      UBDraftRecord6(1).Field2 = "28"           ' Designates Prenote Trans
      UBDraftRecord6(1).Field3 = LEFT$(UBCustRec(1).TRANSIT, 8)
      UBDraftRecord6(1).Field4 = RIGHT$(UBCustRec(1).TRANSIT, 1)
      UBDraftRecord6(1).Field5 = LEFT$(BANKACCT$, 17)
      UBDraftRecord6(1).Field6 = "0000000000"   ' All zero's for Prenote
      UBDraftRecord6(1).Field7 = AcctNumber$
      UBDraftRecord6(1).Field8 = UCASE$(nme$)
      UBDraftRecord6(1).Field9 = "  "
      UBDraftRecord6(1).Field10 = "0"
      UBDraftRecord6(1).Field11 = LEFT$(UBDraftRec(1).BANKORIG, 8) + Trace$
      PUT DraftFileNum, COUNTER, UBDraftRecord6(1)
      hash# = hash# + VAL(LEFT$(UBCustRec(1).TRANSIT, 8))
      UBCustRec(1).PreNoteFlag = 1
      PUT UBCust, Cnt, UBCustRec(1)
      number = number + 1
    END IF
  NEXT Cnt
  CLOSE DraftFileNum
  Step3 = True
  RETURN
  
ProcessStep4:
  hash$ = STR$(hash#)
  hash$ = RIGHT$(hash$, LEN(hash$) - 1)
  
  IF LEN(hash$) < 10 THEN
    hash$ = STRING$(10 - LEN(hash$), "0") + hash$
  END IF
  IF LEN(hash$) > 10 THEN
    hash$ = RIGHT$(hash$, 10)
  END IF
  
  IF LEN(Trace$) > 6 THEN Trace$ = RIGHT$(Trace$, 6)
  
  REDIM UBDraftRecord8(1) AS UBDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT8.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord8(1))
  UBDraftRecord8(1).Field1 = "8"
  UBDraftRecord8(1).Field2 = "200"
  UBDraftRecord8(1).Field3 = Trace$
  UBDraftRecord8(1).Field4 = hash$
  UBDraftRecord8(1).Field5 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field7 = "0" + UBDraftRec(1).FEDID
  UBDraftRecord8(1).Field8 = STRING$(19, 32)    ' Reserved
  UBDraftRecord8(1).Field9 = STRING$(6, 32)     ' Reserved for Federal Reserve Use
  UBDraftRecord8(1).Field10 = LEFT$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord8(1).Field11 = "0000001"
  PUT DraftFileNum, 1, UBDraftRecord8(1)
  CLOSE DraftFileNum
  Step4 = True
  RETURN
  
ProcessStep5:
  TotSize# = VAL(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records
  
  IF BlockSize! <> INT(BlockSize!) THEN
    BlockSize! = INT(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize - 1)))
  ELSE
    FillSize! = 0
  END IF
  
  BlockSize$ = STR$(BlockSize!)
  BlockSize$ = RIGHT$(BlockSize$, LEN(BlockSize$) - 1)
  IF LEN(BlockSize$) < 6 THEN BlockSize$ = STRING$(6 - LEN(BlockSize$), "0") + BlockSize$
  IF LEN(Trace$) < 8 THEN Trace$ = STRING$(8 - LEN(Trace$), "0") + Trace$
  
  REDIM UBDraftRecord9(1) AS UBDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT9.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord9(1))
  UBDraftRecord9(1).Field1 = "9"
  UBDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  UBDraftRecord9(1).Field3 = BlockSize$
  UBDraftRecord9(1).Field4 = Trace$
  UBDraftRecord9(1).Field5 = hash$
  UBDraftRecord9(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord9(1).Field7 = "000000000000"
  UBDraftRecord9(1).Field8 = STRING$(39, 32)    ' Reserved
  PUT DraftFileNum, 1, UBDraftRecord9(1)
  CLOSE DraftFileNum
  
  ' Now Put Them Together In File Name UBDFNOTE
  OutFile = FREEFILE
  OPEN "O", OutFile, "UBDFNOTE.DAT": WIDTH #OutFile, 255
  
  REDIM UBDraftRecord1(1) AS UBDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord1(1))
  GET DraftFileNum, 1, UBDraftRecord1(1)
  PRINT #OutFile, UBDraftRecord1(1).Field1;
  PRINT #OutFile, UBDraftRecord1(1).Field2;
  PRINT #OutFile, UBDraftRecord1(1).Field3;
  PRINT #OutFile, UBDraftRecord1(1).Field4;
  PRINT #OutFile, UBDraftRecord1(1).Field5;
  PRINT #OutFile, UBDraftRecord1(1).Field6;
  PRINT #OutFile, UBDraftRecord1(1).Field7;
  PRINT #OutFile, UBDraftRecord1(1).Field8;
  PRINT #OutFile, UBDraftRecord1(1).Field9;
  PRINT #OutFile, UBDraftRecord1(1).Field10;
  PRINT #OutFile, UBDraftRecord1(1).Field11;
  PRINT #OutFile, UBDraftRecord1(1).Field12;
  PRINT #OutFile, UBDraftRecord1(1).Field13
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord5(1) AS UBDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT5.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord5(1))
  GET DraftFileNum, 1, UBDraftRecord5(1)
  PRINT #OutFile, UBDraftRecord5(1).Field1;
  PRINT #OutFile, UBDraftRecord5(1).Field2;
  PRINT #OutFile, UBDraftRecord5(1).Field3;
  PRINT #OutFile, UBDraftRecord5(1).Field4;
  PRINT #OutFile, UBDraftRecord5(1).Field5;
  PRINT #OutFile, UBDraftRecord5(1).Field6;
  PRINT #OutFile, UBDraftRecord5(1).Field7;
  PRINT #OutFile, UBDraftRecord5(1).Field8;
  PRINT #OutFile, UBDraftRecord5(1).Field9;
  PRINT #OutFile, UBDraftRecord5(1).Field10;
  PRINT #OutFile, UBDraftRecord5(1).Field11;
  PRINT #OutFile, UBDraftRecord5(1).Field12;
  PRINT #OutFile, UBDraftRecord5(1).Field13
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord6(1) AS UBDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT6.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94
  
  FOR Cnt = 1 TO NumOfLines
    GET DraftFileNum, Cnt, UBDraftRecord6(1)
    PRINT #OutFile, UBDraftRecord6(1).Field1;
    PRINT #OutFile, UBDraftRecord6(1).Field2;
    PRINT #OutFile, UBDraftRecord6(1).Field3;
    PRINT #OutFile, UBDraftRecord6(1).Field4;
    PRINT #OutFile, UBDraftRecord6(1).Field5;
    PRINT #OutFile, UBDraftRecord6(1).Field6;
    PRINT #OutFile, UBDraftRecord6(1).Field7;
    PRINT #OutFile, UBDraftRecord6(1).Field8;
    PRINT #OutFile, UBDraftRecord6(1).Field9;
    PRINT #OutFile, UBDraftRecord6(1).Field10;
    PRINT #OutFile, UBDraftRecord6(1).Field11
  NEXT Cnt
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord8(1) AS UBDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT8.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord8(1))
  GET DraftFileNum, 1, UBDraftRecord8(1)
  PRINT #OutFile, UBDraftRecord8(1).Field1;
  PRINT #OutFile, UBDraftRecord8(1).Field2;
  PRINT #OutFile, UBDraftRecord8(1).Field3;
  PRINT #OutFile, UBDraftRecord8(1).Field4;
  PRINT #OutFile, UBDraftRecord8(1).Field5;
  PRINT #OutFile, UBDraftRecord8(1).Field6;
  PRINT #OutFile, UBDraftRecord8(1).Field7;
  PRINT #OutFile, UBDraftRecord8(1).Field8;
  PRINT #OutFile, UBDraftRecord8(1).Field9;
  PRINT #OutFile, UBDraftRecord8(1).Field10;
  PRINT #OutFile, UBDraftRecord8(1).Field11
  CLOSE DraftFileNum
  
  REDIM UBDraftRecord9(1) AS UBDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT9.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord9(1))
  GET DraftFileNum, 1, UBDraftRecord9(1)
  PRINT #OutFile, UBDraftRecord9(1).Field1;
  PRINT #OutFile, UBDraftRecord9(1).Field2;
  PRINT #OutFile, UBDraftRecord9(1).Field3;
  PRINT #OutFile, UBDraftRecord9(1).Field4;
  PRINT #OutFile, UBDraftRecord9(1).Field5;
  PRINT #OutFile, UBDraftRecord9(1).Field6;
  PRINT #OutFile, UBDraftRecord9(1).Field7;
  PRINT #OutFile, UBDraftRecord9(1).Field8
  CLOSE DraftFileNum
  FOR Cnt = 1 TO FillSize! / 94
    PRINT #OutFile, STRING$(94, "9")
  NEXT Cnt
  CLOSE
  Step5 = True
  RETURN
  
  
END SUB

