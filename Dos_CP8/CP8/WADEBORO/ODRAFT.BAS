  DEFINT A-Z
DECLARE SUB UBDraftTest ()
DECLARE SUB UBAcctsToDraft ()
DECLARE SUB UBDraftEdit ()
DECLARE SUB UBBuildTransmitFile ()
DECLARE SUB UBPrenote ()
DECLARE SUB UBDraftListing ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB UBBillMenu ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
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
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
  
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
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'ubcust.bi'
  '$INCLUDE: 'ubloc.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubsetbal.BI'
  
  CONST False = 0, True = NOT False
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  MChoice$(1) = "Accounts To Draft Report"
  MChoice$(2) = "Edit Draft Record"
  MChoice$(3) = "Prepare Draft Transmission File"
  MChoice$(4) = "Prepare Draft Prenote File"
  MChoice$(5) = "Print Draft Customer Listing"
  MChoice$(6) = "Prepare Draft Test File for Bank"
  MChoice$(7) = "Exit to OS "
  
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
    
    TitleBox 3, Col, MaxLen + 3, "ACH - Draft Processing System ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
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
      UBDraftEdit
    CASE 3
      UBBuildTransmitFile
    CASE 4
      UBPrenote
    CASE 5
      UBDraftListing
    CASE 6
      UBDraftTest
    CASE 7
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag
  
  RUN "UBbillin"
  

SUB UBAcctsToDraft
  
  
  
END SUB

SUB UBBuildTransmitFile
END SUB

SUB UBDraftEdit
END SUB

SUB UBDraftListing
  
  TotalDraftCustomers = 0
  Dash80$ = STRING$(80, "-")
  Temp1$ = SPACE$(10)
  Temp2$ = SPACE$(12)
  
  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
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
  
  REDIM UBCustRec(1) AS UBCustRecType
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
        PRINT #UBRpt, CHR$(12);
        GOSUB PrintPreHeader
      END IF
      IF UBCustRec(1).DRAFTYPE = "Y" THEN
        TotalDraftCustomers = TotalDraftCustomers + 1
        PRINT #UBRpt, USING "#####"; AcctRecord;
        PRINT #UBRpt, TAB(10); UBCustRec(1).CustName;
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
  
  IF NOT ExitFlag THEN GOTO ReStart
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
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, Cnt, UBCustRec(1)
    AcctRecord = Cnt
    '  Process Customer Here
    IF UBCustRec(1).STATUS = "A" AND UBCustRec(1).DRAFTYPE = "Y" THEN
      COUNTER = COUNTER + 1
      AcctNumber$ = STR$(AcctRecord)
      AcctNumber$ = RIGHT$(AcctNumber$, LEN(AcctNumber$) - 1)
      IF LEN(AcctNumber$) < 15 THEN
        AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
      END IF
      nme$ = UBCustRec(1).CustName
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
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, Cnt, UBCustRec(1)
    AcctRecord = Cnt
    
    '  Process Customer Here
    IF UBCustRec(1).STATUS = "A" AND UBCustRec(1).DRAFTYPE = "Y" AND UBCustRec(1).PreNoteFlag = 0 THEN
      COUNTER = COUNTER + 1
      AcctNumber$ = STR$(AcctRecord)
      AcctNumber$ = RIGHT$(AcctNumber$, LEN(AcctNumber$) - 1)
      IF LEN(AcctNumber$) < 15 THEN
        AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
      END IF
      nme$ = UBCustRec(1).CustName
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

