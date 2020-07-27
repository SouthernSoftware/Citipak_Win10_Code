DEFINT A-Z
DECLARE SUB KillACHFiles ()
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB DisplayDFScrn (ScrnName$)
DECLARE SUB UBLog (Text2Log$)
DECLARE FUNCTION Check99File% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
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
DECLARE SUB QPrintRC (Text$, Row%, Col%, TheColor%)
DECLARE SUB WaitForAction ()
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION Exist% (FileName$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubpaymnt.BI'
  
  TYPE DraftRptType
    TRANSIT  AS STRING * 9
    BankName AS STRING * 14
    'BANKNAME AS STRING * 34
    CUSTACCT AS STRING * 5
    CustName AS STRING * 25
    BillAmt  AS STRING * 10
    BankAcct AS STRING * 20
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
  Col = ((80 - MaxLen) \ 2) - 1
  
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
  
  IF INSTR(COMMAND$, "TEST") < 1 THEN
    RUN "UBBILLIN"
  END IF

FUNCTION Check99File

  IF Exist%("UBPAY99.DAT") THEN
    UBLog "DRAFT: UBPAY99.DAT ALLREADY EXISTS!"
    BlockClear
    Ok = MsgBox%("UBSETUP", "KILL99FL")

    SELECT CASE Ok
    CASE 1
      Check99File = False
    CASE 2
      Check99File = True
      GOSUB RenameOld99File
    END SELECT
  ELSE
    Check99File = True
  END IF

EXIT FUNCTION

RenameOld99File:
  FOR ZZ = 1 TO 999
    Ext$ = "000" + QPTrim$(STR$(ZZ))
    Ext$ = RIGHT$(Ext$, 3)
    NewName$ = "UBPAY99." + Ext$
    IF NOT Exist%(NewName$) THEN
      UBLog "DRAFT: RENAMED UBPAY99.DAT TO " + "UBPAY99." + Ext$
      NAME "UBPAY99.DAT" AS NewName$
      UBLog "DRAFT: PAYMENT FILE RENAMED SUCCESSFULLY"
      EXIT FOR
    END IF
  NEXT
RETURN
END FUNCTION

SUB DisplayDFScrn (ScrnName$)
  LibFile2Scrn "UBDRAFT", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB KillACHFiles

  REDIM FileSpec$(2)
  FileSpec$(1) = "DS*."
  FileSpec$(2) = "*.ACH"

  FOR KCnt = 1 TO 2
    REDIM TempArray$(0)
    FileCount = 0

    FileName$ = DIR$(FileSpec$(KCnt))
    IF LEN(FileName$) = 0 THEN
      GOTO ExitKill
    ELSE
      FileCount = 1                    'It is, so count files.
      REDIM PRESERVE TempArray$(FileCount)
      TempArray$(FileCount) = FileName$
      DO
        FileName$ = DIR$
        IF LEN(FileName$) = 0 THEN
          EXIT DO
        ELSE
          FileCount = FileCount + 1
          REDIM PRESERVE TempArray$(FileCount)
          TempArray$(FileCount) = FileName$
        END IF
      LOOP
    END IF
    FOR Cnt = 1 TO FileCount
      KillFile TempArray$(Cnt)
    NEXT
ExitKill:
  NEXT
END SUB

SUB UBAcctsToDraft
  REDIM TempScrn(1)

  CycleCnt = 0

  Dash80$ = STRING$(80, "-")

  REDIM Cycle(1 TO 16) AS INTEGER
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM BankTotals(1 TO 1) AS BankTotalsType
  
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen          'load setup file
  TownName$ = UBSetUpRec(1).UTILNAME

  LibName$ = "UBDRAFT"
  ScrnName$ = "DGETCYCL"

'*********************************
  CursorOff
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  BlockClear
  DisplayDFScrn ScrnName$
  QPrintRC "Accounts to Draft ", 5, 21, 14
  QPrintRC "]", 5, 39, 10
  QPrintRC SPACE$(10), 7, 38, 0
  Fld(1).Protected = True

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE 13
      TCyc$ = QPTrim$(Form$(2, 0))
      IF LEN(TCyc$) > 0 THEN
        ThisCycle = Value#(Form$(2, 0), ECode)
        FOR Cnt = 1 TO CycleCnt
          IF ThisCycle = Cycle(Cnt) THEN
            GOTO NotDoneYet
          END IF
        NEXT
        CycleCnt = CycleCnt + 1
        Cycle(CycleCnt) = ThisCycle
        LSET Form$(CycleCnt + 2, 0) = QPTrim$(STR$(ThisCycle))
        Action = 1
      END IF
    CASE F10Key
      'GOSUB CheckInfo1
      'IF InfoOK THEN
        EXIT DO
      'END IF
    CASE ESC
      ExitFlag = True
    END SELECT
NotDoneYet:
  LOOP UNTIL ExitFlag

  IF ExitFlag THEN
    EXIT SUB
  END IF

'*********************************

  MaxLines = 58
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
    CustCycle = UBCustRec(1).BillCycl
    CustOk = False
    IF CycleCnt > 0 THEN
      FOR CCnt = 1 TO CycleCnt
        IF Cycle(CCnt) = 0 THEN
          CustOk = True
        ELSEIF CustCycle = Cycle(CCnt) THEN
          CustOk = True
          EXIT FOR
        END IF
      NEXT
    ELSE
      CustOk = True
    END IF

    IF CustOk THEN
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
    LSET DFTRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
    LSET DFTRec(1).BankAcct = QPTrim$(UBCustRec(1).BankAcct)
    LSET DFTRec(1).BillAmt = FUsing$(STR$(Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)), "#####.##")
    PRINT #UBRpt, DFTRec(1).TRANSIT; " "; DFTRec(1).BankName;
    PRINT #UBRpt, DFTRec(1).CUSTACCT; "  "; DFTRec(1).CustName;
    PRINT #UBRpt, " "; DFTRec(1).BillAmt; "  "; DFTRec(1).BankAcct
    
    LineCnt = LineCnt + 1
    IF AskAbandonPrint% THEN
      ABExit = True
      GOTO NON2PrintExit:
    END IF
    ShowPctComp Cnt, CstCnt
  NEXT
  
  PRINT #UBRpt, CHR$(12)
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Billing Bank Draft Register.                "; QPTrim$(TownName$)
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
  PRINT #UBRpt,
  PRINT #UBRpt, "Cycle List:"
  TabOffSet = 5
  FOR Cnt = 1 TO CycleCnt
    PRINT #UBRpt, TAB(TabOffSet); FUsing$(STR$(Cycle(Cnt)), "#####");
    TabOffSet = TabOffSet + 8
    IF TabOffSet > 70 THEN
      PRINT #UBRpt,
      TabOffSet = 5
    END IF
  NEXT
  PRINT #UBRpt,
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
  
CheckInfo1:
  IF CycleCnt > 0 THEN
    InfoOK = True
  ELSE
    SaveScrn TempScrn()
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Cycle Selection", 10, 28, -1
    QPrintRC "Press any Key to Continue.", 13, 27, -1
    WaitForAction
    RestScrn TempScrn()
    ERASE TempScrn
  END IF
RETURN

PrintBankDFTHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Billing Bank Draft Register.                "; QPTrim$(TownName$)
  PRINT #UBRpt, "Date: "; DATE$; TAB(72); "Page: "; PageNo
  PRINT #UBRpt, "Bank No.  Bank Name  Acct No.  Customer Name            Draft Amt    Bank Acct."
  
  PRINT #UBRpt, Dash80$
  LineCnt = 5
  
  RETURN

END SUB

SUB UBBuildTransmitFile
  
  UBLog " IN: Draft Build Transmit File"

  LibName$ = "UBDRAFT"
  'ScrnName$ = "UBDFTFIN"

  ScrnName$ = "DGETCYCL"

  REDIM TempScrn(0)

  '*****************
  REDIM Cycle(1 TO 16) AS INTEGER
  REDIM DFTRec(1) AS DraftRptType
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBDraftRec(1) AS UBDraftRecType
  DraftFile = FREEFILE
  OPEN "UBSDRAFT.dat" FOR RANDOM ACCESS READ SHARED AS #DraftFile LEN = LEN(UBDraftRec(1))
  GET DraftFile, 1, UBDraftRec(1)
  CLOSE

  CompanyAcct$ = QPTrim$(UBDraftRec(1).CompAcct)
  IF LEN(CompanyAcct$) > 0 THEN
    BalRecFlag = True
  END IF

'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  TownName$ = UCASE$(UBSetUpRec(1).UTILNAME)

  IF INSTR(TownName$, "WARRENTON") > 0 THEN
    WarrFlag = True
  END IF

  IF INSTR(TownName$, "NORWOOD") > 0 THEN
    NWoodFlag = True
  END IF

  IF INSTR(TownName$, "PLYMOUTH") > 0 THEN
    PlyFlag = True
  END IF
  
  REDIM PayRec(1) AS UBPaymentRecType
  PayRecLen = LEN(PayRec(1))
  PayFileName$ = "UBPAY99.DAT"

  IF UBSetUpRec(1).Make99File = "Y" THEN
    IF NOT Check99File THEN
      GOTO DraftExit
    END IF
    Make99Flag = True
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  NumOfRecs& = LOF(UBCust) \ UBCustRecLen
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  'CursorOff
  BlockClear
  DisplayDFScrn ScrnName$
  QPrintRC "Draft Transmission File ", 5, 21, 14
           
  QPrintRC "]", 5, 45, 10
  QPrintRC "Draft Date:", 7, 26, 3

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode

    CASE 13
      TCyc$ = QPTrim$(Form$(2, 0))
      IF LEN(TCyc$) > 0 THEN
        ThisCycle = Value#(Form$(2, 0), ECode)
        FOR Cnt = 1 TO CycleCnt
          IF ThisCycle = Cycle(Cnt) THEN
            GOTO NotDoneHere
          END IF
        NEXT
        CycleCnt = CycleCnt + 1
        Cycle(CycleCnt) = ThisCycle
        LSET Form$(CycleCnt + 2, 0) = QPTrim$(STR$(ThisCycle))
        Action = 1
      END IF
    CASE F10Key
      BDate = Date2Num(Form$(1, 0))
      IF BDate > 0 THEN
        GOSUB ProcessDraft
        Done = True
      ELSE
        SaveScrn TempScrn()
        DisplayUBScrn "ERRSCRN1"
        QPrintRC "Invalid Draft Date!", 10, 30, -1
        QPrintRC "Please correct and try again.", 13, 26, -1
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 1
      END IF
    CASE ESC
      Done = True
    END SELECT
NotDoneHere:
  LOOP UNTIL Done

  EXIT SUB
  
ProcessDraft:

  BlockClear
  KillACHFiles

  DraftDate$ = Form$(1, 0)
  DraftDate$ = LEFT$(DraftDate$, 2) + MID$(DraftDate$, 4, 2) + RIGHT$(DraftDate$, 2)

  IF VAL(DraftDate$) = 0 THEN EXIT SUB

  NewDraftFile$ = QPTrim$(UBDraftRec(1).FileName)

  IF LEN(NewDraftFile$) = 0 THEN
    NewDraftFile$ = "DS" + DraftDate$ + ".ACH"
  END IF


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
      IF Make99Flag THEN
        LSET Form$(6, 0) = "Building Payment File  ..Done!"
        QPrintRC "7. File Name Is: " + NewDraftFile$, 16, 25, 79
      ELSE
        LSET Form$(6, 0) = "File Name Is: " + NewDraftFile$
      END IF
      Action = 1
      GOTO DraftFormLoop
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE ESC
      Done = True
    END SELECT
    
  LOOP UNTIL Done
  
DraftExit:
  CLOSE

  KillFile "UBDRAFT1.dat"
  KillFile "UBDRAFT5.dat"
  KillFile "UBDRAFT6.dat"
  KillFile "UBDRAFT9.dat"
  KillFile "UBDRAFT8.dat"


EXIT SUB
  
'RETURN
  
  
DraftProcessStep1:

  FedIDNum$ = QPTrim$(UBDraftRec(1).FEDPREFX + UBDraftRec(1).FEDID) + "00"

  REDIM UBDraftRecord1(1) AS UBDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN "UBDRAFT1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #DraftFileNum LEN = LEN(UBDraftRecord1(1))
  UBDraftRecord1(1).Field1 = "1"
  UBDraftRecord1(1).Field2 = "01"
  UBDraftRecord1(1).Field3 = " " + QPTrim$(UBDraftRec(1).BANKDEST)
  UBDraftRecord1(1).Field4 = " " + QPTrim$(UBDraftRec(1).BANKORIG)
  UBDraftRecord1(1).Field5 = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  UBDraftRecord1(1).Field6 = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  UBDraftRecord1(1).Field7 = "A"
  UBDraftRecord1(1).Field8 = "094"
  UBDraftRecord1(1).Field9 = "10"
  UBDraftRecord1(1).Field10 = "1"
  UBDraftRecord1(1).Field11 = QPTrim$(UCASE$(UBDraftRec(1).BankName))
  UBDraftRecord1(1).Field12 = QPTrim$(UCASE$(UBDraftRec(1).BANKLOC))
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
  UBDraftRecord5(1).Field3 = LEFT$(TownName$, 16)
  UBDraftRecord5(1).Field4 = "                    "

  UBDraftRecord5(1).Field5 = FedIDNum$
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
  IF Make99Flag THEN
    GOSUB Setup99File
  END IF

  FOR Cnt& = 1 TO NumOfRecs&
    GET UBCust, Cnt&, UBCustRec(1)
    CustCycle = UBCustRec(1).BillCycl
    CustOk = False
    IF CycleCnt > 0 THEN
      FOR CCnt = 1 TO CycleCnt
        IF Cycle(CCnt) = 0 THEN
          CustOk = True
        ELSEIF CustCycle = Cycle(CCnt) THEN
          CustOk = True
          EXIT FOR
        END IF
      NEXT
    ELSE
      CustOk = True
    END IF
    
    IF CustOk THEN
      IF UBCustRec(1).STATUS = "A" OR UBCustRec(1).STATUS = "B" THEN
        IF (UBCustRec(1).USEDRAFT = "Y") OR (LEN(QPTrim$(UBCustRec(1).BankName)) > 0) THEN
          CustBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          IF CustBal# > 0 THEN
            ' Process Customer Here
            BillAmt$ = (STR$(CustBal# * 100))     'Remove the Decimals
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
            nme$ = UBCustRec(1).CustName
            IF LEN(nme$) < 22 THEN
              nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
            ELSE
              nme$ = LEFT$(nme$, 22)
            END IF
          
            'Check for Spaces WithIn Bank Account Numbered as Entered by Customer
            BankAcct$ = QPTrim$(UBCustRec(1).BankAcct)
            SP = INSTR(BankAcct$, " ")
            IF SP > 0 THEN
              BankAcct$ = LEFT$(BankAcct$, SP - 1) + RIGHT$(BankAcct$, LEN(BankAcct$) - SP)
            END IF
          
            IF LEN(BankAcct$) < 17 THEN
              BankAcct$ = BankAcct$ + STRING$(17 - LEN(BankAcct$), 32)
            END IF
            Trace = Trace + 1
            Trace$ = STR$(Trace)
            Trace$ = RIGHT$(Trace$, LEN(Trace$) - 1)

            IF LEN(Trace$) < 7 THEN
              Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
            END IF
            UBDraftRecord6(1).Field1 = "6"
            UBDraftRecord6(1).Field2 = "27"       ' Designates Debit Trans
            UBDraftRecord6(1).Field3 = LEFT$(UBCustRec(1).TRANSIT, 8)
            UBDraftRecord6(1).Field4 = RIGHT$(UBCustRec(1).TRANSIT, 1)
            UBDraftRecord6(1).Field5 = LEFT$(BankAcct$, 17)
            UBDraftRecord6(1).Field6 = BillAmt$   ' All zero's for Prenote
            UBDraftRecord6(1).Field7 = AcctNumber$
            UBDraftRecord6(1).Field8 = UCASE$(nme$)
            UBDraftRecord6(1).Field9 = "  "
            UBDraftRecord6(1).Field10 = "0"
            UBDraftRecord6(1).Field11 = LEFT$(UBDraftRec(1).BANKORIG, 8) + Trace$
            PUT DraftFileNum, COUNTER, UBDraftRecord6(1)
            hash# = hash# + VAL(LEFT$(UBCustRec(1).TRANSIT, 8))
            number = number + 1
            IF Make99Flag THEN
              GOSUB Make99Rec
            END IF
          END IF
        END IF
      END IF
    END IF
  NEXT

  IF BalRecFlag THEN
    ' Process Customer Here
    BillAmt$ = STR$(TotalAmount#)      'Remove the Decimals
    BillAmt$ = LTRIM$(BillAmt$)
    IF LEN(BillAmt$) < 10 THEN BillAmt$ = STRING$(10 - LEN(BillAmt$), "0") + BillAmt$
    'Get Running Billed Amount
    COUNTER = COUNTER + 1

    AcctNumber$ = "0"   'STR$(Cnt&)
    AcctNumber$ = RIGHT$(AcctNumber$, LEN(AcctNumber$) - 1)
    IF LEN(AcctNumber$) < 15 THEN
      AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
    END IF
    nme$ = TownName$
    IF LEN(nme$) < 22 THEN
      nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
    ELSE
      nme$ = LEFT$(nme$, 22)
    END IF

    BankAcct$ = CompanyAcct$
    SP = INSTR(BankAcct$, " ")
    IF SP > 0 THEN
      BankAcct$ = LEFT$(BankAcct$, SP - 1) + RIGHT$(BankAcct$, LEN(BankAcct$) - SP)
    END IF

    IF LEN(BankAcct$) < 17 THEN
      BankAcct$ = BankAcct$ + STRING$(17 - LEN(BankAcct$), 32)
    END IF
    Trace = Trace + 1
    Trace$ = STR$(Trace)
    Trace$ = RIGHT$(Trace$, LEN(Trace$) - 1)

    IF LEN(Trace$) < 7 THEN
      Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
    END IF

    UBDraftRecord6(1).Field1 = "6"
    UBDraftRecord6(1).Field2 = "22"       ' Designates credit checking
    UBDraftRecord6(1).Field3 = LEFT$(UBDraftRec(1).BANKDEST, 8)
    UBDraftRecord6(1).Field4 = RIGHT$(UBDraftRec(1).BANKDEST, 1)
    UBDraftRecord6(1).Field5 = LEFT$(BankAcct$, 17)
    UBDraftRecord6(1).Field6 = BillAmt$   ' All zero's for Prenote
    UBDraftRecord6(1).Field7 = AcctNumber$
    UBDraftRecord6(1).Field8 = UCASE$(nme$)
    UBDraftRecord6(1).Field9 = "  "
    UBDraftRecord6(1).Field10 = "0"
    UBDraftRecord6(1).Field11 = LEFT$(UBDraftRec(1).BANKORIG, 8) + Trace$
    
    PUT DraftFileNum, COUNTER, UBDraftRecord6(1)
    hash# = hash# + VAL(LEFT$(UBDraftRec(1).BANKDEST, 8))
    number = number + 1
  END IF

  CLOSE DraftFileNum

  IF Make99Flag THEN
    CLOSE Pay99File
  END IF
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

  IF BalRecFlag THEN
    UBDraftRecord8(1).Field6 = TotalAmount$       ' zero for prenote
  ELSE
    UBDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  END IF

  UBDraftRecord8(1).Field7 = FedIDNum$
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
    FillSize! = 940 - (TotSize# - (940 * (BlockSize! - 1)))
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

  IF BalRecFlag THEN
    UBDraftRecord9(1).Field7 = TotalAmount$       ' zero for prenote
  ELSE
    UBDraftRecord9(1).Field7 = "000000000000"     ' zero for prenote
  END IF

  UBDraftRecord9(1).Field8 = STRING$(39, 32)    ' Reserved

  PUT DraftFileNum, 1, UBDraftRecord9(1)
  CLOSE DraftFileNum
  
  ' Now Put Them Together In File Name UBDFNOTE
  OutFile = FREEFILE
  OPEN NewDraftFile$ FOR OUTPUT AS OutFile

  WIDTH #OutFile, 255
  
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
  PRINT #OutFile, UBDraftRecord1(1).Field13;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF
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
  PRINT #OutFile, UBDraftRecord5(1).Field13;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF
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
    PRINT #OutFile, UBDraftRecord6(1).Field11;
    IF NWoodFlag = 0 THEN
      PRINT #OutFile,
    END IF
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
  PRINT #OutFile, UBDraftRecord8(1).Field11;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF
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
  PRINT #OutFile, UBDraftRecord9(1).Field8;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF
  CLOSE DraftFileNum
  IF NWoodFlag = 0 THEN
    FOR Cnt = 1 TO FillSize! / 94
      PRINT #OutFile, STRING$(94, "9")
    NEXT Cnt
  END IF
  CLOSE
  Step5 = True
RETURN
  
Make99Rec:

  REDIM PayRec(1) AS UBPaymentRecType

  PayRec(1).OPERNUM = 99
  PayRec(1).PAYDATE = BDate
  PayRec(1).CUSTACCT = Cnt&
  PayRec(1).CustName = UBCustRec(1).CustName
  PayRec(1).CUSTADDR = UBCustRec(1).Addr1
  'PayRec(1).CUSTCMNT= UBBillRec(1).
  PayRec(1).AMTOWED = CustBal#
  PayRec(1).TENDERTY = "BANK DRAFT"
  PayRec(1).CASHAMT = PayRec(1).AMTOWED  'was setting 0
  PayRec(1).CHKAMT = 0
  PayRec(1).AMTRECD = PayRec(1).AMTOWED
  PayRec(1).CHANGE = 0
  PayRec(1).DESC = "DRAFT PAYMENT TRANS"

  FOR ZZCnt = 1 TO 15
    PayRec(1).PaidOwed(ZZCnt).AMTOWE1 = UBCustRec(1).CurrRevAmts(ZZCnt)
    PayRec(1).PaidOwed(ZZCnt).AMTPD1 = PayRec(1).PaidOwed(ZZCnt).AMTOWE1
  NEXT

  PayRec(1).TOTOWED = PayRec(1).AMTOWED
  PayRec(1).AMTPAID = PayRec(1).AMTOWED

  PUT #Pay99File, , PayRec(1)
  
RETURN

Setup99File:
  
  Pay99File = FREEFILE
  OPEN PayFileName$ FOR OUTPUT AS #Pay99File
  CLOSE Pay99File

  Pay99File = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS Pay99File LEN = PayRecLen

RETURN

END SUB

SUB UBDraftListing
  
  TotalDraftCustomers = 0
  Dash80$ = STRING$(80, "-")
  Temp1$ = SPACE$(10)
  Temp2$ = SPACE$(12)
  MaxLines = 60
  FF$ = CHR$(12)

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
    TitleBox 2, Col, MaxLen + 3, "Draft Customer Report ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
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
    IF UsingAcct THEN
      GET UBCust, Cnt, UBCustRec(1)
      AcctRecord = Cnt
    ELSE
      GET UBCust, IndexArray(Cnt).RecNum, UBCustRec(1)
      AcctRecord = IndexArray(Cnt).RecNum
    END IF
    '  Process Customer Here
    IF UBCustRec(1).STATUS = "A" THEN
      IF LineCnt >= MaxLines THEN
        PRINT #UBRpt, FF$;
        GOSUB PrintPreHeader
      END IF
      IF UBCustRec(1).USEDRAFT = "Y" THEN
        TotalDraftCustomers = TotalDraftCustomers + 1
        PRINT #UBRpt, USING "#####"; AcctRecord;
        PRINT #UBRpt, TAB(10); UBCustRec(1).CustName;
        PRINT #UBRpt, TAB(48); LEFT$(UBCustRec(1).TRANSIT, 9);
        PRINT #UBRpt, TAB(59); LEFT$(UBCustRec(1).BankAcct, 20)
        IF UBCustRec(1).PreNoteFlag = 0 THEN
          PRINT #UBRpt, TAB(2); "N";
        ELSE
          PRINT #UBRpt, TAB(2); "Y";
        END IF
        PRINT #UBRpt, TAB(15); QPTrim$(UBCustRec(1).BankName);
        PRINT #UBRpt, TAB(50); LEFT$(UBCustRec(1).BANKLOC, 20); "  "; UBCustRec(1).BillCycl
        PRINT #UBRpt, STRING$(79, "-")
        LineCnt = LineCnt + 3
      END IF
    END IF

    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT Cnt
  PRINT #UBRpt, "Total Draft Customers on File: "; USING "####,#"; TotalDraftCustomers
  PRINT #UBRpt, FF$;
  
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
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TownName$ = UBSetUpRec(1).UTILNAME

  IF INSTR(TownName$, "WARRENTON") > 0 THEN
    TestFileName$ = "UBDFTEST"
    WarrFlag = True
  ELSE
    TestFileName$ = "UBDFTEST.DAT"
  END IF
  
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

  BlockClear
  DisplayUBScrn ScrnName$
  
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
      Form$(6, 0) = "File Name Is: " + TestFileName$
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

  FedIDNum$ = QPTrim$(UBDraftRec(1).FEDPREFX + UBDraftRec(1).FEDID) + "00"

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
  UBDraftRecord1(1).Field11 = QPTrim$(UCASE$(UBDraftRec(1).BankName))
  UBDraftRecord1(1).Field12 = QPTrim$(UCASE$(UBDraftRec(1).BANKLOC))
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
  UBDraftRecord5(1).Field3 = "UTILITY DRAFT TEST"
  UBDraftRecord5(1).Field4 = "                    "
  UBDraftRecord5(1).Field5 = FedIDNum$
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
      nme$ = UBCustRec(1).CustName
      IF LEN(nme$) < 22 THEN
        nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
      ELSE
        nme$ = LEFT$(nme$, 22)
      END IF
      'Check for Spaces WithIn Bank Account Numbered as Entered by Customer
      BankAcct$ = QPTrim$(UBCustRec(1).BankAcct)
      SP = INSTR(BankAcct$, " ")
      IF SP > 0 THEN
        BankAcct$ = LEFT$(BankAcct$, SP - 1) + RIGHT$(BankAcct$, LEN(BankAcct$) - SP)
      END IF
      IF LEN(BankAcct$) < 17 THEN BankAcct$ = BankAcct$ + STRING$(17 - LEN(BankAcct$), 32)
      Trace = Trace + 1
      Trace$ = STR$(Trace): Trace$ = RIGHT$(Trace$, LEN(Trace$) - 1)
      IF LEN(Trace$) < 7 THEN Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
      
      UBDraftRecord6(1).Field1 = "6"
      UBDraftRecord6(1).Field2 = "28"           ' Designates Prenote Trans
      UBDraftRecord6(1).Field3 = LEFT$(UBCustRec(1).TRANSIT, 8)
      UBDraftRecord6(1).Field4 = RIGHT$(UBCustRec(1).TRANSIT, 1)
      UBDraftRecord6(1).Field5 = LEFT$(BankAcct$, 17)
      UBDraftRecord6(1).Field6 = "0000000000"   ' All zero's for Prenote
      UBDraftRecord6(1).Field7 = AcctNumber$
      UBDraftRecord6(1).Field8 = UCASE$(nme$)
      UBDraftRecord6(1).Field9 = "  "
      UBDraftRecord6(1).Field10 = "0"
      UBDraftRecord6(1).Field11 = LEFT$(UBDraftRec(1).BANKORIG, 8) + Trace$
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
  UBDraftRecord8(1).Field7 = FedIDNum$
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
    FillSize! = 940 - (TotSize# - (940 * (BlockSize! - 1)))
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

  OPEN "O", OutFile, TestFileName$: WIDTH #OutFile, 255
  
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
  
  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TownName$ = UCASE$(UBSetUpRec(1).UTILNAME)
  
  'PreNoteFile$ = QPTrim$(UBDraftRec(1).FileName)
  
  IF INSTR(TownName$, "WARRENTON") > 0 THEN
    PreNoteFile$ = "UBDFNOTE"
    WarrFlag = True
  ELSEIF LEN(PreNoteFile$) = 0 THEN
    PreNoteFile$ = "UBDFNOTE.DAT"
  END IF

  IF INSTR(TownName$, "NORWOOD") > 0 THEN
    NWoodFlag = True
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBDFTPRE"
  
  CursorOff
  
  
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

  BlockClear
  DisplayUBScrn ScrnName$
  
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
      Form$(6, 0) = "File Name Is: " + PreNoteFile$
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
  FedIDNum$ = QPTrim$(UBDraftRec(1).FEDPREFX + UBDraftRec(1).FEDID) + "00"

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
  UBDraftRecord1(1).Field11 = QPTrim$(UCASE$(UBDraftRec(1).BankName))
  UBDraftRecord1(1).Field12 = QPTrim$(UCASE$(UBDraftRec(1).BANKLOC))
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
  UBDraftRecord5(1).Field3 = LEFT$(TownName$, 16)
  UBDraftRecord5(1).Field4 = "                    "
  UBDraftRecord5(1).Field5 = FedIDNum$
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
      nme$ = UBCustRec(1).CustName
      IF LEN(nme$) < 22 THEN
        nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
      ELSE
        nme$ = LEFT$(nme$, 22)
      END IF
      
      'Check for Spaces WithIn Bank Account Numbered as Entered by Customer
      BankAcct$ = QPTrim$(UBCustRec(1).BankAcct)
      SP = INSTR(BankAcct$, " ")
      IF SP > 0 THEN
        BankAcct$ = LEFT$(BankAcct$, SP - 1) + RIGHT$(BankAcct$, LEN(BankAcct$) - SP)
      END IF
      
      IF LEN(BankAcct$) < 17 THEN BankAcct$ = BankAcct$ + STRING$(17 - LEN(BankAcct$), 32)
      Trace = Trace + 1
      Trace$ = STR$(Trace): Trace$ = RIGHT$(Trace$, LEN(Trace$) - 1)
      IF LEN(Trace$) < 7 THEN Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
      
      UBDraftRecord6(1).Field1 = "6"
      UBDraftRecord6(1).Field2 = "28"           ' Designates Prenote Trans
      UBDraftRecord6(1).Field3 = LEFT$(UBCustRec(1).TRANSIT, 8)
      UBDraftRecord6(1).Field4 = RIGHT$(UBCustRec(1).TRANSIT, 1)
      UBDraftRecord6(1).Field5 = LEFT$(BankAcct$, 17)
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
  UBDraftRecord8(1).Field7 = FedIDNum$
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

  OPEN "O", OutFile, PreNoteFile$: WIDTH #OutFile, 255
  
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
  PRINT #OutFile, UBDraftRecord1(1).Field13;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF
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
  PRINT #OutFile, UBDraftRecord5(1).Field13;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF

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
    PRINT #OutFile, UBDraftRecord6(1).Field11;
    IF NWoodFlag = 0 THEN
      PRINT #OutFile,
    END IF
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
  PRINT #OutFile, UBDraftRecord8(1).Field11;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF

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
  PRINT #OutFile, UBDraftRecord9(1).Field8;
  IF NWoodFlag = 0 THEN
    PRINT #OutFile,
  END IF
  CLOSE DraftFileNum

  IF NWoodFlag = 0 THEN
    IF FillSize! < 0 THEN
      FillSize! = 940 - ABS(FillSize!)
    END IF
    FOR Cnt = 1 TO ABS(FillSize!) / 94
      PRINT #OutFile, STRING$(94, "9")
    NEXT Cnt
  END IF

  CLOSE
  Step5 = True
  RETURN
  
  
END SUB

