DEFINT A-Z
DECLARE SUB IndexVendorFile ()
DECLARE SUB NewVendorRec (NewVRec)
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB WaitForAction
DECLARE SUB Post2PO (FileName$, BadTrans%)
DECLARE FUNCTION EdCheck4PO% (Form$(), AcctFlds(), Fld() AS ANY)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION EDCheck4Vendor% (Form$(), Terms)
DECLARE FUNCTION oEDCheck4Acct% (Form$(), AcctFldNo%)
DECLARE FUNCTION FindVendorRec% (VendorCode$)
DECLARE SUB PostInvTrans (APType%)
DECLARE SUB OpenAPVendorFile (APVendorFile%, NumVRecs%, VendorRecLen%)
DECLARE SUB OpenVendorIdx (VendorIdxFile%, NumActiveVendors%)
DECLARE FUNCTION ChkForVendor% (ChkVendor$)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB Prn2Scrn2 (TitleText$(), FileName$, TopRow%, BotRow%, Clr%, PrnFlag%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS Config)
DECLARE SUB CancelPO (VoidTransRecNum&, APLedgerFile, NumTrans, APDistFile, NumDistRecs&)
DECLARE SUB GetFYDates (FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate)

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'pageinfo.BI'                  'Form Page info
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'RPTPRN.BI'                    'Report printer routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'SCRAUX.bi'
'$INCLUDE: 'ap.bi'

DECLARE SUB EdEntry (JrnlCode%)
DECLARE SUB OpenAPEditFile (APEditFile%, NumEdTrans%, APType)
DECLARE SUB NewCJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList (APType)
DECLARE SUB DisplayTotals ()
'DECLARE SUB PostTrans (APType)
DECLARE SUB Check4BlankFlds (Form$(), BadFld%, FY1BegDate, FY2EndDate)
DECLARE FUNCTION ListRecs% (APType)
DECLARE SUB ResetForm (Form$(), NumFlds%)
DECLARE SUB ShowEdMode (EdMode%)
DECLARE SUB Rec2Form (Form$(), RecordNumber%, Fld() AS ANY, APType)
DECLARE SUB SaveTrans (Form$(), RecordNumber%, EdMode%, APType)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
'DECLARE FUNCTION FldNum% (FldName$, Fld() AS FieldInfo)

  CONST False = 0, True = NOT False, NumAcctFlds = 24
  
  DIM SHARED TotLen, FundLen, AcctLen, DetLen
  DIM SHARED User$, CashAcct$, APAcct$, EncAcct$
  DIM SHARED AcctIdx AS GLAcctIndexType
  DIM SHARED acct AS GLAcctRecType
  DIM SHARED APEdit AS APInv81Type
  DIM SHARED VendorIdx AS VendorIdxRecType

  TYPE FLen
    V AS STRING * 42
  END TYPE

  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'init.BI'

  TextCursor -2, -2             'Define the Mouse cursor

  '---Master Menu
  DIM MenuText$(1 TO 4)
  MenuText$(1) = "Enter/Edit Invoices for Payment"
  MenuText$(2) = "Print Check Register"
  MenuText$(3) = "Print Checks"
  MenuText$(4) = "Quit"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box to line 20
  Action = 0    '0 means stay in the menu
  Choice = 1    'Pre-load choice to highlight

  FOR Cnt = 1 TO UBOUND(MenuText$)              'Find max menu width
    TLen = LEN(MenuText$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT

  'Center Menu within Screen
  Row = ((25 - (UBOUND(MenuText$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2

  'ShowCursor
  JrnlType = 2

  DO
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode

    Title$ = "Check Printing"
    MnuTitle Title$, Row, Col, MaxLen, 15

    PrintTitle User$
    PrintHelp "Invoice Processing"

    LOCATE Row, Col, 0
    VertMenu MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO

    SELECT CASE Choice
      CASE 1
        ClrBG
        EdEntry JrnlType          ', AcctFlds()
      CASE 2
        PrnEditList JrnlType
      CASE 3
       ' PostInvTrans JrnlType
      CASE 4
        HideCursor
        CLS : END
      CASE ELSE
    END SELECT

  LOOP

  'HideCursor
  '  DispLoad

  'RUN "glmenu"

SUB Check4BlankFlds (Form$(), BadFld, FY1BegDate, FY2EndDate)
  
  FOR zz = 10 TO 12
    IF Date2Num(Form$(zz, 0)) = -32768 THEN
      BadFld = zz
      EXIT SUB
    END IF

    CheckDate = Date2Num(Form$(zz, 0))
    
    IF CheckDate < FY1BegDate OR CheckDate > FY2EndDate THEN
      BadFld = zz
      EXIT SUB
    END IF
  
  NEXT
  
'  FOR FieldNum = 3 TO 5
'
'    IF LEN(QPTrim$(Form$(FieldNum, 0))) = 0 THEN
'      BadFld = FieldNum
'      EXIT FOR
'    END IF
'
'  NEXT
  
END SUB

FUNCTION EdCheck4PO (Form$(), AcctFlds(), Fld() AS FieldInfo)

  NumFlds = UBOUND(Form$)
  GotOne = False

  REDIM TVendor(1)       AS VendorRecType
  REDIM TAPLedgerRec(1)  AS APLedger81RecType
  REDIM TAPDistRec(1)    AS APDistRecType
  
  DistRecLen = LEN(TAPDistRec(1))
  LedgerRecLen = LEN(TAPLedgerRec(1))
  
  VRecNum = CVI(LEFT$(Form$(FldNum("VRECNUM", Fld()), 0), 2))
  PONum$ = QPTrim$(Form$(FldNum("PONUM", Fld()), 0))
  
  OpenVendorFile VendorFile, NumVRecs
  GET VendorFile, VRecNum, TVendor(1)
  CLOSE
  
  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen
  
  NextTrans& = TVendor(1).FrstTran
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, TAPLedgerRec(1)
    IF TAPLedgerRec(1).TrCode = 4 AND QPTrim$(TAPLedgerRec(1).DOCNum) = PONum$ THEN
      GotOne = True

      LSET Form$(NumFlds - 1, 0) = MKL$(NextTrans&)

      'LSET Form$(NumFlds, 0) = "-1"         'set po active flag
      LSET Form$(NumFlds, 0) = STR$(-1)
 
      LSET Form$(7, 0) = STR$(TAPLedgerRec(1).Amt)

      OpenAcctFile AcctFileNum, NumAccts
      OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen
      NextDist& = TAPLedgerRec(1).FrstDist
      DistCnt = 0
      DO UNTIL NextDist& = 0
        GET APDistFile, NextDist&, TAPDistRec(1)
        DistCnt = DistCnt + 1
        GET AcctFileNum, TAPDistRec(1).DistAcctRec, acct
        StartFld = AcctFlds(DistCnt)
        LSET Form$(StartFld, 0) = TAPDistRec(1).DistAcctNum
        LSET Form$(StartFld + 1, 0) = acct.Title
        LSET Form$(StartFld + 2, 0) = MKI$(TAPDistRec(1).DistAcctRec)
        LSET Form$(StartFld + 3, 0) = STR$(TAPDistRec(1).DistAmt)
        NextDist& = TAPDistRec(1).NextDist
      LOOP
    END IF

    IF GotOne THEN
       IF DistCnt < NumAcctFlds THEN
         FOR CDCnt = AcctFlds(DistCnt + 1) TO AcctFlds(NumAcctFlds)'clear all of distrubtions
           LSET Form$(CDCnt, 0) = ""             'fields
           LSET Form$(CDCnt + 1, 0) = ""
           LSET Form$(CDCnt + 2, 0) = ""
           LSET Form$(CDCnt + 3, 0) = ""
         NEXT
       END IF
       EXIT DO
    END IF

    NextTrans& = TAPLedgerRec(1).NextTrans
  LOOP

  CLOSE

'08-12-96
'Added to make sure the POFLAG and PORECNUM
'are cleared to "0" if no PO.  If a field is blank
'QuickScreen stories MAX NEG VALUE in this field. This
'caused a bug in the posting routine
  IF NOT GotOne THEN
    LSET Form$(NumFlds - 1, 0) = "0"     'clear po rec number
    LSET Form$(NumFlds, 0) = "0"         'clear po active flag
  END IF

  EdCheck4PO = GotOne
 
END FUNCTION

FUNCTION EDCheck4Vendor% (Form$(), Terms)
  
  SHARED VendorIdx AS VendorIdxRecType
  SHARED Vendor AS VendorRecType
  
  MatchingVendors = 0
  List$ = SPACE$(65)
  
  VLookup$ = QPTrim$(UCASE$(Form$(2, 0)))
  VLookuplen = LEN(VLookup$)
  
  REDIM ScrBuf(2000)
  
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  
  IF NumActiveVendors = 0 THEN
    CLOSE
    ListVendors = 0
    EXIT FUNCTION
  END IF
  
  OpenVendorFile VendorFile, NumVRecs
  
  REDIM VendorList(1 TO 1) AS FLen               'Dim array to hold trans list
  FoundVendors = 0
  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
      FoundVendors = FoundVendors + 1
      REDIM PRESERVE VendorList(1 TO FoundVendors) AS FLen
      GET VendorFile, VendorIdx.RecNum, Vendor
      IF Vendor.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 13) = Vendor.VNAME
        MID$(VendorList(MatchingVendors).V, 39) = MKI$(Vendor.VTERMS)
        MID$(VendorList(MatchingVendors).V, 41) = MKI$(VendorIdx.RecNum)
      END IF
    END IF
  NEXT
  
  CLOSE
  
  SELECT CASE MatchingVendors
  CASE 0
    Ok = MsgBox("AP", "NOVENDOR")
    VendorPickList = 0
    EDCheck4Vendor = False
  CASE 1
    EDCheck4Vendor = True
    LSET Form$(2, 0) = VendorList(1).V
    LSET Form$(3, 0) = MID$(VendorList(1).V, 13)
    Terms = CVI(MID$(VendorList(1).V, 39, 2))
    LSET Form$(4, 0) = MID$(VendorList(1).V, 41, 2)
  CASE ELSE
    REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
    Row = 5: Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 37
    BoxBot = 18
    ListDescTxt$ = "VendorCode  Name"
    ListDesc$ = SPACE$(MaxLen + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    DO          'call vertmenu to display matching recs to choose from
      VertMenuT VendorList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        CLOSE
        EDCheck4Vendor = False
        EXIT DO
      END IF
      IF Choice > 0 THEN
        LSET Form$(2, 0) = VendorList(Choice).V
        LSET Form$(3, 0) = MID$(VendorList(Choice).V, 13)
        Terms = CVI(MID$(VendorList(Choice).V, 39, 2))
        LSET Form$(4, 0) = MID$(VendorList(Choice).V, 41, 2)
        EDCheck4Vendor = True
        'SaveField 0, Form$(), Fld(), BadFld
        EXIT DO
      END IF
    LOOP
    
  END SELECT
  
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  QPrintRC TName$, 3, 27, -1
  ERASE VendorList, ScrBuf
  
END FUNCTION

SUB EdEntry (APType)

  'STATIC OIDate, ODDate, OPDate
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  SHARED Choice$()
  
  SHARED VendorIdx AS VendorIdxRecType
  SHARED Vendor AS VendorRecType
  List$ = SPACE$(65)

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum
  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType
  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs
  
  REDIM Choice$(0 TO 2, 0)

'  Choice$(0, 0) = "8"
'  Choice$(1, 0) = "1 Open Payable"
  'Choice$(2, 0) = "2 Suspended Inv."
  
  NumOfScrns = 3
  
  '--get first active record number
  OpenAPEditFile APEditFile, NumEdTrans, APType

  CurrRec = 0
  NextRec = 1
  RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
  CLOSE APEditFile
  IF RecNum = 0 THEN
    EdMode = 1
  ELSE
    EdMode = 0
  END IF

Top:
  REDIM Frm(1)    AS FormInfo
  REDIM FormName$(1 TO NumOfScrns)
  CommaFormat$ = "#########,.##"
  
  FOR Scr = 1 TO NumOfScrns
    FormName$(Scr) = "AP" + LTRIM$(STR$(Scr))
  NEXT

  Header = 14  'was 11
  AddedNewVendor = False

  '-- Get total number of fields
  NumFlds = True
  FOR Scr = 1 TO NumOfScrns
    NumFlds = NumFlds + LibNumberOfFields("AP", FormName$(Scr)) + 1
  NEXT
  
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
  
  'for each page, get first and last fields
  StartEl = 0   'Load first form at array start
  FOR Scr = 1 TO NumOfScrns     'Get field defs
    LibGetFldDef "AP", FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT

  FOR F = 1 TO NumFlds          'For all fields on all pages
    LSET Form$(F, 0) = ""       'Clear each field
  NEXT
  
  
  Cnt = 0
  REDIM AcctFlds(1 TO NumAcctFlds)
  FOR F = 1 TO NumFlds
    IF QPTrim$(Fld(F).ScratchS) = "A" THEN
      Cnt = Cnt + 1
      AcctFlds(Cnt) = F
    END IF
  NEXT
  Frm(1).PageDisable = True
  Frm(1).StayOnField = True


'added this block
     POFldNum = FldNum("PONUM", Fld())
  TotalAmtFld = FldNum("INVAMT", Fld())
   DueDateFld = FldNum("DUEDATE", Fld())
   InvDateFld = FldNum("INVDATE", Fld())
  DistDateFld = FldNum("DISTDATE", Fld())
   PayCodeFld = FldNum("PAYCODE", Fld())
   PORecNumFld = FldNum("PORECNUM", Fld())
   GotPOFlagFld = FldNum("GOTPOFLG", Fld())
   

  Scr = 1       'Set screen number to one
  LibFile2Scrn "AP", FormName$(Scr), MonoCode, -1, ErrCode      'display screen
  'GOSUB ShowRecNo
  IF EdMode = 0 THEN
    Rec2Form Form$(), RecNum, Fld(), APType
    GotPOFlag = VAL(Form$(GotPOFlagFld, 0))      'po active flag
  ELSE
    GotPOFlag = 0
  END IF
  
  ShowEdMode EdMode
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = 0
      LSET Form$(PayCodeFld, 0) = "1"
      IF EdMode = 1 AND LEN(KeepDate$) > 0 THEN
        LSET Form$(InvDateFld, 0) = KeepDate$
        SaveField InvDateFld, Form$(), Fld(), BadFld
      END IF
      Action = 2
    END IF
    
    IF Frm(1).FldNo <> POFldNum AND Frm(1).PrevFld = POFldNum AND LEN(QPTrim$(Form$(POFldNum, 0))) > 0 AND NOT GotPOFlag THEN
      IF EdCheck4PO(Form$(), AcctFlds(), Fld()) THEN
        SaveField 0, Form$(), Fld(), BadFld
        SaveField GotPOFlagFld, Form$(), Fld(), BadFld
        SaveField PORecNumFld, Form$(), Fld(), BadFld
        GotPOFlag = True
        Action = 1
        IF Frm(1).FldNo = 4 THEN Frm(1).FldNo = 2
      ELSE
        IF GotPOFlag THEN
          GotPOFlag = False
          LSET Form$(TotalAmtFld, 0) = ""               'clear amt field
          LSET Form$(PORecNumFld, 0) = ""      'clear po rec num
          LSET Form$(GotPOFlagFld, 0) = "0"      'clear po active flag
          FOR TFCnt = AcctFlds(1) TO AcctFlds(NumAcctFlds) 'clear all of distrubtions
            LSET Form$(TFCnt, 0) = ""             'fields
            LSET Form$(TFCnt + 1, 0) = ""
            LSET Form$(TFCnt + 2, 0) = ""
            LSET Form$(TFCnt + 3, 0) = ""
          NEXT
          Action = 2
          SaveField 0, Form$(), Fld(), BadFld
        END IF
      END IF
    ELSEIF Frm(1).FldNo <> POFldNum AND Frm(1).PrevFld = POFldNum AND LEN(QPTrim$(Form$(POFldNum, 0))) = 0 AND NOT GotPOFlag THEN
      LSET Form$(PORecNumFld, 0) = "0"     'clear po rec number
      LSET Form$(GotPOFlagFld, 0) = "0"         'clear po active flag
      SaveField GotPOFlagFld, Form$(), Fld(), BadFld
      SaveField PORecNumFld, Form$(), Fld(), BadFld
      Action = 1
    END IF
    
    IF Frm(1).FldNo <> InvDateFld AND Frm(1).PrevFld = InvDateFld AND LEN(QPTrim$(Form$(InvDateFld, 0))) > 6 THEN
      TInvDate = Date2Num(Form$(InvDateFld, 0))
      TDueDate = TInvDate + Terms
      Form$(DueDateFld, 0) = Num2Date(TDueDate)
      Form$(DistDateFld, 0) = Form$(InvDateFld, 0)
      SaveField DueDateFld, Form$(), Fld(), BadFld
      SaveField DistDateFld, Form$(), Fld(), BadFld
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE -106   'Alt+F3
      
      IF Frm(1).FldNo > Header THEN             'if we're on an account field
        'for each accounting distribution
        FOR Cnt = 1 TO NumAcctFlds
          'if this acct field is less than the current field
          'set the row counter
          IF AcctFlds(Cnt) <= Frm(1).FldNo THEN
            FldRow = FldRow + 1 '+AcctFlds(Cnt)
          ELSE
            FOR Cnt3 = AcctFlds(FldRow) TO AcctFlds(FldRow) + 3
              LSET Form$(Cnt3, 0) = ""
            NEXT
            
            FOR Cnt3 = FldRow TO NumAcctFlds - 1
              LSET Form$(AcctFlds(Cnt3), 0) = Form$(AcctFlds(Cnt3 + 1), 0)
              LSET Form$(AcctFlds(Cnt3) + 1, 0) = Form$((AcctFlds(Cnt3 + 1) + 1), 0)
              LSET Form$(AcctFlds(Cnt3) + 2, 0) = Form$((AcctFlds(Cnt3 + 1) + 2), 0)
              LSET Form$(AcctFlds(Cnt3) + 3, 0) = Form$((AcctFlds(Cnt3 + 1) + 3), 0)
            NEXT
            'Cnt is at 24 after 23 times thru above loop
            LSET Form$(AcctFlds(Cnt3), 0) = ""
            LSET Form$(AcctFlds(Cnt3) + 1, 0) = ""
            LSET Form$(AcctFlds(Cnt3) + 2, 0) = ""
            LSET Form$(AcctFlds(Cnt3) + 3, 0) = ""
            
            Frm(1).FldNo = AcctFlds(FldRow)
            Frm(1).PrevFld = Frm(1).FldNo - 1
            FldRow = 0
            Action = 1
            EXIT FOR
          END IF
        NEXT
      END IF
    CASE F2Key  'Change Editing Mode
      
      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        SELECT CASE EdMode
        CASE 0  'were in Edit Mode
          EdMode = 1            ' switch to add mode
          GotPOFlag = 0
          ShowEdMode EdMode
          ResetForm Form$(), NumFlds
          FirstTime = True
          Frm(1).FldNo = 1
          Action = 1
          'QPrintRC "New  ", 25, 68, -1
          GOSUB ShowDistTot
          
        CASE 1  'currently adding records
          'check to see if there are any active transactions
          OpenAPEditFile APEditFile, NumEdTrans, APType
          CurrRec = 0: NextRec = 1
          RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE APEditFile
          IF RecNum = 0 THEN    'No active records can't edit
            EdMode = 1
          ELSE
            EdMode = 0          'switch to edit mode
            ShowEdMode EdMode
            RecNum = ListRecs(APType)
            IF RecNum > 0 THEN
              Rec2Form Form$(), RecNum, Fld(), APType
              GotPOFlag = VAL(Form$(GotPOFlagFld, 0))      'po active flag
              Frm(1).FldNo = 1
              Action = 1
            ELSE
              CLOSE
              EXIT SUB
            END IF
          END IF
          GOSUB ShowDistTot
        END SELECT              'Edmode
      END IF

'***********
'Look here
    CASE F7Key  'Add new vendor
      REDIM tempScrn2(2000)
      MScrnSave 1, 1, 25, 80, tempScrn2(0)
      NewVendorRec NVRecNum

      IF NVRecNum > 0 THEN
        AddedNewVendor = True
        OpenVendorFile VendorFile, NumVRecs
        GET VendorFile, NVRecNum, Vendor
        CLOSE VendorFile
        LSET Form$(2, 0) = Vendor.VNum
        LSET Form$(3, 0) = Vendor.VNAME
        LSET Form$(4, 0) = MKI$(NVRecNum)
        Terms = Vendor.VTERMS
        SaveField 2, Form$(), Fld(), BadFld
        Frm(1).FldNo = Frm(1).FldNo + 2  'was 5
        Action = 1
      ELSE
        AddedNewVendor = False
        Frm(1).FldNo = 1
        Action = 2
      END IF

      MScrnRest 1, 1, 25, 80, tempScrn2(0)
      Action = 1
'*************
    CASE ELSE
      
    END SELECT
    
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN      'User is moving around the form

      


'Look here
      IF (Frm(1).PrevFld = 2 AND Frm(1).FldNo <> 2) AND NOT AddedNewVendor THEN
      'Check to see if this is a valid vendor
        VendorOK = EDCheck4Vendor(Form$(), Terms)
        IF VendorOK THEN
          SaveField 1, Form$(), Fld(), BadFld
          Frm(1).FldNo = Frm(1).FldNo + 2  'was 5
          Action = 1
        ELSE
          'AddedNewVendor = False
          Frm(1).FldNo = 1
          Action = 2
        END IF
      ELSE
        AddedNewVendor = False
      END IF

      GOSUB ShowDistTot         'Update Distribution totals
      
      '--Check the field we're in to see if it is an acct field.
      'if so and if there is data in the field, remove dashes for user
      'set stripped acct string in form
      
      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
        LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
        Action = 1
      END IF
      
      '--the following checks the field we have just left to see
      'if it is an account field, and '  if the field is not blank,
      'the acct number is formatted with dashes
      'if the field is blank the acct name and distribution amt are blanked
      'and the amt field is protected to prevent user from entering a
      'distribution without an account.
      IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN       '"A" = AcctField
        IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
          TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).PrevFld, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).PrevFld, 0) = TempAcct$
          Fld(Frm(1).PrevFld + 3).Protected = False
        ELSE
          FOR Cnt2 = 1 TO 3
            LSET Form$(Frm(1).PrevFld + Cnt2, 0) = ""
            Fld(Frm(1).PrevFld + Cnt2).Protected = True
          NEXT
          IF Frm(1).FldNo > Frm(1).PrevFld THEN
            Frm(1).FldNo = Frm(1).PrevFld
          END IF
        END IF
        Action = 1
        
      END IF
      
      '--if we've just left an account field, check to see if
      'we have a valid acct number
      IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
        IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN     '= AcctField
          GOSUB CheckForValidAcctNum
        END IF
      END IF
      
    END IF      'End of test for PrevFld <> current FldNo
    
    IF SaveFlag THEN GOSUB SaveRec
    
    '--Process key strokes
    
    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      IF EMode = 0 THEN
        Ok = MsgBox%("GL.QSL", "OK2DELTR")
        IF Ok = 2 THEN          'Pressed Button #2
          Form$(1, 0) = "-1"
          SaveField 1, Form$(), Fld(), BadFld
          SaveTrans Form$(), RecNum, EdMode, APType
          'stuffbuf CHR$(0) + CHR$(F5Key)
          ResetForm Form$(), NumFlds
          RecNum = ListRecs(APType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), APType
            GotPOFlag = VAL(Form$(GotPOFlagFld, 0))      'po active flag
            Action = 1
          ELSE
            EXIT SUB
          END IF
        ELSE
          Action = 2
        END IF
        Frm(1).FldNo = 1
      END IF
      
    CASE F5KEY
      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        IF EdMode = 0 THEN
          RecNum = ListRecs(APType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), APType
            'GotPOFlag = VAL(Form$(NumFlds, 0))        'po active flag
            GotPOFlag = VAL(Form$(GotPOFlagFld, 0))
            Frm(1).FldNo = 1
            Action = 1
          END IF
        END IF
        
      END IF
      
    CASE F10Key

      '--Make sure acct fields are formatted correctly if user
      'Pressed F10 with cursor in an account field
      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
        IF LEN(QPTrim$(Form$(Frm(1).FldNo, 0))) > 0 THEN
          TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).FldNo, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).FldNo, 0) = TempAcct$
          SaveField Frm(1).FldNo, Form$(), Fld(), BadFld   'Write the flag to the form$
          'Fld(Frm(1).PrevFld + 3).Protected = False
          'Frm(1).FldNo = 1
          'Action = 1
        END IF
      END IF
      SaveFlag = True
      
    CASE CtrlL  'Ctrl+LeftArrow               'Get the next record
      
      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          Frm(1).FldNo = 1
          OpenAPEditFile APEditFile, NumEdTrans, APType
          CurrRec = RecNum: NextRec = -1
          RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE APEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), APType
          GotPOFlag = VAL(Form$(GotPOFlagFld, 0))      'po active flag
          GOSUB ShowDistTot
          ShowEdMode EdMode
          Scr = 1
          Action = 1
        END IF
      END IF
      
    CASE CtrlR  'Ctrl+RightArrow             'Get the previous record
      
      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          Frm(1).FldNo = 1
          OpenAPEditFile APEditFile, NumEdTrans, APType
          CurrRec = RecNum: NextRec = 1
          RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE APEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), APType
          GotPOFlag = VAL(Form$(GotPOFlagFld, 0))       'po active flag
          GOSUB ShowDistTot
          ShowEdMode EdMode
          Action = 1
          Scr = 1
        END IF
      END IF
      
    END SELECT

    'End of Checking key codes
    '--Set screen pages
    'IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld AND NOT StayOnPage THEN
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
      
      SELECT CASE Frm(1).FldNo
      CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
        Scr = 1
      CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
        Scr = 2
      CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
        Scr = 3
      END SELECT
      '  Action = 1

      LibFile2Scrn "AP", FormName$(Scr), MonoCode, -2, ErrCode
      
      GOSUB ShowDistTot
      ShowEdMode EdMode
      'GOSUB ShowRecNo
    END IF
    
    '--Buttons are not attached to the form, the following processes
    '--the rodent
    IF Frm(1).Presses THEN
      
      SELECT CASE Frm(1).MRow
      CASE 8
        SELECT CASE Frm(1).MCol
        CASE 71 TO 73
          PressButton -72, 8, 72, 72
        END SELECT
        
      CASE 19
        SELECT CASE Frm(1).MCol
        CASE 71 TO 73
          PressButton -80, 19, 72, 72
        END SELECT
        
      CASE 21
        SELECT CASE Frm(1).MCol
        CASE 9 TO 21
          PressButton F10Key, 21, 9, 21
        CASE 22 TO 37
          PressButton F3Key, 21, 22, 37         'Alt+F3
        CASE 38 TO 56
          PressButton -106, 21, 38, 56
        CASE 57 TO 71
          PressButton ESC, 21, 57, 71
        END SELECT
        
      CASE 25
        SELECT CASE Frm(1).MCol
        CASE 24 TO 38
          PressButton F7Key, 25, 24, 38
        CASE 41 TO 51
          PressButton F2Key, 25, 41, 51
        CASE 54 TO 64
          PressButton F5KEY, 25, 54, 64
        CASE 68 TO 70
          PressButton CtrlL, 25, 68, 70
        CASE 73 TO 75
          PressButton CtrlR, 25, 73, 75
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL Frm(1).KeyCode = ESC
  
  'End of FormEdit Loop
  EXIT SUB
  
  
CheckForValidAcctNum:

  AcctFld = Frm(1).PrevFld
  MatchAccts = 0
  Lookup$ = QPTrim$(UCASE$(Form$(AcctFld, 0)))
  Lookuplen = LEN(Lookup$)

  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

'look here
'The EDCheck4Acct function this gosub replaces was loading the vendor index
'file everytime it checked an account.
'Should help speed this up

  OpenAcctFile AcctFileNum, NumAccts

  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, acct
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).V = List$
      AcctList(MatchAccts).V = TempAcctIdx(Cnt).AcctNum
      MID$(AcctList(MatchAccts).V, TotLen + 4) = acct.Title
      MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
    END IF
  NEXT

  CLOSE AcctFileNum

  SELECT CASE MatchAccts
  CASE 0
    PLAY "L32O3EC"              'Acct NOT found
    Ok = MsgBox("GL.QSL", "NOACCT")
    EDCheck4Acct = False
  CASE 1
    AcctRecNum = CVI(RIGHT$(AcctList(1).V, 2))
    GOSUB EGetAcctInfo
    EDCheck4Acct = True
  CASE ELSE
    Row = 5
    Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
    BoxBot = 18
    ListDescTxt$ = "Account No.   Description"
    ListDesc$ = SPACE$(MaxLen + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    DO          'call vertmenu to display matching recs to choose from
      VertMenuT AcctList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EDCheck4Acct = False
        EXIT DO
      END IF
      IF Choice > 0 THEN
        AcctRecNum = CVI(RIGHT$(AcctList(Choice).V, 2))
        GOSUB EGetAcctInfo
        EDCheck4Acct = True
        EXIT DO
      END IF
    LOOP
  END SELECT

  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen

  ERASE AcctList, ScrBuf

  SELECT CASE EDCheck4Acct
  CASE True
    SaveField AcctFld, Form$(), Fld(), BadFld   'Write the flag to the form$
    SaveField AcctFld + 1, Form$(), Fld(), BadFld               'Write the flag to the form$
    Action = 2
  CASE ELSE
    Frm(1).FldNo = AcctFld      'Put user back on acct field
    LSET Form$(Frm(1).FldNo, 0) = ""
    LSET Form$(Frm(1).FldNo + 1, 0) = ""
    Action = 2
  END SELECT

RETURN

EGetAcctInfo:
  OpenAcctFile AcctFileNum, NumAccts
  GET AcctFileNum, AcctRecNum, acct
  CLOSE AcctFileNum
  LSET Form$(AcctFld, 0) = acct.Num
  LSET Form$(AcctFld + 1, 0) = acct.Title
  LSET Form$(AcctFld + 2, 0) = MKI$(AcctRecNum)
RETURN
  
  
ShowDistTot:
  TotDist# = 0
  FOR Cnt = 1 TO NumFlds        'look thru acct fld list
    IF QPTrim$(Fld(Cnt).ScratchS) = "S" THEN    '= Distribution Field
      TotDist# = Round#(TotDist# + Value(Form$(Cnt, 0), ECode))
    END IF
  NEXT
  
  Amt# = Round#(Value(Form$(TotalAmtFld, 0), ECode))
  UnDist# = Amt# - TotDist#
  
  Dist$ = SPACE$(14)
  UnDist$ = SPACE$(14)
  LSET Dist$ = FUsing(STR$(TotDist#), CommaFormat$)
  LSET UnDist$ = FUsing(STR$(UnDist#), CommaFormat$)
  QPrintRC UnDist$, 20, 24, -1
  QPrintRC Dist$, 20, 56, -1
  
  RETURN
  
  
SaveRec:

  BadFld = 0
  Check4BlankFlds Form$(), BadFld, FY1BegDate, FY2EndDate          'make sure no blank fields

  IF BadFld = 0 THEN            'if no blanks, make sure
    GOSUB ShowDistTot           'distribution are in bal

    IF UnDist# = 0 AND Amt# <> 0 THEN
      Form$(1, 0) = "0"         'Set deleted flag to false
      SaveField 1, Form$(), Fld(), BadFld       'Write the flag to the form$
      SaveTrans Form$(), RecNum, EdMode, APType 'Save the transaction
      KeepDate$ = Form$(10, 0)
      Ok = MsgBox%("GL.QSL", "SAVED")           'tell user we saved it


      IF EdMode = 1 THEN        'if in Add Mode, clear
        'ResetForm Form$(), NumFlds              ' the form for the next entry
        'Frm(1).FldNo = 1
        SaveFlag = False
        GOTO Top
      END IF
      'SaveField 0, Form$(), Fld(), BadFld

      Action = 1
      Scr = 1
      
    ELSE
      Ok = MsgBox%("GL.QSL", "BADDIST")         'Distributions out of balance
      Action = 1

'tell user
    END IF
    
  ELSE
    Ok = MsgBox%("GL.QSL", "BADDATE")           'Tell user no blank fields
    Frm(1).FldNo = BadFld
    
  END IF
  
  Action = 1
  FirstTime = True
  SaveFlag = False

  IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
    LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
    Action = 1
  END IF
  
  RETURN
  
  
CheckFormEdit:
  
  IF Frm(1).Edited THEN
    BEEP
    CanChange = False
    Ok = MsgBox("CMNDLG", "CHANGREC")           'returns 2 for Abandon
    IF Ok = 2 THEN CanChange = True
    Action = 2
  ELSE
    CanChange = True
  END IF
  
  RETURN
  
InitForm:
  REDIM Frm(1)    AS FormInfo
  Frm(1).PageDisable = True
 
  Action = 1
  RETURN
  
  'ShowRecNo:
  '  QPrintRC STR$(RecNum), 25, 67, -1
  '  RETURN

END SUB

FUNCTION FindVendorRec (VendorCode$)
  
  SHARED VendorIdx AS VendorIdxRecType
  
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  
  IF NumActiveVendors = 0 THEN
    FindVendorRec = 0
    CLOSE VendorIdxFile
    EXIT FUNCTION
  END IF
  
  Match = False
  FirstRec = 1
  LastRec = NumActiveVendors
  Lookfor$ = QPTrim$(UCASE$(VendorCode$))
  
  DO UNTIL LastRec < FirstRec
    
    MiddleRec = (LastRec + FirstRec) \ 2
    
    GET VendorIdxFile, MiddleRec, VendorIdx
    
    IF Lookfor$ = QPTrim$(VendorIdx.VendorCode) THEN
      Match = True
      EXIT DO
    ELSEIF Lookfor$ < VendorIdx.VendorCode THEN
      LastRec = MiddleRec - 1
    ELSE
      FirstRec = MiddleRec + 1
    END IF
    
  LOOP
  
  IF Match THEN
    FindVendorRec = VendorIdx.RecNum
  ELSE
    FindVendorRec = 0
  END IF
  
  CLOSE VendorIdxFile
  
END FUNCTION

'***************************************************************************
'Finds the next undeleted record.
'Call with NextRec value of -1 for previous record, +1 for the next record.
'If a record is not found, the function returns the value of CurrRec.
'File should be opened before function is called
'***************************************************************************
FUNCTION GetNextRec (FileNum, NumRecs, CurrRec, NextRec)
  
  Found = 0
  rec = CurrRec
  
  DO
    
    rec = rec + NextRec         'Set file pointer to next record
    
    IF rec > NumRecs OR rec <= 0 THEN           'test for beg or end of file
      Found = 0 'if no more records then get out
      EXIT DO
    END IF
    
    GET FileNum, rec, APEdit    'Get the record
    
    IF APEdit.DELFLAG <> True THEN                'Ok if not deleted
      Found = 1
      EXIT DO   'Get out of loop when we find one
    END IF
    
  LOOP
  
  IF Found = 0 THEN
    GetNextRec = CurrRec
  ELSE
    GetNextRec = rec
  END IF
  
END FUNCTION

FUNCTION ListRecs (APType)
  
  '  SHARED Cnf AS Config
  REDIM ScrBuf(2000)
  
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  
  'QPrintRC "List ", 25, 68, -1
  
  OpenAPEditFile APEditFile, NumEdTrans, APType
  
  REDIM TransList$(1 TO NumEdTrans)             'Dim array to hold trans list
  
  Active = 0    'Counter for Active Transactions
  
  FOR Cnt = 1 TO NumEdTrans     'Read the file
    GET APEditFile, Cnt, APEdit
    IF NOT APEdit.DELFLAG THEN  'if active add to list
      Active = Active + 1
      List$ = SPACE$(75)
      LSET List$ = APEdit.VENDNAME
      MID$(List$, 24) = LEFT$(APEdit.INVNUM, 15)
      MID$(List$, 42) = ConvDateStr1(Num2Date(APEdit.INVDATE))
      MID$(List$, 53) = FUsing$(STR$(APEdit.INVAMT), "#########,.##")
      MID$(List$, 73) = MKI$(Cnt)
      TransList$(Active) = List$
    END IF
    
  NEXT
  CLOSE
  
  IF Active = 0 THEN
    'KILL "CJED.DAT"                     '..init file to zero
    EXIT FUNCTION
  END IF
  
  REDIM PRESERVE TransList$(1 TO Active)        'Redim list to only active trans
  
  LOCATE 3, 6   'set up vertmenu
  Action = 0
  MaxLen = 65
  BoxBot = 23
  
  ListDesc$ = "   Vendor                Invoice           Inv Date       Total Amt  "
  QPrintRC ListDesc$, 2, 6, Cnf.ActivCh
  'MPaintBox 6, 75, 7, 76, 8              'to add shadow to list
  
  DO            'call vertmenu
    VertMenu TransList$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      CLOSE
      Abort = True
      EXIT DO
      EXIT FUNCTION
    END IF
    IF Choice > 0 THEN EXIT DO
  LOOP
  
  IF Abort THEN
    Choice = 0
  ELSE
    RecNum = CVI(MID$(TransList$(Choice), 73, 2))               'Get the record number
  END IF        'off the end of the $
  
  CLOSE APEditFile              'Clean up after ourselves
  ERASE TransList$
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  
  ListRecs = RecNum
  
  
END FUNCTION

SUB MakeInvPOIF (FileName$)


STOP
'REDIM PO2Post(1 TO 1) AS GLTransRecType
'
'POIFFile = FREEFILE
'
'IFLen = LEN(PO2Post(1))
'
'OPEN FileName$ FOR RANDOM AS POIFFile LEN = IFLen
'Num2Get = LOF(POIFFile) \ IFLen
'REDIM PO2Post(1 TO Num2Get)  AS GLTransRecType
'FOR Cnt = 1 TO Num2Get
'  GET POIFFile, Cnt, PO2Post(Cnt)
'NEXT
'CLOSE
'
'FOR Cnt = 1 TO Num2Get
'  PO2Post(Cnt).AcctNum = LEFT$(PO2Post(Cnt).AcctNum, FundLen) + EncAcct$
'
'NEXT
'
'  'Tr2Post(1).Ref = APEdit.InvNum
'  'Tr2Post(1).DrAmt = 0
'  'Tr2Post(1).CrAmt = TrFundSum#(Fund)
'  'Tr2Post(1).Src = "AP" + ConvDateStr$(DATE$)
'EncAcct$

END SUB

SUB MakeTempReport

  '--Posting Report
  DIM Trans2Post AS GLTransRecType
  RecLen = LEN(Trans2Post)
  
  CJ2Post = FREEFILE
  OPEN "APINVIF.dat" FOR RANDOM AS CJ2Post LEN = RecLen
  
  NumTrans = LOF(CJ2Post) \ RecLen
  
  PrnFile = FREEFILE
  OPEN "Trans.prn" FOR OUTPUT AS #PrnFile
  
  'print field names
  ToPrint$ = SPACE$(78)
  LSET ToPrint$ = "Date"
  MID$(ToPrint$, 11) = "Acct Number"
  MID$(ToPrint$, 29) = "Description"
  MID$(ToPrint$, 46) = "Reference"
  MID$(ToPrint$, 61) = "Debit"
  MID$(ToPrint$, 74) = "Credit"
  PRINT #PrnFile, ToPrint$
  
  FOR Cnt = 1 TO NumTrans
    GET CJ2Post, Cnt, Trans2Post
    ToPrint$ = SPACE$(78)
    LSET ToPrint$ = ConvDateStr1(Num2Date$(Trans2Post.TRDate))
    MID$(ToPrint$, 11) = Trans2Post.AcctNum
    MID$(ToPrint$, 29) = LEFT$(Trans2Post.Desc, 15)
    MID$(ToPrint$, 46) = Trans2Post.Ref
    MID$(ToPrint$, 56) = FUsing(STR$(Trans2Post.DrAmt), "#######,.##")
    MID$(ToPrint$, 69) = FUsing(STR$(Trans2Post.CrAmt), "#######,.##")
    'MID$(ToPrint$, 80) = STR$(Cnt&)
    'MID$(ToPrint$, 90) = STR$(Trans2Post.NextTran)
    PRINT #PrnFile, ToPrint$
  NEXT
  
  CLOSE
  
  'Prn2Scrn "Trans.prn", 1, 25, 11, PrnFlag
  KILL "trans.prn"
  
END SUB

SUB NewVendorRec (NVRecNum%)

  REDIM Vendor(1) AS VendorRecType

  REDIM TempVendorRec(1) AS VendorRecType
  Vendor(1) = TempVendorRec(1)

  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  REDIM Frm(1) AS FormInfo           'Form editing info

  LibName$ = "AP.QSL"           'Screen library
  FormName$ = "APVENDOR"
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintTitle User$
  Action = 1

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--User is moving around the form
    '--Check for duplicate vendor when user leaves field 1
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN
      IF Frm(1).PrevFld = 1 THEN
        VRecNum = FindVendorRec(QPTrim$(Form$(1, 0)))
        IF VRecNum > 0 THEN
          Ok = MsgBox("AP.QSL", "DUPVNDR")
          Form$(1, 0) = ""
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB SaveNewRec
      Done = True
      NVRecNum = NewVRecNum
    CASE ESC
      Done = True
      NVRecNum = 0
    CASE ELSE
      Done = False
    END SELECT

  LOOP UNTIL Done

  EXIT SUB

SaveNewRec:
  PrintHelp "Saving to disk.. Please wait."

  OpenVendorFile VendorFile, NumVRecs

  Vendor(1).VNum = Form$(1, 0)
  Vendor(1).VNAME = Form$(2, 0)
  Vendor(1).Addr1 = Form$(3, 0)
  Vendor(1).City = Form$(5, 0)
  Vendor(1).State = Form$(6, 0)
  Vendor(1).Zip = Form$(7, 0)
  Vendor(1).PayToName = Form$(8, 0)
  Vendor(1).PayToAddr = Form$(9, 0)
  Vendor(1).PayToCity = Form$(11, 0)
  Vendor(1).PaytoState = Form$(12, 0)
  Vendor(1).PaytoZip = Form$(13, 0)
  Vendor(1).Addr2 = Form$(4, 0)
  Vendor(1).PaytoAddr2 = Form$(10, 0)
  Vendor(1).VIN = GetNewVendorPIN

  Vendor(1).Phone = Form$(14, 0)
  'Vendor.YTDPurch = value(Form$(15, 0))
  Vendor(1).FedID = Form$(16, 0)
  Vendor(1).CoCode = Form$(17, 0)
  Vendor(1).StCode = Form$(18, 0)
  Vendor(1).DELFLAG = 0

  NewVRecNum = NumVRecs + 1

  PUT VendorFile, NewVRecNum, Vendor(1)

  CLOSE VendorFile

  IndexVendorFile

  RETURN

END SUB

SUB OpenAPEditFile (APEditFile, NumEdTrans, APType)
  
  SELECT CASE APType
  CASE 1
    FileName$ = "APPED.DAT"
  CASE 2
    FileName$ = "APIED.DAT"
  END SELECT
  
  EdLen = LEN(APEdit)
  APEditFile = FREEFILE
  OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS APEditFile LEN = EdLen
  NumEdTrans = LOF(APEditFile) \ EdLen
  
END SUB

SUB OpenAPVendorFile (APVendorFile, NumVRecs, VendorRecLen)
  
  APVendorFile = FREEFILE
  OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS APVendorFile LEN = VendorRecLen
  NumVRecs = LOF(APVendorFile) \ VendorRecLen
  
END SUB

SUB PostInvTrans (APType)

  'GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  'STOP
  SHARED APAcct$
  SHARED Vendor AS VendorRecType

  REDIM APDistRec(1) AS APDistRecType
  REDIM APLedgerRec(1) AS APLedger81RecType

  REDIM TR2Post(1) AS GLTransRecType
  LedgerRecLen = LEN(APLedgerRec(1))
  DistRecLen = LEN(APDistRec(1))
  VendorRecLen = LEN(Vendor)
  
  '--Paint a background
  FOR Cnt = 2 TO 24
    QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
  NEXT
  
  '--Verify that there are transactions
  OpenAPEditFile APEditFile, NumEdTrans, APType
    FOR Cnt = 1 TO NumEdTrans
      GET APEditFile, Cnt, APEdit
      IF NOT APEdit.DELFLAG THEN
        Active = Active + 1
      END IF
    NEXT
  CLOSE
  
  '--Check for no active transactions
  IF Active = 0 THEN
    '--No active transactions - tell user and get out
    Ok = MsgBox%("GL.QSL", "NOTRANS")
    EXIT SUB
  END IF
  
  '--make sure we're ready to post
  Ok = MsgBox%("GL.QSL", "POSTGJ")
  IF Ok = 1 THEN EXIT SUB       '1=Cancel Button
  
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM TrFundSum#(1 TO NumFunds)
  
  AP2Post = FREEFILE
  OPEN "APINVIF.dat" FOR RANDOM AS AP2Post LEN = LEN(TR2Post(1))

  REDIM POIFRec(1) AS GLTransRecType
  PO2Post = FREEFILE
  OPEN "POINVIF.dat" FOR RANDOM AS PO2Post LEN = LEN(POIFRec(1))
  
  POIFRecNum = 0
  RecordNum = 0 'Reset Active counter for posting
  MSrc$ = "AP" + ConvDateStr$(DATE$)

  OpenAPEditFile APEditFile, NumEdTrans, APType
  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen
  OpenAPVendorFile APVendorFile, NumVRecs, VendorRecLen
  OpenAcctFile AcctFileNum, NumAccts

  LogFile = FREEFILE
  OPEN "GLUTIL.LOG" FOR APPEND AS LogFile
  PRINT #LogFile, "Invoice Posting started @ " + TIME$ + " on " + DATE$
  
  '--Post transaction to A/P Ledger and update links
  FOR Cnt = 1 TO NumEdTrans
    GET APEditFile, Cnt, APEdit

    IF NOT APEdit.DELFLAG THEN

      VendRecNum = CVI(APEdit.VRecNum)  'FindVendorRec(APEdit.VENDOR)
      PrintHelp "Posting Vendor: " + STR$(VendRecNum)

      IF VendRecNum > 0 THEN

        PostCnt = PostCnt + 1
        NumLedgerRecs = NumLedgerRecs + 1

        '--update vendor to transaction link
        OpenVendorFile VendorFile, NumVRecs
        GET VendorFile, VendRecNum, Vendor
        PrevVendTrans = Vendor.LastTran
        FrstVendTrans = Vendor.FrstTran
        IF FrstVendTrans = 0 THEN
          Vendor.LastTran = NumLedgerRecs
          Vendor.FrstTran = NumLedgerRecs
        ELSE
          Vendor.LastTran = NumLedgerRecs
          GET APLedgerFile, PrevVendTrans, APLedgerRec(1)
          APLedgerRec(1).NextTrans = NumLedgerRecs
          PUT APLedgerFile, PrevVendTrans, APLedgerRec(1)
        END IF
        PUT VendorFile, VendRecNum, Vendor
        CLOSE VendorFile

        '--write transaction to apledger
        APLedgerRec(1).VRecNum = VendRecNum
        APLedgerRec(1).VendorCode = APEdit.Vendor
        APLedgerRec(1).TRDate = APEdit.INVDATE
        APLedgerRec(1).DOCNum = APEdit.INVNUM
        APLedgerRec(1).PONum = APEdit.PONum
        APLedgerRec(1).DUEDATE = APEdit.DUEDATE
        APLedgerRec(1).TrCode = 1
        APLedgerRec(1).PAYCODE = VAL(APEdit.PAYCODE)
        APLedgerRec(1).GLDistDate = APEdit.DistDate
        APLedgerRec(1).Amt = APEdit.INVAMT
        APLedgerRec(1).NextTrans = 0
        APLedgerRec(1).FrstDist = NumDistRecs& + 1
        '--New File format
        APLedgerRec(1).Get1099 = APEdit.Get1099
        APLedgerRec(1).COMMENt = APEdit.INVDESC
        APLedgerRec(1).PSLFLag = APEdit.PSLFLag

        'IF INSTR(APEdit.INVNUM, "LAB BILL") THEN STOP
        '--write distributions to apdist.dat

        'FOR DistCnt = 1 TO 24 'LastDist
        '  IF LEN(QPTrim$(APEdit.Dist(DistCnt).DACN)) > 0 THEN
        'NEXT
        
        'DistWriteCnt = 0
        FOR DistCnt = 1 TO 24 'LastDist
          IF LEN(QPTrim$(APEdit.Dist(DistCnt).DACN)) > 0 AND APEdit.Dist(DistCnt).DAMT <> 0 THEN
            REDIM APDistRec(1) AS APDistRecType
            NumDistRecs& = NumDistRecs& + 1
            APDistRec(1).APLedgerRec = NumLedgerRecs
            APDistRec(1).DistAcctRec = CVI(APEdit.Dist(DistCnt).DACREC)
            APDistRec(1).DistAcctNum = APEdit.Dist(DistCnt).DACN
            APDistRec(1).DistAmt = APEdit.Dist(DistCnt).DAMT
            APDistRec(1).NextDist = NumDistRecs& + 1
            PUT APDistFile, NumDistRecs&, APDistRec(1)
            '--could put glupdate here
          END IF
        NEXT

        'rewrite the last valid distrubtion to indicate
        'this is the last distrubtion for this invoice
        APDistRec(1).NextDist = 0
        PUT APDistFile, NumDistRecs&, APDistRec(1)

        '--update the last distribution pointer in apledger.dat
        APLedgerRec(1).LastDist = NumDistRecs&
        PUT APLedgerFile, NumLedgerRecs, APLedgerRec(1)
        'CCnt = CCnt + 1 whats this for?
        
        '--Post PO
        'IF APEdit.POFLAG <> False THEN
        IF APEdit.POFLAG = True THEN

          'PORecNum& = CVL(APEdit.POAPLRecNum)
          'CancelPO PORecNum&, APLedgerFile, NumLedgerRecs, APDistFile, NumDistRecs&

          '--Cancel Open PO
          GET APLedgerFile, CVL(APEdit.POAPLRecNum), APLedgerRec(1)
          APLedgerRec(1).TrCode = -4
          PUT APLedgerFile, CVL(APEdit.POAPLRecNum), APLedgerRec(1)

          '--Create PO Liquidation1 entries
          NextDist& = APLedgerRec(1).FrstDist
          DO UNTIL NextDist& = 0
            GET APDistFile, NextDist&, APDistRec(1)

            '--Make Debit side of entry
            POIFRec = POIFRec + 1
            POIFRec(1).Src = "IL" + ConvDateStr$(DATE$)
            AcctNum$ = LEFT$(APDistRec(1).DistAcctNum, FundLen) + EncAcct$
            POIFRec(1).AcctNum = AcctNum$
            POIFRec(1).TRDate = Date2Num(DATE$)
            POIFRec(1).Desc = APEdit.VENDNAME  '"A/P Liquidation"  'APLedger(1).PONum
            POIFRec(1).Ref = APEdit.INVNUM   'APLedgerRec(1).
            POIFRec(1).CrAmt = 0
            POIFRec(1).DrAmt = APDistRec(1).DistAmt
            PUT PO2Post, POIFRec, POIFRec(1)

            POIFRec = POIFRec + 1
            POIFRec(1).AcctNum = APDistRec(1).DistAcctNum
            POIFRec(1).CrAmt = APDistRec(1).DistAmt
            POIFRec(1).DrAmt = 0
            PUT PO2Post, POIFRec, POIFRec(1)

            NextDist& = APDistRec(1).NextDist
          LOOP
        END IF 'po flag
      
      '--could'nt find vendor
      ELSE
        '--Mark trans as deleted and log it.
        BadVendor = BadVendor + 1
        APEdit.DELFLAG = True
        PUT APEditFile, Cnt, APEdit
        GOSUB LogBadTrans

      END IF '--test for good vendor
    END IF   '--test for not deleted trans
  NEXT       'Trans
  
  '--All Done with post to apledger & apdist
  IF BadVendor = 0 THEN
    PRINT #LogFile, "No Posting Errors. Transactions Posted: "; PostCnt
  END IF

  '--Now create interface files for posting to gltrans and potrans
  FOR Cnt = 1 TO NumEdTrans

    REDIM TrFundSum#(1 TO NumFunds)

    '--Reinitialize transaction fund summary elements for next loop
    GET APEditFile, Cnt, APEdit

    '--If the invoice is not deleted post it
    IF NOT APEdit.DELFLAG THEN
      FOR AcctDist = 1 TO 24

        '--Write out only those with an account number
        IF LEN(QPTrim$(APEdit.Dist(AcctDist).DACN)) > 0 THEN

          '--Make the Debits
          RecordNum = RecordNum + 1
          TR2Post(1).AcctNum = APEdit.Dist(AcctDist).DACN
          TR2Post(1).TRDate = APEdit.DistDate    'APEdit.INVDATE
          TR2Post(1).Desc = APEdit.VENDNAME
          TR2Post(1).Ref = APEdit.INVNUM
          TR2Post(1).DrAmt = APEdit.Dist(AcctDist).DAMT
          TR2Post(1).CrAmt = 0
          TR2Post(1).Src = MSrc$ '"AP" + ConvDateStr$(DATE$)
          TR2Post(1).Marked = False
          PUT AP2Post, RecordNum, TR2Post(1)

          '--This creates an entry to ap for every distribution
          '--Make the Credits
          'RecordNum = RecordNum + 1
          'TR2Post(1).AcctNum = LEFT$(APEdit.Dist(AcctDist).DACN, FundLen) + APAcct$
          'TR2Post(1).TRDate = APEdit.DistDate    'APEdit.INVDATE
          'TR2Post(1).Desc = APEdit.VENDNAME
          'TR2Post(1).Ref = APEdit.INVNUM
          'TR2Post(1).DrAmt = 0
          'TR2Post(1).CrAmt = APEdit.Dist(AcctDist).DAMT
          'TR2Post(1).Src = MSrc$ '"AP" + ConvDateStr$(DATE$)
          'TR2Post(1).Marked = False
          'PUT AP2Post, RecordNum, TR2Post(1)
          
          '--This summarizes ap entry by fund
          '--Add the distribution to the proper fund
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(APEdit.Dist(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(APEdit.Dist(AcctDist).DAMT)
              'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(APEdit.DIST(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT
         END IF  '--test for blank distribution line
      NEXT      '--Acct'g Distribution

      '--Make the A/P Credits
      FOR Fund = 1 TO NumFunds
        IF TrFundSum#(Fund) <> 0 THEN
          RecordNum = RecordNum + 1
          TR2Post(1).AcctNum = FundList$(Fund) + APAcct$
          TR2Post(1).TRDate = APEdit.DistDate
          TR2Post(1).Desc = APEdit.VENDNAME
          TR2Post(1).Ref = APEdit.INVNUM
          TR2Post(1).DrAmt = 0
          TR2Post(1).CrAmt = TrFundSum#(Fund)
          TR2Post(1).Src = MSrc$ '"AP" + ConvDateStr$(DATE$)
          TR2Post(1).Marked = False
          PUT AP2Post, RecordNum, TR2Post(1)
        END IF  '--Fund summary <> 0
      NEXT      '--fund check for balance

    END IF      '--not deleted
  NEXT          '--transaction

  CLOSE
  
  '--Post Distributions to General Ledger Accts
  Post2GL "APINVIF.dat", BadGLTrans%

  '--Post PO liquidations
  IF FileSize("POINVIF.dat") > 0 THEN
    Post2PO "POINVIF.dat", BadPOTrans%
  END IF
  
  '--Tell user if we have any problems
  IF BadVendor > 0 THEN
    CLS
    PRINT "Error: "; BadVendor; " Invoice(s) NOT posted."
    PRINT "Unable to locate vendor.  Review Posting Log for details."
    PRINT "Press any key to continue."
    k$ = INPUT$(1)
  END IF
  
  IF BadGLTrans > 0 THEN
    CLS
    PRINT "Error: "; BadGLTrans; " Invoice Distribution(s) NOT posted."
    PRINT "Review GL Posting Log."
    PRINT "Press any key to continue."
    k$ = INPUT$(1)
  END IF

  IF BadPOTrans > 0 THEN
    CLS
    PRINT "Error: "; BadPOTrans; " PO liquidation(s) NOT posted."
    PRINT "Review PO Posting Log."
    PRINT "Press any key to continue."
    k$ = INPUT$(1)
  END IF
  
  KILL "APIED.DAT"
  KILL "APINVIF.DAT"
  KILL "POINVIF.DAT"

EXIT SUB

LogBadTrans:
  PRINT #LogFile, "Unable to find Vendor. Transaction deleted."
  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 1) = APEdit.Vendor
  MID$(ToPrint$, 12) = Num2Date(APEdit.INVDATE)
  MID$(ToPrint$, 24) = APEdit.INVNUM
  MID$(ToPrint$, 44) = STR$(Round(APEdit.INVAMT))
  PRINT #LogFile, ToPrint$
RETURN

END SUB

SUB PrnEditList (APType)

  FF$ = CHR$(12)
  MaxLines = 56
  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "#########,.##"   'ten millions
  DistSumLine$ = "-------------"
  TransTotal# = 0
  TranCnt = 0

  SELECT CASE APType
  CASE 1
    '--Purchase Orders
    FileName$ = "APPOREG.PRN"
    RegTitle$ = "Purchase Orders"
    TranCol = CreditCol
    CashCol = DebitCol
  CASE 2
    '--Invoices
    FileName$ = "APINVREG.PRN"
    RegTitle$ = "A/P Purchases Journal"
    TranCol = CreditCol
    'CashCol = CreditCol
  END SELECT

  REDIM Title$(7)
  Title$(1) = User$
  Title$(2) = RegTitle$
  Title$(3) = "Run Date: " + DATE$'""
  Title$(4) = ""
  Title$(5) = "Vendor                       Comment                                PSL  1099 "
  Title$(6) = "Invoice             PO       Date      Due Date  Post Date                 Amt"
  Title$(7) = STRING$(78, "=")

  SHARED Choice$()              'Multiple choice array

  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array

  REDIM Choice$(0 TO 3, 1)      'Choices for entry type field
  Choice$(0, 0) = "2"           'Field Number
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"

  '--Get user input on where to send report
  LibLoadDisplayForm "GL.QSL", "PRNGJ", Form$(), Fld()
  PrintHelp "Print A/P Purchases Journal"

  DIM Frm AS FormInfo
  Action = 1

  '--Get the report options from user
  DO

    EditForm Form$(), Fld(), Frm, Cnf, Action

    '--Check the user's key presses
    SELECT CASE Frm.KeyCode

    CASE F10Key 'SaveButton
      IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(2, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
      END IF

    CASE ESC
      EXIT SUB

    END SELECT

  LOOP UNTIL Frm.KeyCode = F10Key

  '==Start processing register=>
  LOCATE 1, 1, 0

  OpenAPEditFile APEditFile, NumEdTrans, APType
  PrnFileNum = FREEFILE

  OPEN FileName$ FOR OUTPUT AS #PrnFileNum

  '--Get a list of active funds
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  REDIM FundGrdTot#(1 TO NumFunds)

  GOSUB PrintHeader

  FOR Transaction = 1 TO NumEdTrans
    GET APEditFile, Transaction, APEdit

    IF NOT APEdit.DELFLAG THEN
      TranCnt = TranCnt + 1
      TransTotal# = TransTotal# + Round#(APEdit.INVAMT)

      '--Print 1st Line - Transaction details
      ToPrint$ = SPACE$(78)
      LSET ToPrint$ = APEdit.VENDNAME
      MID$(ToPrint$, 30) = APEdit.INVDESC
      MID$(ToPrint$, 70) = APEdit.PSLFLag
      MID$(ToPrint$, 75) = APEdit.Get1099
      PRINT #PrnFileNum, ToPrint$

      ToPrint$ = SPACE$(78)
      'MID$(ToPrint$, 1) = LEFT$(APEdit.VENDNAME, 16)
      MID$(ToPrint$, 1) = APEdit.INVNUM
      MID$(ToPrint$, 20) = LEFT$(APEdit.PONum, 9)
      MID$(ToPrint$, 30) = ConvDateStr1$(Num2Date(APEdit.INVDATE))
      MID$(ToPrint$, 40) = ConvDateStr1$(Num2Date(APEdit.DUEDATE))
      MID$(ToPrint$, 50) = ConvDateStr1$(Num2Date(APEdit.DistDate))
      'MID$(ToPrint$, 60) = LEFT$(APEdit.INVNUM, 6)
      MID$(ToPrint$, 66) = FUsing$(STR$(APEdit.INVAMT), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$

      '--Blank line between detail and acct'g distributions
      PRINT #PrnFileNum,

      '--Print Distribution Label
      ToPrint$ = SPACE$(78)
      MID$(ToPrint$, 2) = "Accounting Distribution:"
      PRINT #PrnFileNum, ToPrint$

      '--Print Field Titles
      ToPrint$ = SPACE$(78)
      MID$(ToPrint$, 4) = "Account Number   Name                                  Distribution"
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 4
      IF LineCnt >= MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintHeader
      END IF

      '--Print Accounting Distributions
      TotTranDist# = 0

      '--Loop Thru distributions to print and summarize
      FOR AcctDist = 1 TO 24
        '--no more distributions when we find a blank Acct Number field
        IF LEN(QPTrim$(APEdit.Dist(AcctDist).DACN)) > 0 THEN
          '--Add distribution to total
          TotTranDist# = TotTranDist# + Round#(APEdit.Dist(AcctDist).DAMT)
          '--Add distribution to proper fund
          Found = False
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(APEdit.Dist(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              Found = True
              FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(APEdit.Dist(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT
          IF NOT Found THEN
            Ok = MsgBox("GL", "NOFUND")
            EXIT FOR
          END IF
          '--Print this distribution
          ToPrint$ = SPACE$(78)
          MID$(ToPrint$, 4) = APEdit.Dist(AcctDist).DACN
          MID$(ToPrint$, 21) = APEdit.Dist(AcctDist).DACNM
          MID$(ToPrint$, TranCol) = FUsing$(STR$(APEdit.Dist(AcctDist).DAMT), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          LineCnt = LineCnt + 1
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintHeader
          END IF
        'ELSE
        '  BEEP
        END IF  'Active transaction test
      NEXT      'Distribution

      '--Summary line after last distribution
      ToPrint$ = SPACE$(78)
      MID$(ToPrint$, TranCol) = DistSumLine$
      PRINT #PrnFileNum, ToPrint$

      '--Transaction Distribution Totals
      ToPrint$ = SPACE$(78)
      MID$(ToPrint$, 4) = "Total Distributed"
      MID$(ToPrint$, TranCol) = FUsing$(STR$(TotTranDist#), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 2
      IF LineCnt >= MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintHeader
      END IF
      
      '--2 blank lines before next distribution
      ToPrint$ = STRING$(78, "-")
      PRINT #PrnFileNum, ToPrint$
      PRINT #PrnFileNum,
      LineCnt = LineCnt + 2
      IF LineCnt >= MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintHeader
      END IF

    END IF      'Not deleted test

  NEXT          'Transaction

  IF LineCnt > 45 THEN
    PRINT #PrnFileNum, FF$
  END IF

  '--Summary
  ToPrint$ = SPACE$(78)
  LSET ToPrint$ = "File Totals:"
  PRINT #PrnFileNum, ToPrint$

  ToPrint$ = SPACE$(78)
  LSET ToPrint$ = "Number of Transactions"
  MID$(ToPrint$, 31) = FUsing$(STR$(TranCnt), "####")
  PRINT #PrnFileNum, ToPrint$

  ToPrint$ = SPACE$(78)
  LSET ToPrint$ = "Grand Totals"
  MID$(ToPrint$, 25) = FUsing$(STR$(TransTotal#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$

  PRINT #PrnFileNum,

  ToPrint$ = SPACE$(78)
  LSET ToPrint$ = "Summary by Fund:"
  PRINT #PrnFileNum, ToPrint$

  FOR Cnt = 1 TO NumFunds
    IF FundGrdTot#(Cnt) <> 0 THEN
      ToPrint$ = SPACE$(78)
      LSET ToPrint$ = "Fund" + " " + FundList$(Cnt)
      MID$(ToPrint$, 25) = FUsing$(STR$(FundGrdTot#(Cnt)), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      GrdTot# = GrdTot# + Round#(FundGrdTot#(Cnt))
    END IF
  NEXT

  ToPrint$ = SPACE$(78)
  LSET ToPrint$ = "Total All Funds"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrdTot#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  PRINT #PrnFileNum, FF$

  CLOSE

  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT

  PrintRptFile RegTitle$, FileName$, LPTNo, RetCode%, EntryPoint
  KILL FileName$

EXIT SUB


PrintHeader:
  FOR HCnt = 1 TO 7
    PRINT #PrnFileNum, Title$(HCnt)
  NEXT
  LineCnt = 7
RETURN





  
END SUB

SUB Rec2Form (Form$(), RecordNumber%, Fld() AS FieldInfo, APType)
  
  OpenAPEditFile APEditFile, NumEdTrans, APType
  
  GET APEditFile, RecordNumber, APEdit
  
  Form$(0, 0) = SPACE$(LEN(APEdit))
  
  BCopy VARSEG(APEdit), VARPTR(APEdit), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  
  'Form$(1, 0) = STR$(APEdit.DelFlag)
  'Form$(2, 0) = Num2Date(APEdit.InvDate)
  'Form$(3, 0) = APEdit.InvNum
  'Form$(4, 0) = APEdit.PONum
  'Form$(5, 0) = STR$(APEdit.InvAMT)
  'Form$(6, 0) = APEdit.BatchNum
  'Form$(7, 0) = STR$(APEdit.TCode)
  
  CLOSE APEditFile              'close the file
  
  'Rec$ = SPACE$(5)
  'LSET Rec$ = STR$(RecordNumber%)
  'QPrintRC Rec$, 25, 67, -1
  
  
END SUB

SUB ResetForm (Form$(), NumFlds)
  
  FOR F = 1 TO NumFlds
    Form$(F, 0) = ""

  NEXT
  
END SUB

SUB SaveTrans (Form$(), RecordNumber, EdMode, APType)
  
  OpenAPEditFile APEditFile, NumEdTrans, APType
  
  IF EdMode = 1 THEN
    rec = NumEdTrans + 1
  ELSE
    rec = RecordNumber
  END IF
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(APEdit), VARPTR(APEdit), LEN(Form$(0, 0)), 0

  PUT APEditFile, rec, APEdit

  CLOSE APEditFile

END SUB

SUB ShowEdMode (EdMode%)
  
'  Status$ = SPACE$(24)
'  F2Ed$ = "  F2=Edit  "
'  F2New$ = "  F2=New   "
'  ActiveColor = 112
'  InActiveColor = 115
'
'  SELECT CASE EdMode
'  CASE 0
'    LSET Status$ = "Edit an Existing Record"
'    F2BtnCap$ = F2New$
'    F5BtnColor = ActiveColor
'
'  CASE 1
'    LSET Status$ = "Add a New Record"
'    F2BtnCap$ = F2Ed$
'    F5BtnColor = InActiveColor
'
'  END SELECT
'
'  QPrintRC Status$, 25, 2, -1
'  QPrintRC F2BtnCap$, 25, 28, -1
'  MPaintBox 25, 41, 25, 51, F5BtnColor          'F5 Button enable / disable
  
  Status$ = SPACE$(16)
  F2Ed$ = "  F2=Edit  "
  F2New$ = "  F2=New   "
  ActiveColor = 112
  InActiveColor = 115

  SELECT CASE EdMode
  CASE 0
    LSET Status$ = "Editing"
    F2BtnCap$ = F2New$
    F5BtnColor = ActiveColor

  CASE 1
    LSET Status$ = "New Transaction"
    F2BtnCap$ = F2Ed$
    F5BtnColor = InActiveColor

  END SELECT

  QPrintRC Status$, 25, 2, -1
  QPrintRC F2BtnCap$, 25, 41, -1
  MPaintBox 25, 54, 25, 64, F5BtnColor          'F5 Button enable / disable
  
END SUB

