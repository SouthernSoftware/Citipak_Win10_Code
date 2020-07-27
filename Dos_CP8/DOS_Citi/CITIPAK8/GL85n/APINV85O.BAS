DEFINT A-Z
DECLARE FUNCTION EdCheck4PO! (Form$(), AcctFlds!(), Fld() AS ANY)
DECLARE SUB InitTaxInfo (Array() AS ANY, TaxActualAdj%)
DECLARE SUB DisplayAPScrn (ScrnName$)
DECLARE SUB WPickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag%, Cnf AS ANY)
DECLARE SUB SelectInvPOs (Form$(), AcctFlds%(), Fld() AS ANY, ChangeFlag)
DECLARE SUB CancelInv ()
DECLARE SUB VoidInvTrans (VoidTransRecNum&)
DECLARE SUB IndexVendorFile ()
DECLARE SUB NewVendorRec (NewVRec)
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB WaitForAction
DECLARE SUB Post2PO (FileName$, BadTrans%)
DECLARE FUNCTION EDCheckforDupInv% (Form$(), AcctFlds(), Fld() AS ANY)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION EDCheck4Vendor% (Form$(), Terms)
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
DECLARE SUB QPrintRC (t$, R%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS Config)
  
DECLARE SUB OpenDefDistFile (VDefDist AS ANY, APDefDistFile%, NumDefRecs%)
DECLARE SUB EdCheck4DefDist (VRECNUM%, Form$(), AcctFlds%(), GotDef%)
  
DECLARE SUB EdEntry (JrnlCode%)
DECLARE SUB OpenAPEditFile (APEditFile%, NumEdTrans%, APType)
DECLARE SUB NewCJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList (APType)
DECLARE SUB DisplayTotals ()
  'DECLARE SUB PostTrans (APType)
DECLARE FUNCTION ListRecs% (APType)
DECLARE SUB ResetForm (Form$(), NumFlds%)
DECLARE SUB ShowEdMode (EdMode%)
DECLARE SUB Rec2Form (Form$(), RecordNumber%, Fld() AS ANY, APType)
DECLARE SUB SaveTrans (Form$(), RecordNumber%, EdMode%, APType)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetPostDates (LPDate, HPDate)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
  
  '$INCLUDE: 'DefCnf.bi'                    'monitor and color info
  '$INCLUDE: 'FieldInf.bi'                  'field information
  '$INCLUDE: 'FormEdit.bi'                  'form editing information
  '$INCLUDE: 'pageinfo.BI'                  'Form Page info
  '$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
  '$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
  '$INCLUDE: 'GL.BI'                        'GL File Types
  '$INCLUDE: 'SCRAUX.bi'
  '$INCLUDE: 'ap.bi'
  '$INCLUDE: 'taxinfo.bi'

  CONST False = 0, True = NOT False, NumAcctFlds = 24
  
  DIM SHARED TotLen, FundLen, AcctLen, DetLen
  DIM SHARED User$, CashAcct$, APAcct$, EncAcct$

  DIM SHARED InvTaxRec(1)  AS InvoiceTaxInfoType

  DIM SHARED TState, TCounty, StaTaxFlag, CtyTaxFlag, GotTaxFile
  
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED APEdit AS APInv85Type
  DIM SHARED VendorIdx AS VendorIdxRecType
  
  TYPE FLen
    V AS STRING * 42
  END TYPE
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  JrnlType = 2
  
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'init.bi'                      'Get GL system info
  
  TextCursor -2, -2             'Define the Mouse cursor
  
  DIM MenuText$(1 TO 5)
  MenuText$(1) = "Enter/Edit Invoices"
  MenuText$(2) = "Print Invoice Register"
  MenuText$(3) = "Post Entries"
  MenuText$(4) = "Void An Open Invoice"
  MenuText$(5) = "Quit"
  
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
  
  ShowCursor
  
  DO
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    Title$ = "Invoice Processing"
    MnuTitle Title$, Row, Col, MaxLen, 15
    
    PrintTitle User$
    
    PrintHelp "Invoice Processing"
    
    LOCATE Row, Col, 0
    VertMenu MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
    CASE 1
      ClrBG
      EdEntry JrnlType
    CASE 2
      PrnEditList JrnlType
    CASE 3
      PostInvTrans JrnlType
    CASE 4
      CancelInv
    CASE 5
      HideCursor
      CLS
      END
    CASE ELSE
    END SELECT
    
  LOOP
  
  ClrBG
  RUN "apmenu"

SUB CancelInv
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  LdRecLen = LEN(APLedgerRec(1))
  SHARED VendorIdx AS VendorIdxRecType
  SHARED VENDOR AS VendorRecType
  
  FormName$ = "VOIDINV"
  LibName$ = "AP"
  REDIM Frm(1) AS FormInfo
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  Action = 1
  FirstTime = True
  
  LibFile2Scrn LibName$, FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).FldNo > 1 AND Frm(1).PrevFld = 1 THEN             'User is moving around the form
      ChkVendor$ = QPTrim$(Form$(1, 0))
      GOSUB Check4Vendor
      IF NOT Abort THEN
        GOSUB FindInv
      ELSE
        Abort = False
      END IF
    END IF
    
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = True
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = True
    END IF
    
  LOOP UNTIL ExitFlag OR OkFlag
  
  IF ExitFlag THEN EXIT SUB
  
  IF OkFlag THEN
    temp$ = InputBox$(LibName$, "OK2VOID", Cancel%)
  END IF
  
  IF Cancel THEN
    EXIT SUB
  ELSE
    'VoidChkNum& = VAL(Form$(2, 0)) 'QPValL(Form$(2, 0))
    VoidInvTrans VoidTransRecNum&
  END IF
  
  EXIT SUB
  
  
FindInv:
  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  InvCnt = 0
  List$ = SPACE$(45)
  
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  
  GET VendorFile, VoidVendorRecNum&, VENDOR
  NextTrans& = VENDOR.FrstTran
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, APLedgerRec(1)
    IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).PAYCODE = 1 THEN
      InvCnt = InvCnt + 1
    END IF
    NextTrans& = APLedgerRec(1).NextTrans
    
  LOOP
  
  IF InvCnt = 0 THEN
    Ok = MsgBox(LibName$, "NOCHECKS")
    Frm(1).FldNo = 1
    Action = 1
    GOTO VoidInvExit
  END IF
  
  REDIM InvList(1 TO InvCnt) AS FLen            'Dim array to hold trans list
  
  MatchCnt = 0
  NextTrans& = VENDOR.FrstTran
  
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, APLedgerRec(1)
    IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).PAYCODE = 1 THEN
      MatchCnt = MatchCnt + 1
      LSET InvList(MatchCnt).V = List$
      InvList(MatchCnt).V = QPTrim$(APLedgerRec(1).DOCNum)
      MID$(InvList(MatchCnt).V, 11) = Num2Date(APLedgerRec(1).TrDate)
      MID$(InvList(MatchCnt).V, 22) = FUsing(STR$(APLedgerRec(1).Amt), "$$########,.##")
      MID$(InvList(MatchCnt).V, 39) = MKL$(NextTrans&)
    END IF
    NextTrans& = APLedgerRec(1).NextTrans
  LOOP
  
  CLOSE
  
  Row = 12
  Col = 15
  LOCATE Row, Col
  Action = 0
  MaxWidth = 37
  BoxBot = 18
  ListDescTxt$ = "P O Num       Date          Amount"
  ListDesc$ = SPACE$(MaxWidth + 4)
  MID$(ListDesc$, 3) = ListDescTxt$
  QPrintRC ListDesc$, Row - 1, Col, 112         'Cnf.ActivCh
  Choice = 1
  DO
    'call vertmenu to display matching recs to choose from
    VertMenuT InvList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      Abort = True
      Frm(1).FldNo = 1
      Action = 1
      GOTO VoidInvExit
    END IF
    IF Choice > 0 THEN
      LSET Form$(2, 0) = QPTrim$(LEFT$(InvList(Choice).V, 10))
      SaveField 2, Form$(), Fld(), BadFld
      VoidTransRecNum& = CVL(RIGHT$(InvList(Choice).V, 4))
      Frm(1).FldNo = 3
      EXIT DO
    END IF
  LOOP
  
VoidInvExit:
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  
  RETURN
  
  
Check4Vendor:
  MatchingVendors = 0
  List$ = SPACE$(65)
  
  VLookup$ = QPTrim$(UCASE$(ChkVendor$))
  VLookuplen = LEN(VLookup$)
  
  REDIM ScrBuf(2000)
  
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  
  IF NumActiveVendors = 0 THEN
    CLOSE
    ListVendors = 0
    EXIT SUB
  END IF
  
  OpenVendorFile VendorFile, NumVRecs
  
  REDIM VendorList(1 TO NumVRecs) AS FLen       'Dim array to hold trans list
  
  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
      GET VendorFile, VendorIdx.RecNum, VENDOR
      IF VENDOR.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 14) = VENDOR.VNAME
        MID$(VendorList(MatchingVendors).V, 41) = MKI$(VendorIdx.RecNum)
      END IF
    END IF
  NEXT
  
  CLOSE
  
  IF MatchingVendors = 0 THEN
    Ok = MsgBox(LibName$, "NOVENDOR")
    VendPickList = 0
    Frm(1).FldNo = 1
    Action = 1
    GOTO VoidChkReturn
  END IF
  
  IF MatchingVendors = 1 THEN
    LSET Form$(Frm(1).PrevFld, 0) = VendorList(1).V
    SaveField Frm(1).PrevFld, Form$(), Fld(), BadFld
    Action = 1
    VoidVendorRecNum& = CVI(RIGHT$(VendorList(1).V, 2))
    GOTO VoidChkReturn
  END IF
  
  IF MatchingVendors > 1 THEN
    REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
    Row = 5: Col = 15
    LOCATE Row, Col
    Action = 0
    MaxWidth = 39
    BoxBot = 18
    ListDescTxt$ = "VendorCode  Name"
    ListDesc$ = SPACE$(MaxWidth + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    Choice = 1
    DO
      'call vertmenu to display matching recs to choose from
      VertMenuT VendorList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        CLOSE
        Abort = True
        Frm(1).FldNo = 1
        Action = 1
        GOTO VoidChkReturn
      END IF
      IF Choice > 0 THEN
        VoidVendorRecNum& = CVI(RIGHT$(VendorList(Choice).V, 2))
        EXIT DO
      END IF
    LOOP
    '  MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen
  ELSE
  END IF
  
  IF Choice > 0 THEN
    LSET Form$(Frm(1).PrevFld, 0) = VendorList(Choice).V
    SaveField Frm(1).PrevFld, Form$(), Fld(), BadFld
    Choice = 0
  END IF
  
VoidChkReturn:
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  QPrintRC TName$, 3, 27, -1
  ERASE VendorList
  
  RETURN
  
END SUB

SUB EdCheck4DefDist (VRECNUM, Form$(), AcctFlds(), GotDef)
  
  LibName$ = "AP"
  FormName$ = "GETDST"
  
  REDIM Vndr(1) AS VendorRecType
  OpenVendorFile VendorFile, NumVRecs
  GET VendorFile, VRECNUM, Vndr(1)
  
  IF Vndr(1).DefDist > 0 THEN
    Amt$ = InputBox$(LibName$, FormName$, Cancel)
    IF Cancel THEN
      CLOSE
      EXIT SUB
    ELSE
      DistAmt# = VAL(Amt$)
      IF DistAmt# <> 0 THEN
        REDIM VDefDist(1) AS VendorDefDistRecType
        DefRecLen = LEN(VDefDist(1))
        OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs
        GET APDefDistFile, Vndr(1).DefDist, VDefDist(1)
        FOR D = 1 TO 8
          StartFld = AcctFlds(D)
          IF LEN(QPTrim$(VDefDist(1).DefDist(D).DefAcct)) > 0 THEN
            GotDef = D          'set the flag to the last field
            LSET Form$(StartFld, 0) = VDefDist(1).DefDist(D).DefAcct
            LSET Form$(StartFld + 1, 0) = VDefDist(1).DefDist(D).DefAcctName
            DistPct# = Round#((VDefDist(1).DefDist(D).DefPct * .01))
            AmtDist# = Round#(DistPct# * DistAmt#)
            LSET Form$(StartFld + 3, 0) = STR$(AmtDist#)
            'LSET Form$(StartFld + 3, 0) = STR$(TAPDistRec(1).DistAmt)
          ELSE
            EXIT FOR
          END IF
        NEXT
      END IF
    END IF
  END IF
  
  CLOSE
  
END SUB

FUNCTION EDCheck4Vendor% (Form$(), Terms)
  
  SHARED VendorIdx AS VendorIdxRecType
  SHARED VENDOR AS VendorRecType
  
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
  
  REDIM VendorList(1 TO 1) AS FLen              'Dim array to hold trans list
  FoundVendors = 0
  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
      FoundVendors = FoundVendors + 1
      REDIM PRESERVE VendorList(1 TO FoundVendors) AS FLen
      GET VendorFile, VendorIdx.RecNum, VENDOR
      IF VENDOR.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 13) = VENDOR.VNAME
        MID$(VendorList(MatchingVendors).V, 38) = VENDOR.GET1099
        MID$(VendorList(MatchingVendors).V, 39) = MKI$(VENDOR.VTERMS)
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
    LSET Form$(14, 0) = MID$(VendorList(MatchingVendors).V, 38)
    Terms = CVI(MID$(VendorList(1).V, 39, 2))
    LSET Form$(4, 0) = MID$(VendorList(1).V, 41, 2)
  CASE ELSE
    REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
    Row = 5: Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 36
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
        'LSET Form$(14, 0) = MID$(VendorList(Choice).V, 38, 1)
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

FUNCTION EDCheckforDupInv (Form$(), AcctFlds(), Fld() AS FieldInfo)
  
  NumFlds = UBOUND(Form$)
  GotOne = False
  
  REDIM TVendor(1)       AS VendorRecType
  REDIM TAPLedgerRec(1)  AS APLedger81RecType
  REDIM TAPDistRec(1)    AS APDistRecType
  
  DistRecLen = LEN(TAPDistRec(1))
  LedgerRecLen = LEN(TAPLedgerRec(1))
  
  VRECNUM = CVI(LEFT$(Form$(FldNum("VRECNUM", Fld()), 0), 2))
  INVNUM$ = QPTrim$(Form$(5, 0))
  
  OpenVendorFile VendorFile, NumVRecs
  GET VendorFile, VRECNUM, TVendor(1)
  CLOSE
  
  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen
  
  NextTrans& = TVendor(1).FrstTran
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, TAPLedgerRec(1)
    'TAPLedgerRec(1)
    IF TAPLedgerRec(1).TrCode = 1 AND QPTrim$(TAPLedgerRec(1).DOCNum) = INVNUM$ THEN
      GotOne = True
      EXIT DO
    END IF
    NextTrans& = TAPLedgerRec(1).NextTrans
  LOOP
  
  CLOSE
  
  EDCheckforDupInv = GotOne
  
END FUNCTION

SUB EdEntry (APType)

  REDIM PRESERVE TaxDist(1 TO 2)  AS DistSumType

  InitTaxInfo TaxDist(), TaxActualAdj

  TaxRecLen = LEN(InvTaxRec(1))
  
  Adding = 1
  Editing = 0

  GetPostDates LPDate, HPDate

  SHARED VendorIdx AS VendorIdxRecType
  SHARED VENDOR AS VendorRecType

  List$ = SPACE$(65)
  CommaFormat$ = "#########,.##"
  GotPOFlag = 0
  
  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum

  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType

  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs
  
  '--get first active record number
  CurrRec = 0
  NextRec = 1
  OpenAPEditFile APEditFile, NumEdTrans, APType
  RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
  CLOSE APEditFile
  
  IF RecNum = 0 THEN
    EdMode = Adding
  ELSE
    EdMode = Editing
  END IF
  '===
  
Top:
  
'if the field count changes MUST Adjust "Header" below
  Header = 24

  NumOfScrns = 3
  
  REDIM Frm(1)    AS FormInfo
  REDIM FormName$(1 TO NumOfScrns)
  FOR Scr = 1 TO NumOfScrns
    FormName$(Scr) = "AP" + LTRIM$(STR$(Scr))
  NEXT
  
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
  TCnt = 0
  REDIM AcctFlds(1 TO NumAcctFlds)
  REDIM TaxFlds(1 TO 2)
  FOR F = 1 TO NumFlds
    IF QPTrim$(Fld(F).ScratchS) = "A" THEN
      Cnt = Cnt + 1
      AcctFlds(Cnt) = F
    END IF
    IF QPTrim$(Fld(F).ScratchS) = "T" THEN
      TCnt = TCnt + 1
      TaxFlds(TCnt) = F
    END IF

  NEXT

  DoTaxFlag = False
  Frm(1).PageDisable = True
  Frm(1).StayOnField = True
  
'******* Field Numbers
  POFldNum = FldNum("PONUM", Fld())
  TotalAmtFld = FldNum("INVAMT", Fld())
  DueDateFld = FldNum("DUEDATE", Fld())
  TaxFlagFld = FldNum("TAXYN", Fld())
  TaxTotFld = FldNum("TAXTOTAL", Fld())
  InvNumbFld = FldNum("INVNUM", Fld())
  InvDateFld = FldNum("INVDATE", Fld())
  DistDateFld = FldNum("DISTDATE", Fld())
  PayCodeFld = FldNum("PAYCODE", Fld())
  PORecNumFld = FldNum("PORECNUM", Fld())
  GotPOFlagFld = FldNum("GOTPOFLG", Fld())
'*********************
  
  Scr = 1       'Set screen number to one
  LibFile2Scrn "AP", FormName$(Scr), MonoCode, -1, ErrCode      'display screen
  
  IF EdMode = Editing THEN
    Rec2Form Form$(), RecNum, Fld(), APType
    GotPOFlag = VAL(Form$(GotPOFlagFld, 0))     'po active flag
    OldVend$ = QPTrim$(Form$(2, 0))
    OPONum$ = QPTrim$(Form$(POFldNum, 0))
  END IF
  
  ShowEdMode EdMode
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF (Frm(1).KeyCode = -72) OR (Frm(1).KeyCode = -15) THEN
      FOR zz = Frm(1).FldNo TO 1 STEP -1   'loop backwards
        IF Fld(zz).Protected = False THEN
          Frm(1).FldNo = zz
          Action = 1
          EXIT FOR
        END IF
      NEXT
    END IF

    IF FirstTime THEN
      'Fld(90).Protected = True
      GOSUB GetInvTotal
      FirstTime = False
      GotPOFlag = 0
      LSET Form$(PayCodeFld, 0) = "1"
      IF EdMode = Adding AND LEN(KeepDate$) > 0 THEN
        LSET Form$(InvDateFld, 0) = KeepDate$
        SaveField InvDateFld, Form$(), Fld(), BadFld
      END IF
      IF NOT GotTaxFile THEN
        LSET Form$(TaxFlagFld, 0) = "N"
        LSET Form$(TaxTotFld, 0) = "0.00"
        Fld(TaxFlagFld).Protected = True
      ELSEIF EdMode = Adding THEN
        LSET Form$(TaxFlagFld, 0) = "Y"
        TaxZeroFlag = False
      ELSE

        IF Form$(TaxFlagFld, 0) = "N" THEN
          TaxZeroFlag = True
        END IF
        GOSUB FixTaxDist
      END IF
      Action = 2
    END IF
    
    'IF InvChanged OR (Frm(1).FldNo <> TotalAmtFld AND Frm(1).PrevFld = TotalAmtFld) THEN
    IF InvChanged OR (Frm(1).FldNo <> Frm(1).PrevFld) THEN
      TmpInvAmt# = Round#(Value(Form$(TotalAmtFld, 0), ECode))
      IF TmpInvAmt# <> InvTotalAmt# THEN
        InvTotalAmt# = TmpInvAmt#
        IF GotTaxFile THEN
          GOSUB FixTaxDist
        END IF
      END IF
    END IF

    IF Frm(1).FldNo <> InvNumbFld AND Frm(1).PrevFld = InvNumbFld THEN
      'if the field just exited was the invoice number field
      GOSUB Check4DupeInvoice
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
'here
    IF Frm(1).FldNo <> TaxFlagFld AND Frm(1).PrevFld = TaxFlagFld THEN
      IF GotTaxFile THEN
        SELECT CASE Form$(TaxFlagFld, 0)
        CASE "Y"
          TaxZeroFlag = False
        CASE ELSE
          TaxZeroFlag = True
        END SELECT
        GOSUB FixTaxDist
        Action = 1
      END IF
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F4KEY
      IF Frm(1).FldNo = POFldNum THEN           'if there in the PO field
        SelectInvPOs Form$(), AcctFlds(), Fld(), ChangeFlag
        Action = 1
        IF ChangeFlag THEN
          Frm(1).FldNo = Frm(1).FldNo + 1
          IF GotTaxFile THEN
            GOSUB FixTaxDist
          END IF
        END IF
      END IF
    CASE -106   'Alt+F3
      IF Frm(1).FldNo > Header THEN             'if we're on an account field
        GOSUB DelInvDist
      END IF
    CASE F2Key  'Change Editing Mode

      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        SELECT CASE EdMode
        CASE Editing            ' were in Edit Mode
          EdMode = Adding       ' switch to add mode
          ShowEdMode EdMode
          ResetForm Form$(), NumFlds
          FirstTime = True
          Frm(1).FldNo = 1
          Action = 1
          GOSUB ShowDistTot
        CASE Adding             'currently adding records
          FirstTime = True
          'check to see if there are any active transactions
          OpenAPEditFile APEditFile, NumEdTrans, APType
          CurrRec = 0
          NextRec = 1
          RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE APEditFile
          IF RecNum = 0 THEN    'No active records can't edit
            EdMode = Adding
          ELSE
            EdMode = Editing    'switch to edit mode
            ShowEdMode EdMode
            RecNum = ListRecs(APType)
            IF RecNum > 0 THEN
              Rec2Form Form$(), RecNum, Fld(), APType
              GotPOFlag = VAL(Form$(GotPOFlagFld, 0))           'po active flag
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
      
    CASE F7KEY  'Add new vendor
      IF Frm(1).FldNo = 2 THEN
        REDIM tempScrn2(2000)
        MScrnSave 1, 1, 25, 80, tempScrn2(0)
        NewVendorRec NVRecNum
        IF NVRecNum > 0 THEN
          OpenVendorFile VendorFile, NumVRecs
          GET VendorFile, NVRecNum, VENDOR
          CLOSE VendorFile
          LSET Form$(2, 0) = VENDOR.VNum
          LSET Form$(3, 0) = VENDOR.VNAME
          LSET Form$(4, 0) = MKI$(NVRecNum)
          Terms = VENDOR.VTERMS
          SaveField 2, Form$(), Fld(), BadFld
          Frm(1).FldNo = Frm(1).FldNo + 1         'was 5
          Action = 1
        ELSE
          Frm(1).FldNo = 1
          Action = 2
        END IF
        MScrnRest 1, 1, 25, 80, tempScrn2(0)
        Action = 1
      END IF

    CASE ELSE
      
    END SELECT
      IF Frm(1).FldNo = 2 AND Frm(1).PrevFld <> 2 THEN            'if there in the PO field
        MPaintBox 25, 56, 25, 65, 112
      ELSEIF Frm(1).FldNo <> 2 AND Frm(1).PrevFld = 2 THEN            'if there in the PO fieldFrm(1).FldNo = 2 AND Frm(1).PrevFld = 2 THEN            'if there in the PO field
        MPaintBox 25, 56, 25, 65, 115
      END IF
    
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN      'User is moving around the form
      IF Frm(1).FldNo = POFldNum THEN           'if there in the PO field
        MPaintBox 22, 54, 22, 62, 112
      ELSE
        MPaintBox 22, 54, 22, 62, 115
      END IF
      
      IF Frm(1).FldNo = 15 AND Frm(1).PrevFld = 14 THEN
        IF GotDef = False THEN
          EdCheck4DefDist CVI(Form$(4, 0)), Form$(), AcctFlds(), GotDef
          IF GotDef > 0 THEN
            Frm(1).FldNo = AcctFlds(1) + (GotDef * 4)
          END IF
        END IF
      END IF
      
'Vendor process block   '061098  Debuged checked out OK
      IF (Frm(1).PrevFld = 2 AND Frm(1).FldNo <> 2) THEN
        ThisVend$ = QPTrim$(Form$(2, 0))        'get vendor field
        IF LEN(ThisVend$) > 0 THEN              'if there is one
          IF ThisVend$ <> OldVend$ THEN         'if it's not the same
            IF EDCheck4Vendor(Form$(), Terms) THEN              'check for valid vendor
              SaveField 1, Form$(), Fld(), BadFld               'if it was save to form
              Frm(1).FldNo = Frm(1).FldNo + 2   'skip over vend disc field
              Action = 1        'set action to update
              OldVend$ = QPTrim$(Form$(2, 0))   'store vendor after lookup
            ELSE                'not a valid vendor
              Frm(1).FldNo = 1  'back to vendor field
              Action = 2        '
            END IF
          END IF                'vendor not changed, do nothing
        ELSE    'no vendor code entered
          Frm(1).FldNo = 1      'back to vendor field
          Action = 2            'set action to resume
        END IF
      END IF
'Vendor process END

      GOSUB ShowDistTot         'Update Distribution totals
      
      'IF Frm(1).FldNo > AcctFlds(TaxActualAdj) + 3 THEN
      '--Check the field we're in to see if it is an acct field.
      '  if so and if there is data in the field, remove dashes for user
      '  set stripped acct string in form
      
        IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
          LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
          Action = 1
        END IF

        '--the following checks the field we have just left to see
        '  if it is an account field, and '  if the field is not blank,
        '  the acct number is formatted with dashes
        '  if the field is blank the acct name and distribution amt are blanked
        '  and the amt field is protected to prevent user from entering a
        '  distribution without an account.
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
        '  we have a valid acct number
        IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
          IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN     '= AcctField
            GOSUB CheckForValidAcctNum
          END IF
        END IF
      'END IF
      
    END IF      'End of test for PrevFld <> current FldNo
    
    IF SaveFlag THEN GOSUB SaveRec
    
    SELECT CASE Frm(1).KeyCode      '--Process key strokes
    CASE F3Key
      IF EdMode = 0 THEN
        Ok = MsgBox%("GL.QSL", "OK2DELTR")
        IF Ok = 2 THEN          'Pressed Button #2
          Form$(1, 0) = "-1"
          SaveField 1, Form$(), Fld(), BadFld
          SaveTrans Form$(), RecNum, EdMode, APType
          ResetForm Form$(), NumFlds
          RecNum = ListRecs(APType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), APType
            GotPOFlag = VAL(Form$(GotPOFlagFld, 0))             'po active flag
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
        IF EdMode = Editing THEN
          RecNum = ListRecs(APType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), APType
            'GotPOFlag = VAL(Form$(NumFlds, 0))        'po active flag
            GotPOFlag = VAL(Form$(GotPOFlagFld, 0))
            OldVend$ = QPTrim$(Form$(2, 0))     'store vendor
            Frm(1).FldNo = 1
            Action = 1
          END IF
        END IF
      END IF
      
    CASE F10Key
      '--Make sure acct fields are formatted correctly if user
      '  Pressed F10 with cursor in an account field

      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
        IF LEN(QPTrim$(Form$(Frm(1).FldNo, 0))) > 0 THEN
          TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).FldNo, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).FldNo, 0) = TempAcct$
          SaveField Frm(1).FldNo, Form$(), Fld(), BadFld        'Write the flag to the form$
        END IF
      END IF
      SaveFlag = True
      
    CASE CtrlL  'Ctrl+LeftArrow               'Get the next record
      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          EdMode = Editing
          Frm(1).FldNo = 1
          OpenAPEditFile APEditFile, NumEdTrans, APType
          CurrRec = RecNum
          NextRec = -1
          RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE APEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), APType
          GotPOFlag = VAL(Form$(GotPOFlagFld, 0))               'po active flag
          GOSUB ShowDistTot
          ShowEdMode EdMode
          Scr = 1
          Action = 1
          OldVend$ = QPTrim$(Form$(2, 0))       'store vendor
        END IF
      END IF
      
    CASE CtrlR  'Ctrl+RightArrow             'Get the previous record
      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          EdMode = Editing
          Frm(1).FldNo = 1
          OpenAPEditFile APEditFile, NumEdTrans, APType
          CurrRec = RecNum
          NextRec = 1
          RecNum = GetNextRec(APEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE APEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), APType
          GotPOFlag = VAL(Form$(GotPOFlagFld, 0))               'po active flag
          GOSUB ShowDistTot
          ShowEdMode EdMode
          Action = 1
          Scr = 1
          OldVend$ = QPTrim$(Form$(2, 0))       'store vendor
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
      CASE ELSE

      END SELECT
      Action = 1
      '  Action = 1
      LibFile2Scrn "AP", FormName$(Scr), MonoCode, -2, ErrCode

      GOSUB ShowDistTot
      ShowEdMode EdMode
    END IF
    
    '--Buttons are not attached to the form, the following processes
    '--the rodent
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21
        SELECT CASE Frm(1).MCol
        CASE 7 TO 18
          PressButton ESC, 21, 7, 18
        CASE 19 TO 32
          PressButton F3Key, 21, 19, 32
        CASE 33 TO 49
          PressButton -106, 21, 33, 49          'Alt-F3
        CASE 50 TO 63
          PressButton F4KEY, 21, 50, 63
        CASE 64 TO 73
          PressButton F10Key, 21, 64, 73
        END SELECT
        
      CASE 25
        SELECT CASE Frm(1).MCol
        CASE 34 TO 42
          PressButton F2Key, 25, 34, 42
        CASE 43 TO 51
          PressButton F5KEY, 25, 43, 51
        CASE 52 TO 66
          PressButton F7KEY, 25, 52, 66
        CASE 67 TO 69
          PressButton CtrlL, 25, 67, 69
        CASE 70 TO 72
          PressButton CtrlR, 25, 70, 72
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
  'The EDCheck4Acct function this gosub replaces, was loading the vendor index
  'file everytime it checked an account. Should help speed this up
  
  OpenAcctFile AcctFileNum, NumAccts
  
  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, Acct
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).V = List$
      AcctList(MatchAccts).V = TempAcctIdx(Cnt).AcctNum
      MID$(AcctList(MatchAccts).V, TotLen + 4) = Acct.Title
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
  GET AcctFileNum, AcctRecNum, Acct
  CLOSE AcctFileNum
  LSET Form$(AcctFld, 0) = Acct.Num
  LSET Form$(AcctFld + 1, 0) = Acct.Title
  LSET Form$(AcctFld + 2, 0) = MKI$(AcctRecNum)
RETURN
  
  
ShowDistTot:
  TotDist# = 0
  FOR Cnt = AcctFlds(1) TO AcctFlds(NumAcctFlds)    'look thru acct fld list
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
  QPrintRC UnDist$, 21, 24, -1
  QPrintRC Dist$, 21, 56, -1
RETURN
  
SaveRec:
  BadFld = 0
  
  'Check invoice date fields
  FOR DFld = 11 TO 13
    D = Date2Num(Form$(DFld, 0))
    IF D < LPDate OR D > HPDate THEN
      BadFld = DFld
      EXIT FOR
    END IF
  NEXT
  
  IF BadFld = 0 THEN            'if no blanks, make sure
    GOSUB ShowDistTot           'distribution are in bal
    IF UnDist# = 0 AND Amt# <> 0 THEN
      '--Set deleted flag to false
      Form$(1, 0) = "0"
      SaveField 1, Form$(), Fld(), BadFld       'Write the flag to the form$
      
      '--Make sure the farkin po flag gets set to false and NOT -32767!!!
      IF Form$(GotPOFlagFld, 0) <> "-1" THEN
        Form$(GotPOFlagFld, 0) = "0"
        SaveField GotPOFlagFld, Form$(), Fld(), BadFld          'Write the flag to the form$
      END IF
      
      SaveTrans Form$(), RecNum, EdMode, APType 'Save the transaction
      
      KeepDate$ = Form$(InvDateFld, 0)
      GotDef = False
      Ok = MsgBox%("GL.QSL", "SAVED")           'tell user we saved it
      IF EdMode = Adding THEN   'if in Add Mode, clear
        SaveFlag = False
        'EditedFlag = False
        GOTO Top
      END IF
      Action = 1
      Scr = 1
    ELSE
      Ok = MsgBox%("GL.QSL", "BADDIST")         'Distributions out of balance
      Action = 1
    END IF
  ELSE
    Ok = MsgBox%("GL.QSL", "BADDATE")           'Tell user no blank fields
    Frm(1).FldNo = BadFld
    
  END IF
  
  Action = 1
  FirstTime = True
  SaveFlag = False
  'EditedFlag = False
  IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN             '= AcctField
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
  GotDef = False                '--Flag for Default Distribution
  Action = 1
  RETURN
  
Check4DupeInvoice:
  IF EDCheckforDupInv(Form$(), AcctFlds(), Fld()) THEN
    'mod to allow use of dupe if OK
    Ok = MsgBox%("AP.QSL", "DUPEINV2")
    IF Ok = 2 THEN
      Frm(1).FldNo = 4
      Action = 1
    ELSE
      Frm(1).FldNo = 6
      Action = 1
    END IF
  END IF
  
  OpenAPEditFile APEditFile, NumEdTrans, APType
  NumofEditRecs = LOF(APEditFile) / LEN(APEdit)
  FOR CheckCnt = 1 TO NumofEditRecs
    GET APEditFile, CheckCnt, APEdit            'write it
    IF APEdit.DELFLAG = 0 THEN
    IF APEdit.INVNUM = Form$(5, 0) AND EdMode = Adding AND VAL(APEdit.VRECNUM) = VAL(Form$(4, 0)) THEN
      Form$(5, 0) = ""
      Ok = MsgBox%("AP.QSL", "DUPEINV")
      Frm(1).FldNo = 4
      Action = 1
    END IF
    END IF
  NEXT CheckCnt
  CLOSE APEditFile
  RETURN

DelInvDist:
  'for each accounting distribution
  FOR Cnt = 1 TO NumAcctFlds
    'if this acct field is less than the current field
    'set the row counter
    IF AcctFlds(Cnt) <= Frm(1).FldNo THEN
      FldRow = FldRow + 1       '+AcctFlds(Cnt)
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
  RETURN

GetInvTotal:
  InvTotalAmt# = Value#(Form$(TotalAmtFld, 0), ErrCode)
RETURN

FixTaxDist:

  TaxTotal# = 0
  
    TInvAmt# = Round#(Value(Form$(TotalAmtFld, 0), ECode))
    IF StaTaxFlag THEN
      IF TaxZeroFlag THEN
        LSET Form$(TaxFlds(1), 0) = TaxDist(1).DistAcctNum
        LSET Form$(TaxFlds(1) + 1, 0) = TaxDist(1).AcctTitle
        TaxDist(1).DistAmt = 0
        TaxTotal# = 0
        LSET Form$(TaxFlds(1) + 3, 0) = FUsing$(STR$(TaxTotal#), "####.##")
      ELSE
        TCalc# = InvTaxRec(1).InvTax(1).TaxAmt
        IF TCalc# > 0 THEN
          StaTax# = Round#(TInvAmt# * (TCalc# / 100))
          LSET Form$(TaxFlds(1), 0) = TaxDist(1).DistAcctNum
          LSET Form$(TaxFlds(1) + 1, 0) = TaxDist(1).AcctTitle
          TaxDist(1).DistAmt = StaTax#
          TaxTotal# = Round#(TaxTotal# + StaTax#)
          LSET Form$(TaxFlds(1) + 3, 0) = FUsing$(STR$(StaTax#), "####.##")
        END IF
    END IF
    IF CtyTaxFlag THEN
      IF TaxZeroFlag THEN
        LSET Form$(TaxFlds(TaxActualAdj), 0) = TaxDist(TaxActualAdj).DistAcctNum
        LSET Form$(TaxFlds(TaxActualAdj) + 1, 0) = TaxDist(TaxActualAdj).AcctTitle
        TaxDist(TaxActualAdj).DistAmt = 0
        TaxTotal# = 0
        Form$(TaxFlds(TaxActualAdj) + 3, 0) = FUsing$(STR$(TaxTotal#), "####.##")
      ELSE
        TCalc# = InvTaxRec(1).InvTax(2).TaxAmt
        IF TCalc# > 0 THEN
          CtyTax# = Round#(TInvAmt# * (TCalc# / 100))
          LSET Form$(TaxFlds(TaxActualAdj), 0) = TaxDist(TaxActualAdj).DistAcctNum
          LSET Form$(TaxFlds(TaxActualAdj) + 1, 0) = TaxDist(TaxActualAdj).AcctTitle
          TaxDist(TaxActualAdj).DistAmt = CtyTax#
          TaxTotal# = Round#(TaxTotal# + CtyTax#)
          Form$(TaxFlds(TaxActualAdj) + 3, 0) = FUsing$(STR$(CtyTax#), "####.##")
        END IF
      END IF
    END IF
    'IF (TaxTotal# > 0) OR ZeroTaxFlag THEN
      Form$(TaxTotFld, 0) = FUsing$(STR$(TaxTotal#), "####.##")
    'END IF
  END IF
  Action = 1

RETURN


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
      BEEP
      Found = 0 'if no more records then get out
      EXIT DO
    END IF
    
    GET FileNum, rec, APEdit    'Get the record
    
    IF APEdit.DELFLAG <> True THEN              'Ok if not deleted
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

SUB InitTaxInfo (TaxDist() AS DistSumType, TaxActualAdj)

  TaxRecLen = LEN(InvTaxRec(1))
  TaxFile = FREEFILE
  OPEN "APINVTAX.DAT" FOR RANDOM AS TaxFile LEN = TaxRecLen

  IF LOF(TaxFile) > 0 THEN
    GET TaxFile, 1, InvTaxRec(1)
    IF (LEN(QPTrim$(InvTaxRec(1).InvTax(1).AcctNo)) > 0) AND (InvTaxRec(1).InvTax(1).TaxAmt > 0) THEN
      TState = 1
      StaTaxFlag = True
      GotTaxFile = True
    END IF
    IF (LEN(QPTrim$(InvTaxRec(1).InvTax(2).AcctNo)) > 0) AND (InvTaxRec(1).InvTax(2).TaxAmt > 0) THEN
      CtyTaxFlag = True
      GotTaxFile = True
      IF TState = 1 THEN
        TCounty = 2
      ELSE
        TCounty = 1
      END IF
    END IF
  END IF
  CLOSE TaxFile

  IF GotTaxFile THEN
    IF StaTaxFlag THEN
      TaxActualAdj = TaxActualAdj + 1
      STAcctRec = FindAcct(QPTrim$(InvTaxRec(1).InvTax(1).AcctNo))
      REDIM PRESERVE TaxDist(1 TO TaxActualAdj)  AS DistSumType
      TaxDist(TaxActualAdj).DistAcctNum = InvTaxRec(1).InvTax(1).AcctNo
      TaxDist(TaxActualAdj).AcctTitle = GetAcctTitle$(STAcctRec)
    END IF
    IF CtyTaxFlag THEN
      TaxActualAdj = TaxActualAdj + 1
      CTAcctRec = FindAcct(QPTrim$(InvTaxRec(1).InvTax(2).AcctNo))
      REDIM PRESERVE TaxDist(1 TO TaxActualAdj)  AS DistSumType
      TaxDist(TaxActualAdj).DistAcctNum = InvTaxRec(1).InvTax(2).AcctNo
      TaxDist(TaxActualAdj).AcctTitle = GetAcctTitle$(CTAcctRec)
    END IF
  END IF

END SUB

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
    
    '    FOR zz = 1 TO 6
    '      APEdit.PORecs(zz) = 0
    '    NEXT
    '    PUT APEditFile, Cnt, APEdit
    
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
  
  REDIM PO2Post(1 TO 1) AS GLTransRecType
  
  POIFFile = FREEFILE
  
  IFLen = LEN(PO2Post(1))
  
  OPEN FileName$ FOR RANDOM AS POIFFile LEN = IFLen
  Num2Get = LOF(POIFFile) \ IFLen
  REDIM PO2Post(1 TO Num2Get)  AS GLTransRecType
  FOR Cnt = 1 TO Num2Get
    GET POIFFile, Cnt, PO2Post(Cnt)
  NEXT
  CLOSE
  
  FOR Cnt = 1 TO Num2Get
    PO2Post(Cnt).AcctNum = LEFT$(PO2Post(Cnt).AcctNum, FundLen) + EncAcct$
    
  NEXT
  
  'Tr2Post(1).Ref = APEdit.InvNum
  'Tr2Post(1).DrAmt = 0
  'Tr2Post(1).CrAmt = TrFundSum#(Fund)
  'Tr2Post(1).Src = "AP" + ConvDateStr$(DATE$)
  
  
  
  
  
  
  
  
  
  
  
  
  
  
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
    LSET ToPrint$ = ConvDateStr1(Num2Date$(Trans2Post.TrDate))
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
  
  REDIM VENDOR(1) AS VendorRecType
  
  REDIM TempVendorRec(1) AS VendorRecType
  VENDOR(1) = TempVendorRec(1)
  
  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  REDIM Frm(1) AS FormInfo      'Form editing info
  
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
        VRECNUM = FindVendorRec(QPTrim$(Form$(1, 0)))
        IF VRECNUM > 0 THEN
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
  
  VENDOR(1).VNum = Form$(1, 0)
  VENDOR(1).VNAME = Form$(2, 0)
  VENDOR(1).Addr1 = Form$(3, 0)
  VENDOR(1).City = Form$(5, 0)
  VENDOR(1).State = Form$(6, 0)
  VENDOR(1).Zip = Form$(7, 0)
  VENDOR(1).PayToName = Form$(8, 0)
  VENDOR(1).PayToAddr = Form$(9, 0)
  VENDOR(1).PayToCity = Form$(11, 0)
  VENDOR(1).PaytoState = Form$(12, 0)
  VENDOR(1).PaytoZip = Form$(13, 0)
  VENDOR(1).Addr2 = Form$(4, 0)
  VENDOR(1).PaytoAddr2 = Form$(10, 0)
  VENDOR(1).VIN = GetNewVendorPIN
  
  
  VENDOR(1).phone = Form$(20, 0)
  
  VENDOR(1).FedID = Form$(15, 0)
  VENDOR(1).CoCode = Form$(16, 0)
  VENDOR(1).StCode = Form$(17, 0)
  VENDOR(1).DELFLAG = 0
  VENDOR(1).GET1099 = Form$(18, 0)
  VENDOR(1).Fax = Form$(21, 0)
  VENDOR(1).VTERMS = Value(Form$(14, 0), 0)
  VENDOR(1).YTDPay = 0
  VENDOR(1).CurrBal = 0
  VENDOR(1).FrstTran = 0
  VENDOR(1).LastTran = 0
  VENDOR(1).FrstPO = 0
  VENDOR(1).LastPO = 0
  VENDOR(1).DefDist = 0
  VENDOR(1).Contact = Form$(19, 0)
  
  
  
  
  
  
  
  
  
  
  
  NewVRecNum = NumVRecs + 1
  
  PUT VendorFile, NewVRecNum, VENDOR(1)
  
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
  
  SHARED APAcct$
  SHARED VENDOR AS VendorRecType
  
  REDIM APDistRec(1) AS APDistRecType
  REDIM APLedgerRec(1) AS APLedger81RecType
  REDIM TR2Post(1) AS GLTransRecType
  
  LedgerRecLen = LEN(APLedgerRec(1))
  DistRecLen = LEN(APDistRec(1))
  VendorRecLen = LEN(VENDOR)
  
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

  IF Active = 0 THEN '--No active transactions - tell user and get out
    Ok = MsgBox%("GL.QSL", "NOTRANS")
    EXIT SUB
  END IF

  Ok = MsgBox%("GL.QSL", "POSTGJ")   '--make sure we're ready to post
  IF Ok = 1 THEN EXIT SUB            '1=Cancel Button
  
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM TrFundSum#(1 TO NumFunds)
  
  AP2Post = FREEFILE
  OPEN "APINVIF.dat" FOR RANDOM AS AP2Post LEN = LEN(TR2Post(1))
  
  PO2Post = FREEFILE
  OPEN "POINVIF.dat" FOR RANDOM AS PO2Post LEN = LEN(TR2Post(1))
  
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
  FOR Cnt = 1 TO NumEdTrans '--Post transaction to A/P Ledger, update links
    GET APEditFile, Cnt, APEdit
    IF NOT APEdit.DELFLAG THEN
      VendRecNum = CVI(APEdit.VRECNUM)
      PrintHelp "Posting Vendor: " + STR$(VendRecNum)
      IF VendRecNum > 0 THEN
        PostCnt = PostCnt + 1
        NumLedgerRecs = NumLedgerRecs + 1
        '--update vendor to transaction link
        OpenVendorFile VendorFile, NumVRecs
        GET VendorFile, VendRecNum, VENDOR
        PrevVendTrans = VENDOR.LastTran
        FrstVendTrans = VENDOR.FrstTran
        IF FrstVendTrans = 0 THEN
          VENDOR.LastTran = NumLedgerRecs
          VENDOR.FrstTran = NumLedgerRecs
        ELSE
          VENDOR.LastTran = NumLedgerRecs
          GET APLedgerFile, PrevVendTrans, APLedgerRec(1)
          APLedgerRec(1).NextTrans = NumLedgerRecs
          PUT APLedgerFile, PrevVendTrans, APLedgerRec(1)
        END IF
        PUT VendorFile, VendRecNum, VENDOR
        CLOSE VendorFile
'--write transaction to apledger
        APLedgerRec(1).VRECNUM = VendRecNum
        APLedgerRec(1).VendorCode = APEdit.VENDOR
        APLedgerRec(1).TrDate = APEdit.INVDATE
        APLedgerRec(1).DOCNum = APEdit.INVNUM
        APLedgerRec(1).PONUM = APEdit.PONUM
        APLedgerRec(1).DUEDATE = APEdit.DUEDATE
        APLedgerRec(1).TrCode = 1
        APLedgerRec(1).PAYCODE = VAL(APEdit.PAYCODE)
        APLedgerRec(1).GLDistDate = APEdit.DISTDATE

'071398 'added Tax amount into invoice total
        'APLedgerRec(1).Amt = APEdit.INVAMT
        'APLedgerRec(1).Amt = Round#(APEdit.INVAMT + APEdit.TaxTotal)
'071398
        APLedgerRec(1).NextTrans = 0
        APLedgerRec(1).FrstDist = NumDistRecs& + 1
        APLedgerRec(1).GET1099 = APEdit.GET1099
        APLedgerRec(1).Comment = APEdit.INVDESC   '--New File format
        APLedgerRec(1).PSLFlag = APEdit.PSLFlag
        APLedgerRec(1).TaxAmt = APEdit.TaxTotal
'TAX***************
'        FOR DistCnt = 1 TO 2   'LastDist
'          TaxAcct$ = QPTrim$(APEdit.ITax(DistCnt).DACN)
'          IF LEN(TaxAcct$) > 0 AND APEdit.ITax(DistCnt).DAMT <> 0 THEN
'            REDIM APDistRec(1) AS APDistRecType
'            AcctRec = FindAcct(TaxAcct$)
'            NumDistRecs& = NumDistRecs& + 1
'            APDistRec(1).APLedgerRec = NumLedgerRecs
'            APDistRec(1).DistAcctRec = AcctRec
'            APDistRec(1).DistAcctNum = APEdit.ITax(DistCnt).DACN
'            APDistRec(1).DistAmt = APEdit.ITax(DistCnt).DAMT
'            APDistRec(1).NextDist = NumDistRecs& + 1
'            PUT APDistFile, NumDistRecs&, APDistRec(1)
'          END IF
'        NEXT
'TAX***************
        FOR DistCnt = 1 TO 24   'LastDist
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
        APDistRec(1).NextDist = 0    'Last distrubtion for this invoice
        PUT APDistFile, NumDistRecs&, APDistRec(1)
        '--update the last distribution pointer in apledger.dat
        APLedgerRec(1).LastDist = NumDistRecs&
        PUT APLedgerFile, NumLedgerRecs, APLedgerRec(1)
'Changed to look for up to "6" PO's
        FOR POCnt = 1 TO 6                '--Set the PO's Flag to closed
          IF LEN(QPTrim$(APEdit.PORecs(POCnt))) > 0 THEN
            GotPOs = True
            ThisPO$ = APEdit.PORecs(POCnt)
            WhatPO& = CVL(ThisPO$)
            GET APLedgerFile, WhatPO&, APLedgerRec(1)
            APLedgerRec(1).TrCode = -4
            PUT APLedgerFile, WhatPO&, APLedgerRec(1)
          END IF
        NEXT
      ELSE  '--could not find vendor, Mark trans as deleted and log it.
        BadVendor = BadVendor + 1
        APEdit.DELFLAG = True
        PUT APEditFile, Cnt, APEdit
        GOSUB LogBadTrans
      END IF                    '--test for good vendor
    END IF                      '--test for not deleted trans
  NEXT
  '--All Done with post to apledger & apdist
  IF BadVendor = 0 THEN
    PRINT #LogFile, "No Posting Errors. Transactions Posted: "; PostCnt
  END IF
  FOR Cnt = 1 TO NumEdTrans '--Create interface files for gltrans & potrans
    REDIM TrFundSum#(1 TO NumFunds)
    '--Reinitialize transaction fund summary elements for next loop
    GET APEditFile, Cnt, APEdit
    IF NOT APEdit.DELFLAG THEN  '--If invoice not deleted post it
      GotPOs = False
      FOR POCnt = 1 TO 6  'See if this Invoice had a PO
        IF LEN(QPTrim$(APEdit.PORecs(POCnt))) > 0 THEN
          GotPOs = True
          EXIT FOR
        END IF
      NEXT
'****************************
'      FOR AcctDist = 1 TO 2   'LastDist
'        TaxAcct$ = QPTrim$(APEdit.ITax(AcctDist).DACN)
'        IF LEN(TaxAcct$) > 0 AND APEdit.ITax(AcctDist).DAMT <> 0 THEN
'          RecordNum = RecordNum + 1
'          TR2Post(1).AcctNum = APEdit.ITax(AcctDist).DACN
'          TR2Post(1).TrDate = APEdit.DISTDATE
'          TR2Post(1).Desc = APEdit.VENDNAME
'          TR2Post(1).Ref = APEdit.INVNUM
'          TR2Post(1).DrAmt = APEdit.ITax(AcctDist).DAMT
'          TR2Post(1).CrAmt = 0
'          TR2Post(1).Src = MSrc$                '"AP" + ConvDateStr$(DATE$)
'          TR2Post(1).Marked = False
'          PUT AP2Post, RecordNum, TR2Post(1)
'          FOR Fund = 1 TO NumFunds '--Add distribution to the proper fund
'            FundNum$ = LEFT$(APEdit.ITax(AcctDist).DACN, FundLen)
'            IF FundNum$ = FundList$(Fund) THEN
'              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(APEdit.ITax(AcctDist).DAMT)
'              EXIT FOR
'            END IF
'          NEXT

'        END IF
'      NEXT
'******************************
      FOR AcctDist = 1 TO 24    '--Write only those with account numbers
        IF LEN(QPTrim$(APEdit.Dist(AcctDist).DACN)) > 0 THEN
          RecordNum = RecordNum + 1
          TR2Post(1).AcctNum = APEdit.Dist(AcctDist).DACN
          TR2Post(1).TrDate = APEdit.DISTDATE
          TR2Post(1).Desc = APEdit.VENDNAME
          TR2Post(1).Ref = APEdit.INVNUM
          TR2Post(1).DrAmt = APEdit.Dist(AcctDist).DAMT
          TR2Post(1).CrAmt = 0
          TR2Post(1).Src = MSrc$                '"AP" + ConvDateStr$(DATE$)
          TR2Post(1).Marked = False
          PUT AP2Post, RecordNum, TR2Post(1)
'*******************************************************************
          '--If there was a purchase order create a
          '--potrans record to liquidate encumbrance.
          IF GotPOs THEN
            POIFRecNum = POIFRecNum + 1
            TR2Post(1).AcctNum = APEdit.Dist(AcctDist).DACN
            TR2Post(1).TrDate = APEdit.DISTDATE
            TR2Post(1).Desc = APEdit.VENDNAME
            TR2Post(1).Ref = APEdit.INVNUM
            TR2Post(1).DrAmt = 0
            TR2Post(1).CrAmt = APEdit.Dist(AcctDist).DAMT
            TR2Post(1).Src = MSrc$              '"PO" + ConvDateStr$(DATE$)
            TR2Post(1).Marked = False
            PUT PO2Post, POIFRecNum, TR2Post(1)
          END IF
'*******************************************************************
          FOR Fund = 1 TO NumFunds '--Add distribution to the proper fund
            FundNum$ = LEFT$(APEdit.Dist(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(APEdit.Dist(AcctDist).DAMT)
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
          TR2Post(1).TrDate = APEdit.DISTDATE   'APEdit.INVDATE
          TR2Post(1).Desc = APEdit.VENDNAME
          TR2Post(1).Ref = APEdit.INVNUM
          TR2Post(1).DrAmt = 0
          TR2Post(1).CrAmt = TrFundSum#(Fund)
          TR2Post(1).Src = MSrc$                '"AP" + ConvDateStr$(DATE$)
          TR2Post(1).Marked = False
          PUT AP2Post, RecordNum, TR2Post(1)
'*************************************************************
          '--IF PO then Create an interface rec for potrans
          IF GotPOs THEN          'If there was one
            POIFRecNum = POIFRecNum + 1
            TR2Post(1).AcctNum = FundList$(Fund) + EncAcct$
            TR2Post(1).TrDate = APEdit.DISTDATE
            TR2Post(1).Desc = APEdit.VENDNAME
            TR2Post(1).Ref = APEdit.INVNUM
            TR2Post(1).DrAmt = TrFundSum#(Fund)
            TR2Post(1).CrAmt = 0
            TR2Post(1).Src = MSrc$
            TR2Post(1).Marked = False
            PUT PO2Post, POIFRecNum, TR2Post(1)
          END IF
'**************************************************************
        END IF  '--Fund summary <> 0
      NEXT      '--fund check for balance
    END IF      '--not deleted
  NEXT          'transaction
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
    K$ = INPUT$(1)
  END IF
  
  IF BadGLTrans > 0 THEN
    CLS
    PRINT "Error: "; BadGLTrans; " Invoice Distribution(s) NOT posted."
    PRINT "Review GL Posting Log."
    PRINT "Press any key to continue."
    K$ = INPUT$(1)
  END IF
  
  IF BadPOTrans > 0 THEN
    CLS
    PRINT "Error: "; BadPOTrans; " PO liquidation(s) NOT posted."
    PRINT "Review PO Posting Log."
    PRINT "Press any key to continue."
    K$ = INPUT$(1)
  END IF
  
  KILL "APIED.DAT"
  KILL "APINVIF.DAT"
  KILL "POINVIF.DAT"
  
  EXIT SUB
  
LogBadTrans:
  PRINT #LogFile, "Unable to find Vendor. Transaction deleted."
  ToPrint$ = SPACE$(80)
  MID$(ToPrint$, 1) = APEdit.VENDOR
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
  Comma2Fmt$ = "#####,.##"   'ten millions
  DistSumLine$ = "-------------"
  TransTotal# = 0
  TranCnt = 0
  
  SHARED VENDOR AS VendorRecType

  REDIM TAPLedgerRec(1)  AS APLedger81RecType
  LedgerRecLen = LEN(TAPLedgerRec(1))

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
  Title$(3) = "Run Date: " + DATE$              '""
  Title$(4) = ""
  Title$(5) = "Vendor Code & Name                    Comment                     TAX  PSL 1099"
  Title$(6) = "Invoice             PO       Date      Due Date  Post Date              Amt"
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
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen

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
    GET VendorFile, CVI(APEdit.VRECNUM), VENDOR
    
    IF NOT APEdit.DELFLAG THEN
      IF APEdit.POFLAG = -32767 THEN
        PRINT #PrnFileNum, "************* Error: Corrupt PO Flag *************"
        PRINT #PrnFileNum, "DO NOT POST!!  Contact Customer Support"
        EXIT FOR
      END IF
      
      TranCnt = TranCnt + 1
      TransTotal# = TransTotal# + Round#(APEdit.INVAMT)
      
      '--Print 1st Line - Transaction details
      ToPrint$ = SPACE$(78)
      
      LSET ToPrint$ = VENDOR.VNum
      MID$(ToPrint$, 12) = APEdit.VENDNAME
      MID$(ToPrint$, 39) = APEdit.INVDESC
      MID$(ToPrint$, 68) = APEdit.TAXYN
      MID$(ToPrint$, 73) = APEdit.PSLFlag
      MID$(ToPrint$, 78) = APEdit.GET1099
      PRINT #PrnFileNum, ToPrint$
      
      ToPrint$ = SPACE$(78)
      'MID$(ToPrint$, 1) = LEFT$(APEdit.VENDNAME, 16)
      MID$(ToPrint$, 1) = APEdit.INVNUM
      MID$(ToPrint$, 20) = LEFT$(APEdit.PONUM, 9)
      MID$(ToPrint$, 30) = ConvDateStr1$(Num2Date(APEdit.INVDATE))
      MID$(ToPrint$, 40) = ConvDateStr1$(Num2Date(APEdit.DUEDATE))
      MID$(ToPrint$, 50) = ConvDateStr1$(Num2Date(APEdit.DISTDATE))
      'MID$(ToPrint$, 60) = LEFT$(APEdit.INVNUM, 6)
      MID$(ToPrint$, 66) = FUsing$(STR$(APEdit.INVAMT), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      IF INSTR(APEdit.PONUM, "Multi") > 0 THEN
        FOR zz = 1 TO 6
          TNum$ = APEdit.PORecs(zz)
          IF LEN(QPTrim$(TNum$)) > 0 THEN
            LedgerRec& = CVL(TNum$)
            GET APLedgerFile, LedgerRec&, TAPLedgerRec(1)
            ToPrint$ = SPACE$(78)
            MID$(ToPrint$, 20) = LEFT$(TAPLedgerRec(1).PONUM, 9)
            MID$(ToPrint$, 66) = FUsing$(STR$(TAPLedgerRec(1).Amt), CommaFmt$)
            MID$(ToPrint$, 50) = ConvDateStr1$(Num2Date(TAPLedgerRec(1).TrDate))
            'TrDate
            PRINT #PrnFileNum, ToPrint$
          END IF
        NEXT
      END IF
      '--Blank line between detail and acct'g distributions
      PRINT #PrnFileNum,
      
      '--Print Distribution Label
      ToPrint$ = SPACE$(78)
      LSET ToPrint$ = " Accounting Distribution:"
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
      IF LineCnt >= MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintHeader
      END IF

'      IF APEdit.TaxYN = "Y" THEN
'        PRINT #PrnFileNum,
'        ToPrint$ = SPACE$(78)
'        LSET ToPrint$ = " Tax Distribution:"
'        PRINT #PrnFileNum, ToPrint$
'        ToPrint$ = SPACE$(78)
'        FOR zz = 1 TO 2
'          MID$(ToPrint$, 4) = APEdit.ITax(zz).DACN
'          MID$(ToPrint$, 21) = APEdit.ITax(zz).DACNM
'          MID$(ToPrint$, TranCol) = FUsing$(STR$(APEdit.ITax(zz).DAMT), CommaFmt$)
'          IF APEdit.ITax(zz).DAMT > 0 THEN
'            TotalTax# = Round#(TotalTax# + APEdit.ITax(zz).DAMT)
'          END IF
'          PRINT #PrnFileNum, ToPrint$
'          LineCnt = LineCnt + 1
'        NEXT
'      END IF
      ToPrint$ = STRING$(78, "-")
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 2
      PRINT #PrnFileNum,
      
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
  'LSET ToPrint$ = "Grand Totals"
  'PRINT #PrnFileNum, ToPrint$
  LSET ToPrint$ = "Distrubtions:"
  MID$(ToPrint$, 25) = FUsing$(STR$(TransTotal#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  LSET ToPrint$ = "        Taxs:"
  MID$(ToPrint$, 25) = FUsing$(STR$(TotalTax#), CommaFmt$)
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
  
  CLOSE APEditFile              'close the file
  
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
  
  PUT APEditFile, rec, APEdit   'write it
  CLOSE APEditFile              'close the file
  
END SUB

SUB SelectInvPOs (Form$(), AcctFlds(), Fld() AS FieldInfo, ChangeFlag)

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  'GotOld=

  REDIM TVendor(1)       AS VendorRecType
  REDIM TAPLedgerRec(1)  AS APLedger81RecType
  REDIM TAPDistRec(1)    AS APDistRecType

  DistRecLen = LEN(TAPDistRec(1))
  LedgerRecLen = LEN(TAPLedgerRec(1))

  Title$ = SPACE$(68)
  LSET Title$ = " PO Num        TR Date       Req Num            Amount"

  'This array points to the po record number fields
  REDIM POFld(1 TO 6) AS INTEGER

  'This array hold the actual po record numbers
  REDIM OPORecs(1 TO 6) AS LONG

  FOR zz = 1 TO 6
    POFld(zz) = FldNum("POREC" + QPTrim$(STR$(zz)), Fld())

    TNum$ = Form$(POFld(zz), 0)
    IF LEN(QPTrim$(TNum$)) > 0 THEN
      GotOld = True
      OPORecs(zz) = CVL(TNum$)
    ELSE
      OPORecs(zz) = 0
    END IF
  NEXT

  PONumFld = FldNum("PONUM", Fld())

  MaxLen = 59   'Set menu width to zero
  Action = 0    '0 let vertmenu have control until they select or cancel
  BoxBot = 15
  AmtOffSet = 43
  Row = 6
  NumFlds = UBOUND(Form$)
  GotSome = False
  POCnt = 0

  VRECNUM = CVI(LEFT$(Form$(FldNum("VRECNUM", Fld()), 0), 2))
  PONumFld = FldNum("PONUM", Fld())

  OpenVendorFile VendorFile, NumVRecs
  GET VendorFile, VRECNUM, TVendor(1)
  CLOSE

  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen

  LedgerRec& = TVendor(1).FrstTran
  DO WHILE LedgerRec& > 0
    GET APLedgerFile, LedgerRec&, TAPLedgerRec(1)
    IF TAPLedgerRec(1).TrCode = 4 THEN
      GotSome = True
      POCnt = POCnt + 1
      REDIM PRESERVE MTChoice(1 TO POCnt) AS FLen2  'Items to display
      REDIM PRESERVE Picked(1 TO POCnt)             'Selected items list
      REDIM PRESERVE PORecs(1 TO POCnt) AS LONG     'PO's rec numbers
      REDIM PRESERVE PODist(1 TO POCnt) AS LONG     'PO's first distribution
      LSET MTChoice(POCnt).V = TAPLedgerRec(1).PONUM
      MID$(MTChoice(POCnt).V, 14) = Num2Date(TAPLedgerRec(1).TrDate)
      MID$(MTChoice(POCnt).V, 28) = LEFT$(QPTrim$(TAPLedgerRec(1).DOCNum), 10)
      MID$(MTChoice(POCnt).V, AmtOffSet) = FUsing$(STR$(TAPLedgerRec(1).Amt), "#######.##")
      PORecs(POCnt) = LedgerRec&
      PODist(POCnt) = TAPLedgerRec(1).FrstDist
'This retags any previous selected po's
      IF GotOld THEN
        FOR zz = 1 TO 6
          IF OPORecs(zz) = LedgerRec& THEN
            OPCnt = OPCnt + 1
            Picked(OPCnt) = POCnt
            EXIT FOR
          END IF
        NEXT
      END IF
'***************************************
    END IF
    LedgerRec& = TAPLedgerRec(1).NextTrans
  LOOP
  CLOSE

  IF GotSome THEN
    'MPaintBox 3, 6, 22, 74, 8
    'MPaintBox 4, 7, 20, 74, 8              'this is ok
    MPaintBox 2, 1, 24, 80, 8   'to add shadow to list
    DisplayAPScrn "PICKOPTI"
    QPrintRC Title$, Row - 1, 7, 112
    WPickList MTChoice(), Picked(), PickedCnt, Row, UpDateFlag, Cnf
  ELSE
    DisplayAPScrn "NOVENDPO"
    WaitForAction
    GOTO ExitPOSelect
  END IF

  IF UpDateFlag THEN
    ChangeFlag = True
    TotalAmt# = 0
    DistCnt = 0
    OpenAcctFile AcctFileNum, NumAccts
    OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen

    FOR zz = 1 TO PickedCnt

      Form$(POFld(zz), 0) = MKL$(PORecs(Picked(zz))) 'store po rec number
      SaveField POFld(zz), Form$(), Fld(), BadFld
      NextDist& = PODist(Picked(zz))
      DO UNTIL NextDist& = 0
        GET APDistFile, NextDist&, TAPDistRec(1)
        DistCnt = DistCnt + 1
        GET AcctFileNum, TAPDistRec(1).DistAcctRec, Acct
        REDIM PRESERVE DistSum(1 TO DistCnt)  AS DistSumType
        DistSum(DistCnt).DistAcctNum = TAPDistRec(1).DistAcctNum
        DistSum(DistCnt).AcctTitle = Acct.Title
        DistSum(DistCnt).DistAmt = TAPDistRec(1).DistAmt
        NextDist& = TAPDistRec(1).NextDist
      LOOP
    NEXT

    CLOSE

    IF DistCnt >= 1 THEN
      SortT DistSum(1), DistCnt, 0, 44, 0, 16
    ELSEIF DistCnt < 1 THEN
      GOTO Skip2Here:
    END IF

    ActualCnt = 1
    FirstAcct$ = QPTrim$(DistSum(1).DistAcctNum)

    REDIM ActualDist(1 TO 1)  AS DistSumType

    ActualDist(1).DistAcctNum = DistSum(1).DistAcctNum
    ActualDist(1).AcctTitle = DistSum(1).AcctTitle
    ActualDist(1).DistAmt = DistSum(1).DistAmt

    FOR zz = 2 TO DistCnt
      ThisAcct$ = QPTrim$(DistSum(zz).DistAcctNum)
      IF ThisAcct$ = FirstAcct$ THEN
        ActualDist(ActualCnt).DistAmt = Round#(ActualDist(ActualCnt).DistAmt + DistSum(zz).DistAmt)
      ELSE
        ActualCnt = ActualCnt + 1
        REDIM PRESERVE ActualDist(1 TO ActualCnt) AS DistSumType
        ActualDist(ActualCnt).DistAcctNum = DistSum(zz).DistAcctNum
        ActualDist(ActualCnt).AcctTitle = DistSum(zz).AcctTitle
        ActualDist(ActualCnt).DistAmt = DistSum(zz).DistAmt
        FirstAcct$ = QPTrim$(DistSum(zz).DistAcctNum)
      END IF
    NEXT

    FOR zz = 1 TO ActualCnt
      StartFld = AcctFlds(zz)
      LSET Form$(StartFld, 0) = ActualDist(zz).DistAcctNum
      LSET Form$(StartFld + 1, 0) = ActualDist(zz).AcctTitle
      LSET Form$(StartFld + 3, 0) = STR$(ActualDist(zz).DistAmt)
      TotalAmt# = Round#(TotalAmt# + ActualDist(zz).DistAmt)
      FOR SAVFld = StartFld TO StartFld + 3
        SaveField SAVFld, Form$(), Fld(), BadFld
      NEXT
    NEXT

Skip2Here:

    IF PickedCnt > 1 THEN
      LSET Form$(PONumFld, 0) = "Multi"
    ELSEIF PickedCnt = 1 THEN
      LSET Form$(PONumFld, 0) = QPTrim$(LEFT$(MTChoice(Picked(1)).V, 13))
    ELSE
      LSET Form$(PONumFld, 0) = ""
    END IF

    FOR zz = PickedCnt + 1 TO 6
      LSET Form$(POFld(zz), 0) = MKL$(0)
      SaveField POFld(zz), Form$(), Fld(), BadFld
    NEXT

    LSET Form$(7, 0) = QPTrim$(STR$(TotalAmt#))
    SaveField 7, Form$(), Fld(), BadFld

    IF ActualCnt < NumAcctFlds THEN
      FOR CDCnt = AcctFlds(ActualCnt + 1) TO AcctFlds(NumAcctFlds)                'clear all of distrubtions
        LSET Form$(CDCnt, 0) = ""               'fields
        LSET Form$(CDCnt + 1, 0) = ""
        LSET Form$(CDCnt + 2, 0) = ""
        LSET Form$(CDCnt + 3, 0) = ""
        FOR SAVFld = CDCnt TO CDCnt + 3
          SaveField SAVFld, Form$(), Fld(), BadFld
        NEXT
      NEXT
    END IF
  END IF

ExitPOSelect:

  RestScrn TempScrn()
  ERASE TempScrn

  
END SUB

SUB ShowEdMode (EdMode%)
  
  Status$ = SPACE$(16)
  ActiveColor = 112
  InActiveColor = 115
  
  SELECT CASE EdMode
  CASE 0
    LSET Status$ = "Editing"
    F2BtnCap$ = "New "
    F5BtnColor = ActiveColor
  CASE 1
    LSET Status$ = "New Transaction"
    F2BtnCap$ = "Edit"
    F5BtnColor = InActiveColor
    
  END SELECT
  HideCursor
  QPrintRC Status$, 25, 2, -1
  QPrintRC F2BtnCap$, 25, 38, -1
  MPaintBox 25, 47, 25, 50, F5BtnColor          'F5 Button enable / disable
  ShowCursor
  
END SUB

SUB VoidInvTrans (VoidTransRecNum&)
  
  'STOP
  LOCATE , , 0
  
  '-If this file exists this routine did not finish last time.
  InterfaceFile$ = "APVDIF.DAT"
  IF Exist(InterfaceFile$) THEN
    'KILL InterfaceFile$
    CLOSE
    CLS
    QPrintRC "Whoa!", 1, 1, 12
    QPrintRC "The file APVDIF.DAT already exits. I Cannot continue.", 2, 1, -1
    QPrintRC "Press any key to return to the prior menu.", 3, 1, -1
    KeyPress$ = INPUT$(1)
    EXIT SUB
  END IF
  
  REDIM APLedger(1) AS APLedger81RecType
  REDIM DistRec(1) AS APDistRecType
  REDIM GLifRec(1) AS GLTransRecType
  
  LdRecLen = LEN(APLedger(1))
  DistRecLen = LEN(DistRec(1))
  GLIFRecLen = LEN(GLifRec(1))
  
  PrnFileName$ = "VOIDINV.PRN"
  PrnTitle$ = "Voided Invoice Report"
  Fmt$ = "######,.##"
  LPTNo = 1
  
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  
  GET APLedgerFile, VoidTransRecNum&, APLedger(1)
  APLedger(1).TrCode = -1
  PUT APLedgerFile, VoidTransRecNum&, APLedger(1)
  CLOSE APLedgerFile
  
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen
  
  GLIFFile = FREEFILE
  OPEN InterfaceFile$ FOR RANDOM AS GLIFFile LEN = GLIFRecLen
  
  PrnFile = FREEFILE
  OPEN PrnFileName$ FOR OUTPUT AS PrnFile
  GOSUB PrintReportHeader
  
  NextDist& = APLedger(1).FrstDist
  DO UNTIL NextDist& = 0
    GET APDistFile, NextDist&, DistRec(1)
    IFRec = IFRec + 1
    
    '--Set Entry Description Defaults
    GLifRec(1).Src = "VI" + ConvDateStr$(DATE$)
    GLifRec(1).TrDate = APLedger(1).TrDate
    GLifRec(1).Desc = APLedger(1).DOCNum
    GLifRec(1).Ref = "TO VOID"
    
    '--Make Debit side of entry
    GLifRec(1).AcctNum = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
    GLifRec(1).CrAmt = 0
    GLifRec(1).DrAmt = DistRec(1).DistAmt
    PUT GLIFFile, IFRec, GLifRec(1)
    GTDebit# = GTDebit# + GLifRec(1).DrAmt
    GOSUB PrintGLIF
    
    '--Make Credit side
    IFRec = IFRec + 1
    GLifRec(1).AcctNum = DistRec(1).DistAcctNum
    GLifRec(1).CrAmt = DistRec(1).DistAmt
    GLifRec(1).DrAmt = 0
    PUT GLIFFile, IFRec, GLifRec(1)
    GTCredit# = GTCredit# + GLifRec(1).CrAmt
    GOSUB PrintGLIF
    
    NextDist& = DistRec(1).NextDist
    
  LOOP
  
  PRINT #PrnFile,
  PRINT #PrnFile, "Grand Total";
  PRINT #PrnFile, TAB(58); USING Fmt$; GTDebit#;
  PRINT #PrnFile, TAB(70); USING Fmt$; GTCredit#
  
  CLOSE
  
  Post2GL InterfaceFile$, BadAcct
  KILL InterfaceFile$
  
  IF BadAcct THEN
    CLS
    QPrintRC "Posting Error.  Review GL Utility Log for details." + STR$(BadAcct), 2, 1, -1
    QPrintRC "Uposted Transactions: " + STR$(BadAcct), 2, 1, -1
    QPrintRC "Press any key for posting report", 3, 1, -1
    KeyPress$ = INPUT$(1)
  END IF
  
  PrintRptFile PrnTitle$, PrnFileName$, LPTNo, RetCode%, EntryPoint
  
  EXIT SUB
  
  
PrintReportHeader:
  PRINT #PrnFile, PrnTitle$
  PRINT #PrnFile, "Run Date: " + DATE$
  PRINT #PrnFile,
  PRINT #PrnFile, "Acct No         Date        Description        Ref            Debit      Credit"
  PRINT #PrnFile, "==============================================================================="
  RETURN
  
PrintGLIF:
  PRINT #PrnFile, GLifRec(1).AcctNum;
  PRINT #PrnFile, TAB(17); Num2Date(GLifRec(1).TrDate);
  PRINT #PrnFile, TAB(29); LEFT$(GLifRec(1).Desc, 17);
  PRINT #PrnFile, TAB(48); GLifRec(1).Ref;
  PRINT #PrnFile, TAB(58); USING Fmt$; GLifRec(1).DrAmt;
  PRINT #PrnFile, TAB(70); USING Fmt$; GLifRec(1).CrAmt
  RETURN
  
END SUB

