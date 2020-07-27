DEFINT A-Z

DECLARE SUB POApprove ()
DECLARE SUB PrnNAEditList ()
DECLARE SUB PrnAPEditList ()
DECLARE SUB POControl ()
DECLARE SUB VoidPOTrans (VoidTransRecNum&)
DECLARE SUB CancelPO ()
DECLARE SUB OpenPOTransFile (TransFileNum%, NumTrans&)
DECLARE SUB Post2PO (FileName$, BadTrans%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION EDCheck4Vendor% (Form$(), Fld() AS ANY, Terms$)
DECLARE FUNCTION EDCheck4Acct% (Form$(), AcctFldNo%)
DECLARE FUNCTION FindVendorRec% (VendorCode$)
DECLARE SUB PostPOTrans ()
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPVendorFile (APVendorFile%, NumVRecs%, VendorRecLen%)
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenVendorIdx (VendorIdxFile%, NumActiveVendors%)
DECLARE FUNCTION ChkForVendor% (ChkVendor$)
DECLARE SUB VertMenuT (ITEMS() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB PickList (ITEMS() AS ANY, Picked%(), NPicked%, Cnf AS ANY)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (x$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB EdEntry ()
DECLARE SUB OpenPOEditFile (POEditFile%, NumEdTrans%)
DECLARE SUB NewCJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList ()
DECLARE SUB DisplayTotals ()
DECLARE SUB ValidateForm (Form$(), AcctFlds() AS ANY, BadFld%)
DECLARE FUNCTION ListRecs% (APType)
DECLARE SUB ResetForm (Form$(), NumFlds%)
DECLARE SUB ShowEdMode (EDMode%)
DECLARE SUB Rec2Form (Form$(), RecordNumber%, Fld() AS ANY, APType)
DECLARE SUB SaveTrans (Form$(), RecordNumber%, EDMode%, APType)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB POMenu ()
DECLARE SUB PrintForms ()
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, x%, Y%, Wdth%, TitColor%)

  '$INCLUDE: 'DefCnf.bi'                    'monitor and color info
  '$INCLUDE: 'FieldInf.bi'                  'field information
  '$INCLUDE: 'FormEdit.bi'                  'form editing information
  '$INCLUDE: 'pageinfo.BI'                  'Form Page info
  '$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
  '$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
  '$INCLUDE: 'GL.BI'                        'GL File Types
  '$INCLUDE: 'SCRAUX.bi'
  '$INCLUDE: 'ap.bi'
  ''$INCLUDE: 'apwb.bi'
  '$INCLUDE: 'Poform.bi'
  
  CONST False = 0, TRUE = NOT False
  CONST VNameFld = 5
  CONST VRecFld = 6
  
  DIM SHARED TotLen, FundLen, AcctLen, DetLen
  DIM SHARED CashAcct$, APAcct$, EncAcct$, User$                ', FBAcct$
  DIM SHARED AcctIdx AS GLAcctIndexType
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED POEdit AS POFORMRecType2
  DIM SHARED POTrans AS GLTransRecType
  DIM SHARED POCont(1) AS POControlRecType
  
  '$INCLUDE: 'init.bi'                      'Get GL system info
  
  TYPE FLen
    V AS STRING * 42
  END TYPE
  
  TYPE PickListArray
    Item AS STRING * 64
  END TYPE

  TYPE AcctPOChkType
    Acct    AS STRING * 14
    Bgt     AS DOUBLE
    Encumb  AS DOUBLE
    Bal     AS DOUBLE
    POTotal AS DOUBLE
    Fill    AS STRING * 18
  END TYPE


  POMenu
  ClrBG
  RUN "APMENU"

SUB CancelPO
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  LdRecLen = LEN(APLedgerRec(1))
  DIM VendorIdx AS VendorIdxRecType
  DIM VENDOR AS VendorRecType
  
  FormName$ = "VOIDPO"
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
  FirstTime = TRUE
  
  LibFile2Scrn LibName$, FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).FldNo > 1 AND Frm(1).PrevFld = 1 THEN             'User is moving around the form
      ChkVendor$ = QPTrim$(Form$(1, 0))
      GOSUB Check4Vendor
      IF NOT Abort THEN
        GOSUB FindPO
      ELSE
        Abort = False
      END IF
    END IF
    
    IF Frm(1).KeyCode = EscKey THEN
      ExitFlag = TRUE
    ELSEIF Frm(1).KeyCode = F10Key THEN
      OkFlag = TRUE
    END IF
    
  LOOP UNTIL ExitFlag OR OkFlag
  
  IF ExitFlag THEN EXIT SUB
  
  IF OkFlag THEN
    Temp$ = InputBox$(LibName$, "OK2VOID", Cancel%)
  END IF
  
  IF Cancel THEN
    EXIT SUB
  ELSE
    'VoidChkNum& = VAL(Form$(2, 0)) 'QPValL(Form$(2, 0))
    VoidPOTrans VoidTransRecNum&
  END IF
  
  EXIT SUB
  
  
FindPO:
  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  POCnt = 0
  List$ = SPACE$(45)
  
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  
  GET VendorFile, VoidVendorRecNum&, VENDOR
  NextTrans& = VENDOR.FrstTran
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, APLedgerRec(1)
    IF APLedgerRec(1).TrCode = 4 THEN
      POCnt = POCnt + 1
    END IF
    NextTrans& = APLedgerRec(1).NextTrans
    
  LOOP
  
  IF POCnt = 0 THEN
    Ok = MsgBox(LibName$, "NOCHECKS")
    Frm(1).FldNo = 1
    Action = 1
    GOTO VoidPOExit
  END IF
  
  REDIM POList(1 TO POCnt) AS FLen              'Dim array to hold trans list
  
  MatchCnt = 0
  NextTrans& = VENDOR.FrstTran
  
  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, APLedgerRec(1)
    IF APLedgerRec(1).TrCode = 4 THEN
      MatchCnt = MatchCnt + 1
      LSET POList(MatchCnt).V = List$
      POList(MatchCnt).V = QPTrim$(APLedgerRec(1).PONUM)
      MID$(POList(MatchCnt).V, 11) = Num2Date(APLedgerRec(1).TrDate)
      MID$(POList(MatchCnt).V, 22) = FUsing(STR$(APLedgerRec(1).Amt), "$$########,.##")
      MID$(POList(MatchCnt).V, 39) = MKL$(NextTrans&)
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
    VertMenuT POList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      Abort = TRUE
      Frm(1).FldNo = 1
      Action = 1
      GOTO VoidPOExit
    END IF
    IF Choice > 0 THEN
      LSET Form$(2, 0) = QPTrim$(LEFT$(POList(Choice).V, 10))
      SaveField 2, Form$(), Fld(), BadFld
      VoidTransRecNum& = CVL(RIGHT$(POList(Choice).V, 4))
      Frm(1).FldNo = 3
      EXIT DO
    END IF
  LOOP
  
VoidPOExit:
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
        MID$(VendorList(MatchingVendors).V, 14) = VENDOR.VName
        MID$(VendorList(MatchingVendors).V, 41) = MKI$(VendorIdx.RecNum)
      END IF
    END IF
  NEXT
  
  CLOSE
  
  IF MatchingVendors = 0 THEN
    Ok = MsgBox(LibName$, "NOVENDOR")
    VendPickList = 0
    Frm(1).FldNo = 1
    Abort = TRUE
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
        Abort = TRUE
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

FUNCTION EDCheck4Acct% (Form$(), AcctFld%)
  
  MatchAccts = 0
  List$ = SPACE$(65)
  
  Lookup$ = QPTrim$(UCASE$(Form$(AcctFld, 0)))
  Lookuplen = LEN(Lookup$)
  
  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  
  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum
  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType
  
  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs
  
  OpenAcctFile AcctFileNum, NumAccts
  
  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, Acct
      IF Acct.Typ = "E" OR Acct.Typ = "A" THEN
        'IF Acct.Typ = "A" THEN STOP
        MatchAccts = MatchAccts + 1
        LSET AcctList(MatchAccts).V = List$
        AcctList(MatchAccts).V = TempAcctIdx(Cnt).AcctNum
        MID$(AcctList(MatchAccts).V, TotLen + 4) = Acct.Title
        MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
      END IF
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
    GOSUB GetAcctInfo
    EDCheck4Acct = TRUE
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
        GOSUB GetAcctInfo
        EDCheck4Acct = TRUE
        EXIT DO
      END IF
    LOOP
  END SELECT
  
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  
  ERASE AcctList, TempAcctIdx, ScrBuf
  
  
  EXIT FUNCTION
  
GetAcctInfo:
  OpenAcctFile AcctFileNum, NumAccts
  GET AcctFileNum, AcctRecNum, Acct
  CLOSE AcctFileNum
  LSET Form$(AcctFld, 0) = Acct.Num
  'LSET Form$(AcctFld + 1, 0) = Acct.Title
  LSET Form$(AcctFld + 1, 0) = MKI$(AcctRecNum)
  RETURN
  
  
  
END FUNCTION

FUNCTION EDCheck4Vendor% (Form$(), Fld() AS FieldInfo, Terms$)
  
  SHARED VendorIdx AS VendorIdxRecType
  SHARED VENDOR AS VendorRecType
  
  MatchingVendors = 0
  List$ = SPACE$(65)
  
  VLookup$ = QPTrim$(UCASE$(Form$(VNameFld, 0)))
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
  
  REDIM VendorList(1 TO NumActiveVendors) AS FLen               'Dim array to hold trans list
  
  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
      GET VendorFile, VendorIdx.RecNum, VENDOR
      IF VENDOR.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 13) = VENDOR.VName
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
    EDCheck4Vendor = TRUE
    LSET Form$(VNameFld, 0) = VendorList(1).V
    'LSET Form$(6, 0) = STR$(CVI(MID$(VendorList(1).V, 13, 28)))
    VRecNum = CVI(MID$(VendorList(1).V, 41, 2))
    GOSUB GetVendorAddress
  CASE ELSE
    REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
    Row = 5: Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
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
        LSET Form$(VNameFld, 0) = VendorList(Choice).V
        VRecNum = CVI(MID$(VendorList(Choice).V, 41, 2))
        GOSUB GetVendorAddress
        EDCheck4Vendor = TRUE
        '  SaveField 2, Form$(), Fld(), BadFld
        EXIT DO
      END IF
    LOOP
    
  END SELECT
  
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  QPrintRC TName$, 3, 27, -1
  ERASE VendorList, ScrBuf
  
  EXIT FUNCTION
  
GetVendorAddress:
  OpenVendorFile VendorFile, NumVRecs
  GET VendorFile, VRecNum, VENDOR
  
  Form$(VRecFld, 0) = STR$(VRecNum)
  SaveField VRecFld, Form$(), Fld(), BadFld
  
  Form$(VRecFld + 1, 0) = VENDOR.VName
  SaveField VRecFld + 1, Form$(), Fld(), BadFld
  
  Form$(VRecFld + 2, 0) = VENDOR.Addr1
  SaveField VRecFld + 2, Form$(), Fld(), BadFld
  
  Form$(VRecFld + 3, 0) = VENDOR.Addr2
  SaveField VRecFld + 3, Form$(), Fld(), BadFld
  
  Form$(VRecFld + 4, 0) = VENDOR.City
  SaveField VRecFld + 4, Form$(), Fld(), BadFld
  
  Form$(VRecFld + 5, 0) = VENDOR.State + " " + LEFT$(VENDOR.Zip, 5) + "-" + RIGHT$(VENDOR.Zip, 4)
  SaveField VRecFld + 5, Form$(), Fld(), BadFld
  
  CLOSE
  RETURN
  
END FUNCTION

SUB EdEntry
  
  SHARED VendorIdx AS VendorIdxRecType
  SHARED VENDOR AS VendorRecType
  
mainpobody:
  
  'REDIM POCont(1) AS POControlRecType
  
  LibName$ = "PO"
  FormName$ = "POFORM"
  NumOfScrns = 3
  Cnt = 0
  Saved$ = "N"
  
  REDIM FormName$(1 TO NumOfScrns)
  
  FOR scr = 1 TO NumOfScrns
    FormName$(scr) = FormName$ + LTRIM$(STR$(scr))
  NEXT
  
  '-- Get total number of fields
  NumFlds = -1
  FOR scr = 1 TO NumOfScrns
    NumFlds = NumFlds + LibNumberOfFields(LibName$, FormName$(scr)) + 1
  NEXT
  
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
  REDIM Frm(1)    AS FormInfo
  
  '--for each page, get first and last fields
  StartEl = 0   'Load first form at array start
  FOR scr = 1 TO NumOfScrns     'Get field defs
    LibGetFldDef LibName$, FormName$(scr), StartEl, Fld(), Form$(), ErrCode
    PageInfo(scr).PageNo = scr
    PageInfo(scr).FirstFld = StartEl + 1
    StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
    PageInfo(scr).LastFld = StartEl - 1
  NEXT
  
  '--Clear all fields and keep a list of the acct number fields
  REDIM AcctFlds(1 TO 12)
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
    IF QPTrim$(Fld(F).ScratchS) = "A" THEN
      Cnt = Cnt + 1
      AcctFlds(Cnt) = F
    END IF
  NEXT
  
  'Frm(1).PageDisable = True
  '--get first active record number and set Editing Mode
  OpenPOEditFile POEditFile, NumEdTrans
  CurrRec = 0
  NextRec = 1
  RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
  CLOSE POEditFile
  
  IF RecNum = 0 THEN
    EDMode = 1
  ELSE
    EDMode = 0
  END IF
  
  IF EDMode = 0 THEN
    Rec2Form Form$(), RecNum, Fld(), APType
  END IF
  ShowEdMode EDMode
  IF EDMode = 1 THEN
    POFile = FREEFILE
    OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFile LEN = LEN(POCont(1))
    IF LOF(POFile) > 0 THEN
      GET POFile, 1, POCont(1)
      Form$(2, 0) = "N/A"
      Form$(12, 0) = POCont(1).Shipto1
      Form$(13, 0) = POCont(1).Shipto2
      Form$(14, 0) = POCont(1).Shipto3
      Form$(15, 0) = POCont(1).Shipto4
      Form$(16, 0) = POCont(1).Shipto5
      Form$(17, 0) = POCont(1).FOB
      Form$(18, 0) = POCont(1).Shipvia
      Form$(19, 0) = POCont(1).Terms
      Form$(21, 0) = POCont(1).Addinst1
      Form$(22, 0) = POCont(1).Addinst2
      Form$(23, 0) = POCont(1).Addinst3
      CLOSE POFile
    END IF
  END IF
  
  Action = 1
  Frm(1).FldNo = 3
  Fld(2).Protected = -1

  IF QPTrim$(PONumber$) <> "N/A" THEN
    FOR ll = 2 TO 8
      Fld(ll).Protected = -1
    NEXT ll
    FOR ll = 24 TO 122
      Fld(ll).Protected = -1
    NEXT ll
  ELSE
    FOR ll = 2 TO 8
      Fld(ll).Protected = 0
    NEXT ll
    
    FOR ll = 24 TO 122
      Fld(ll).Protected = 0
    NEXT ll
    Frm(1).FldNo = 3
  END IF
  Fld(6).Protected = -1
  
  CommaFormat$ = "#########,.##"
  GOSUB ShowDistTot

  GOSUB ProtectFields

  scr = 1
  LibFile2Scrn LibName$, FormName$(scr), MonoCode, -1, ErrCode

  DO
    IF FirstTime THEN
      FirstTime = 0
      scr = 1
      Action = 1
    END IF
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF scr = 1 THEN PONumber$ = Form$(2, 0)
    
    IF QPTrim$(PONumber$) <> "N/A" THEN
      FOR ll = 24 TO 122
        Fld(ll).Protected = 0
      NEXT ll
      Fld(6).Protected = -1
    ELSE
      IF Fld(2).Protected = -1 THEN
        FOR ll = 2 TO 8
          Fld(ll).Protected = 0
        NEXT ll
        FOR ll = 24 TO 122
          Fld(ll).Protected = 0
        NEXT ll
        Frm(1).FldNo = 3
      END IF
      Fld(6).Protected = -1
    END IF

    GOSUB ProtectFields
    
    'Alt+F3 code was here
    
    'f2 key was here
    'SELECT CASE Frm(1).KeyCode
    'END SELECT
    
    '--User is moving around the form
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN
      '--Check for vendor
      IF Frm(1).PrevFld = VNameFld AND Frm(1).FldNo <> VNameFld THEN
        VendorOK = EDCheck4Vendor(Form$(), Fld(), Terms$)

        IF VendorOK THEN
          'Frm(1).FldNo = 11
          Action = 1
        ELSE
          Frm(1).FldNo = VNameFld
          Action = 2
        END IF
      END IF
      
      '--strip out the dashes if we are in an account field
      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN
        LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
        Action = 1
      END IF
      
      '--check for valid acct
      IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
        IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN
          FAcct$ = FmtAcct$(QPTrim$(Form$(Frm(1).PrevFld, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).PrevFld, 0) = FAcct$
          GOSUB CheckForValidAcctNum
        END IF
      END IF
      
      IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "S" THEN
        GOSUB ShowDistTot
      END IF
      
    END IF      'End of test for PrevFld <> current FldNo
    
    IF SaveFlag THEN
      GOSUB SaveRec
      IF Saved$ = "Y" THEN
        Action = 1
        GOTO mainpobody
      END IF
    END IF
    
    '--Process key strokes
    SELECT CASE Frm(1).KeyCode
    CASE F2Key  'Change Editing Mode
      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        SELECT CASE EDMode
        CASE 0  'we're in Edit Mode
          EDMode = 1            'switch to add mode
          ShowEdMode EDMode
          ResetForm Form$(), NumFlds
          Frm(1).FldNo = 1
          POFile = FREEFILE
          OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFile LEN = LEN(POCont(1))
          IF LOF(POFile) > 0 THEN
            GET POFile, 1, POCont(1)
            Form$(2, 0) = "N/A"
            Form$(12, 0) = POCont(1).Shipto1
            Form$(13, 0) = POCont(1).Shipto2
            Form$(14, 0) = POCont(1).Shipto3
            Form$(15, 0) = POCont(1).Shipto4
            Form$(16, 0) = POCont(1).Shipto5
            Form$(17, 0) = POCont(1).FOB
            Form$(18, 0) = POCont(1).Shipvia
            Form$(19, 0) = POCont(1).Terms
            Form$(21, 0) = POCont(1).Addinst1
            Form$(22, 0) = POCont(1).Addinst2
            Form$(23, 0) = POCont(1).Addinst3
          END IF
          CLOSE POFile
          
          Action = 1
          GOSUB ShowDistTot
        CASE 1  'currently adding records
          '--check to see if there are any active transactions
          OpenPOEditFile POEditFile, NumEdTrans
          CurrRec = 0: NextRec = 1
          RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE POEditFile
          IF RecNum = 0 THEN    'No active records can't edit
            EDMode = 1
          ELSE
            EDMode = 0          'switch to edit mode
            ShowEdMode EDMode
            RecNum = ListRecs(APType)
            IF RecNum > 0 THEN
              Rec2Form Form$(), RecNum, Fld(), APType
              Frm(1).FldNo = 1
              Action = 1
            ELSE
              CLOSE
              EXIT SUB
            END IF
          END IF
          GOSUB ShowDistTot
        CASE ELSE
        END SELECT              'Edmode
      END IF
      
    CASE F3Key
      IF EMode = 0 THEN
        Ok = MsgBox%("GL.QSL", "OK2DELTR")
        IF Ok = 2 THEN          'Pressed yes button #2
          Form$(1, 0) = "-1"
          SaveField 1, Form$(), Fld(), BadFld
          SaveTrans Form$(), RecNum, EDMode, APType
          ResetForm Form$(), NumFlds
          RecNum = ListRecs(APType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), APType
          ELSE
            EXIT SUB
          END IF
        ELSE
          Action = 2
        END IF
        Frm(1).FldNo = 1
        Action = 1
      END IF
      
    CASE F5KEY
      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        IF EDMode = 0 THEN
          RecNum = ListRecs(APType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), APType
            Frm(1).FldNo = 1
            Action = 1
          END IF
        END IF
      END IF
      
    CASE F10Key
      '--Make sure acct fields are formatted correctly if user
      '--Preses F10 with cursor in an account field
      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN
        IF LEN(QPTrim$(Form$(Frm(1).FldNo, 0))) > 0 THEN
          FAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).FldNo, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).FldNo, 0) = FAcct$
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      '--send user back thru loop to verify acct before saving
      Action = 1
      
      SaveFlag = TRUE
      
    CASE CtrlL  'Ctrl+LeftArrow               'Get the next record
      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          Frm(1).FldNo = 1
          OpenPOEditFile POEditFile, NumEdTrans
          CurrRec = RecNum: NextRec = -1
          RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE POEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), APType
          GOSUB ShowDistTot
          ShowEdMode EDMode
          'Frm(1).PrevFld = 1
          scr = 1
          Action = 1
        END IF
      END IF
      
    CASE CtrlR  'Ctrl+RightArrow             'Get the previous record
      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          Frm(1).FldNo = 1
          OpenPOEditFile POEditFile, NumEdTrans
          CurrRec = RecNum: NextRec = 1
          RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE POEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), APType
          GOSUB ShowDistTot
          ShowEdMode EDMode
          ' Frm(1).PrevFld = 1
          Action = 1
          scr = 1
        END IF
      END IF
    END SELECT
    
    '--Process screen page changes
    IF Frm(1).FldNo > PageInfo(scr).LastFld OR Frm(1).FldNo < PageInfo(scr).FirstFld THEN
      SELECT CASE Frm(1).FldNo
      CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
        scr = 1
      CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
        scr = 2
      CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
        scr = 3
      END SELECT
      LibFile2Scrn LibName$, FormName$(scr), MonoCode, -2, ErrCode
      GOSUB ShowDistTot
      ShowEdMode EDMode
      Action = 1
    END IF
    
    '--Buttons are not attached to the form, the following processes the mouse
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21
        SELECT CASE Frm(1).MCol
        CASE 34 TO 44
          PressButton F10Key, 21, 34, 44
        CASE 47 TO 60
          PressButton F3Key, 21, 47, 60
        CASE 63 TO 75
          PressButton ESC, 21, 63, 75
        END SELECT
      CASE 25
        SELECT CASE Frm(1).MCol
        CASE 40 TO 50
          PressButton F2Key, 25, 40, 50
        CASE 53 TO 63
          PressButton F5KEY, 25, 53, 63
        CASE 66 TO 70
          PressButton CtrlL, 25, 66, 70
        CASE 74 TO 78
          PressButton CtrlR, 25, 74, 78
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL Frm(1).KeyCode = ESC
  
  EXIT SUB
  
  
CheckForValidAcctNum:
  AcctFld = Frm(1).PrevFld
  SELECT CASE EDCheck4Acct%(Form$(), AcctFld%)
  CASE TRUE
    SaveField AcctFld, Form$(), Fld(), BadFld
    Action = 2
  CASE ELSE
    Frm(1).FldNo = AcctFld
    LSET Form$(Frm(1).FldNo, 0) = ""
    Action = 2
    SaveFlag = False
  END SELECT
RETURN
  
  
ShowDistTot:
  TotDist# = 0
  FOR Cnt = 1 TO NumFlds        'look thru acct fld list
    IF QPTrim$(Fld(Cnt).ScratchS) = "S" THEN    '= Distribution Field
      TotDist# = Round#(TotDist# + Value(Form$(Cnt, 0), ECode))
    END IF
  NEXT
  Dist$ = FUsing(STR$(TotDist#), CommaFormat$)
  QPrintRC Dist$, 19, 48, -1
RETURN

SaveRec:
  GOSUB ShowDistTot
  IF TotDist# <> Value(Form$(24, 0), ECode) THEN
    Ok = MsgBox%(LibName$, "BADPO2")
    Frm(1).FldNo = 1
    GOTO JumpOut
  END IF
  
  BadFld = 0
  ValidateForm Form$(), AcctFlds(), BadFld      'Make sure we have req fields
  IF BadFld = 0 THEN
    GOSUB ShowDistTot
    Form$(1, 0) = "0"                           'Set deleted flag to false
    SaveField 1, Form$(), Fld(), BadFld         'Write the flag to the form$
    SaveTrans Form$(), RecNum, EDMode, APType   'Save the transaction
    Ok = MsgBox%("GL.QSL", "SAVED")             'tell user we saved it

    IF EDMode = 1 THEN                      'if in Add Mode, clear
      ResetForm Form$(), NumFlds            'the form for the next entry
      CALL StuffBuf(CHR$(0) + CHR$(60))
    END IF
    scr = 1
    Frm(1).FldNo = 1
    Saved$ = "Y"
  ELSE
    Ok = MsgBox%(LibName$, "BADPO1")
    BEEP
    Frm(1).FldNo = BadFld
  END IF
JumpOut:
  Action = 1
  FirstTime = -1
  SaveFlag = False
  RETURN
  
  
CheckFormEdit:
  IF Frm(1).Edited THEN
    BEEP
    CanChange = False
    Ok = MsgBox("CMNDLG", "CHANGREC")           'returns 2 for Abandon
    IF Ok = 2 THEN CanChange = TRUE
    Action = 2
  ELSE
    CanChange = TRUE
  END IF
  RETURN
  
  
InitForm:
  REDIM Frm(1)    AS FormInfo
  'Frm(1).PageDisable = True
  Action = 1
RETURN

ProtectFields:
  FOR zz = 1 TO 12
    Fld(AcctFlds(zz) + 1).Protected = -1
  NEXT
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
  LastRec = NumActiveVendors
  Lookfor$ = QPTrim$(UCASE$(VendorCode$))

  FOR Cnt = 1 TO LastRec
    GET VendorIdxFile, Cnt, VendorIdx
    VName$ = QPTrim$(VendorIdx.VendorCode)
    IF Lookfor$ = VName$ THEN
      Match = TRUE
      EXIT FOR
    END IF
  NEXT
  
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
  Rec = CurrRec
  
  DO
    
    Rec = Rec + NextRec         'Set file pointer to next record
    
    IF Rec > NumRecs OR Rec <= 0 THEN           'test for beg or end of file
      Found = 0 'if no more records then get out
      EXIT DO
    END IF
    
    GET FileNum, Rec, POEdit    'Get the record
    
    IF POEdit.Deleted <> -1 THEN                'Ok if not deleted
      Found = 1
      EXIT DO   'Get out of loop when we find one
    END IF
    
  LOOP
  
  IF Found = 0 THEN
    GetNextRec = CurrRec
  ELSE
    GetNextRec = Rec
  END IF
  
END FUNCTION

FUNCTION ListRecs (APType)
  
  SHARED Cnf AS Config
  REDIM ScrBuf(2000)
  
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  
  'QPrintRC "List ", 25, 68, -1
  
  OpenPOEditFile POEditFile, NumEdTrans
  
  REDIM TransList$(1 TO NumEdTrans)             'Dim array to hold trans list
  
  Active = 0    'Counter for Active Transactions
  
  FOR Cnt = 1 TO NumEdTrans     'Read the file
    GET POEditFile, Cnt, POEdit
    IF NOT POEdit.Deleted THEN  'if active add to list
      Active = Active + 1
      
      List$ = SPACE$(75)
      
      LSET List$ = POEdit.VNDRCODE
      MID$(List$, 12) = POEdit.PONUM
      MID$(List$, 38) = ConvDateStr1(Num2Date(POEdit.PODATE))
      MID$(List$, 53) = FUsing$(STR$(POEdit.POAMT), "#########,.##")
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
  
  ListDesc$ = "   Vendor    P O Number                Date            Total Amt     "
  QPrintRC ListDesc$, 2, 6, Cnf.ActivCh
  'MPaintBox 6, 75, 7, 76, 8              'to add shadow to list
  
  DO            'call vertmenu
    VertMenu TransList$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      CLOSE
      Abort = TRUE
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
  
  CLOSE POEditFile              'Clean up after ourselves
  ERASE TransList$
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  
  ListRecs = RecNum
  
  
END FUNCTION

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

SUB OpenAPVendorFile (APVendorFile, NumVRecs, VendorRecLen)
  
  APVendorFile = FREEFILE
  OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS APVendorFile LEN = VendorRecLen
  NumVRecs = LOF(APVendorFile) \ VendorRecLen
  
END SUB

SUB OpenPOEditFile (POEditFile, NumEdTrans)
  
  FileName$ = "APPED.DAT"
  
  EdLen = LEN(POEdit)
  POEditFile = FREEFILE
  OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS POEditFile LEN = EdLen
  NumEdTrans = LOF(POEditFile) \ EdLen
  
END SUB

SUB POApprove

  SHARED User$
  
  SHARED Choice$()              'Multiple choice array
  
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
  
  '--Get user input on where to send report
  LibLoadDisplayForm "AP.QSL", "POAPPROV", Form$(), Fld()
  PrintHelp "Purchase Order Approval"
  
  DIM Frm AS FormInfo
  
StartOver:
  Form$(1, 0) = ""
  Form$(2, 0) = ""
  Frm.FldNo = 1
  Action = 1
  
  '--Get the report options from user
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    Password$ = QPTrim$(Form$(2, 0))
    PwdLen = LEN(Password$)
    IF PwdLen > 0 AND PwdLen <> OldPwdLen THEN
      LOCATE 13, 39: PRINT STRING$(15, " ")
      LOCATE 13, 39:
      PRINT STRING$(PwdLen, "*")
      OldPwdLen = PwdLen
    END IF
    
    SELECT CASE Frm.KeyCode
      
    CASE F10Key 'SaveButton
      
      DeptNumber$ = Form$(1, 0)
      DeptNumber$ = QPTrim$(DeptNumber$)
      Password$ = QPTrim$(Form$(2, 0))
      PwdLen = LEN(Password$)
      IF PwdLen > 0 AND PwdLen <> OldPwdLen THEN
        LOCATE 13, 39: PRINT STRING$(15, " ")
        LOCATE 13, 39:
        PRINT STRING$(PwdLen, "*")
        OldPwdLen = PwdLen
      END IF
      
    CASE ESC
      EXIT SUB
    END SELECT
    
    IF Frm.FldNo = 3 AND NOldPwdLen <> PwdLen THEN
      LOCATE 13, 39: PRINT STRING$(15, " ")
      LOCATE 13, 39:
      PRINT STRING$(PwdLen, "*")
      NOldPwdLen = PwdLen
      
    END IF
    
  LOOP UNTIL Frm.KeyCode = F10Key
  
  '==Start processing register=>
  LOCATE 1, 1, 0
  IF Password$ <> "GOODPO" THEN GOTO StartOver
  
  
  'Now Get All Open PO's Into Pick List for Picking
  
  'DIM PO2Approve AS POApproveRecType
  DIM PO AS POFORMRecType2
  OpenPOEditFile POEditFile, NumEdTrans
  
  'Found Out How Many in PO File to Pick From
  NumberPO! = 0
  FOR Transaction = 1 TO NumEdTrans
    GET POEditFile, Transaction, PO
    PODept$ = QPTrim$(PO.REQNUM)
    
    IF NOT PO.Deleted THEN
      
      IF PODept$ = DeptNumber$ OR DeptNumber$ = "ALL" THEN
        
        IF LEFT$(PO.PONUM, 3) = "N/A" THEN
          
          NumberPO! = NumberPO! + 1
          
        END IF  'Not deleted test
      END IF    'Test for Dept Number Match
    END IF      'Test for Approved PO
    
    
  NEXT          'Transaction
  
  REDIM POApproveList(NumberPO!) AS PickListArray
  REDIM Picked(NumberPO!)
  
  POName$ = SPACE$(64)
  
  FOR Transaction = 1 TO NumEdTrans
    GET POEditFile, Transaction, PO
    PODept$ = QPTrim$(PO.REQNUM)
    
    IF NOT PO.Deleted THEN
      IF PODept$ = DeptNumber$ OR DeptNumber$ = "ALL" THEN
        IF LEFT$(PO.PONUM, 3) = "N/A" THEN
          GoodPO! = GoodPO! + 1
          LSET POName$ = STR$(Transaction) + " Dept# " + PO.REQNUM + "  " + LEFT$(PO.VNDRINF1, 25) + "  $" + STR$(PO.POAMT)
          POApproveList(GoodPO!).Item = POName$
        END IF  'Not deleted test
      END IF    'Test for Dept Number Match
    END IF      'Test for Approved P
  NEXT          'Transaction
  
  CLOSE
  ClrBG
  QPrintRC "  F10 To Continue, SpaceBar To Toggle, Alt-M Marks All, Alt-C Clear All Marks", 25, 1, -1
  
  IF GoodPO! = 0 THEN
    PRINT CHR$(7);
    QPrintRC "  No PO's on File to Approve                                                 ", 25, 1, -1
    SLEEP 2
    EXIT SUB
  END IF
  
  PickList POApproveList(), Picked(), NPicked, Cnf
  
  IF NPicked = 0 THEN EXIT SUB
  
  ClrBG
  QPrintRC "  Now Assigning the PO Numbers to Selected PO's ", 25, 1, -1
  
  OpenPOEditFile POEditFile, NumEdTrans
  
  POFile = FREEFILE
  OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFile LEN = LEN(POCont(1))
  
  FOR PCnt = 1 TO NPicked
    IF Picked(PCnt) THEN
      
      Transaction! = VAL(LEFT$(POApproveList(PCnt).Item, 6))
      GET POEditFile, Transaction!, PO
      GET POFile, 1, POCont(1)
      PO.PONUM = POCont(1).PONumber
      PO.Deleted = 1
      PUT POEditFile, Transaction!, PO
      PONumber$ = POCont(1).PONumber
      KK = INSTR(PONumber$, "-")
      IF KK > 0 THEN
        OLDSTUFF$ = LEFT$(PONumber$, KK)
        PONumber$ = MID$(PONumber$, KK + 1, LEN(PONumber$) - KK + 1)
      END IF
      PONumber& = VAL(PONumber$)
      PONumber& = PONumber& + 1
      PONumber$ = LTRIM$(STR$(PONumber&))
      IF KK > 0 THEN
        PONumber$ = OLDSTUFF$ + PONumber$
      END IF
      POCont(1).PONumber = PONumber$
      PUT POFile, 1, POCont(1)
    END IF
  NEXT PCnt
  CLOSE
  
  EXIT SUB
  
END SUB

SUB POControl
  'REDIM POCont(1) AS POControlRecType
  
  LibName$ = "PO.QSL"
  FormName$ = "POCTRL"
  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  DIM Frm AS FormInfo           'Form editing info
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintTitle User$
  PrintHelp "PO Control"
  OrgVend$ = ""
  Action = 1
  GOSUB CHECKPREV
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode
    CASE F10Key
      IF LEN(QPTrim$(Form$(1, 0))) = 0 THEN
        '--Don't allow to save with a blank vendor code
        BEEP
        LSET Form$(1, 0) = ""
        Frm.FldNo = 1
        Action = 1
        Done = False
      ELSE
        GOSUB SaveControl
        
        Done = TRUE
      END IF
    CASE ESC
      Done = TRUE
    CASE ELSE
    END SELECT
    
  LOOP UNTIL Done
  EXIT SUB
CHECKPREV:
  POFile = FREEFILE
  OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFile LEN = LEN(POCont(1))
  IF LOF(POFile) > 0 THEN
    GET POFile, 1, POCont(1)
    Form$(1, 0) = POCont(1).PONumber
    Form$(2, 0) = POCont(1).Header1
    Form$(3, 0) = POCont(1).Header2
    Form$(4, 0) = POCont(1).Header3
    Form$(5, 0) = POCont(1).Header4
    Form$(6, 0) = POCont(1).Shipto1
    Form$(7, 0) = POCont(1).Shipto2
    Form$(8, 0) = POCont(1).Shipto3
    Form$(9, 0) = POCont(1).Shipto4
    Form$(10, 0) = POCont(1).Shipto5
    Form$(11, 0) = POCont(1).FOB
    Form$(12, 0) = POCont(1).Shipvia
    Form$(13, 0) = POCont(1).Terms
    Form$(14, 0) = POCont(1).Addinst1
    Form$(15, 0) = POCont(1).Addinst2
    Form$(16, 0) = POCont(1).Addinst3
  END IF
  CLOSE POFile
  RETURN
  
  
SaveControl:
  POFile = FREEFILE
  OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFile LEN = LEN(POCont(1))
  
  POCont(1).PONumber = Form$(1, 0)
  POCont(1).Header1 = Form$(2, 0)
  POCont(1).Header2 = Form$(3, 0)
  POCont(1).Header3 = Form$(4, 0)
  POCont(1).Header4 = Form$(5, 0)
  POCont(1).Shipto1 = Form$(6, 0)
  POCont(1).Shipto2 = Form$(7, 0)
  POCont(1).Shipto3 = Form$(8, 0)
  POCont(1).Shipto4 = Form$(9, 0)
  POCont(1).Shipto5 = Form$(10, 0)
  POCont(1).FOB = Form$(11, 0)
  POCont(1).Shipvia = Form$(12, 0)
  POCont(1).Terms = Form$(13, 0)
  POCont(1).Addinst1 = Form$(14, 0)
  POCont(1).Addinst2 = Form$(15, 0)
  POCont(1).Addinst3 = Form$(16, 0)
  POCont(1).Pading = ""
  PUT POFile, 1, POCont(1)
  CLOSE POFile
  RETURN
  
  
  
  
END SUB

SUB POMenu
  
  '$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
  
  TextCursor -2, -2             'Define the Mouse cursor
  'ShowCursor    'Show it
  
  DIM MenuText$(1 TO 9)
  MenuText$(1) = "Enter/Edit Purchase Orders"
  MenuText$(2) = "Approve Purchase Orders"
  MenuText$(3) = "Print Purchase Orders Non-Approved"
  MenuText$(4) = "Print Purchase Orders Approved"
  MenuText$(5) = "Print Purchase Order Forms"
  MenuText$(6) = "Post Approved PO's"
  MenuText$(7) = "Cancel Open PO's"
  MenuText$(8) = "Purchase Order Control"
  MenuText$(9) = "Quit"
  
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
    Title$ = "Purchase Order Processing "
    
    MnuTitle Title$, Row, Col, MaxLen, 15
    PrintTitle User$
    PrintHelp "P.O. Processing"
    LOCATE Row, Col, 0
    VertMenu MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
    CASE 1
      ClrBG
      EdEntry
    CASE 2
      ClrBG
      POApprove
    CASE 3
      ClrBG
      'BgtvsAct User$, FundLen, AcctLen          'unrem
      'BgtvsAct
      PrnNAEditList
    CASE 4
      ClrBG
      PrnAPEditList
    CASE 5
      ClrBG
      PrintForms
    CASE 6
      ClrBG
      PostPOTrans
    CASE 7
      ClrBG
      CancelPO
    CASE 8
      ClrBG
      POControl
    CASE 9
      CLS : END
    CASE ELSE
    END SELECT
    
  LOOP
  
END SUB

SUB PostPOTrans
  
  '  SHARED APAcct$
  
  SHARED VENDOR AS VendorRecType
  REDIM APDistRec(1) AS APDistRecType
  REDIM APLedgerRec(1) AS APLedger81RecType
  REDIM Tr2Post(1) AS GLTransRecType
  DIM PO AS POFORMRecType2
  
  LedgerRecLen = LEN(APLedgerRec(1))
  DistRecLen = LEN(APDistRec(1))
  VendorRecLen = LEN(VENDOR)
  
  '--Paint a background
  FOR Cnt = 2 TO 24
    QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
  NEXT
  
  '--Verify that there are transactions
  OpenPOEditFile POEditFile, NumEdTrans
  
  '--Test for valid Vendor Numbers
  FOR Cnt = 1 TO NumEdTrans
    GET POEditFile, Cnt, PO
    
    IF PO.Deleted = 1 THEN
      VendRecNum = FindVendorRec(PO.VNDRCODE)
      IF VendRecNum > 0 THEN
        Active = Active + 1
      ELSE
        AbortFlag = TRUE
        COLOR 12: CLS
        BEEP
        PRINT "Error: Invalid Vendor Code"
        PRINT "Unable to locate vendor: "; PO.VNDRCODE
        PRINT "Operation Aborted."
        PRINT
        COLOR 7
        PRINT "Press any key to continue."
        k$ = INPUT$(1)
        EXIT FOR
      END IF
    END IF
  NEXT
  CLOSE
  
  IF AbortFlag THEN EXIT SUB
  
  '--Check for no active transactions
  IF Active = 0 THEN
  '--No active transactions - tell user and get out
    Ok = MsgBox%("GL.QSL", "NOTRANS")
    EXIT SUB
  END IF
  
  '--make sure we're ready to post
  Ok = MsgBox%("GL.QSL", "POSTGJ")
  IF Ok = 1 THEN EXIT SUB       '1=No
  
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM TrFundSum#(1 TO NumFunds)
  
  OpenPOEditFile POEditFile, NumEdTrans
  
  AP2Post = FREEFILE
  OPEN "APPOIF.DAT" FOR RANDOM AS AP2Post LEN = LEN(Tr2Post(1))
  
  RecordNum = 0 'Reset Active counter for posting
  
  FOR Cnt = 1 TO NumEdTrans     'number of invoices to process
    REDIM TrFundSum#(1 TO NumFunds)
    'Reinitialize transaction fund summary elements for next loop
    GET POEditFile, Cnt, PO
    
    IF PO.Deleted = 1 THEN
      FOR AcctDist = 1 TO 12
        IF LEN(QPTrim$(PO.ITEMS(AcctDist).ACCTNO)) THEN
          RecordNum = RecordNum + 1
          Tr2Post(1).AcctNum = PO.ITEMS(AcctDist).ACCTNO
          Tr2Post(1).TrDate = PO.PODATE
          Tr2Post(1).Desc = PO.VNDRCODE
          Tr2Post(1).Ref = PO.PONUM
          Tr2Post(1).DrAmt = PO.ITEMS(AcctDist).EXT
          Tr2Post(1).CrAmt = 0
          Tr2Post(1).Src = "PO" + ConvDateStr$(DATE$)
          PUT AP2Post, RecordNum, Tr2Post(1)
          '--Add this distribution to proper fund
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(PO.ITEMS(AcctDist).ACCTNO, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(PO.ITEMS(AcctDist).EXT)
              'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(PO.Items(AcctDist).Ext)
              EXIT FOR
            END IF
          NEXT
          
        END IF  'test for blank distribution line
      NEXT      'Acct'g Distribution
    END IF      'not deleted
    
  NEXT          'transaction
  
  CLOSE
  
  '--common post & link sub in comnaux
  Post2PO "APPOIF.DAT", BadTrans%
  IF BadTrans = 0 THEN          'posted ok
    ' MakeTempReport
    ' KILL "APINVIF.dat"
  ELSE          '
    '--Couldn't find an account.
    '--Account was possibly deleted after entry made?
    
    COLOR 12: CLS
    BEEP
    PRINT "Error: PO Transaction Posting"
    PRINT "Unable to locate one or more G/L Accounts."
    PRINT "Unposted transactions listed in POPOST.LOG"
    PRINT
    COLOR 7
    PRINT "Press any key to continue."
    k$ = INPUT$(1)
    
  END IF
  
  '--Now post transaction to apledger.dat
  OpenPOEditFile POEditFile, NumEdTrans
  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen
  OpenAPVendorFile APVendorFile, NumVRecs, VendorRecLen
  
  FOR Cnt = 1 TO NumEdTrans
    GET POEditFile, Cnt, PO
    IF PO.Deleted = 1 THEN
      '--Update PO in vendor link to apledger
      NumLedgerRecs = NumLedgerRecs + 1
      VendRecNum = PO.VNDRREC
      
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
      
      '--Post transaction to apledger
      APLedgerRec(1).VRecNum = PO.VNDRREC
      APLedgerRec(1).VendorCode = PO.VNDRCODE
      APLedgerRec(1).TrDate = PO.PODATE
      APLedgerRec(1).DOCNum = PO.PONUM
      APLedgerRec(1).PONUM = PO.PONUM
      APLedgerRec(1).TrCode = 4
      
      APLedgerRec(1).DeptNumb = VAL(PO.REQNUM)
      
      'APLedgerRec(1).PayCode = VAL(PO.PayCode)
      APLedgerRec(1).GLDistDate = PO.PODATE
      APLedgerRec(1).Amt = PO.POAMT
      APLedgerRec(1).NextTrans = 0
      APLedgerRec(1).FrstDist = NumDistRecs& + 1
      
      FOR D = 1 TO 12
        IF LEN(QPTrim$(PO.ITEMS(D).ACCTNO)) THEN
          '--Post and link distributions
          REDIM APDistRec(1) AS APDistRecType
          NumDistRecs& = NumDistRecs& + 1
          APDistRec(1).APLedgerRec = NumLedgerRecs
          APDistRec(1).DistAcctRec = PO.ITEMS(D).AcctRec
          APDistRec(1).DistAcctNum = PO.ITEMS(D).ACCTNO
          APDistRec(1).DistAmt = PO.ITEMS(D).EXT
          'APDistRec(1).DistCRAmt = 0
          APDistRec(1).NextDist = NumDistRecs& + 1
          PUT APDistFile, NumDistRecs&, APDistRec(1)
        END IF
      NEXT
      'Update the last distribution pointer
      GET APDistFile, NumDistRecs&, APDistRec(1)
      APDistRec(1).NextDist = 0
      PUT APDistFile, NumDistRecs&, APDistRec(1)
      
      APLedgerRec(1).LastDist = NumDistRecs&
      PUT APLedgerFile, NumLedgerRecs, APLedgerRec(1)
      PO.Deleted = -1
      PUT POEditFile, Cnt, PO
    END IF
  NEXT

  OK2Kill = -1
  FOR Cnt = 1 TO NumEdTrans
    GET POEditFile, Cnt, PO
    IF PO.Deleted <> -1 THEN
      OK2Kill = 0
      EXIT FOR
    END IF
  NEXT
  CLOSE

  IF OK2Kill = -1 THEN
    KILL "APPED.DAT"
  END IF
  KILL "APPOIF.DAT"
END SUB

SUB PrintForms
  
  LOCATE , , 0
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "3"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibName$ = "AP.QSL"
  FormName$ = "APAPPRN"
  Title$ = "PO Forms"
  FF$ = CHR$(12)
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintHelp "Print PO Forms"
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
    CASE F10Key 'SaveButton
      DeptNumber$ = Form$(1, 0)
      DeptNumber$ = QPTrim$(DeptNumber$)
      
      IF LEN(QPTrim$(Form$(3, 0))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(3, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(3, 0)), 1))
      END IF
    CASE ESC
      EXIT SUB
    END SELECT
  LOOP UNTIL Frm.KeyCode = F10Key
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  '--Start printing forms
  
  POFile = FREEFILE
  OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFile LEN = LEN(POCont(1))
  IF LOF(POFile) > 0 THEN
    GET POFile, 1, POCont(1)
  END IF
  
  DIM PO AS POFORMRecType2
  OpenPOEditFile POEditFile, NumEdTrans
  PrnFileName$ = "POFORMS.PRN"
  PrnFile = FREEFILE
  OPEN PrnFileName$ FOR OUTPUT AS #PrnFile
  
  FOR t = 1 TO NumEdTrans
    GET POEditFile, t, PO
    
    PODept$ = QPTrim$(PO.REQNUM)
    
    IF NOT PO.Deleted THEN
      IF LEFT$(PO.PONUM, 3) <> "N/A" THEN
        IF PODept$ = DeptNumber$ OR DeptNumber$ = "ALL" THEN
          
          PO.Deleted = 1        'Mark as Printed
          PUT POEditFile, t, PO
          
          ItemsPrinted = 0
          
          GOSUB PrintPOHeader
          
          
          '--Loop Thru Items
          DetailLines = 0
          FOR i = 1 TO 12
            IF LEN(QPTrim$(PO.ITEMS(i).ACCTNO)) THEN
              PRINT #PrnFile, TAB(1); PO.ITEMS(i).STKNO;
              PRINT #PrnFile, TAB(10); PO.ITEMS(i).DESCA;
              IF PO.ITEMS(i).QUAN > -999999.99# THEN
                PRINT #PrnFile, TAB(34); USING "####.##"; PO.ITEMS(i).QUAN;
              END IF
              IF PO.ITEMS(i).PRICE > -999999.99# THEN
                PRINT #PrnFile, TAB(44); USING "######.##"; PO.ITEMS(i).PRICE;
              END IF
              PRINT #PrnFile, TAB(54); USING "######.##"; PO.ITEMS(i).EXT;
              PRINT #PrnFile, TAB(65); StripAcct$(PO.ITEMS(i).ACCTNO)
              PRINT #PrnFile, TAB(10); PO.ITEMS(i).DESCB
              DetailLines = DetailLines + 2
            END IF              'Active transaction test
            IF LEN(QPTrim$(PO.ITEMS(i).ACCTNO)) = 0 AND LEN(QPTrim$(PO.ITEMS(i).DESCA)) THEN
              PRINT #PrnFile, TAB(12); PO.ITEMS(i).DESCA
              PRINT #PrnFile, TAB(12); PO.ITEMS(i).DESCB
              DetailLines = DetailLines + 2
            END IF              'Active transaction test
          NEXT
          
          
          PRINT #PrnFile,
          PRINT #PrnFile, TAB(20); "Total Purchase Order Amount "; TAB(50); USING "$$######,#.##"; PO.POAMT
          PRINT #PrnFile,
          PRINT #PrnFile, TAB(5); "Additional Instruction"
          PRINT #PrnFile, TAB(5); PO.Addinst1
          PRINT #PrnFile, TAB(5); PO.Addinst2
          PRINT #PrnFile, TAB(5); PO.Addinst3
          PRINT #PrnFile, ""
          PRINT #PrnFile, TAB(5); "Provision for payment of monies to fall due under this agreement"
          PRINT #PrnFile, TAB(5); "has been made by appropriation duly made or by bonds or notes"
          PRINT #PrnFile, TAB(5); "duly authorized, as required by the LOCAL GOVERNMENT ACT."
          PRINT #PrnFile, ""
          PRINT #PrnFile, TAB(5); "Department Head Signature: _________________________________________"
          PRINT #PrnFile, ""
          PRINT #PrnFile, TAB(5); "Finance Officer Signature: _________________________________________"
          PRINT #PrnFile, FF$
        END IF
      END IF
    END IF
  NEXT
  
  CLOSE
  
  PrintRptFile Title$, PrnFileName$, LPTNo, RetCode%, EntryPoint
  EXIT SUB
  
PrintPOHeader:
  PRINT #PrnFile, ""
  Header1$ = POCont(1).Header1
  Header1$ = QPTrim$(Header1$)
  Header2$ = POCont(1).Header2
  Header2$ = QPTrim$(Header2$)
  Header3$ = POCont(1).Header3
  Header3$ = QPTrim$(Header3$)
  Header4$ = POCont(1).Header4
  Header4$ = QPTrim$(Header4$)
  PRINT #PrnFile, TAB(33); "PURCHASE ORDER"
  PRINT #PrnFile, ""
  PRINT #PrnFile, TAB(40 - (LEN(Header1$) / 2)); Header1$
  PRINT #PrnFile, TAB(40 - (LEN(Header2$) / 2)); Header2$
  PRINT #PrnFile, TAB(40 - (LEN(Header3$) / 2)); Header3$
  PRINT #PrnFile, TAB(40 - (LEN(Header4$) / 2)); Header4$
  PRINT #PrnFile, TAB(5); "PO Date: "; Num2Date(PO.PODATE); TAB(55); "PO # "; PO.PONUM
  PRINT #PrnFile, ""
  PRINT #PrnFile, TAB(5); "Vendor:"; TAB(45); "Ship To:"
  PRINT #PrnFile, TAB(5); PO.VNDRINF1; TAB(45); PO.SHPLINE1
  PRINT #PrnFile, TAB(5); PO.VNDRINF2; TAB(45); PO.SHPLINE2
  PRINT #PrnFile, TAB(5); PO.VNDRINF3; TAB(45); PO.SHPLINE3
  PRINT #PrnFile, TAB(5); PO.VNDRINF4; TAB(45); PO.SHPLINE4
  PRINT #PrnFile, TAB(5); PO.VNDRINF5
  PRINT #PrnFile, ""
  PRINT #PrnFile, STRING$(79, "-")
  PRINT #PrnFile, "Ship Via: "; PO.Shipvia; TAB(35); "  FOB: "; PO.FOB; TAB(62); "Dept # "; PO.REQNUM
  PRINT #PrnFile, " Ship By: "; PO.SHIPON; TAB(35); "Terms: "; PO.Terms
  PRINT #PrnFile, STRING$(79, "-")
  PRINT #PrnFile, TAB(47); "Unit"; TAB(57); "Total"
  PRINT #PrnFile, "Stock #"; TAB(10); "Description"; TAB(37); "Quan"; TAB(47); "Price"; TAB(57); "Price"; TAB(65); "Acct Number"
  PRINT #PrnFile, STRING$(79, "-")
  RETURN
  
END SUB

SUB PrnAPEditList
  SHARED User$
  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "#########,.##"   'ten millions
  DistSumLine$ = "-------------"
  FF$ = CHR$(12)
  REDIM Title$(5)
  MaxLines = 56
  TransTotal# = 0
  TranCnt = 0
  ToPrint$ = SPACE$(78)
  
  SHARED Choice$()              'Multiple choice array
  REDIM Choice$(0 TO 3, 1)      'Choices for entry type field
  Choice$(0, 0) = "3"           'Field Number
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  
  FileName$ = "APPOREG.PRN"
  RegTitle$ = "Purchase Orders"
  TranCol = CreditCol
  CashCol = DebitCol
  
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
  '--Get user input on where to send report

  LibFile2Scrn "AP.QSL", "BG", MonoCode%, Attribute%, ErrorCode%
  LibLoadDisplayForm "AP.QSL", "APAPPRN", Form$(), Fld()
  PrintHelp "Print A/P Purchases Journal"
  
  DIM Frm AS FormInfo
  Action = 1
  
  '--Get the report options from user
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
    CASE F10Key 'SaveButton
      
      DeptNumber$ = Form$(1, 0)
      DeptNumber$ = QPTrim$(DeptNumber$)
      
      IF LEN(LTRIM$(RTRIM$(Form$(3, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(3, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(3, 0)), 1))
      END IF
      
    CASE ESC
      EXIT SUB
    END SELECT
  LOOP UNTIL Frm.KeyCode = F10Key
  
  '==Start processing register=>
  LOCATE 1, 1, 0
  
  DIM PO AS POFORMRecType2
  OpenPOEditFile POEditFile, NumEdTrans
  PrnFileNum = FREEFILE
  
  OPEN FileName$ FOR OUTPUT AS #PrnFileNum
  
  '--Get a list of active funds
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM FundGrdTot#(1 TO NumFunds)
  
  GOSUB PrintHeader1
  
  FOR Transaction = 1 TO NumEdTrans
    GET POEditFile, Transaction, PO
    PODept$ = QPTrim$(PO.REQNUM)
    
    IF NOT PO.Deleted THEN
      
      IF PODept$ = DeptNumber$ OR DeptNumber$ = "ALL" THEN
        
        IF LEFT$(PO.PONUM, 3) <> "N/A" THEN
          
          TranCnt = TranCnt + 1
          TransTotal# = TransTotal# + Round#(PO.POAMT)
          
          '--Print 1st Line - Transaction details
          LSET ToPrint$ = ""    'SPACE$(78)
          
          LSET ToPrint$ = PO.VNDRCODE
          MID$(ToPrint$, 30) = ConvDateStr1$(Num2Date(PO.PODATE))
          MID$(ToPrint$, 42) = PO.PONUM
          MID$(ToPrint$, 66) = FUsing$(STR$(PO.POAMT), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          
          '--Blank line between detail and acct'g distributions
          PRINT #PrnFileNum, "Dept #"; PO.REQNUM
          
          '--Print Distribution Label
          LSET ToPrint$ = ""
          MID$(ToPrint$, 2) = "Accounting Distribution:"
          PRINT #PrnFileNum, ToPrint$
          
          '--Print Field Titles
          LSET ToPrint$ = ""
          MID$(ToPrint$, 4) = "Account Number                                         Distribution"
          PRINT #PrnFileNum, ToPrint$
          
          '--Print Accounting Distributions
          TotTranDist# = 0
          
          '--Loop Thru distributions to print and summarize
          FOR AcctDist = 1 TO 12
            IF LEN(QPTrim$(PO.ITEMS(AcctDist).ACCTNO)) THEN
              
              '--Print this distribution
              LSET ToPrint$ = ""
              MID$(ToPrint$, 4) = PO.ITEMS(AcctDist).ACCTNO
              MID$(ToPrint$, TranCol) = FUsing$(STR$(PO.ITEMS(AcctDist).EXT), CommaFmt$)
              PRINT #PrnFileNum, ToPrint$
              
              TotTranDist# = TotTranDist# + PO.ITEMS(AcctDist).EXT
              
              LineCnt = LineCnt + 1
              IF LineCnt >= MaxLines THEN
                PRINT #PrnFileNum, FF$
                GOSUB PrintHeader1
              END IF
              
              'Sum by fund
              Found = False
              FOR Fund = 1 TO NumFunds
                FundNum$ = LEFT$(PO.ITEMS(AcctDist).ACCTNO, FundLen)
                IF FundNum$ = FundList$(Fund) THEN
                  Found = TRUE
                  FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(PO.ITEMS(AcctDist).EXT)
                  EXIT FOR
                END IF
              NEXT
              
              IF NOT Found THEN 'Oh.Shit = True
                Ok = MsgBox("GL", "NOFUND")
              END IF
            END IF              'Active transaction test
            
            
          NEXT  'Distribution
          
          '--Summary line after last distribution
          LSET ToPrint$ = ""
          MID$(ToPrint$, TranCol) = DistSumLine$
          PRINT #PrnFileNum, ToPrint$
          
          '--Transaction Distribution Totals
          LSET ToPrint$ = ""
          MID$(ToPrint$, 4) = "Total Distributed"
          MID$(ToPrint$, TranCol) = FUsing$(STR$(TotTranDist#), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          LineCnt = LineCnt + 2
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintHeader1
          END IF
          
          '--2 blank lines before next distribution
          ToPrint$ = STRING$(78, "=")
          PRINT #PrnFileNum, ToPrint$
          PRINT #PrnFileNum,
          LineCnt = LineCnt + 2
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintHeader1
          END IF
        END IF  'Not deleted test
      END IF    'Test for Dept Number Match
    END IF      'Test for Approved PO
    
    
  NEXT          'Transaction
  
  'PRINT #PrnFileNum, STRING$(80, "-")
  IF LineCnt > 45 THEN
    PRINT #PrnFileNum, FF$
  END IF
  
  '--Summary
  LSET ToPrint$ = ""
  LSET ToPrint$ = "File Totals:"
  PRINT #PrnFileNum, ToPrint$
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Number of Transactions"
  MID$(ToPrint$, 31) = FUsing$(STR$(TranCnt), "####")
  PRINT #PrnFileNum, ToPrint$
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Grand Totals"
  MID$(ToPrint$, 25) = FUsing$(STR$(TransTotal#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  
  PRINT #PrnFileNum,
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Summary by Fund:"
  PRINT #PrnFileNum, ToPrint$
  
  FOR Cnt = 1 TO NumFunds
    IF FundGrdTot#(Cnt) > 0 THEN
      LSET ToPrint$ = ""
      LSET ToPrint$ = "Fund" + " " + FundList$(Cnt)
      MID$(ToPrint$, 25) = FUsing$(STR$(FundGrdTot#(Cnt)), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      GrdTot# = GrdTot# + Round#(FundGrdTot#(Cnt))
    END IF
  NEXT
  
  LSET ToPrint$ = ""
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
  
  EXIT SUB
  
PrintHeader1:
  Title$(1) = "Appoved PO's "
  Title$(2) = "Run Date: " + DATE$
  Title$(3) = "Dept #" + DeptNumber$
  Title$(4) = "Vendor                       Date        PO No                             Amt"
  Title$(5) = STRING$(78, "-")
  
  PRINT #PrnFileNum, User$
  FOR HCnt = 1 TO 5
    PRINT #PrnFileNum, Title$(HCnt)
  NEXT
  LineCnt = 6
  RETURN
  
END SUB

SUB PrnNAEditList
  
  SHARED User$
  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "#########,.##"   'ten millions
  OverFmt$ = "#######.##"   'ten millions
  DistSumLine$ = "-------------"
  FF$ = CHR$(12)
  REDIM Title$(5)
  MaxLines = 56
  TransTotal# = 0
  TranCnt = 0
  ToPrint$ = SPACE$(78)
  
  SHARED Choice$()              'Multiple choice array
  REDIM Choice$(0 TO 3, 1)      'Choices for entry type field
  Choice$(0, 0) = "2"           'Field Number
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  
  FileName$ = "APPOREG.PRN"
  RegTitle$ = "Purchase Orders"
  TranCol = CreditCol
  CashCol = DebitCol
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
  '--Get user input on where to send report
  LibLoadDisplayForm "AP.QSL", "NAPOPRN", Form$(), Fld()

  PrintHelp "Print A/P Purchases Journal"
  
  DIM Frm AS FormInfo
  Action = 1
  
  '--Get the report options from user
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
    CASE F10Key 'SaveButton
      
      DeptNumber$ = Form$(1, 0)
      DeptNumber$ = QPTrim$(DeptNumber$)
      
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
  
  OpenAcctIdx IdxFile, NumIdxRecs
  OpenAcctFile AcctFile, NumAccts

  REDIM POChk(1 TO NumIdxRecs) AS AcctPOChkType

  FOR Cnt = 1 TO NumIdxRecs
    GET IdxFile, Cnt, AcctIdx
    IF AcctIdx.RecNum > 0 THEN
      DidCnt = DidCnt + 1
      GET AcctFile, AcctIdx.RecNum, Acct
      IF Acct.Deleted = 0 THEN
        POChk(DidCnt).Acct = Acct.Num
        POChk(DidCnt).Bgt = Acct.Bgt
        POChk(DidCnt).Encumb = Acct.Encumb
        POChk(DidCnt).Bal = Acct.Bal
      END IF
    END IF
  NEXT
  CLOSE AcctFile, IdxFile

  '==Start processing register=>
  LOCATE 1, 1, 0
  
  DIM PO AS POFORMRecType2
  OpenPOEditFile POEditFile, NumEdTrans
  PrnFileNum = FREEFILE
  
  OPEN FileName$ FOR OUTPUT AS #PrnFileNum
  
  '--Get a list of active funds
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  REDIM FundGrdTot#(1 TO NumFunds)
  
  GOSUB PrintHeader
  
  FOR Transaction = 1 TO NumEdTrans
    GET POEditFile, Transaction, PO
    PODept$ = QPTrim$(PO.REQNUM)
    
    IF NOT PO.Deleted THEN
      IF PODept$ = DeptNumber$ OR DeptNumber$ = "ALL" THEN
        IF LEFT$(PO.PONUM, 3) = "N/A" THEN
          TranCnt = TranCnt + 1
          TransTotal# = TransTotal# + Round#(PO.POAMT)
          '--Print 1st Line - Transaction details
          LSET ToPrint$ = ""    'SPACE$(78)
          
          LSET ToPrint$ = PO.VNDRCODE + " " + STR$(PO.VNDRREC)
          MID$(ToPrint$, 30) = ConvDateStr1$(Num2Date(PO.PODATE))
          MID$(ToPrint$, 42) = PO.PONUM
          MID$(ToPrint$, 66) = FUsing$(STR$(PO.POAMT), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          
          '--Blank line between detail and acct'g distributions
          PRINT #PrnFileNum, "Dept # "; PO.REQNUM
          '--Print Distribution Label
          LSET ToPrint$ = ""
          MID$(ToPrint$, 2) = "Accounting Distribution:"
          PRINT #PrnFileNum, ToPrint$
          '--Print Field Titles
          LSET ToPrint$ = ""
          MID$(ToPrint$, 4) = "Account Number                                         Distribution"
          PRINT #PrnFileNum, ToPrint$
          '--Print Accounting Distributions
          TotTranDist# = 0
          '--Loop Thru distributions to print and summarize
          FOR AcctDist = 1 TO 12
            ThisAcct$ = QPTrim$(PO.ITEMS(AcctDist).ACCTNO)
            IF LEN(ThisAcct$) THEN
              FOR AcctCnt = 1 TO DidCnt
                IF INSTR(POChk(AcctCnt).Acct, ThisAcct$) > 0 THEN
                  WhatAcct = AcctCnt
                  ThisDist# = PO.ITEMS(AcctDist).EXT
                  POChk(WhatAcct).POTotal = Round#(POChk(WhatAcct).POTotal + ThisDist#)
                  EXIT FOR
                END IF
              NEXT
              '--Print this distribution
              LSET ToPrint$ = ""
              MID$(ToPrint$, 4) = ThisAcct$
              MID$(ToPrint$, TranCol) = FUsing$(STR$(ThisDist#), CommaFmt$)

              AcctTotal# = Round#(POChk(WhatAcct).Bal + POChk(WhatAcct).POTotal + POChk(WhatAcct).Encumb)
              IF AcctTotal# > POChk(WhatAcct).Bgt THEN
                Over$ = "OVER BUDGET:" + FUsing$(STR$(ABS(POChk(WhatAcct).Bgt - AcctTotal#)), OverFmt$)
                MID$(ToPrint$, 25) = Over$
              END IF
              PRINT #PrnFileNum, ToPrint$
              
              TotTranDist# = TotTranDist# + ThisDist#
              
              LineCnt = LineCnt + 1
              IF LineCnt >= MaxLines THEN
                PRINT #PrnFileNum, FF$
                GOSUB PrintHeader
              END IF
              
              'Sum by fund
              Found = False
              FOR Fund = 1 TO NumFunds
                FundNum$ = LEFT$(PO.ITEMS(AcctDist).ACCTNO, FundLen)
                IF FundNum$ = FundList$(Fund) THEN
                  Found = TRUE
                  FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(ThisDist#)
                  EXIT FOR
                END IF
              NEXT
              
              IF NOT Found THEN 'Oh.Shit = True
                Ok = MsgBox("GL", "NOFUND")
              END IF
              
            END IF              'Active transaction test
            
          NEXT  'Distribution
          
          '--Summary line after last distribution
          LSET ToPrint$ = ""
          MID$(ToPrint$, TranCol) = DistSumLine$
          PRINT #PrnFileNum, ToPrint$
          
          '--Transaction Distribution Totals
          LSET ToPrint$ = ""
          MID$(ToPrint$, 4) = "Total Distributed"
          MID$(ToPrint$, TranCol) = FUsing$(STR$(TotTranDist#), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          LineCnt = LineCnt + 2
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintHeader
          END IF
          
          '--2 blank lines before next distribution
          ToPrint$ = STRING$(78, "=")
          PRINT #PrnFileNum, ToPrint$
          PRINT #PrnFileNum,
          LineCnt = LineCnt + 2
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintHeader
          END IF
        END IF  'Not deleted test
      END IF    'Not Correct Dept
    END IF      'Not Approved Flag
  NEXT          'Transaction
  
  'PRINT #PrnFileNum, STRING$(80, "-")
  IF LineCnt > 45 THEN
    PRINT #PrnFileNum, FF$
  END IF
  
  '--Summary
  LSET ToPrint$ = ""
  LSET ToPrint$ = "File Totals:"
  PRINT #PrnFileNum, ToPrint$
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Number of Transactions"
  MID$(ToPrint$, 31) = FUsing$(STR$(TranCnt), "####")
  PRINT #PrnFileNum, ToPrint$
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Grand Totals"
  MID$(ToPrint$, 25) = FUsing$(STR$(TransTotal#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  
  PRINT #PrnFileNum,
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Summary by Fund:"
  PRINT #PrnFileNum, ToPrint$
  
  FOR Cnt = 1 TO NumFunds
    IF FundGrdTot#(Cnt) > 0 THEN
      LSET ToPrint$ = ""
      LSET ToPrint$ = "Fund" + " " + FundList$(Cnt)
      MID$(ToPrint$, 25) = FUsing$(STR$(FundGrdTot#(Cnt)), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      GrdTot# = GrdTot# + Round#(FundGrdTot#(Cnt))
    END IF
  NEXT
  
  LSET ToPrint$ = ""
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
  
  EXIT SUB
  
PrintHeader:
  Title$(1) = "Purchase Orders : NOT YET APPROVED !!!"
  Title$(2) = "Run Date: " + DATE$
  Title$(3) = "Dept #: " + DeptNumber$
  Title$(4) = "Vendor                       Date        PO No                             Amt"
  Title$(5) = STRING$(78, "-")
  
  PRINT #PrnFileNum, User$
  FOR HCnt = 1 TO 5
    PRINT #PrnFileNum, Title$(HCnt)
  NEXT
  LineCnt = 6
  RETURN
  
END SUB

SUB Rec2Form (Form$(), RecordNumber%, Fld() AS FieldInfo, APType)
  
  OpenPOEditFile POEditFile, NumEdTrans
  
  GET POEditFile, RecordNumber, POEdit
  
  Form$(0, 0) = SPACE$(LEN(POEdit))
  
  BCopy VARSEG(POEdit), VARPTR(POEdit), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  
  'Form$(1, 0) = STR$(POEdit.DelFlag)
  'Form$(2, 0) = Num2Date(POEdit.InvDate)
  'Form$(3, 0) = POEdit.InvNum
  'Form$(4, 0) = POEdit.Ref
  'Form$(5, 0) = STR$(POEdit.InvAMT)
  'Form$(6, 0) = POEdit.BatchNum
  'Form$(7, 0) = STR$(POEdit.TCode)
  
  CLOSE POEditFile              'close the file
  
  'Rec$ = SPACE$(5)
  'LSET Rec$ = STR$(RecordNumber%)
  'QPrintRC Rec$, 25, 67, -1
  
  
END SUB

SUB ResetForm (Form$(), NumFlds)
  
  FOR F = 1 TO NumFlds
    Form$(F, 0) = ""
  NEXT
  
END SUB

SUB SaveTrans (Form$(), RecordNumber, EDMode, APType)

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(POEdit), VARPTR(POEdit), LEN(Form$(0, 0)), 0
  
  OpenPOEditFile POEditFile, NumEdTrans
  
  IF EDMode = 1 THEN
    Rec = NumEdTrans + 1
  ELSE
    Rec = RecordNumber
  END IF
'POFORMRecType2
  
  PUT POEditFile, Rec, POEdit   'write it
  CLOSE POEditFile              'close the file
  
END SUB

SUB ShowEdMode (EDMode%)
  
  Status$ = SPACE$(24)
  F2Ed$ = "  F2=Edit  "
  F2New$ = "  F2=New   "
  ActiveColor = 112
  InActiveColor = 115
  
  SELECT CASE EDMode
  CASE 0
    LSET Status$ = "Edit an Existing Record"
    F2BtnCap$ = F2New$
    F5BtnColor = ActiveColor
    
  CASE 1
    LSET Status$ = "Add a New Record"
    F2BtnCap$ = F2Ed$
    F5BtnColor = InActiveColor
    
  END SELECT
  
  QPrintRC Status$, 25, 2, -1
  QPrintRC F2BtnCap$, 25, 40, -1
  MPaintBox 25, 53, 25, 63, F5BtnColor          'F5 Button enable / disable
  
END SUB

SUB ValidateForm (Form$(), AcctFlds(), BadFld)
  'Use this sub to validate fields
  
  '--Check PONum
  IF LEN(QPTrim$(Form$(2, 0))) = 0 THEN
    BadFld = 2
    EXIT SUB
  END IF
  
  '--Check Dept Number
  IF LEN(QPTrim$(Form$(2, 0))) = 0 THEN
    BadFld = 2
    EXIT SUB
  END IF
  
  
  '--Check Date
  IF Date2Num(Form$(4, 0)) < -32000 THEN
    BadFld = 4
    EXIT SUB
  END IF
  
  '--Check Distributions
  FOR F = 1 TO 12               'UBOUND(AcctFlds, 1)
    AcctFldNum = AcctFlds(F)
    AmtFldNum = AcctFldNum - 1
    
    '--check for an account number without a distribution amount
    'IF LEN(Form$(AcctFldNum, 0)) > 0 THEN
    '  IF Value(Form$((AmtFldNum), 0), ECode) = 0 THEN
    '    BadFld = AcctFldNum
    '    EXIT FOR
    '  END IF
    'END IF
    
    '--check for a distribution without an account number
    IF Value(Form$(AmtFldNum, 0), ECode) <> 0 THEN
      IF LEN(QPTrim$(Form$(AcctFldNum, 0))) = 0 THEN
        BadFld = (AcctFldNum - 1)
      END IF
    END IF
  NEXT
  
  
END SUB

SUB VoidPOTrans (VoidTransRecNum&)
  
  REDIM APLedger(1) AS APLedger81RecType
  REDIM DistRec(1) AS APDistRecType
  LdRecLen = LEN(APLedger(1))
  DistRecLen = LEN(DistRec(1))
  
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  GET APLedgerFile, VoidTransRecNum&, APLedger(1)
  APLedger(1).TrCode = -4
  PUT APLedgerFile, VoidTransRecNum&, APLedger(1)
  CLOSE APLedgerFile
  
  POIFFile$ = "POVDIF.DAT"
  KillFile POIFFile$
  REDIM GLIFRec(1) AS GLTransRecType
  GLIFRecLen = LEN(GLIFRec(1))
  GLIFFile = FREEFILE
  OPEN POIFFile$ FOR RANDOM AS GLIFFile LEN = GLIFRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen
  NextDist& = APLedger(1).FrstDist
  DistCnt = 0
  
  DO UNTIL NextDist& = 0
    GET APDistFile, NextDist&, DistRec(1)
    IFRec = IFRec + 1
    
    '--Make Debit side of entry
    GLIFRec(1).Src = "VD" + ConvDateStr$(DATE$)
    AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + EncAcct$
    GLIFRec(1).AcctNum = AcctNum$
    GLIFRec(1).TrDate = Date2Num(DATE$)
    GLIFRec(1).Desc = "CANCELLED PO"            'APLedger(1).PONum
    GLIFRec(1).Ref = APLedger(1).PONUM
    GLIFRec(1).CrAmt = 0
    GLIFRec(1).DrAmt = DistRec(1).DistAmt
    PUT GLIFFile, IFRec, GLIFRec(1)
    
    IFRec = IFRec + 1
    'AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
    'GLIFRec(1).AcctNum = AcctNum$
    GLIFRec(1).AcctNum = DistRec(1).DistAcctNum
    GLIFRec(1).CrAmt = DistRec(1).DistAmt
    GLIFRec(1).DrAmt = 0
    PUT GLIFFile, IFRec, GLIFRec(1)
    
    NextDist& = DistRec(1).NextDist
    
  LOOP
  
  CLOSE
  
  Post2PO POIFFile$, BadAcct
  KillFile POIFFile$
  
END SUB

SUB XXALTF3
  
  '--This is just holding this code temporarily
  'This chunk of code deletes a distribution line when Alt+F3 is pressed
  'This needs to occur before prev fld test
  'IF Frm(1).FldNo <> 7 AND Frm(1).PrevFld = 7 AND LEN(QPTrim$(Form$(7, 0))) = 0 THEN
  '  Frm(1).FldNo = 7
  '  Action = 2
  '  Frm(1).KeyCode = 0
  'END IF
  
  'SELECT CASE Frm(1).KeyCode '!!!
  'CASE -106   'Alt+F3
  '  IF Frm(1).FldNo > Header THEN     'if we're on an account field
  '    'for each accounting distribution
  '    FOR Cnt = 1 TO 36
  '      'if this acct field is less than the current field
  '      'set the row counter
  '      IF AcctFlds(Cnt) <= Frm(1).FldNo THEN
  '        FldRow = FldRow + 1 '+AcctFlds(Cnt)
  '      ELSE
  '        FOR Cnt3 = AcctFlds(FldRow) TO AcctFlds(FldRow) + 3
  '          LSET Form$(Cnt3, 0) = ""
  '        NEXT
  '
  '        FOR Cnt3 = FldRow TO 35
  '          LSET Form$(AcctFlds(Cnt3), 0) = Form$(AcctFlds(Cnt3 + 1), 0)
  '          LSET Form$(AcctFlds(Cnt3) + 1, 0) = Form$((AcctFlds(Cnt3 + 1) + 1), 0)
  '          LSET Form$(AcctFlds(Cnt3) + 2, 0) = Form$((AcctFlds(Cnt3 + 1) + 2), 0)
  '          LSET Form$(AcctFlds(Cnt3) + 3, 0) = Form$((AcctFlds(Cnt3 + 1) + 3), 0)
  '        NEXT
  '        'Cnt is at 36 after 35 times thru above loop
  '
  '        LSET Form$(AcctFlds(Cnt3), 0) = ""
  '        LSET Form$(AcctFlds(Cnt3) + 1, 0) = ""
  '        LSET Form$(AcctFlds(Cnt3) + 2, 0) = ""
  '        LSET Form$(AcctFlds(Cnt3) + 3, 0) = ""
  '
  '        Frm(1).FldNo = AcctFlds(FldRow)
  '        Frm(1).PrevFld = Frm(1).FldNo - 1
  '        FldRow = 0
  '        Action = 1
  '        EXIT FOR
  '      END IF
  '    NEXT
  '  END IF
  
  
END SUB

