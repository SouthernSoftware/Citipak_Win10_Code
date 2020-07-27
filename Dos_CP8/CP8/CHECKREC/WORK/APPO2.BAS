DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, X%, Y%, Wdth%, TitColor%)
DEFINT A-Z
DECLARE SUB VoidPOTrans (VoidTransRecNum&)
DECLARE SUB CancelPO ()
DECLARE SUB OpenPOTransFile (TransFileNum%, NumTrans&)
DECLARE SUB Post2PO (FileName$, BadTrans%)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION EDCheck4Vendor% (Form$(), Terms$)
DECLARE FUNCTION EDCheck4Acct% (Form$(), AcctFldNo%)
DECLARE FUNCTION FindVendorRec% (VendorCode$)
DECLARE SUB PostPOTrans ()
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPVendorFile (APVendorFile%, NumVRecs%, VendorRecLen%)
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenVendorIdx (VendorIdxFile%, NumActiveVendors%)
DECLARE FUNCTION ChkForVendor% (ChkVendor$)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
'DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
'DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
'DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'pageinfo.BI'                  'Form Page info
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'SCRAUX.bi'
'$INCLUDE: 'ap.bi'

DECLARE SUB EdEntry (JrnlCode%)
DECLARE SUB OpenPOEditFile (POEditFile%, NumEdTrans%)
DECLARE SUB NewCJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList (APType)
DECLARE SUB DisplayTotals ()
'DECLARE SUB PostTrans (APType)
DECLARE SUB Check4BlankFlds (Form$(), BadFld%)
DECLARE FUNCTION ListRecs% (APType)
DECLARE SUB ResetForm (Form$(), NumFlds%)
DECLARE SUB ShowEdMode (EdMode%)
DECLARE SUB Rec2Form (Form$(), RecordNumber%, Fld() AS ANY, APType)
DECLARE SUB SaveTrans (Form$(), RecordNumber%, EdMode%, APType)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
  
  CONST False = 0, True = NOT False
  
  DIM SHARED TotLen, FundLen, AcctLen, DetLen
  DIM SHARED CashAcct$, APAcct$, EncAcct$, User$ ', FBAcct$
  DIM SHARED AcctIdx AS GLAcctIndexType
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED POEdit AS APPOType
  DIM SHARED POTrans AS GLTransRecType

  '$INCLUDE: 'init.bi'                      'Get GL system info
  
  TYPE FLen
    V AS STRING * 42
  END TYPE
   
  'STOP
  'Post2PO "TR2POST.DAT", BadTrans
  'IF BadTrans THEN
  '  STOP
  'END IF
  'END
  
  PIProcessMenu 1
  ClrBG
  RUN "APMENU"

SUB CancelPO

  REDIM APLedgerRec(1) AS APLedger81RecType
  LdRecLen = LEN(APLedgerRec(1))
  DIM VendorIdx AS VendorIdxRecType
  DIM Vendor AS VendorRecType

  FormName$ = "VOIDPO"
  LibName$ = "AP"
  REDIM Frm(1) AS FormInfo
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                         'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  Frm(1).FldNo = 1                    'Start editing on field #1
  Frm(1).InsStat = False              'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0                  'Set form starting element to 0 and
  Action = 1
  FirstTime = True

  LibFile2Scrn LibName$, FormName$, MonoCode, -1, ErrCode
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).FldNo > 1 AND Frm(1).PrevFld = 1 THEN   'User is moving around the form
      ChkVendor$ = QPTrim$(Form$(1, 0))
      GOSUB Check4Vendor
      IF NOT Abort THEN
        GOSUB FindPO
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

  GET VendorFile, VoidVendorRecNum&, Vendor
  NextTrans& = Vendor.FrstTran
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

  REDIM POList(1 TO POCnt) AS FLen       'Dim array to hold trans list

  MatchCnt = 0
  NextTrans& = Vendor.FrstTran

  DO UNTIL NextTrans& = 0
    GET APLedgerFile, NextTrans&, APLedgerRec(1)
      IF APLedgerRec(1).TrCode = 4 THEN
        MatchCnt = MatchCnt + 1
        LSET POList(MatchCnt).V = List$
        POList(MatchCnt).V = QPTrim$(APLedgerRec(1).PONum)
        MID$(POList(MatchCnt).V, 11) = Num2Date(APLedgerRec(1).TRDate)
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
  QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
  Choice = 1
  DO
    'call vertmenu to display matching recs to choose from
    VertMenuT POList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      Abort = True
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
  MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen

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
      GET VendorFile, VendorIdx.RecNum, Vendor
      IF Vendor.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 14) = Vendor.VNAME
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

SUB Check4BlankFlds (Form$(), BadFld)
  
  IF Date2Num(Form$(6, 0)) < -32000 THEN
    BadFld = 6
    EXIT SUB
  END IF
  
  FOR FieldNum = 2 TO 5
    
    IF LEN(QPTrim$(Form$(FieldNum, 0))) = 0 THEN
      BadFld = FieldNum
      EXIT FOR
    END IF
    
  NEXT
  
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
  FOR Cnt = 1 TO NumIdxRecs               'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, Acct
      IF Acct.Typ = "E" THEN
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
      PLAY "L32O3EC"            'Acct NOT found
      Ok = MsgBox("GL.QSL", "NOACCT")
      EDCheck4Acct = False
    CASE 1
      AcctRecNum = CVI(RIGHT$(AcctList(1).V, 2))

      GOSUB GetAcctInfo
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
      DO         'call vertmenu to display matching recs to choose from
        VertMenuT AcctList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
        IF Ky$ = CHR$(27) THEN
          EDCheck4Acct = False
          EXIT DO
        END IF
        IF Choice > 0 THEN
          AcctRecNum = CVI(RIGHT$(AcctList(Choice).V, 2))
          GOSUB GetAcctInfo
          EDCheck4Acct = True
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
  LSET Form$(AcctFld + 1, 0) = Acct.Title
  LSET Form$(AcctFld + 2, 0) = MKI$(AcctRecNum)
RETURN



END FUNCTION

FUNCTION EDCheck4Vendor% (Form$(), Terms$)

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

  REDIM VendorList(1 TO NumActiveVendors) AS FLen       'Dim array to hold trans list
  
  FOR Cnt = 1 TO NumActiveVendors               'Read the file
    GET VendorIdxFile, Cnt, VendorIdx
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN

      GET VendorFile, VendorIdx.RecNum, Vendor
      IF Vendor.DELFLAG = 0 THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 13) = Vendor.VNAME
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
    LSET Form$(3, 0) = MID$(VendorList(1).V, 13, 28)
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
    DO         'call vertmenu to display matching recs to choose from
      VertMenuT VendorList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        CLOSE
        EDCheck4Vendor = False
        EXIT DO
      END IF
      IF Choice > 0 THEN
        LSET Form$(2, 0) = VendorList(Choice).V
        LSET Form$(3, 0) = MID$(VendorList(Choice).V, 13, 28)
        EDCheck4Vendor = True
        '  SaveField 2, Form$(), Fld(), BadFld
        EXIT DO
      END IF
    LOOP

  END SELECT

  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  QPrintRC TName$, 3, 27, -1
  ERASE VendorList, ScrBuf

END FUNCTION

SUB EdEntry (APType)

  'STATIC  OIDate, ODDate, OPDate
  
  SHARED Choice$()
  
  SHARED VendorIdx AS VendorIdxRecType
  SHARED Vendor AS VendorRecType
  
  REDIM Choice$(0 TO 2, 0)
  '  TName$ = SPACE$(20)
  
  Choice$(0, 0) = "7"
  Choice$(1, 0) = "1 Open Payable"
  'Choice$(2, 0) = "2 Suspended Inv."
  
  NumOfScrns = 3
  
  REDIM Frm(1)    AS FormInfo
  REDIM FormName$(1 TO NumOfScrns)
  CommaFormat$ = "#########,.##"
  

  'SELECT CASE APType
  'CASE 1        'Cash Reciepts
    FOR Scr = 1 TO NumOfScrns
      FormName$(Scr) = "PO" + LTRIM$(STR$(Scr))
    NEXT
    Header = 6
    TotalAmtFld = 5

  'CASE 2        'Cash Disbursements
  '  FOR Scr = 1 TO NumOfScrns
  '    FormName$(Scr) = "AP" + LTRIM$(STR$(Scr))
  '  NEXT
  '  Header = 6
  '  TotalAmtFld = 5
  '  'CASE 3  'AP
  '  'AP stuff here
  'END SELECT
  
  '-- Get total number of fields
  NumFlds = -1
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
  
  ' This is broken in my formedit
  REDIM AcctFlds(1 TO 36)
  FOR F = 1 TO NumFlds
    IF QPTrim$(Fld(F).ScratchS) = "A" THEN
      Cnt = Cnt + 1
      AcctFlds(Cnt) = F
    END IF
  NEXT
  Frm(1).PageDisable = True
  
  '--get first active record number

  OpenPOEditFile POEditFile, NumEdTrans
  
  CurrRec = 0
  NextRec = 1
  RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
  CLOSE POEditFile
  IF RecNum = 0 THEN
    EdMode = 1
  ELSE
    EdMode = 0
  END IF
  '===
  
  Scr = 1       'Set screen number to one
  LibFile2Scrn "AP", FormName$(Scr), MonoCode, -1, ErrCode      'display screen
  'GOSUB ShowRecNo
  IF EdMode = 0 THEN
    Rec2Form Form$(), RecNum, Fld(), APType
  END IF
  
  ShowEdMode EdMode
  
  Action = 1
  FirstTime = -1

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = 0
     ' Action = 2
    END IF

    'This chunk of code deletes a distribution line when Alt+F3 is pressed
    'This needs to occur before prev fld test
    'Gets a list of acct fields from disk file.

    IF Frm(1).FldNo <> 7 AND Frm(1).PrevFld = 7 AND LEN(QPTrim$(Form$(7, 0))) = 0 THEN
      Frm(1).FldNo = 7
      Action = 2
      Frm(1).KeyCode = 0
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE -106   'Alt+F3
      
      IF Frm(1).FldNo > Header THEN             'if we're on an account field
        
        'for each accounting distribution
        FOR Cnt = 1 TO 36
          'if this acct field is less than the current field
          'set the row counter
          IF AcctFlds(Cnt) <= Frm(1).FldNo THEN
            FldRow = FldRow + 1 '+AcctFlds(Cnt)
          ELSE
            FOR Cnt3 = AcctFlds(FldRow) TO AcctFlds(FldRow) + 3
              LSET Form$(Cnt3, 0) = ""
            NEXT
            
            FOR Cnt3 = FldRow TO 35
              LSET Form$(AcctFlds(Cnt3), 0) = Form$(AcctFlds(Cnt3 + 1), 0)
              LSET Form$(AcctFlds(Cnt3) + 1, 0) = Form$((AcctFlds(Cnt3 + 1) + 1), 0)
              LSET Form$(AcctFlds(Cnt3) + 2, 0) = Form$((AcctFlds(Cnt3 + 1) + 2), 0)
              LSET Form$(AcctFlds(Cnt3) + 3, 0) = Form$((AcctFlds(Cnt3 + 1) + 3), 0)
            NEXT
            'Cnt is at 36 after 35 times thru above loop
            
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
          ShowEdMode EdMode
          ResetForm Form$(), NumFlds
          FirstTime = -1
          Frm(1).FldNo = 1
          Action = 1
          QPrintRC "New  ", 25, 68, -1
          GOSUB ShowDistTot
          
        CASE 1  'currently adding records
          'check to see if there are any active transactions
          OpenPOEditFile POEditFile, NumEdTrans
          CurrRec = 0: NextRec = 1
          RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE POEditFile
          
          IF RecNum = 0 THEN    'No active records can't edit
            EdMode = 1
          ELSE
            EdMode = 0          'switch to edit mode
            ShowEdMode EdMode
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
        END SELECT              'Edmode
      END IF
    CASE ELSE
      
    END SELECT
    
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN      'User is moving around the form
      IF Frm(1).PrevFld = 2 AND Frm(1).FldNo <> 2 THEN
        VendorOK = EDCheck4Vendor(Form$(), Terms$)
        IF VendorOK THEN
'          Formula$ = CalcDate$ + Terms$
'          Form$(9, 2) = Formula$
'          SaveField 8, Form$(), Fld(), BadFld
'          IF EdMode = 0 THEN
'            CALL CalcFields(1, 8, Form$(), Fld())
'          END IF
          Frm(1).FldNo = 3
          Action = 2
        ELSE
          Frm(1).FldNo = 1
          Action = 2
        END IF
      END IF
      GOSUB ShowDistTot         'Update Distribution totals
      
      '--Check the field we're in to see if it is an acct field.
      '  if so and if there is data in the field, remove dashes for user
      '  set stripped acct string in form

      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
        LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
        Action = 1
      END IF
      
      '--the following checks the field we have just left to see
      '  if it is an account field. if it is then
      '  if the field is not blank, the acct number is formatted with dashes
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
      END IF
      
    CASE F5KEY
      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        IF EdMode = 0 THEN
          RecNum = ListRecs(APType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), APType
            Frm(1).FldNo = 1
            Action = 1
          END IF
        END IF
        
      END IF
      
    CASE F7KEY
      IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN       '= AcctField
        BEEP
      END IF
      
    CASE F10Key
      '--Make sure acct fields are formatted correctly if user
      '  Preses F10 with cursor in an account field
      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
        IF LEN(QPTrim$(Form$(Frm(1).FldNo, 0))) > 0 THEN
          TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).FldNo, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).FldNo, 0) = TempAcct$
          'Fld(Frm(1).PrevFld + 3).Protected = False
          Frm(1).FldNo = 1
          'Action = 1
        END IF
      END IF
      'send user back thru loop to verify acct before saving
      Action = 1
      'OIDate = Date2Num(Form$(8, 0))
      'ODDate = Date2Num(Form$(9, 0))
      'OPDate = Date2Num(Form$(10, 0))
      
      SaveFlag = True
      
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
          ShowEdMode EdMode
          'Frm(1).PrevFld = 1
          Scr = 1
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
          ShowEdMode EdMode
          ' Frm(1).PrevFld = 1
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
      LibFile2Scrn "GL", FormName$(Scr), MonoCode, -2, ErrCode
      
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
          PressButton F3Key, 21, 22, 37         'F3
        CASE 38 TO 56
          PressButton -106, 21, 38, 56      'Alt+F3
        CASE 57 TO 71
          PressButton ESC, 21, 57, 71
        END SELECT
        
      CASE 25
        SELECT CASE Frm(1).MCol
        CASE 27 TO 39
          PressButton F2Key, 25, 27, 39
        CASE 40 TO 52
          PressButton F5KEY, 25, 40, 52
        CASE 53 TO 59
          PressButton CtrlL, 25, 53, 59
        CASE 73 TO 79
          PressButton CtrlR, 25, 73, 79
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL Frm(1).KeyCode = ESC
  
  'End of FormEdit Loop
  EXIT SUB
  
  
CheckForValidAcctNum:
  AcctFld = Frm(1).PrevFld
  SELECT CASE EDCheck4Acct%(Form$(), AcctFld%)
    CASE True
      SaveField AcctFld, Form$(), Fld(), BadFld       'Write the flag to the form$
      SaveField AcctFld + 1, Form$(), Fld(), BadFld     'Write the flag to the form$
      SaveField AcctFld + 2, Form$(), Fld(), BadFld     'Write the flag to the form$
      Action = 2
    CASE ELSE
      Frm(1).FldNo = AcctFld             'Put user back on acct field
      LSET Form$(Frm(1).FldNo, 0) = ""
      LSET Form$(Frm(1).FldNo + 1, 0) = ""
      LSET Form$(Frm(1).FldNo + 2, 0) = ""
      Action = 2
  END SELECT
  
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
  Check4BlankFlds Form$(), BadFld               'make sure no blank fields
  IF BadFld = 0 THEN            'if no blanks, make sure
    GOSUB ShowDistTot           'distribution are in bal
    
    IF UnDist# = 0 AND Amt# <> 0 THEN
      Form$(1, 0) = "0"         'Set deleted flag to false
      SaveField 1, Form$(), Fld(), BadFld       'Write the flag to the form$
      SaveTrans Form$(), RecNum, EdMode, APType 'Save the transaction
      Ok = MsgBox%("GL.QSL", "SAVED")           'tell user we saved it
      
      IF EdMode = 1 THEN        'if in Add Mode, clear
        ResetForm Form$(), NumFlds              ' the form for the next entry
        Frm(1).FldNo = 1
      END IF
      
    ELSE
      Ok = MsgBox%("GL.QSL", "BADDIST")         'Distributions out of balance
      'tell user
    END IF
    
  ELSE
    Ok = MsgBox%("GL.QSL", "NOBLANK")           'Tell user no blank fields
    Frm(1).FldNo = BadFld
    
  END IF
  
  Action = 1
  FirstTime = -1
  SaveFlag = False
  
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
  Rec = CurrRec
  
  DO
    
    Rec = Rec + NextRec         'Set file pointer to next record
    
    IF Rec > NumRecs OR Rec <= 0 THEN           'test for beg or end of file
      Found = 0 'if no more records then get out
      EXIT DO
    END IF
    
    GET FileNum, Rec, POEdit    'Get the record
    
    IF POEdit.DELFLAG <> -1 THEN                'Ok if not deleted
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
  
  QPrintRC "List ", 25, 68, -1
  
  OpenPOEditFile POEditFile, NumEdTrans
  
  REDIM TransList$(1 TO NumEdTrans)             'Dim array to hold trans list
  
  Active = 0    'Counter for Active Transactions
  
  FOR Cnt = 1 TO NumEdTrans     'Read the file
    GET POEditFile, Cnt, POEdit
    IF NOT POEdit.DELFLAG THEN  'if active add to list
      Active = Active + 1
      
      List$ = SPACE$(75)
      
      LSET List$ = POEdit.Vendor
      MID$(List$, 12) = POEdit.PONum
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
  
  ListDesc$ = "   Vendor    Invoice                   Inv Date        Total Amt     "
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

SUB PIProcessMenu (JrnlType)
  
  '$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
  
  TextCursor -2, -2             'Define the Mouse cursor
  'ShowCursor    'Show it
  
  '***Master Menu
  
  DIM MenuText$(1 TO 5)
  
  'SELECT CASE JrnlType
  'CASE 1
    MenuText$(1) = "Enter/Edit Purchase Orders"
    MenuText$(2) = "Print Purchase Order Journal"
    MenuText$(3) = "Post Entries"
    MenuText$(4) = "Cancel Open Purchase Order"
    MenuText$(5) = "Quit"

  'CASE 2
  '  MenuText$(1) = "Enter New Journal Entries"
  '  MenuText$(1) = "Enter/Edit Invoices"
  '  MenuText$(2) = "Print Invoice Register"
  '  MenuText$(3) = "Post Entries"
  '  MenuText$(4) = "Quit"

  'END SELECT
  
  
  MaxLen = 0  'Set menu width to zero
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

    'TitleBox 6, Col, MaxLen + 3, Title$, Cnf
    'TitleBox 17, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    PrintTitle User$
    
    'IF JrnlType = 2 THEN
    '  PrintHelp "Invoice Processing"
    'ELSEIF JrnlType = 1 THEN
      PrintHelp "P.O. Processing"
    'END IF
    
    LOCATE Row, Col, 0
    VertMenu MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
    CASE 1
      ClrBG
      EdEntry JrnlType          ', AcctFlds()
    CASE 2
      ClrBG
      PrnEditList JrnlType
    CASE 3
      ClrBG
      PostPOTrans
    CASE 4
      ClrBG
      CancelPO
    CASE 5
      CLS : END
    CASE ELSE
    END SELECT
    
  LOOP
  
  'HideCursor
'  DispLoad
  
  'RUN "glmenu"

  
END SUB

SUB PostPOTrans

'  SHARED APAcct$

  SHARED Vendor AS VendorRecType
  REDIM APDistRec(1) AS APDistRecType
  REDIM APLedgerRec(1) AS APLedger81RecType
  REDIM Tr2Post(1) AS GLTransRecType
  
  LedgerRecLen = LEN(APLedgerRec(1))
  DistRecLen = LEN(APDistRec(1))
  VendorRecLen = LEN(Vendor)

  '--Paint a background
  FOR Cnt = 2 TO 24
    QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
  NEXT
  
  '--Verify that there are transactions
  OpenPOEditFile POEditFile, NumEdTrans
  
  '--Test for valid Vendor Numbers
  FOR Cnt = 1 TO NumEdTrans     'summarize the file totals
    GET POEditFile, Cnt, POEdit
    IF NOT POEdit.DELFLAG THEN
      VendRecNum = FindVendorRec(POEdit.Vendor)
      IF VendRecNum > 0 THEN
        Active = Active + 1
      ELSE
        AbortFlag = True
        COLOR 12: CLS
        BEEP
        PRINT "Error: Invalid Vendor Code"
        PRINT "Unable to locate vendor: "; POEdit.Vendor
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
    GET POEditFile, Cnt, POEdit
    
    IF NOT POEdit.DELFLAG THEN
      FOR AcctDist = 1 TO 36
        IF LEN(QPTrim$(POEdit.Dist(AcctDist).DACN)) = 0 THEN
          '--No more distributions now create other side of transaction
          FOR Fund = 1 TO NumFunds
            IF TrFundSum#(Fund) <> 0 THEN
              RecordNum = RecordNum + 1
              'STOP
              Tr2Post(1).AcctNum = FundList$(Fund) + EncAcct$
              Tr2Post(1).TRDate = POEdit.PODATE
              Tr2Post(1).Desc = POEdit.VENDNAME
          
              Tr2Post(1).Ref = POEdit.PONum
              Tr2Post(1).DrAmt = 0
              Tr2Post(1).CrAmt = TrFundSum#(Fund)
              Tr2Post(1).Src = "PO" + ConvDateStr$(DATE$)
              PUT AP2Post, RecordNum, Tr2Post(1)
            END IF              'Fund summary <> 0
          NEXT  'fund
          
          EXIT FOR              'NumFunds
          
        ELSE
          '--Process each accounting distribution
          RecordNum = RecordNum + 1
          
          Tr2Post(1).AcctNum = POEdit.Dist(AcctDist).DACN
          Tr2Post(1).TRDate = POEdit.PODATE
          'Tr2Post(1).Desc = POEdit.VENDOR
          Tr2Post(1).Desc = POEdit.VENDNAME
          Tr2Post(1).Ref = POEdit.PONum
          Tr2Post(1).DrAmt = POEdit.Dist(AcctDist).DAMT
          Tr2Post(1).CrAmt = 0
          Tr2Post(1).Src = "PO" + ConvDateStr$(DATE$)
          
          PUT AP2Post, RecordNum, Tr2Post(1)
          
          '--Add this distribution to proper fund
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(POEdit.Dist(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(POEdit.Dist(AcctDist).DAMT)
              'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(POEdit.DIST(AcctDist).DAMT)
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
    GET POEditFile, Cnt, POEdit
    IF NOT POEdit.DELFLAG THEN

      '--Update PO in vendor link to apledger
      NumLedgerRecs = NumLedgerRecs + 1

      '--????????????????????????????????????????????????????????????
      '--Get the record number instead of finding it
      '--add error trap in case they deleted a vendor prior to post
      VendRecNum = FindVendorRec(POEdit.Vendor)
      OpenVendorFile VendorFile, NumVRecs
      GET VendorFile, VendRecNum, Vendor
      
      'PrevVendTrans = VENDOR.LastPO
      'FrstVendTrans = VENDOR.FrstPO
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

      '--Post transaction to apledger
      APLedgerRec(1).VRecNum = VendRecNum
      APLedgerRec(1).VendorCode = POEdit.Vendor
      APLedgerRec(1).TRDate = POEdit.PODATE
      APLedgerRec(1).DOCNum = POEdit.PONum
      APLedgerRec(1).PONum = POEdit.PONum
      'APLedgerRec(1).DueDate = POEdit.DueDate
      APLedgerRec(1).TrCode = 4
      'APLedgerRec(1).PayCode = VAL(POEdit.PayCode)
      'APLedgerRec(1).GLDistDate = POEdit.DISTDATE
      APLedgerRec(1).Amt = POEdit.POAMT
      APLedgerRec(1).NextTrans = 0
      APLedgerRec(1).FrstDist = NumDistRecs& + 1
      
      '--Find the number of children to tag along
      FOR Cnt2 = 1 TO 36
        IF LEN(QPTrim$(POEdit.Dist(Cnt2).DACN)) = 0 THEN
          LastActive = Cnt2 - 1
          EXIT FOR
        END IF
      NEXT

      '--Post and link distributions
      FOR Cnt2 = 1 TO LastActive
        REDIM APDistRec(1) AS APDistRecType
        NumDistRecs& = NumDistRecs& + 1
        APDistRec(1).APLedgerRec = NumLedgerRecs
        APDistRec(1).DistAcctRec = CVI(POEdit.Dist(Cnt2).DACREC)
        APDistRec(1).DistAcctNum = POEdit.Dist(Cnt2).DACN
        APDistRec(1).DistAmt = POEdit.Dist(Cnt2).DAMT
        'APDistRec(1).DistCRAmt = 0
        IF Cnt2 < LastActive THEN
          APDistRec(1).NextDist = NumDistRecs& + 1
        ELSE
          APDistRec(1).NextDist = 0
        END IF
        PUT APDistFile, NumDistRecs&, APDistRec(1)
      NEXT

      APLedgerRec(1).LastDist = NumDistRecs&
      PUT APLedgerFile, NumLedgerRecs, APLedgerRec(1)

      'CCnt = CCnt + 1 what is this count for

    END IF

  NEXT
  CLOSE

  '--Post ok.  Kill edit and interface files
  'IF BadTrans = 0 THEN
    KILL "APPED.DAT"
    KILL "APPOIF.DAT"
  'END IF

END SUB

SUB PrnEditList (APType)

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
  
  SHARED Choice$()              'Multiple choice array
  REDIM Choice$(0 TO 3, 1)      'Choices for entry type field
  Choice$(0, 0) = "2"           'Field Number
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  
  SELECT CASE APType
  CASE 1
    'Purchase Orders
    FileName$ = "APPOREG.PRN"
    RegTitle$ = "Purchase Orders"
    TranCol = CreditCol
    CashCol = DebitCol
  CASE 2
    'Invoices
    FileName$ = "APINVREG.PRN"
    RegTitle$ = "A/P Purchases Journal"
    TranCol = CreditCol
    'CashCol = CreditCol
  END SELECT
  
  Title$(1) = RegTitle$
  Title$(2) = "Run Date: " + DATE$
  Title$(3) = ""
  Title$(4) = "Vendor                       Date        PO No                             Amt"
  Title$(5) = STRING$(78, "-")
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
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
    GET POEditFile, Transaction, POEdit
    
    IF NOT POEdit.DELFLAG THEN
      TranCnt = TranCnt + 1
      TransTotal# = TransTotal# + Round#(POEdit.POAMT)
      
      '--Print 1st Line - Transaction details
      ToPrint$ = SPACE$(78)
      LSET ToPrint$ = POEdit.Vendor
      MID$(ToPrint$, 12) = LEFT$(POEdit.VENDNAME, 16)
      MID$(ToPrint$, 30) = ConvDateStr1$(Num2Date(POEdit.PODATE))
      MID$(ToPrint$, 42) = POEdit.PONum
      MID$(ToPrint$, 66) = FUsing$(STR$(POEdit.POAMT), CommaFmt$)
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
      
      '--Print Accounting Distributions
      TotTranDist# = 0
      
      '--Loop Thru distributions to print and summarize
      FOR AcctDist = 1 TO 36
        
        '--no more distributions when we find a blank Acct Number field
        IF LEN(QPTrim$(POEdit.Dist(AcctDist).DACN)) = 0 THEN
          
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
          '--Last distribution for this transaction - exit loop
          EXIT FOR
          
        ELSE    '--we have active transactions
          '--Add distribution to total
          TotTranDist# = TotTranDist# + Round#(POEdit.Dist(AcctDist).DAMT)
          
          '--Add distribution to proper fund
          Found = False
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(POEdit.Dist(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              Found = True
              FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(POEdit.Dist(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT
          
          IF NOT Found THEN  'Oh.Shit = True
            Ok = MsgBox("GL", "NOFUND")
          END IF
          
          '--Print this distribution
          ToPrint$ = SPACE$(78)
          MID$(ToPrint$, 4) = POEdit.Dist(AcctDist).DACN
          MID$(ToPrint$, 21) = POEdit.Dist(AcctDist).DACNM
          MID$(ToPrint$, TranCol) = FUsing$(STR$(POEdit.Dist(AcctDist).DAMT), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          LineCnt = LineCnt + 1
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintHeader
          END IF
          
        END IF  'Active transaction test
        
      NEXT      'Distribution
      
      '--2 blank lines before next distribution
      ToPrint$ = STRING$(78, "=")
      PRINT #PrnFileNum, ToPrint$
      PRINT #PrnFileNum,
      LineCnt = LineCnt + 2
        IF LineCnt >= MaxLines THEN
          PRINT #PrnFileNum, FF$
          GOSUB PrintHeader
        END IF
      END IF      'Not deleted test
    
  NEXT          'Transaction
  
  'PRINT #PrnFileNum, STRING$(80, "-")
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
    IF FundGrdTot#(Cnt) > 0 THEN
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

EXIT SUB

PrintHeader:
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

SUB SaveTrans (Form$(), RecordNumber, EdMode, APType)
  
  OpenPOEditFile POEditFile, NumEdTrans
  
  IF EdMode = 1 THEN
    Rec = NumEdTrans + 1
  ELSE
    Rec = RecordNumber
  END IF
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(POEdit), VARPTR(POEdit), LEN(Form$(0, 0)), 0
  
  'POEdit.DelFlag = VAL(Form$(1, 0))
  'POEdit.InvDate = Date2Num(Form$(2, 0))
  'POEdit.InvNum = Form$(3, 0)
  'POEdit.Ref = Form$(4, 0)
  'POEdit.InvAMT = Round#(Value(Form$(5, 0), ErrCode))
  'POEdit.BatchNum = Form$(6, 0)
  'POEdit.TCode = VAL(Form$(7, 0))
  'STOP
  
  PUT POEditFile, Rec, POEdit   'write it
  CLOSE POEditFile              'close the file
  
  
  '--stuff--->
  '   POEdit.Src = "CR" + ConvDateStr$(DATE$)
  '
  '   Amt# = Value(Form$(Amt, 0), ErrCode)      'Convert Amt$ to Double
  '   IF ErrCode THEN
  '      CLS
  '      CLOSE
  '      PRINT "An error occured in the value function!"
  '      EXIT SUB
  '   END IF
  '                                             'Only one amount field in form!
  '   SELECT CASE POEdit.EType                  'Is amount debit or credit
  '      CASE "D"                               'Debit
  '         POEdit.DrAmt = Amt#                 'assign amt to debit amt
  '         POEdit.CrAmt = 0
  '      CASE "C"                               'Credit
  '         POEdit.CrAmt = Amt#                 'assign amt to credit amt
  '         POEdit.DrAmt = 0
  '   END SELECT
  '
  '--< end of stuff.
  
  
END SUB

SUB ShowEdMode (EdMode%)
  
  Status$ = SPACE$(24)
  F2Ed$ = "  F2=Edit  "
  F2New$ = "  F2=New   "
  ActiveColor = 112
  InActiveColor = 115
  
  SELECT CASE EdMode
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
  QPrintRC F2BtnCap$, 25, 28, -1
  MPaintBox 25, 41, 25, 51, F5BtnColor          'F5 Button enable / disable
  
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
    GLIFRec(1).TRDate = Date2Num(DATE$)
    GLIFRec(1).Desc = "CANCELLED PO"  'APLedger(1).PONum
    GLIFRec(1).Ref = APLedger(1).PONum
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

