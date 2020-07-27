DEFINT A-Z
'DECLARE SUB VendorHistory ()
DECLARE FUNCTION GetNewVendorPIN% ()
DECLARE SUB CursorOff ()
DECLARE SUB PrintTitle (Title$)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTNo, RetCode%, EntryPoint%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTNo, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (x$, Colr%, Page%)
DECLARE SUB QPrintRC (t$, R%, C%, Clr%)
DECLARE FUNCTION QPTrim$ (x$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
'ECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB Form2Rec (VRec%, Form$(), OrgVend$)

TYPE FLen
  V AS STRING * 64
END TYPE

TYPE ALen
  V AS STRING * 42
END TYPE

DECLARE SUB VertMenuT2 (Items() AS FLen, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB StuffBuf (Ky$)

DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, x%, y%, Wdth%, TitColor%)
DECLARE FUNCTION StripAcct$ (AcctNum$)
DECLARE FUNCTION FmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE SUB Import ()
DECLARE SUB OpenDefDistFile (VdefDist AS ANY, DefDistFileNum%, NumDefRecs%)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)

'$INCLUDE: 'ap.bi'
'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'SCRAUX.BI'
'$INCLUDE: 'GL.BI'

DECLARE SUB EditVendorRec ()
DECLARE SUB Rec2Form (VRec%, Form$())
DECLARE SUB IndexVendorFile ()
DECLARE SUB NewVendorRec ()
DECLARE SUB PrintVendorList ()
DECLARE SUB PrintVendorLabels ()
DECLARE SUB ResortVendorIdx ()
DECLARE SUB SetDefaultDist ()
DECLARE SUB PrintVendorDefDist ()
DECLARE SUB PrintVendorFile ()

DECLARE FUNCTION VendorPickList (Lookfor$, Abort%)
DECLARE FUNCTION FindVendorRec (VendorCode$)
DECLARE SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)
DECLARE SUB OpenVendorFile (VendorFile, NumVRecs)
  
  CONST False = 0, True = NOT False
  
  DIM SHARED VendorIdx AS VendorIdxRecType
  DIM SHARED Vendor AS VendorRecType
  DIM SHARED FundLen, AcctLen, DetLen, TotLen

  DIM SHARED User$
  
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'init.bi'

'Import
'END
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 9)
  
  MChoice$(1) = "Add a New Vendor Record"
  MChoice$(2) = "Edit a Vendor Record"
  MChoice$(3) = "Set Default Distribution"
  MChoice$(4) = "Print Vendor List"
  MChoice$(5) = "Print Vendor File"
  MChoice$(6) = "Print Vendor Labels"
  MChoice$(7) = "Print Vendor Default Distributions"
  MChoice$(8) = "Sort Vendor Indexes"
  MChoice$(9) = "Quit"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2

  DO

    LibFile2Scrn "AP.QSL", "BG", MonoCode, 0, ErrorCode
    'TitleBox 6, Col, MaxLen + 3, "Vendor Maintenance ", Cnf
    'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    MnuTitle "Vendor Maintenance  ", Row, Col, MaxLen, 15
    PrintHelp "Vendor Maintenance Menu"
    QPrintRC SPACE$(80), 1, 1, 112
    PrintTitle User$
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      NewVendorRec
    CASE 2
      EditVendorRec
    CASE 3
      SetDefaultDist
    CASE 4
      PrintVendorList
    CASE 5
      PrintVendorFile
    CASE 6
      PrintVendorLabels
    CASE 7
      PrintVendorDefDist
    CASE 8
      ResortVendorIdx
    CASE 9
      HideCursor
      CLS
      END
    CASE ELSE
    END SELECT
    
  LOOP
  RUN "APMENU"

END

SUB comments

'   VIN        AS LONG
'   VNum       AS STRING * 10
'   VNAME      AS STRING * 30
'   Addr1      AS STRING * 30
'   Addr2      AS STRING * 30
'   City       AS STRING * 22
'   State      AS STRING * 2
'   Zip        AS STRING * 10
'   PayToName  AS STRING * 30
'   PayToAddr  AS STRING * 30
'   PaytoAddr2 AS STRING * 30
'   PayToCity  AS STRING * 22
'   PaytoState AS STRING * 2
'   PaytoZip   AS STRING * 10
'   phone      AS STRING * 14
'   VTERMS     AS INTEGER
'   pad2 AS STRING * 5
'   FedID      AS STRING * 12
'   CoCode     AS STRING * 3
'   StCode     AS STRING * 2
'   YTDPay     AS DOUBLE
'   Get1099  AS STRING * 1
'   CurrBal    AS DOUBLE
'   FrstTran   AS LONG
'   LastTran   AS LONG
'   FrstPO     AS LONG
'   LastPO     AS LONG
'   DefDist    AS LONG
'   DELFLAG    AS INTEGER
'   'Pad        AS STRING * 45
'   Contact    AS STRING * 30
'   fax      AS STRING * 14
'   pad AS STRING * 1

END SUB

SUB EditVendorRec
  
  REDIM ScrnArray(2000)
  MScrnSave 1, 1, 25, 80, ScrnArray(0)

top:

  REDIM TempVendorRec(1) AS VendorRecType
  Vendor = TempVendorRec(1)

  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  DIM Frm AS FormInfo           'Form editing info
  LibName$ = "AP.QSL"
  
  DO
    '--Show a background screen
    FormName$ = "BG"
    LibFile2Scrn LibName$, FormName$, MonoCode, 0, ErrorCode
    PrintHelp "Enter the vendor code. "
    
    '--Get the vendor code user wants to edit
    DO
      FormName$ = "LOOKUPV"
      VCode$ = InputBox$(LibName$, FormName$, Cancel%)
      IF Cancel THEN GOTO ExitVendorEdit

      LOCATE , , 0
      FormName$ = "LOADVLST"
      LibFile2Scrn LibName$, FormName$, MonoCode, 0, ErrorCode
      PrintHelp "Search List"

      VRecNum = VendorPickList(VCode$, Cancel%)
      IF Cancel OR (VRecNum = 0) THEN
        Cancel = False
        EXIT DO
      END IF
      
      '--found the record.. setup editform to edit the record
      FormName$ = "APVENDOR"
      LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
      PrintHelp "Editing"
      
      Rec2Form VRecNum, Form$()
      OrgVend$ = Form$(1, 0)
      Frm.FldNo = 1
      Action = 1
      Fld(1).Protected = -1

      DO
        EditForm Form$(), Fld(), Frm, Cnf, Action
        IF Frm.PrevFld <> Frm.FldNo THEN
          IF Frm.PrevFld = 1 AND OrgVend$ <> Form$(1, 0) THEN
            NRecNum = FindVendorRec(QPTrim$(Form$(1, 0)))
            IF NRecNum > 0 THEN
              FormName$ = "DUPVNDR"
              Ok = MsgBox(LibName$, FormName$)
              Form$(1, 0) = OrgVend$
              Frm.FldNo = 1
              Action = 1
            END IF
          END IF
        END IF
        
        SELECT CASE Frm.KeyCode
          CASE F3Key
            GOSUB DelRec
            GOTO top
          CASE F10Key
            Form2Rec VRecNum, Form$(), OrgVend$
            Done = True
            GOTO top
          CASE ESC
            Done = True
            GOTO top
          CASE ELSE
            Done = False
        END SELECT
        
      LOOP UNTIL Done
    LOOP
  LOOP
  
ExitVendorEdit:
EXIT SUB

  
DelRec:
  IF Vendor.FrstTran = 0 THEN
    FormName$ = "OKDEL"
    Answer = MsgBox(LibName$, FormName$)
    SELECT CASE Answer
    CASE 1      'No
      Action = 1
    CASE 2      'Yes
      Vendor.DELFLAG = True
      OpenVendorFile VendorFile, NumVRecs
      PUT VendorFile, VRecNum, Vendor
      CLOSE VendorFile
      Done = True
      IndexVendorFile
      MScrnRest 1, 1, 25, 80, ScrnArray(0)
      GOTO ExitVendorEdit:
    END SELECT
  ELSE
    FormName$ = "NODEL"
    Answer = MsgBox(LibName$, FormName$)
    Action = 1
  END IF
RETURN
  
END SUB

FUNCTION FindVendorRec (VendorCode$)
  
  'SHARED VendorIdx AS VendorIdxRecType
  
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

SUB Form2Rec (VRec%, Form$(), OrgVend$)

  PrintHelp "Saving to disk.. Please wait."

  Vendor.VNum = Form$(1, 0)
  Vendor.VNAME = Form$(2, 0)
  Vendor.Addr1 = Form$(3, 0)
  Vendor.Addr2 = Form$(4, 0)
  Vendor.City = Form$(5, 0)
  Vendor.State = Form$(6, 0)
  Vendor.Zip = Form$(7, 0)
  Vendor.PayToName = Form$(8, 0)
  Vendor.PayToAddr = Form$(9, 0)
  Vendor.PayToAddr2 = Form$(10, 0)
  Vendor.PayToCity = Form$(11, 0)
  Vendor.PayToState = Form$(12, 0)
  Vendor.PayToZip = Form$(13, 0)
  Vendor.VTerms = VAL(Form$(14, 0))
  Vendor.FedID = Form$(15, 0)
  Vendor.StCode = Form$(16, 0)
  Vendor.CoCode = Form$(17, 0)
  Vendor.Get1099 = Form$(18, 0)
  Vendor.Contact = Form$(19, 0)
  Vendor.Phone = Form$(20, 0)
  Vendor.Fax = Form$(21, 0)

  OpenVendorFile VendorFile, NumVRecs
  IF VRec = 0 THEN   'VRec = 0 if called from Add New Vendor Sub
    VRec = NumVRecs + 1
  END IF
  PUT VendorFile, VRec, Vendor
  CLOSE VendorFile

  IF OrgVend$ <> Form$(1, 0) THEN
    IndexVendorFile
  END IF
  
END SUB

FUNCTION GetNewVendorPIN

  VPinHandle = FREEFILE
  OPEN "VendrPIN.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #VPinHandle LEN = 4
  GET #VPinHandle, 1, VPin&
  VPin& = VPin& + 1
  PUT #VPinHandle, 1, VPin&
  CLOSE #VPinHandle
  GetNewVendorPIN = VPin&

END FUNCTION

SUB Import

'Imports a comma-delimited file

   ImportFileName$ = "APV.ASC"
   ImportFile = FREEFILE
   OPEN ImportFileName$ FOR INPUT AS #ImportFile

   OpenVendorFile VendorFile, NumVRecs

   CLS
   LOCATE 1, 1
   PRINT "Importing Vendor File"

   DO

      RecordsIn = RecordsIn + 1
      LOCATE 2, 1, 0
      PRINT "Processing Record: "; RecordsIn

      INPUT #ImportFile, VCode$, VNAME$, Add1$, Add2$, CitySt$, Zip$, CoCode$, Get1099$

      Vendor.VNum = VCode$
      Vendor.VNAME = VNAME$
      Vendor.Addr1 = Add1$
      Vendor.Addr2 = Add2$
      Vendor.City = CitySt$
      Vendor.State = ""
      Vendor.Zip = Zip$
      Vendor.PayToName = VNAME$
      Vendor.PayToAddr = Add1$
      Vendor.PayToAddr2 = Add2$
      Vendor.PayToCity = CitySt$
      Vendor.PayToState = ""
      Vendor.PayToZip = Zip$

      Vendor.Phone = ""
      Vendor.VTerms = 0
      Vendor.FedID = ""
      Vendor.CoCode = CoCode$
      Vendor.StCode = CoCode$
      Vendor.Get1099 = Get1099$

      PUT VendorFile, RecordsIn, Vendor

   LOOP UNTIL EOF(ImportFile)

   PRINT "File import complete."; RecordsIn; "records were imported."
   PRINT "Press any key to continue."
   k$ = INPUT$(1)

   CLOSE

END SUB

SUB NewVendorRec

  REDIM TempVendorRec(1) AS VendorRecType
  Vendor = TempVendorRec(1)
  LibName$ = "AP.QSL"
  FormName$ = "APVENDOR"
  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  DIM Frm AS FormInfo           'Form editing info
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintTitle User$
  PrintHelp "New Vendor"
  OrgVend$ = ""
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

    '--Check for duplicate vendor when user leaves field 1
    IF Frm.PrevFld <> Frm.FldNo THEN
      IF Frm.PrevFld = 1 THEN
        VRecNum = FindVendorRec(QPTrim$(Form$(1, 0)))
        IF VRecNum > 0 THEN
          Ok = MsgBox(LibName$, "DUPVNDR")
          LSET Form$(1, 0) = ""
          Frm.FldNo = 1
          Frm.KeyCode = 0
          Action = 1
        END IF
      END IF
    END IF
    
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
          Form2Rec VRec, Form$(), OrgVend$
          Done = True
        END IF
      CASE ESC
        Done = True
      CASE ELSE
    END SELECT
    
  LOOP UNTIL Done

END SUB

SUB PrintVendorDefDist

  LOCATE , , 0

  SHARED User$
  FF$ = CHR$(12)
  MaxLines = 55
  LineCnt = 0
  Page = 0
  AcctHeader$ = "         Acct              Name                      Percent"
  
  REDIM Desc$(1)
  REDIM Title$(1 TO 5)
  User$ = QPTrim$(User$)
  Title$(1) = User$
  Title$(2) = "Vendor Default Distribution Listing"
  Title$(3) = ""
  Title$(4) = "Vendor Code  Name"
  Title$(5) = STRING$(78, "-")

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibName$ = "AP.QSL"
  FormName$ = "PRNVNDR"

  LibFile2Scrn LibName$, "BG", MonoCode, 0, ErrorCode
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintHelp "Print Vendor Listing"

  Action = 1

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
      CASE F10Key 'SaveButton
        Lookfor$ = RTRIM$(Form$(1, 0))
        IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
          Dev$ = "S"
          LPTNo = 1
        ELSE
          Dev$ = LEFT$(Form$(1, 0), 1)
          LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
        END IF
      CASE ESC
        EXIT SUB
    END SELECT
  LOOP UNTIL Frm.KeyCode = F10Key

  LibFile2Scrn LibName$, "BG", MonoCode, Attribute, ErrorCode
  PrintHelp "Processing report. Please wait."

  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs

  DIM VDefRec AS VendorDefDistRecType
  DefRecLen = LEN(VDefRec)
  OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs

  PrnFile = FREEFILE
  ReportFile$ = "vndrlst.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  ToPrint$ = SPACE$(80)

  GOSUB PrintDefDistHeader

  FOR cnt = 1 TO NumActiveVendors
    GET VendorIdxFile, cnt, VendorIdx
    GET VendorFile, VendorIdx.RecNum, Vendor
    
    IF Vendor.DefDist > 0 THEN
      HowMany = HowMany + 1

      '--Find the number of distributions for this vendor
      GET APDefDistFile, Vendor.DefDist, VDefRec
      N = 0
      FOR d = 1 TO 8
        IF LEN(QPTrim$(VDefRec.DefDist(d).DefAcct)) THEN
          N = N + 1
        END IF
      NEXT
      
      LinesToPrint = N + 3   '3 lines for name, header & dividing line
      IF LinesToPrint + LineCnt >= MaxLines THEN
        PRINT #PrnFile, FF$
        GOSUB PrintDefDistHeader
      END IF
      
      LSET ToPrint$ = ""
      MID$(ToPrint$, 2) = Vendor.VNum
      MID$(ToPrint$, 14) = Vendor.VNAME
      PRINT #PrnFile, ToPrint$
      LineCnt = LineCnt + 1
      
      PRINT #PrnFile, AcctHeader$
      LineCnt = LineCnt + 1

      FOR d = 1 TO 8
        IF LEN(QPTrim$(VDefRec.DefDist(d).DefAcct)) THEN
          LSET ToPrint$ = ""
          MID$(ToPrint$, 10) = VDefRec.DefDist(d).DefAcct
          MID$(ToPrint$, 28) = VDefRec.DefDist(d).DefAcctName
          MID$(ToPrint$, 55) = STR$(VDefRec.DefDist(d).DefPct) + "%"
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
        END IF
      NEXT

      PRINT #PrnFile, Title$(5)
      LineCnt = LineCnt + 1

    END IF
  NEXT

  PRINT #PrnFile,
  PRINT #PrnFile, HowMany; "Vendors listed."
  PRINT #PrnFile, FF$

  CLOSE

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT

  PrintRptFile Title$(2), ReportFile$, LPTNo, RetCode%, EntryPoint
  KILL ReportFile$

EXIT SUB

PrintDefDistHeader:
Page = Page + 1
  PRINT #PrnFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrnFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrnFile,
  PRINT #PrnFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #PrnFile, "Vendor Code  Name                            City                 Default Dist"
  PRINT #PrnFile, STRING$(79, "-")

  LineCnt = 5
RETURN

  
RETURN




END SUB

SUB PrintVendorFile

  LOCATE , , 0

  SHARED User$
  FF$ = CHR$(12)
  MaxLines = 60
  LineCnt = 0
  Page = 0
  REDIM Desc$(1)
  Header$ = "Vendor File"
  REDIM Title$(1 TO 5)
  User$ = QPTrim$(User$)
  Title$(1) = User$
  Title$(2) = Header$
  Title$(3) = ""
  Title$(4) = "Vendor Code  Name                            Pay To Address"
  Title$(5) = STRING$(78, "-")

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibName$ = "AP.QSL"
  FormName$ = "PRNVNDR"

  LibFile2Scrn LibName$, "BG", MonoCode, 0, ErrorCode
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintHelp "Print Vendor File"

  Action = 1

  DO            'Allow user to edit the form

    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode     'Check the user's key presses
      CASE F10Key 'SaveButton
        Lookfor$ = RTRIM$(Form$(1, 0))
        IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
          Dev$ = "S"
          LPTNo = 1
        ELSE
          Dev$ = LEFT$(Form$(1, 0), 1)
          LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
        END IF
      CASE ESC
        EXIT SUB
    END SELECT

  LOOP UNTIL Frm.KeyCode = F10Key

  LOCATE , , 0

  LibFile2Scrn LibName$, "BG", MonoCode, Attribute, ErrorCode
  PrintHelp "Processing report. Please wait."

  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs

  PrnFile = FREEFILE
  ReportFile$ = "vndrlst.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  ToPrint$ = SPACE$(80)
  DivLine$ = STRING$(78, "-")
  
  GOSUB PrintVRptHeader

  FOR cnt = 1 TO NumActiveVendors
    GET VendorIdxFile, cnt, VendorIdx
    GET VendorFile, VendorIdx.RecNum, Vendor
    HowMany = HowMany + 1

    'IF Vendor.DELFLAG <> 0 THEN STOP

    IF LineCnt > MaxLines - 8 THEN
      PRINT #PrnFile, FF$
      GOSUB PrintVRptHeader
    END IF
    
    '--Line 1
    LSET ToPrint$ = ""
    MID$(ToPrint$, 2) = Vendor.VNum
    MID$(ToPrint$, 14) = Vendor.VNAME
    MID$(ToPrint$, 46) = Vendor.PayToName
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1
    
    '--Line 2
    A1Len = LEN(QPTrim$(Vendor.Addr1))
    PA1Len = LEN(QPTrim$(Vendor.PayToAddr))
    IF A1Len > 0 OR PA1Len > 0 THEN
      LSET ToPrint$ = ""
      MID$(ToPrint$, 14) = Vendor.Addr1
      MID$(ToPrint$, 46) = Vendor.PayToAddr
      PRINT #PrnFile, ToPrint$
      LineCnt = LineCnt + 1
    END IF

    '--Line 3
    a2Len = LEN(QPTrim$(Vendor.Addr2))
    PA2Len = LEN(QPTrim$(Vendor.PayToAddr2))
    IF a2Len > 0 OR PA2Len > 0 THEN
      LSET ToPrint$ = ""
      MID$(ToPrint$, 14) = Vendor.Addr2
      MID$(ToPrint$, 46) = Vendor.PayToAddr2
      PRINT #PrnFile, ToPrint$
      LineCnt = LineCnt + 1
    END IF
    
    '--Line 4
    LSET ToPrint$ = ""
    MID$(ToPrint$, 14) = QPTrim$(Vendor.City) + " " + Vendor.State + " " + Vendor.Zip
    MID$(ToPrint$, 46) = QPTrim$(Vendor.PayToCity) + " " + Vendor.PayToState + " " + Vendor.PayToZip
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1

    '--Line 5
    LSET ToPrint$ = ""
    MID$(ToPrint$, 14) = "ID: " + Vendor.FedID
    MID$(ToPrint$, 30) = "Terms: " + QPTrim$(STR$(Vendor.VTerms))
    MID$(ToPrint$, 41) = "State: " + Vendor.StCode
    MID$(ToPrint$, 53) = "County: " + Vendor.CoCode
    MID$(ToPrint$, 68) = "1099? " + Vendor.Get1099
    'IF Vendor.Get1099 = "Y" THEN STOP
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1

    '--Line 6
    LSET ToPrint$ = ""
    MID$(ToPrint$, 14) = "Contact: " + Vendor.Contact
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1

    '--Line 7
    LSET ToPrint$ = ""
    MID$(ToPrint$, 14) = "Phone: " + Vendor.Phone
    MID$(ToPrint$, 46) = "Fax: " + Vendor.Fax
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1
    
    '--Line 8
    PRINT #PrnFile, DivLine$
    LineCnt = LineCnt + 1

  NEXT
  
  PRINT #PrnFile,
  PRINT #PrnFile, HowMany; "Vendors listed."
  PRINT #PrnFile, FF$

  CLOSE

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT

  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  KILL ReportFile$

EXIT SUB

PrintVRptHeader:
Page = Page + 1
  PRINT #PrnFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrnFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrnFile,
  PRINT #PrnFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #PrnFile, "Vendor Code  Name                            City                 Default Dist"
  PRINT #PrnFile, STRING$(79, "-")

  LineCnt = 5
RETURN


END SUB

SUB PrintVendorLabels

  LOCATE , , 0
   
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibName$ = "AP.QSL"
  FormName$ = "VNDRLBL"
  REDIM ScrBuf(2000)
   

  ToPrint$ = SPACE$(30)

  LibFile2Scrn LibName$, "BG", MonoCode, 0, ErrorCode
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintHelp "Print Vendor Labels"
  
  Action = 1

  DO            'Allow user to edit the form

    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode     'Check the user's key presses

    CASE F5Key
      MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
      GOSUB PrintMask
      MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

    CASE F10Key 'SaveButton
      'Lookfor$ = RTRIM$(Form$(1, 0))
      IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(1, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
      END IF
      
    CASE ESC
      EXIT SUB

    END SELECT

  LOOP UNTIL Frm.KeyCode = F10Key

  LOCATE , , 0
  PrnFile = FREEFILE
  ReportFile$ = "vndrlbl.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  LibFile2Scrn LibName$, "BG", MonoCode, Attribute, ErrorCode
  PrintHelp "Processing report. Please wait."

  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs
  
  '--fix for del not to show
  FOR cnt = 1 TO NumActiveVendors
    GET VendorIdxFile, cnt, VendorIdx
    GET VendorFile, VendorIdx.RecNum, Vendor
    HowMany = HowMany + 1

    LSET ToPrint$ = Vendor.VNum
    PRINT #PrnFile, ToPrint$
    
    LSET ToPrint$ = Vendor.VNAME
    PRINT #PrnFile, ToPrint$

    LSET ToPrint$ = Vendor.Addr1
    PRINT #PrnFile, ToPrint$

    LSET ToPrint$ = Vendor.Addr2
    PRINT #PrnFile, ToPrint$

    LSET ToPrint$ = QPTrim$(Vendor.City) + " " + Vendor.State + " " + Vendor.Zip
    PRINT #PrnFile, ToPrint$
  
    LSET ToPrint$ = ""
    PRINT #PrnFile, ToPrint$

  NEXT

  CLOSE

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT

  Header$ = "Vendor Labels"
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  'KILL ReportFile$

EXIT SUB


PrintMask:

  LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
  IF LPTNo = 0 THEN LPTNo = 1
  EntryPoint = 5
  
  MaskFile = FREEFILE
  MaskReportFile$ = "LBLMASK.PRN"
  OPEN MaskReportFile$ FOR OUTPUT AS #MaskFile
  ToPrint$ = SPACE$(30)
  
  LSET ToPrint$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  PRINT #MaskFile, ToPrint$

  LSET ToPrint$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  PRINT #MaskFile, ToPrint$

  LSET ToPrint$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  PRINT #MaskFile, ToPrint$

  LSET ToPrint$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  PRINT #MaskFile, ToPrint$

  LSET ToPrint$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  PRINT #MaskFile, ToPrint$

  LSET ToPrint$ = ""
  PRINT #MaskFile, ToPrint$

  CLOSE MaskFile

  Header$ = "Label Alignent Test"
  PrintRptFile Header$, MaskReportFile$, LPTNo, RetCode%, EntryPoint
  
RETURN

END SUB

SUB PrintVendorList

  LOCATE , , 0

  SHARED User$
  FF$ = CHR$(12)
  MaxLines = 55
  LineCnt = 0
  Page = 1

  REDIM Desc$(1)
  Header$ = "Vendor Listing"
  REDIM Title$(1 TO 5)
  User$ = QPTrim$(User$)
  Title$(1) = User$
  Title$(2) = Header$
  Title$(3) = ""
  Title$(4) = "Vendor Code  Name                            City                 Default Dist"
  Title$(5) = STRING$(78, "-")
   
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibName$ = "AP.QSL"
  FormName$ = "PRNVNDR"
  
  LibFile2Scrn LibName$, "BG", MonoCode, 0, ErrorCode
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintHelp "Print Vendor Listing"
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
      CASE F10Key 'SaveButton
        Lookfor$ = RTRIM$(Form$(1, 0))
        IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
          Dev$ = "S"
          LPTNo = 1
        ELSE
          Dev$ = LEFT$(Form$(1, 0), 1)
          LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
        END IF
      CASE ESC
        EXIT SUB
    END SELECT
  LOOP UNTIL Frm.KeyCode = F10Key
  
  LibFile2Scrn LibName$, "BG", MonoCode, Attribute, ErrorCode
  PrintHelp "Processing report. Please wait."
  
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs
  PrnFile = FREEFILE
  ReportFile$ = "vndrlst.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  GOSUB PrintRptHeader

  FOR cnt = 1 TO NumActiveVendors
    GET VendorIdxFile, cnt, VendorIdx
    GET VendorFile, VendorIdx.RecNum, Vendor

'    IF Vendor.Get1099 <> "Y" THEN
'      GOTO SkipEm1099
'    END IF

    HowMany = HowMany + 1
    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 2) = Vendor.VNum
    MID$(ToPrint$, 14) = Vendor.VNAME
    MID$(ToPrint$, 46) = QPTrim$(Vendor.City) + " " + Vendor.State

    IF Vendor.DefDist > 0 THEN
      MID$(ToPrint$, 75) = "Y"
    ELSE
      MID$(ToPrint$, 75) = "N"
    END IF

'    MID$(ToPrint$, 75) = Vendor.Get1099

    'MID$(ToPrint$, 75) = STR$(VendorIdx.RecNum)
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1
    IF LineCnt >= MaxLines THEN
      PRINT #PrnFile, FF$
      GOSUB PrintRptHeader
    END IF

SkipEm1099:
  NEXT
  
  PRINT #PrnFile,
  PRINT #PrnFile, HowMany; "Vendors listed."
  PRINT #PrnFile, FF$

  CLOSE
  
  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT

  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  KILL ReportFile$

EXIT SUB

PrintRptHeader:
Page = Page + 1
  PRINT #PrnFile, TAB(40 - (INT(LEN(User$) / 2))); User$
  PRINT #PrnFile, TAB(40 - (INT(LEN(Header$) / 2))); Header$
  PRINT #PrnFile,
  PRINT #PrnFile, "Report Date: "; DATE$; TAB(67); "Page #"; Page
  PRINT #PrnFile, "Vendor Code  Name                            City                 Default Dist"
  PRINT #PrnFile, STRING$(79, "-")


  LineCnt = 5
RETURN

  
END SUB

SUB Rec2Form (VRec%, Form$())
  
  OpenVendorFile VendorFile, NumVRecs
  
  GET VendorFile, VRec, Vendor
  
  Form$(1, 0) = Vendor.VNum
  Form$(2, 0) = Vendor.VNAME
  Form$(3, 0) = Vendor.Addr1
  Form$(4, 0) = Vendor.Addr2
  Form$(5, 0) = Vendor.City
  Form$(6, 0) = Vendor.State
  Form$(7, 0) = Vendor.Zip
  Form$(8, 0) = Vendor.PayToName
  Form$(9, 0) = Vendor.PayToAddr
  Form$(10, 0) = Vendor.PayToAddr2
  Form$(11, 0) = Vendor.PayToCity
  Form$(12, 0) = Vendor.PayToState
  Form$(13, 0) = Vendor.PayToZip
  Form$(14, 0) = QPTrim$(STR$(Vendor.VTerms))
  Form$(15, 0) = Vendor.FedID
  Form$(16, 0) = Vendor.StCode
  Form$(17, 0) = Vendor.CoCode
  Form$(18, 0) = Vendor.Get1099
  Form$(19, 0) = Vendor.Contact
  Form$(20, 0) = Vendor.Phone
  Form$(21, 0) = Vendor.Fax
  
  CLOSE
  
END SUB

SUB ResortVendorIdx
  
  LOCATE 1, 1, 0
  LibFile2Scrn "AP.QSL", "BG", MonoCode, 0, ErrorCode
  LibFile2Scrn "AP.QSL", "SORTVNDR", MonoCode, 0, ErrorCode
  
  k$ = INPUT$(1)
  IF k$ = CHR$(27) THEN EXIT SUB
  
  QPrintRC "Sorting...                   ", 11, 25, -1
  IndexVendorFile
  
  QPrintRC "Sorting complete.", 12, 25, -1
  QPrintRC "Press any key to exit.", 13, 25, -1
  
  k$ = INPUT$(1)
  
END SUB

SUB SetDefaultDist

Dtop:
  
  REDIM TempVendorRec(1) AS VendorRecType
  Vendor = TempVendorRec(1)
  REDIM acct(1) AS GLAcctRecType
  REDIM VdefDist(1) AS VendorDefDistRecType
  DefRecLen = LEN(VdefDist(1))

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum
  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType
  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs
  
  List$ = SPACE$(42)

  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  REDIM Frm(1) AS FormInfo           'Form editing info
  LibName$ = "AP.QSL"

  DO

    '--Show a background screen
    LibFile2Scrn "AP.QSL", "BG", MonoCode, 0, ErrorCode
    PrintHelp "Enter the vendor code. "
    DoExit = False

    DO
      FormName$ = "LOOKUPV"
      VCode$ = InputBox$(LibName$, FormName$, Cancel%)
      IF Cancel THEN
        DoExit = True      'user Escaped then exit
        EXIT DO
      END IF

      LOCATE , , 0
      LibFile2Scrn "AP.QSL", "LOADVLST", MonoCode, 0, ErrorCode
      PrintHelp "Search List"
      VRecNum = VendorPickList(VCode$, Cancel%)
      '--Exit if user pressed Esc while browsing list
      IF Cancel OR (VRecNum = 0) THEN
        Cancel = 0
        EXIT DO
      END IF
      
      FormName$ = "SETDEFD"
      LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
      
      IF VRecNum > 0 THEN
        OpenVendorFile VendorFile, NumVRecs
        GET VendorFile, VRecNum, Vendor
        CSZ$ = QPTrim$(Vendor.City) + " " + Vendor.State + " " + Vendor.Zip
        VDefRec = Vendor.DefDist
        QPrintRC Vendor.VNAME, 5, 16, -1
        QPrintRC Vendor.Addr1, 6, 16, -1
        QPrintRC Vendor.Addr2, 7, 16, -1
        QPrintRC CSZ$, 8, 16, -1
        IF VDefRec > 0 THEN
          GOSUB GetDefRec
        END IF
      END IF
      Action = 1
      

      DO
        EditForm Form$(), Fld(), Frm(1), Cnf, Action

        IF Frm(1).PrevFld <> Frm(1).FldNo THEN

          IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
            LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
            Action = 1
          END IF
          
          IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN       '"A" = AcctField
            IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
              TempAcct$ = StripAcct$(Form$(Frm(1).PrevFld, 0))
              FAcct$ = FmtAcct$(TempAcct$, FundLen%, AcctLen%, DetLen%)
              LSET Form$(Frm(1).PrevFld, 0) = FAcct$
              'Fld(Frm(1).PrevFld + 3).Protected = False
            ELSE
              FOR Cnt2 = 1 TO 2
                LSET Form$(Frm(1).PrevFld + Cnt2, 0) = ""
                'Fld(Frm(1).PrevFld + Cnt2).Protected = True
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
        END IF

        SELECT CASE Frm(1).KeyCode
          CASE F3Key
            GOSUB DelDefRec
            'Done = True
            GOTO Dtop
          CASE F10Key
            GOSUB SaveDefRec
            'Done = True
            GOTO Dtop
          CASE ESC
            'Done = True
            GOTO Dtop
          CASE ELSE
            Done = False
        END SELECT
      LOOP UNTIL Done
    LOOP
  IF DoExit THEN EXIT DO
  LOOP

EXIT SUB

CheckForValidAcctNum:

  AcctFld = Frm(1).PrevFld
  MatchAccts = 0
  Lookup$ = QPTrim$(UCASE$(Form$(AcctFld, 0)))
  Lookuplen = LEN(Lookup$)

  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
   
  OpenAcctFile AcctFileNum, NumAccts

  REDIM AcctList(1 TO NumIdxRecs) AS ALen       'Dim array to hold trans list
  FOR cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(cnt).RecNum, acct(1)
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).V = List$
      AcctList(MatchAccts).V = TempAcctIdx(cnt).AcctNum
      MID$(AcctList(MatchAccts).V, TotLen + 4) = acct(1).Title
      MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(cnt).RecNum)
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
  GET AcctFileNum, AcctRecNum, acct(1)
  CLOSE AcctFileNum
  LSET Form$(AcctFld, 0) = acct(1).Num
  LSET Form$(AcctFld + 1, 0) = acct(1).Title
  'LSET Form$(AcctFld + 2, 0) = MKI$(AcctRecNum)
RETURN

SaveDefRec:
  IF VAL(Form$(25, 0)) > 100 THEN
    BEEP
  END IF

  OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs
  VdefDist(1).VRecNum = VRecNum
  FOR cnt = 1 TO 22 STEP 3
    DCnt = DCnt + 1
    VdefDist(1).DefDist(DCnt).DefAcct = Form$(cnt, 0)
    VdefDist(1).DefDist(DCnt).DefAcctName = Form$(cnt + 1, 0)
    VdefDist(1).DefDist(DCnt).DefPct = VAL(Form$(cnt + 2, 0))
  NEXT
  IF VDefRec THEN
    DefRecNum = VDefRec
    PUT APDefDistFile, DefRecNum, VdefDist(1)
  ELSE
    DefRecNum = NumDefRecs + 1
    PUT APDefDistFile, DefRecNum, VdefDist(1)
    OpenVendorFile VendorFile, NumVRecs
    GET VendorFile, VRecNum, Vendor
    Vendor.DefDist = DefRecNum
    PUT VendorFile, VRecNum, Vendor
  END IF
  CLOSE
  cnt = 0: DCnt = 0
RETURN

GetDefRec:
  OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs
  GET APDefDistFile, VDefRec, VdefDist(1)
  TotPct# = 0
  FOR cnt = 1 TO 22 STEP 3
    DCnt = DCnt + 1
    Form$(cnt, 0) = VdefDist(1).DefDist(DCnt).DefAcct
    Form$(cnt + 1, 0) = VdefDist(1).DefDist(DCnt).DefAcctName
    IF VdefDist(1).DefDist(DCnt).DefPct <> 0 THEN
      Form$(cnt + 2, 0) = STR$(VdefDist(1).DefDist(DCnt).DefPct)
      TotPct# = TotPct# + VdefDist(1).DefDist(DCnt).DefPct
    ELSE
      Form$(cnt + 2, 0) = ""
    END IF
  NEXT
  'CalcFields 0, 3, Form$(), Fld()
  Form$(25, 0) = STR$(TotPct#)
  Frm(1).FldNo = 1
  Action = 1
  CLOSE
  cnt = 0: DCnt = 0
RETURN

DelDefRec:
  OpenVendorFile VendorFile, NumVRecs
  GET VendorFile, VRecNum, Vendor
  Vendor.DefDist = 0
  PUT VendorFile, VRecNum, Vendor
RETURN

END SUB

FUNCTION VendorPickList (Lookfor$, Abort%)
  
  
  SHARED VendorIdx AS VendorIdxRecType
  SHARED Vendor AS VendorRecType
  
  List$ = SPACE$(65)
  
  VLookup$ = QPTrim$(UCASE$(Lookfor$))
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
  IF NumVRecs = 0 THEN EXIT FUNCTION
  REDIM VendorList(1 TO NumVRecs) AS FLen       'Dim array to hold trans list
  
  '  FOR Cnt = 1 TO NumActiveVendors               'Read the file
  FOR cnt = 1 TO NumVRecs       'Read the file
    GET VendorIdxFile, cnt, VendorIdx
    
    IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
    IF VendorIdx.RecNum > 0 THEN
      GET VendorFile, VendorIdx.RecNum, Vendor
      IF NOT Vendor.DELFLAG THEN
        MatchingVendors = MatchingVendors + 1
        LSET VendorList(MatchingVendors).V = List$
        VendorList(MatchingVendors).V = VendorIdx.VendorCode
        MID$(VendorList(MatchingVendors).V, 12) = Vendor.VNAME
        MID$(VendorList(MatchingVendors).V, 44) = QPTrim$(Vendor.City) + " " + Vendor.State
        MID$(VendorList(MatchingVendors).V, 63) = MKI$(VendorIdx.RecNum)
      END IF
    END IF
    END IF
    
  NEXT
  
  CLOSE
  
  IF MatchingVendors > 1 THEN
    REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
    
    Row = 6: Col = 10
    
    LOCATE Row, Col
    Action = 0
    MaxWidth = 60
    BoxBot = 18
    
    'Print a title bar above the menu
    ListDescTxt$ = "VendorCode  Name"
    ListDesc$ = SPACE$(MaxWidth + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    
    QPrintRC ListDesc$, Row - 1, Col, 11        'Cnf.ActivCh
    'MPaintBox Row + 1, MaxWidth + 2, 7, 76, 8 'to add shadow to list
    
    DO
      'call vertmenu to display matching recs to choose from
      VertMenuT2 VendorList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        CLOSE
        Abort = True
        EXIT DO
      END IF
      IF Choice > 0 THEN EXIT DO
    LOOP
    
    
    IF Abort THEN
      'user pressed Esc then exit with no choice
      VendorPickList = 0
    ELSE
      'user chose a vendor from the list
      'Get the record number off the end of the VendorList
      VendorPickList = CVI(MID$(VendorList(Choice).V, 63, 2))
      
    END IF
    
    ERASE VendorList
    MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen
    
    
    
  ELSEIF MatchingVendors = 1 THEN
    VendorPickList = CVI(MID$(VendorList(1).V, 63, 2))
  ELSE
    Ok = MsgBox("AP", "NOVENDOR")
    VendorPickList = 0
    
  END IF
  
END FUNCTION

