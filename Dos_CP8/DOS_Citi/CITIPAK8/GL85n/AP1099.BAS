DEFINT A-Z
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans&, RecLen%)
DECLARE SUB OpenVendorIdx (VendorIdxFile%, NumActiveVendors%)
DECLARE SUB Extract1099Info ()
DECLARE SUB Print1099Report ()
DECLARE SUB Print1099Forms ()
DECLARE SUB EditPayer ()
DECLARE SUB Edit1099Rec (RecNum%)
DECLARE SUB Edit1099Forms ()
DECLARE FUNCTION VendorPickList% (Lookfor$, Abort%)
DECLARE SUB IndexVendorFile ()
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)

TYPE FLen
  V AS STRING * 64
END TYPE

TYPE ALen
  V AS STRING * 42
END TYPE

DECLARE SUB CursorOff ()
DECLARE SUB PrintTitle (Title$)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
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
DECLARE SUB VertMenuT2 (Items() AS FLen, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, x%, y%, Wdth%, TitColor%)
DECLARE FUNCTION Round# (D#)
'DECLARE FUNCTION Value# (E$, ErrCode%)

'$INCLUDE: 'ap.bi'
'$INCLUDE: 'ap1099f.bi'
'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'SCRAUX.BI'
'$INCLUDE: 'GL.BI'

DECLARE SUB Open1099File (FRecLen, NumRecs, Fed1099File)
DECLARE SUB Form2Rec (Form$(), AP1099() AS AP1099RecType)
DECLARE SUB Rec2Form (Form$(), AP1099() AS AP1099RecType)
DECLARE SUB New1099Rec ()

CONST False = 0, True = NOT False

DIM SHARED VendorIdx AS VendorIdxRecType
DIM SHARED Vendor AS VendorRecType

'$INCLUDE: 'SetCnf.bi'
'$INCLUDE: 'init.bi'

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
                
  MChoice$(1) = "Setup Payer Information"
  MChoice$(2) = "Extract 1099 information"
  MChoice$(3) = "Add a New 1099 Record"
  MChoice$(4) = "Edit/View a 1099 Record"
  MChoice$(5) = "Print 1099 Report"
  MChoice$(6) = "Print 1099 Forms"
  MChoice$(7) = "Quit"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2

  DO

    LibFile2Scrn "AP.QSL", "BG", MonoCode, 0, ErrorCode
    MnuTitle "1099 Misc Forms", Row, Col, MaxLen, 15
    PrintHelp "Federal 1099 Form Processing"
    QPrintRC SPACE$(80), 1, 1, 112
    PrintTitle User$
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    SELECT CASE Choice
    CASE 1
      EditPayer
    CASE 2
      Extract1099Info
    CASE 3
      New1099Rec
    CASE 4
      Edit1099Forms
    CASE 5
      Print1099Report
    CASE 6
      Print1099Forms
    CASE 7
      HideCursor
      CLS
      END
    CASE ELSE
    END SELECT
    
  LOOP
  ClrBG
  RUN "APMENU"

END

SUB comments

'1099 printing created 01/26/97.  Cuttin' it close huh?
'completed Jan 28, 1997 10:35 pm

END SUB

SUB Edit1099Forms
  
  SHARED User$
  REDIM AP1099(1) AS AP1099RecType
  FRecLen = LEN(AP1099(1))
  Open1099File FRecLen, NumRecs, Fed1099File
  IF NumRecs = 0 THEN
    CLOSE
    EXIT SUB
  END IF

  REDIM ActiveList(1 TO NumRecs) AS FLen
  L$ = SPACE$(42)

  FOR R = 1 TO NumRecs
    GET Fed1099File, R, AP1099(1)
    IF NOT AP1099(1).Deleted THEN
      F = F + 1
      ActiveList(F).V = L$
      MID$(ActiveList(F).V, 2) = AP1099(1).RecNAME
      MID$(ActiveList(F).V, 40) = MKI$(R)
    END IF
  NEXT
  CLOSE

  IF F > 0 THEN
    REDIM PRESERVE ActiveList(1 TO F) AS FLen
    Row = 6: Col = 20
    MaxWidth = 38
    BoxBot = 18

    LOCATE Row, Col
    Action = 0
    DO
      LOCATE Row, Col
      LibFile2Scrn "AP.QSL", "BG", MonoCode, 0, ErrorCode
      MnuTitle "Recipient Name", Row, Col, MaxWidth, 15
      PrintHelp "Edit 1099 Forms"
      QPrintRC SPACE$(80), 1, 1, 112
      PrintTitle User$

      VertMenuT2 ActiveList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        CLOSE
        Abort = True
        EXIT DO
      END IF
      RecNum = CVI(MID$(ActiveList(Choice).V, 40, 2))
      Edit1099Rec RecNum

    LOOP
  END IF
  ERASE ActiveList

END SUB

SUB Edit1099Rec (RecNum)

  REDIM AP1099(1) AS AP1099RecType
  FRecLen = LEN(AP1099(1))
  LibName$ = "AP.QSL"
  FormName$ = "AP1099"
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  Open1099File FRecLen, Num1099Recs, Fed1099File
  GET Fed1099File, RecNum, AP1099(1)
  Rec2Form Form$(), AP1099()
  CLOSE
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode
      CASE F10Key
        Form2Rec Form$(), AP1099()
        Open1099File FRecLen, NumRecs, Fed1099File
        PUT Fed1099File, RecNum, AP1099(1)
        CLOSE
        Done = True
      CASE F3Key
        AP1099(1).Deleted = -1
        Open1099File FRecLen, NumRecs, Fed1099File
        PUT Fed1099File, RecNum, AP1099(1)
        CLOSE
        Done = True
      CASE ESC
        Done = True
      CASE ELSE
    END SELECT

  LOOP UNTIL Done
  
  
END SUB

SUB EditPayer

  REDIM Payer(1) AS AP1099PayerRecType
  PRecLen = LEN(Payer(1))
  LibName$ = "AP.QSL"
  FormName$ = "AP1099U"
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  GOSUB GetPayerRec
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode
      CASE F10Key
        GOSUB SavePayerRec
        Done = True
      CASE ESC
        Done = True
      CASE ELSE
    END SELECT

  LOOP UNTIL Done

EXIT SUB


GetPayerRec:
  OPEN "APPAYER.DAT" FOR RANDOM AS #1 LEN = PRecLen
  IF LOF(1) > 0 THEN
    GET #1, 1, Payer(1)
    Form$(1, 0) = Payer(1).FedID
    Form$(2, 0) = Payer(1).NAME
    Form$(3, 0) = Payer(1).ADDR
    Form$(4, 0) = Payer(1).ADDR2
    Form$(5, 0) = Payer(1).CSZ
  END IF
  CLOSE
  RETURN


SavePayerRec:
  OPEN "APPAYER.DAT" FOR RANDOM AS #1 LEN = PRecLen
  Payer(1).FedID = Form$(1, 0)
  Payer(1).NAME = Form$(2, 0)
  Payer(1).ADDR = Form$(3, 0)
  Payer(1).ADDR2 = Form$(4, 0)
  Payer(1).CSZ = Form$(5, 0)
  PUT #1, 1, Payer(1)
  CLOSE
  RETURN

END SUB

SUB Extract1099Info

  Year$ = QPTrim$(STR$(VAL(RIGHT$(DATE$, 4)) - 1))

  LibName$ = "AP.QSL"
  FormName$ = "EXT1099"
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  QPrintRC "For Year: " + Year$, 6, 33, 15

  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode
      CASE F10Key
        Done = True
      CASE ESC
        Cancel = True
        Done = True
      CASE ELSE
    END SELECT

  LOOP UNTIL Done
  
  IF Cancel THEN EXIT SUB
  
  LibFile2Scrn "AP.QSL", "BG", MonoCode, 0, ErrorCode

  LDate = Date2Num("01/01/" + Year$)
  HDate = Date2Num("12/31/" + Year$)
  
  SHARED VendorIdx AS VendorIdxRecType
  OpenVendorIdx VendorIdxFile, NumActiveVendors

  SHARED Vendor AS VendorRecType
  OpenVendorFile VendorFile, NumVRecs
  
  DIM APLedger AS APLedger81RecType
  LdRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans&, LdRecLen
  
  DIM AP1099 AS AP1099RecType
  FRecLen = LEN(AP1099)
  Open1099File FRecLen, Num1099Recs, Fed1099File

  IF Num1099Recs > 0 THEN
    CLOSE Fed1099File
    KILL "AP1099.DAT"
    Open1099File FRecLen, Num1099Recs, Fed1099File
  END IF
  
  FOR Cnt = 1 TO NumActiveVendors
    GET VendorIdxFile, Cnt, VendorIdx
    GET VendorFile, VendorIdx.RecNum, Vendor
    QPrintRC "Searching: " + Vendor.VName, 25, 2, -1
    IF Vendor.Get1099 = "Y" THEN
      'IF INSTR(Vendor.VName, "ROY") > 0 THEN STOP
      NextTran& = Vendor.FrstTran
      IF NextTran& > 0 THEN
        V1099Amt# = 0
        DO
          GET APLedgerFile, NextTran&, APLedger
          IF APLedger.TrCode = 1 THEN
            IF APLedger.Get1099 = "Y" THEN
            IF APLedger.TrDate > LDate AND APLedger.TrDate < HDate THEN
              V1099Amt# = V1099Amt# + APLedger.Amt
            END IF
            END IF
          END IF
          NextTran& = APLedger.NextTrans
        LOOP UNTIL NextTran& = 0
        IF V1099Amt# > 600 THEN
          GOSUB Send2File
        END IF
      END IF
    END IF
  NEXT

  CLOSE

EXIT SUB

Send2File:
  Num1099Recs = Num1099Recs + 1

  AP1099.Deleted = 0
  AP1099.RecID = Vendor.FedID
  AP1099.RecNAME = Vendor.VName
  AP1099.RecADDR = ""
  AP1099.RecADDR2 = Vendor.Addr1
  AP1099.RecCSZ = QPTrim$(Vendor.City) + " " + Vendor.State + " " + Vendor.Zip
  AP1099.RecACCT = ""
  AP1099.NOTICE = ""
  AP1099.BOX1 = 0
  AP1099.BOX2 = 0
  AP1099.BOX3 = 0
  AP1099.BOX4 = 0
  AP1099.BOX5 = 0
  AP1099.BOX6 = 0
  AP1099.BOX7 = V1099Amt#
  AP1099.BOX8 = 0
  AP1099.BOX9 = ""
  AP1099.BOX10 = 0
  AP1099.BOX11 = 0
  AP1099.BOX12 = ""

  PUT Fed1099File, Num1099Recs, AP1099

  RETURN


END SUB

SUB Form2Rec (Form$(), AP1099() AS AP1099RecType)
  
  AP1099(1).Deleted = Value#(Form$(1, 0), ErrCode)
  AP1099(1).RecID = Form$(2, 0)
  AP1099(1).RecNAME = Form$(3, 0)
  AP1099(1).RecADDR = Form$(4, 0)
  AP1099(1).RecADDR2 = Form$(5, 0)
  AP1099(1).RecCSZ = Form$(6, 0)
  AP1099(1).RecACCT = Form$(7, 0)
  AP1099(1).NOTICE = Form$(8, 0)
  AP1099(1).BOX1 = Value#(Form$(9, 0), ErrCode)
  AP1099(1).BOX2 = Value#(Form$(10, 0), ErrCode)
  AP1099(1).BOX3 = Value#(Form$(11, 0), ErrCode)
  AP1099(1).BOX4 = Value#(Form$(12, 0), ErrCode)
  AP1099(1).BOX5 = Value#(Form$(13, 0), ErrCode)
  AP1099(1).BOX6 = Value#(Form$(14, 0), ErrCode)
  AP1099(1).BOX7 = Value#(Form$(15, 0), ErrCode)
  AP1099(1).BOX8 = Value#(Form$(16, 0), ErrCode)
  AP1099(1).BOX9 = Form$(17, 0)
  AP1099(1).BOX10 = Value#(Form$(18, 0), ErrCode)
  AP1099(1).BOX11 = Value#(Form$(19, 0), ErrCode)
  AP1099(1).BOX12 = Form$(20, 0)
    
END SUB

SUB New1099Rec
                     
  REDIM AP1099(1) AS AP1099RecType
  FRecLen = LEN(AP1099(1))
  LibName$ = "AP.QSL"
  FormName$ = "AP1099"
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  'PrintTitle User$
  'PrintHelp "New 1099"
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode
      CASE F10Key
        Form2Rec Form$(), AP1099()
        Open1099File FRecLen, NumRecs, Fed1099File
        PUT Fed1099File, NumRecs + 1, AP1099(1)
        CLOSE
        Done = True
      CASE ESC
        Done = True
      CASE ELSE
    END SELECT
    
  LOOP UNTIL Done
  
END SUB

SUB Open1099File (FRecLen, Num1099Recs, Fed1099File)

   Fed1099File = FREEFILE
   OPEN "AP1099.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS Fed1099File LEN = FRecLen
   Num1099Recs = LOF(Fed1099File) \ FRecLen
   
END SUB

SUB Print1099Forms

  LOCATE , , 0
  REDIM ScrnArray(2000)
  
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
  FormName$ = "PRN1099"
  BGScrn$ = "BG"
  LibFile2Scrn LibName$, BGScrn$, MonoCode, 0, ErrorCode
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintHelp "Print Vendor Listing"
  Action = 1

  DO

    EditForm Form$(), Fld(), Frm, Cnf, Action
    IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
      Dev$ = "S"
      LPTNo = 1
    ELSE
      Dev$ = LEFT$(Form$(1, 0), 1)
      LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
    END IF

    SELECT CASE Frm.KeyCode
      CASE F5KEY
        LOCATE , , 0
        MScrnSave 1, 1, 25, 80, ScrnArray(0)
        EntryPoint = 5
        PrintRptFile Header$, "AP1099.MSK", LPTPort%, RetCode%, EntryPoint
        MScrnRest 1, 1, 25, 80, ScrnArray(0)
      CASE F10Key
        Done = True
      CASE ESC
        ExitFlag = True
        Done = True
    END SELECT
  LOOP UNTIL Done ' Frm.KeyCode = F10Key

  IF ExitFlag THEN EXIT SUB

  LibFile2Scrn LibName$, BGScrn$, MonoCode, Attribute, ErrorCode
  PrintHelp "Processing report. Please wait."
  
  '--Get the Payer info
  REDIM Payer(1) AS AP1099PayerRecType
  PRecLen = LEN(Payer(1))
  PayerFile = FREEFILE
  OPEN "APPAYER.DAT" FOR RANDOM AS PayerFile LEN = PRecLen
  GET PayerFile, 1, Payer(1)
  CLOSE PayerFile
  
  '--Open the 1099 data file
  REDIM AP1099(1) AS AP1099RecType
  FRecLen = LEN(AP1099(1))
  Open1099File FRecLen, Num1099Recs, Fed1099File
  
  '--Open the print file
  PrnFile = FREEFILE
  ReportFile$ = "AP1099.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  Fmt$ = "######.##"
  
  '--Print the Forms
  FOR R = 1 TO Num1099Recs
    GET Fed1099File, R, AP1099(1)
    IF NOT AP1099(1).Deleted THEN
      GOSUB PrintForm
'      didcnt = didcnt + 1
'      IF didcnt > 0 THEN
'        EXIT FOR
'      END IF
    END IF
  NEXT
  PRINT #PrnFile, CHR$(12)

  CLOSE
 
  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE ELSE
      EntryPoint = 5
  END SELECT
  

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  'KILL ReportFile$

EXIT SUB


PrintForm:
  PRINT #PrnFile, "~"; TAB(70); "~"
  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile, TAB(6); Payer(1).NAME;
  IF AP1099(1).BOX1 > 0 THEN
    PRINT #PrnFile, TAB(42); USING Fmt$; AP1099(1).BOX1
  ELSE
    PRINT #PrnFile,
  END IF
  PRINT #PrnFile, TAB(6); Payer(1).ADDR
  PRINT #PrnFile, TAB(6); Payer(1).ADDR2
  PRINT #PrnFile, TAB(6); Payer(1).CSZ;
  IF AP1099(1).BOX2 > 0 THEN
    PRINT #PrnFile, TAB(42); USING Fmt$; AP1099(1).BOX2
  ELSE
    PRINT #PrnFile,
  END IF
  PRINT #PrnFile,
  PRINT #PrnFile, 'TAB(6); LEFT$(Payer(1).FedID, 16);

  IF AP1099(1).BOX3 > 0 THEN
    PRINT #PrnFile, TAB(42); USING Fmt$; AP1099(1).BOX3;
  END IF
  IF AP1099(1).BOX4 > 0 THEN
    PRINT #PrnFile, TAB(57); USING Fmt$; AP1099(1).BOX4;
  END IF
  PRINT #PrnFile,

  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile, TAB(6); LEFT$(Payer(1).FedID, 16); TAB(24); QPTrim$(AP1099(1).RecID);

  IF AP1099(1).BOX5 > 0 THEN
    PRINT #PrnFile, TAB(42); USING Fmt$; AP1099(1).BOX5;
  END IF
  IF AP1099(1).BOX6 > 0 THEN
    PRINT #PrnFile, TAB(57); USING Fmt$; AP1099(1).BOX6;
  END IF
  PRINT #PrnFile,


  PRINT #PrnFile,
  PRINT #PrnFile, TAB(6); AP1099(1).RecNAME
  PRINT #PrnFile,

  IF AP1099(1).BOX7 > 0 THEN
    PRINT #PrnFile, TAB(42); USING Fmt$; AP1099(1).BOX7;
  END IF
  IF AP1099(1).BOX8 > 0 THEN
    PRINT #PrnFile, TAB(57); USING Fmt$; AP1099(1).BOX8;
  END IF
  PRINT #PrnFile,

  PRINT #PrnFile,

  PRINT #PrnFile, TAB(6); AP1099(1).RecADDR2; TAB(57);
  IF AP1099(1).BOX10 > 0 THEN
    PRINT #PrnFile, USING Fmt$; AP1099(1).BOX10
  ELSE
    PRINT #PrnFile,
  END IF

  PRINT #PrnFile,
  PRINT #PrnFile, TAB(36); AP1099(1).NOTICE
  PRINT #PrnFile, TAB(6); AP1099(1).RecCSZ
  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile,

  IF AP1099(1).BOX11 > 0 THEN
    PRINT #PrnFile, TAB(42); USING Fmt$; AP1099(1).BOX11;
  END IF

  PRINT #PrnFile, TAB(57); AP1099(1).BOX12

  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile,
  PRINT #PrnFile, "~"; TAB(70); "~"
  RETURN

  
END SUB

SUB Print1099Report

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
  FormName$ = "PRN1099R"
  BGScrn$ = "BG"
  LibFile2Scrn LibName$, BGScrn$, MonoCode, 0, ErrorCode
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  PrintHelp "Print Vendor Listing"
  Action = 1

  DO

    EditForm Form$(), Fld(), Frm, Cnf, Action
    IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
      Dev$ = "S"
      LPTNo = 1
    ELSE
      Dev$ = LEFT$(Form$(1, 0), 1)
      LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
    END IF

    SELECT CASE Frm.KeyCode
      CASE F10Key
        Done = True
      CASE ESC
        ExitFlag = True
        Done = True
    END SELECT
  LOOP UNTIL Done ' Frm.KeyCode = F10Key

  IF ExitFlag THEN EXIT SUB

  LibFile2Scrn LibName$, BGScrn$, MonoCode, Attribute, ErrorCode
  PrintHelp "Processing report. Please wait."

 '--Open the 1099 data file
  REDIM AP1099(1) AS AP1099RecType
  FRecLen = LEN(AP1099(1))
  Open1099File FRecLen, Num1099Recs, Fed1099File

  '--Open the print file
  PrnFile = FREEFILE
  ReportFile$ = "AP1099.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  Fmt$ = "$$#######,#.##"
  Header$ = "1099 Misc Report Totals"

  '--Print the Forms
  Page = 0
  GOSUB Heading

  FOR R = 1 TO Num1099Recs
    GET Fed1099File, R, AP1099(1)
    IF lc >= 60 THEN
     PRINT #PrnFile, CHR$(12);
     GOSUB Heading
    END IF
   IF NOT AP1099(1).Deleted THEN

      TotAmtRep# = Round#(TotAmtRep# + AP1099(1).BOX1 + AP1099(1).BOX2)
      TotAmtRep# = Round#(TotAmtRep# + AP1099(1).BOX3 + AP1099(1).BOX5)
      TotAmtRep# = Round#(TotAmtRep# + AP1099(1).BOX6 + AP1099(1).BOX7)
      TotAmtRep# = Round#(TotAmtRep# + AP1099(1).BOX8 + AP1099(1).BOX10)
      TaxesWH# = Round#(TaxesWH# + AP1099(1).BOX4)
      GrandTotAmtRep# = GrandTotAmtRep# + TotAmtRep#

    PRINT #PrnFile, AP1099(1).RecID;
      PRINT #PrnFile, TAB(20); RTRIM$(AP1099(1).RecNAME); TAB(65); USING "$$#####,#.##"; TotAmtRep#
      PRINT #PrnFile, RTRIM$(AP1099(1).RecADDR) + " " + RTRIM$(AP1099(1).RecADDR2);
      PRINT #PrnFile, TAB(50); RTRIM$(AP1099(1).RecCSZ)
      PRINT #PrnFile, STRING$(79, "-")
      lc = lc + 3
      TotForms = TotForms + 1
      TotAmtRep# = 0
    END IF
  NEXT

  PRINT #PrnFile, "1099 Misc Report Totals"
  PRINT #PrnFile, "---------------------------------"
  PRINT #PrnFile, "Total Forms:"; TAB(24); USING "###"; TotForms
  PRINT #PrnFile, "Federal Tax W/H:"; TAB(20); USING Fmt$; TaxesWH#
  PRINT #PrnFile, "Total Amt Reported:"; TAB(20); USING Fmt$; GrandTotAmtRep#
  PRINT #PrnFile, CHR$(12)

  CLOSE

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  KILL ReportFile$
  EXIT SUB

Heading:
  Page = Page + 1
  PRINT #PrnFile, TAB(30); "1099 Misc Report"
  PRINT #PrnFile,
  PRINT #PrnFile, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #PrnFile, "Acct #"; TAB(20); "Acct Name"; TAB(65); "Amount"
  PRINT #PrnFile, "Address"
  PRINT #PrnFile, STRING$(79, "="): lc = 7
  RETURN
  
END SUB

SUB Rec2Form (Form$(), AP1099() AS AP1099RecType)
    
  Form$(1, 0) = STR$(AP1099(1).Deleted)
  Form$(2, 0) = AP1099(1).RecID
  Form$(3, 0) = AP1099(1).RecNAME
  Form$(4, 0) = AP1099(1).RecADDR
  Form$(5, 0) = AP1099(1).RecADDR2
  Form$(6, 0) = AP1099(1).RecCSZ
  Form$(7, 0) = AP1099(1).RecACCT
  Form$(8, 0) = AP1099(1).NOTICE
  Form$(9, 0) = STR$(AP1099(1).BOX1)
  Form$(10, 0) = STR$(AP1099(1).BOX2)
  Form$(11, 0) = STR$(AP1099(1).BOX3)
  Form$(12, 0) = STR$(AP1099(1).BOX4)
  Form$(13, 0) = STR$(AP1099(1).BOX5)
  Form$(14, 0) = STR$(AP1099(1).BOX6)
  Form$(15, 0) = STR$(AP1099(1).BOX7)
  Form$(16, 0) = STR$(AP1099(1).BOX8)
  Form$(17, 0) = AP1099(1).BOX9
  Form$(18, 0) = STR$(AP1099(1).BOX10)
  Form$(19, 0) = STR$(AP1099(1).BOX11)
  Form$(20, 0) = AP1099(1).BOX12
                 
END SUB

