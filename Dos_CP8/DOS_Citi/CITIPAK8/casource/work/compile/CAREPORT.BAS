DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SmallPause ()
DEFINT A-Z
DECLARE SUB MailingLabels ()
DECLARE SUB CustomerBalances ()
DECLARE SUB OpenCaCustFile (NumOfCaRecs%, CaFile%)
DECLARE SUB DisplayCaScrn (ScrnName$)
DECLARE SUB CustomerListing ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (N#)
'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'CaCust.BI'
  
  DIM SHARED CaCust(1) AS CACUSTType
  DIM SHARED CaBegTr(1) AS CaBegTransType
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  STACK 5000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = "Master Customer Listing"
  MChoice$(2) = "Master Balance Listing"
  MChoice$(3) = "Mailing Labels"
  'MChoice$() = "Transaction Journal"
  
  MChoice$(4) = "Exit to DOS"
  
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
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
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Assessments Report Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      CustomerListing
    CASE 2
      CustomerBalances
    CASE 3
      MailingLabels
    CASE 4

      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF

  IF Exist("Camenu" + Ext$) THEN
    RUN "CaMenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB CustomerBalances

  SHARED Choice$()

  ReportFile$ = "CaCusBal.PRN"   'Report File Name

  Dash80$ = STRING$(80, "=")

  size = 5500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:

  REDIM array(1 TO size) AS Struct

  FF$ = CHR$(12)

  MaxLines = 60
  LineCnt = 0
  CustCnt = 0

  GOSUB SelectBalanceCustomerOutput

  IF Canceled$ = "Y" THEN EXIT SUB
  IF AbortFlag THEN EXIT SUB

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintBalanceCustomerRptHeader

  ' Print Main Body

  OpenCaCustFile NumOfCaRecs, CaFile

  IF UsingIndex THEN
    GOSUB GetNameIndex1
  END IF

  ShowProcessingScrn ("Customer Balance Listing")
  FOR Cnt = 1 TO NumOfCaRecs

    IF UsingIndex THEN
      CustRecNo = array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF

    GET CaFile, CustRecNo, CaCust(1)


    help$ = "Processing Record # " + STR$(Cnt!)
     'Main Print Processing Here

    IF NOT (CaCust(1).Deleted) THEN

      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintBalanceCustomerRptHeader
      END IF


   'Determine if Customer HAS A BALANCE & PRINT IF THEY DO!

        CustBalance# = Round#(CaCust(1).RPBStreet + CaCust(1).RPBSide + CaCust(1).RPBDrive + CaCust(1).RPBDemo + CaCust(1).RPBWater + CaCust(1).RPBSewer)
        CustBalance# = Round#(CustBalance# + CaCust(1).RIBStreet + CaCust(1).RIBSide + CaCust(1).RIBDrive + CaCust(1).RIBDemo + CaCust(1).RIBWater + CaCust(1).RIBSewer)

        IF CustBalance# <= 0 THEN GOTO CUSTSKIP

        RunPrinTotal# = 0
        RunIntTotal# = 0
        
        PRINT #RptHandle, CustRecNo; TAB(10); LEFT$(RTRIM$(CaCust(1).CustName), 28);
         
        IF CaCust(1).RPBStreet + CaCust(1).RIBStreet > 0 THEN
         PRINT #RptHandle, TAB(40); "Streets";
         PRINT #RptHandle, TAB(50); USING "#####.##"; CaCust(1).RPBStreet;
          PRINT #RptHandle, TAB(60); USING "#####.##"; CaCust(1).RIBStreet;
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBStreet + CaCust(1).RIBStreet
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBStreet
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBStreet
          RunPrinStreet# = RunPrinStreet# + CaCust(1).RPBStreet
          RunIntStreet# = RunIntStreet# + CaCust(1).RIBStreet
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBSide + CaCust(1).RIBSide > 0 THEN
         PRINT #RptHandle, TAB(40); "SideWalks";
          PRINT #RptHandle, TAB(50); USING "#####.##"; CaCust(1).RPBSide;
          PRINT #RptHandle, TAB(60); USING "#####.##"; CaCust(1).RIBSide;
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBSide + CaCust(1).RIBSide
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBSide
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBSide
          RunPrinSide# = RunPrinSide# + CaCust(1).RPBSide
          RunIntSide# = RunIntSide# + CaCust(1).RIBSide
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBDrive + CaCust(1).RIBDrive > 0 THEN
         PRINT #RptHandle, TAB(40); "DriveWays";
          PRINT #RptHandle, TAB(50); USING "#####.##"; CaCust(1).RPBDrive;
          PRINT #RptHandle, TAB(60); USING "#####.##"; CaCust(1).RIBDrive;
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBDrive + CaCust(1).RIBDrive
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBDrive
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBDrive
          RunPrinDrive# = RunPrinDrive# + CaCust(1).RPBDrive
          RunIntDrive# = RunIntDrive# + CaCust(1).RIBDrive
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBDemo + CaCust(1).RIBDemo > 0 THEN
         PRINT #RptHandle, TAB(40); "Dredging";
          PRINT #RptHandle, TAB(50); USING "#####.##"; CaCust(1).RPBDemo;
          PRINT #RptHandle, TAB(60); USING "#####.##"; CaCust(1).RIBDemo;
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBDemo + CaCust(1).RIBDemo
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBDemo
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBDemo
          RunPrinDemo# = RunPrinDemo# + CaCust(1).RPBDemo
          RunIntDemo# = RunIntDemo# + CaCust(1).RIBDemo
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBWater + CaCust(1).RIBWater > 0 THEN
         PRINT #RptHandle, TAB(40); "Water";
          PRINT #RptHandle, TAB(50); USING "#####.##"; CaCust(1).RPBWater;
          PRINT #RptHandle, TAB(60); USING "#####.##"; CaCust(1).RIBWater;
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBWater + CaCust(1).RIBWater
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBWater
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBWater
          RunPrinWater# = RunPrinWater# + CaCust(1).RPBWater
          RunIntWater# = RunIntWater# + CaCust(1).RIBWater
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBSewer + CaCust(1).RIBSewer > 0 THEN
         PRINT #RptHandle, TAB(40); "Sewer";
          PRINT #RptHandle, TAB(50); USING "#####.##"; CaCust(1).RPBSewer;
          PRINT #RptHandle, TAB(60); USING "#####.##"; CaCust(1).RIBSewer;
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBSewer + CaCust(1).RIBSewer
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBSewer
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBSewer
          RunPrinSewer# = RunPrinSewer# + CaCust(1).RPBSewer
          RunIntSewer# = RunIntSewer# + CaCust(1).RIBSewer
          LineCnt = LineCnt + 1
        END IF

        IF RunPrinTotal# <> 0 OR RunIntTotal# <> 0 THEN
          PRINT #RptHandle, "Legal Desc:"; LEFT$(CaCust(1).LegalDes, 25);
          PRINT #RptHandle, TAB(39); "*TOTAL DUE";
          PRINT #RptHandle, TAB(50); USING "#####.##"; RunPrinTotal#;
          PRINT #RptHandle, TAB(60); USING "#####.##"; RunIntTotal#;
          PRINT #RptHandle, TAB(70); USING "#####.##"; RunPrinTotal# + RunIntTotal#
          PRINT #RptHandle, STRING$(80, "-")
          LineCnt = LineCnt + 4
        END IF
      END IF
    CustCnt = CustCnt + 1
CUSTSKIP:
    SmallPause
    ShowPctComp Cnt, NumOfCaRecs
    SmallPause
    SmallPause
    NEXT Cnt
  PRINT #RptHandle, FF$
  GOSUB PrintBalanceCustomerRptEnding

  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
'**
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
'**
  ELSE
    EntryPoint = 1
  END IF

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintBalanceCustomerRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(30); "Assessment Balance Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, "Acct #"; TAB(10); "Customer Name"; TAB(50); "Prin Due"; TAB(60); "Int Due"; TAB(70); "Total Due"
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN

PrintBalanceCustomerRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, "Assessment Totals Breakdown"
  PRINT #RptHandle,
  PRINT #RptHandle, "Assessment"; TAB(19); "Prin Balance"; TAB(35); "Int Balance"; TAB(52); "Total Due"
  PRINT #RptHandle, "Street"; TAB(20); USING "######,#.##"; RunPrinStreet#;
   PRINT #RptHandle, TAB(35); USING "######,#.##"; RunIntStreet#;
   PRINT #RptHandle, TAB(50); USING "######,#.##"; RunPrinStreet# + RunIntStreet#
  PRINT #RptHandle, "SideWalks"; TAB(20); USING "######,#.##"; RunPrinSide#;
   PRINT #RptHandle, TAB(35); USING "######,#.##"; RunIntSide#;
   PRINT #RptHandle, TAB(50); USING "######,#.##"; RunPrinSide# + RunIntSide#
  PRINT #RptHandle, "DriveWays"; TAB(20); USING "######,#.##"; RunPrinDrive#;
   PRINT #RptHandle, TAB(35); USING "######,#.##"; RunIntDrive#;
   PRINT #RptHandle, TAB(50); USING "######,#.##"; RunPrinDrive# + RunIntDrive#
  PRINT #RptHandle, "Dredging"; TAB(20); USING "######,#.##"; RunPrinDemo#;
   PRINT #RptHandle, TAB(35); USING "######,#.##"; RunIntDemo#;
   PRINT #RptHandle, TAB(50); USING "######,#.##"; RunPrinDemo# + RunIntDemo#
  PRINT #RptHandle, "Water"; TAB(20); USING "######,#.##"; RunPrinWater#;
   PRINT #RptHandle, TAB(35); USING "######,#.##"; RunIntWater#;
   PRINT #RptHandle, TAB(50); USING "######,#.##"; RunPrinWater# + RunIntWater#
  PRINT #RptHandle, "Sewer"; TAB(20); USING "######,#.##"; RunPrinSewer#;
   PRINT #RptHandle, TAB(35); USING "######,#.##"; RunIntSewer#;
   PRINT #RptHandle, TAB(50); USING "######,#.##"; RunPrinSewer# + RunIntSewer#
  PrinTotal# = Round#(RunPrinStreet# + RunPrinSide# + RunPrinDrive# + RunPrinDemo# + RunPrinWater# + RunPrinSewer#)
  IntTotal# = Round#(RunIntStreet# + RunIntSide# + RunIntDrive# + RunIntDemo# + RunIntWater# + RunIntSewer#)
  PRINT #RptHandle, STRING$(62, "=")
  PRINT #RptHandle, "Totals"; TAB(20); USING "######,#.##"; PrinTotal#;
   PRINT #RptHandle, TAB(35); USING "######,#.##"; IntTotal#;
   PRINT #RptHandle, TAB(50); USING "######,#.##"; PrinTotal# + IntTotal#



 PRINT #RptHandle, FF$
RETURN

SelectBalanceCustomerOutput:
  LibName$ = "TAX"
  ScrnName$ = "CUSTRPT1"

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

  REDIM Choice$(0 TO 2, 0 TO 2)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"

  Action = 1
  ShowCursor
  ClearBack
  DisplayCaScrn ScrnName$
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  LOCATE 8, 33: COLOR 14: PRINT "Balances";
  'printhelp help$
  Form$(2, 0) = "Summary Only"
  Fld(2).Protected = True
  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key

      IF LEFT$(Form$(1, 0), 4) = "Name" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF

      Order$ = LEFT$(Form$(1, 0), 1)

      IF LEFT$(Form$(2, 0), 1) = "S" THEN
        DetailFlag = False
      ELSE
        DetailFlag = True
      END IF
      Detail$ = Form$(2, 0)

      DevSpec$ = LEFT$(Form$(3, 0), 1)
      ExitFlag = True
      RETURN
    CASE EscKey
      AbortFlag = True
'Canceled$ = "Y"
      ExitFlag = True 'EXIT DO
'RETURN
    END SELECT
  LOOP UNTIL ExitFlag

RETURN

GetNameIndex1:
  FOR SCnt = 1 TO NumOfCaRecs
    GET CaFile, SCnt, CaCust(1)
    array(SCnt).who = UCASE$(CaCust(1).CustName)
    array(SCnt).RecNum = SCnt
    Count = NumOfCaRecs
  NEXT SCnt

'Sort Them Here
  SortT array(Start), Count, Dir, SSize, MOff, MSize
  RETURN

END SUB

SUB CustomerListing
  
  SHARED Choice$()

  ReportFile$ = "CaCust.PRN"   'Report File Name

  Dash80$ = STRING$(80, "=")

  size = 5500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:

  REDIM array(1 TO size) AS Struct
  
  FF$ = CHR$(12)

  MaxLines = 60
  LineCnt = 0
  CustCnt = 0
  
  GOSUB SelectDetailCustomerOutput

  IF Canceled$ = "Y" THEN EXIT SUB
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintDetailCustomerRptHeader
  
  ' Print Main Body

  OpenCaCustFile NumOfCaRecs, CaFile
  
  IF UsingIndex THEN
    GOSUB GetNameIndex
  END IF
  ShowProcessingScrn "Master Customer Listing"
  FOR Cnt = 1 TO NumOfCaRecs

    IF UsingIndex THEN
      CustRecNo = array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF

    GET CaFile, CustRecNo, CaCust(1)
    

    help$ = "Processing Record # " + STR$(Cnt!)
     'Main Print Processing Here
    
    IF NOT (CaCust(1).Deleted) THEN

      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintDetailCustomerRptHeader
      END IF
      
      IF NOT DetailFlag THEN
        PRINT #RptHandle, "Acct #"; CustRecNo; TAB(15); RTRIM$(CaCust(1).CustName); TAB(42); RTRIM$(CaCust(1).LegalDes)
        PRINT #RptHandle, TAB(15); RTRIM$(CaCust(1).Add1); TAB(42); RTRIM$(CaCust(1).Parcel)
        PRINT #RptHandle, TAB(15); RTRIM$(CaCust(1).Add2); TAB(42); RTRIM$(CaCust(1).PropDesc)
        PRINT #RptHandle, STRING$(80, "-")
        LineCnt = LineCnt + 4
      ELSE
        PRINT #RptHandle, "Acct #"; CustRecNo; TAB(15); RTRIM$(CaCust(1).CustName); TAB(42); RTRIM$(CaCust(1).LegalDes)
        PRINT #RptHandle, TAB(15); RTRIM$(CaCust(1).Add1); TAB(42); RTRIM$(CaCust(1).Parcel)
        PRINT #RptHandle, TAB(15); RTRIM$(CaCust(1).Add2); TAB(42); RTRIM$(CaCust(1).PropDesc)
        PRINT #RptHandle, ""
        PRINT #RptHandle, "Assessment"; TAB(22); "Amount"; TAB(31); "Term"; TAB(40); "Rate"; TAB(50); "Prin Paid"; TAB(60); "Int Paid"; TAB(70); "Prin Bal"; TAB(81); "Int Bal"; TAB(90); "Total Due"
        LineCnt = LineCnt + 5
      ' Now Go Thru Each of Six Assessments and Determine Which is Due and Print Them
        RunPrinTotal# = 0
        RunIntTotal# = 0

        IF CaCust(1).RPBStreet + CaCust(1).RIBStreet > 0 THEN
         PRINT #RptHandle, "Streets"; TAB(20); USING "#####.##"; CaCust(1).StrAmt;
          PRINT #RptHandle, TAB(30); USING "####"; CaCust(1).StrTerm;
          PRINT #RptHandle, TAB(40); USING "#.##"; CaCust(1).StrRate;
          PRINT #RptHandle, TAB(50); USING "#####.##"; (CaCust(1).StrAmt - CaCust(1).RPBStreet);
          PRINT #RptHandle, TAB(60); USING "#####.##"; (CaCust(1).IStreet - CaCust(1).RIBStreet);
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBStreet;
          PRINT #RptHandle, TAB(80); USING "#####.##"; CaCust(1).RIBStreet;
          PRINT #RptHandle, TAB(90); USING "#####.##"; CaCust(1).RPBStreet + CaCust(1).RIBStreet
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBStreet
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBStreet
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBSide + CaCust(1).RIBSide > 0 THEN
         PRINT #RptHandle, "SideWalks"; TAB(20); USING "#####.##"; CaCust(1).SideAmt;
          PRINT #RptHandle, TAB(30); USING "####"; CaCust(1).SideTerm;
          PRINT #RptHandle, TAB(40); USING "#.##"; CaCust(1).SideRate;
          PRINT #RptHandle, TAB(50); USING "#####.##"; (CaCust(1).SideAmt - CaCust(1).RPBSide);
          PRINT #RptHandle, TAB(60); USING "#####.##"; (CaCust(1).ISide - CaCust(1).RIBSide);
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBSide;
          PRINT #RptHandle, TAB(80); USING "#####.##"; CaCust(1).RIBSide;
          PRINT #RptHandle, TAB(90); USING "#####.##"; CaCust(1).RPBSide + CaCust(1).RIBSide
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBSide
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBSide
          LineCnt = LineCnt + 1
        END IF
        
        IF CaCust(1).RPBDrive + CaCust(1).RIBDrive > 0 THEN
         PRINT #RptHandle, "DriveWays"; TAB(20); USING "#####.##"; CaCust(1).DriveAmt;
          PRINT #RptHandle, TAB(30); USING "####"; CaCust(1).DriveTerm;
          PRINT #RptHandle, TAB(40); USING "#.##"; CaCust(1).DriveRate;
          PRINT #RptHandle, TAB(50); USING "#####.##"; (CaCust(1).DriveAmt - CaCust(1).RPBDrive);
          PRINT #RptHandle, TAB(60); USING "#####.##"; (CaCust(1).IDrive - CaCust(1).RIBDrive);
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBDrive;
          PRINT #RptHandle, TAB(80); USING "#####.##"; CaCust(1).RIBDrive;
          PRINT #RptHandle, TAB(90); USING "#####.##"; CaCust(1).RPBDrive + CaCust(1).RIBDrive
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBDrive
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBDrive
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBDemo + CaCust(1).RIBDemo > 0 THEN
         PRINT #RptHandle, "Dredging"; TAB(20); USING "#####.##"; CaCust(1).DemoAmt;
          PRINT #RptHandle, TAB(30); USING "####"; CaCust(1).DemoTerm;
          PRINT #RptHandle, TAB(40); USING "#.##"; CaCust(1).DemoRate;
          PRINT #RptHandle, TAB(50); USING "#####.##"; (CaCust(1).DemoAmt - CaCust(1).RPBDemo);
          PRINT #RptHandle, TAB(60); USING "#####.##"; (CaCust(1).IDemo - CaCust(1).RIBDemo);
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBDemo;
          PRINT #RptHandle, TAB(80); USING "#####.##"; CaCust(1).RIBDemo;
          PRINT #RptHandle, TAB(90); USING "#####.##"; CaCust(1).RPBDemo + CaCust(1).RIBDemo
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBDemo
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBDemo
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBWater + CaCust(1).RIBWater > 0 THEN
         PRINT #RptHandle, "Water"; TAB(20); USING "#####.##"; CaCust(1).WaterAmt;
          PRINT #RptHandle, TAB(30); USING "####"; CaCust(1).WaterTerm;
          PRINT #RptHandle, TAB(40); USING "#.##"; CaCust(1).WaterRate;
          PRINT #RptHandle, TAB(50); USING "#####.##"; (CaCust(1).WaterAmt - CaCust(1).RPBWater);
          PRINT #RptHandle, TAB(60); USING "#####.##"; (CaCust(1).IWater - CaCust(1).RIBWater);
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBWater;
          PRINT #RptHandle, TAB(80); USING "#####.##"; CaCust(1).RIBWater;
          PRINT #RptHandle, TAB(90); USING "#####.##"; CaCust(1).RPBWater + CaCust(1).RIBWater
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBWater
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBWater
          LineCnt = LineCnt + 1
        END IF

        IF CaCust(1).RPBSewer + CaCust(1).RIBSewer > 0 THEN
         PRINT #RptHandle, "Sewer"; TAB(20); USING "#####.##"; CaCust(1).SewerAmt;
          PRINT #RptHandle, TAB(30); USING "####"; CaCust(1).SewerTerm;
          PRINT #RptHandle, TAB(40); USING "#.##"; CaCust(1).SewerRate;
          PRINT #RptHandle, TAB(50); USING "#####.##"; (CaCust(1).SewerAmt - CaCust(1).RPBSewer);
          PRINT #RptHandle, TAB(60); USING "#####.##"; (CaCust(1).ISewer - CaCust(1).RIBSewer);
          PRINT #RptHandle, TAB(70); USING "#####.##"; CaCust(1).RPBSewer;
          PRINT #RptHandle, TAB(80); USING "#####.##"; CaCust(1).RIBSewer;
          PRINT #RptHandle, TAB(90); USING "#####.##"; CaCust(1).RPBSewer + CaCust(1).RIBSewer
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBSewer
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBSewer
          LineCnt = LineCnt + 1
        END IF
        
        IF RunPrinTotal# <> 0 OR RunIntTotal# <> 0 THEN
          IF CaCust(1).LastPayDate <> 0 THEN
          PRINT #RptHandle, "Last Paid: "; Num2Date$(CaCust(1).LastPayDate);
          END IF
          IF CaCust(1).LastIntDate <> 0 THEN
          PRINT #RptHandle, "  Last Int: "; Num2Date$(CaCust(1).LastIntDate);
          END IF

          PRINT #RptHandle, TAB(60); "Total Due";
          PRINT #RptHandle, TAB(70); USING "#####.##"; RunPrinTotal#;
          PRINT #RptHandle, TAB(80); USING "#####.##"; RunIntTotal#;
          PRINT #RptHandle, TAB(90); USING "#####.##"; RunPrinTotal# + RunIntTotal#
         LineCnt = LineCnt + 1
        END IF
        PRINT #RptHandle, STRING$(100, "-")
      END IF
      CustCnt = CustCnt + 1
   END IF
   ShowPctComp Cnt, NumOfCaRecs
   SmallPause
  
NEXT Cnt

  GOSUB PrintDetailCustomerRptEnding
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
'**
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
'**
  ELSE
    EntryPoint = 1
  END IF
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintDetailCustomerRptHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(30); "Assessment Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
 
  IF NOT DetailFlag THEN
    PRINT #RptHandle, "Summary Format"
    PRINT #RptHandle, TAB(15); "Owner Name"; TAB(42); "Legal Desc"
    PRINT #RptHandle, TAB(15); "Address Line 1"; TAB(42); "Parcel #"
    PRINT #RptHandle, TAB(15); "Address Line 2"; TAB(42); "Legal Address"
  PRINT #RptHandle, Dash80$
    LineCnt = 4
  ELSE
    PRINT #RptHandle, "Detail Format"
    PRINT #RptHandle, STRING$(132, "=")
    LineCnt = 4
  END IF
    
RETURN
  
PrintDetailCustomerRptEnding:
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, FF$
RETURN
  
SelectDetailCustomerOutput:
  LibName$ = "TAX"
  ScrnName$ = "CUSTRPT1"
  
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
  
  REDIM Choice$(0 TO 2, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  
  Action = 1
  ShowCursor
  ClearBack
  DisplayCaScrn ScrnName$
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  'printhelp help$
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      
      IF LEFT$(Form$(1, 0), 4) = "Name" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF

      Order$ = LEFT$(Form$(1, 0), 1)

      IF LEFT$(Form$(2, 0), 1) = "S" THEN
        DetailFlag = False
      ELSE
        DetailFlag = True
      END IF
      Detail$ = Form$(2, 0)

      DevSpec$ = LEFT$(Form$(3, 0), 1)
      ExitFlag = True
      RETURN
    CASE EscKey
      AbortFlag = True
'Canceled$ = "Y"
      ExitFlag = True 'EXIT DO
'RETURN
    END SELECT
  LOOP UNTIL ExitFlag

RETURN
  
GetNameIndex:
  FOR SCnt = 1 TO NumOfCaRecs
    GET CaFile, SCnt, CaCust(1)
    array(SCnt).who = UCASE$(CaCust(1).CustName)
    array(SCnt).RecNum = SCnt
    Count = NumOfCaRecs
  NEXT SCnt

'Sort Them Here
  SortT array(Start), Count, Dir, SSize, MOff, MSize
  RETURN
  
  
END SUB

SUB MailingLabels

  SHARED Choice$()

  ReportFile$ = "CaML.PRN"   'Report File Name

  Dash80$ = STRING$(80, "=")

  size = 5500
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:

  REDIM array(1 TO size) AS Struct

  FF$ = CHR$(12)

  MaxLines = 60
  LineCnt = 0
  CustCnt = 0

  GOSUB SelectMLOutput

  IF Canceled$ = "Y" THEN EXIT SUB
  IF AbortFlag THEN EXIT SUB

  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintMLRptHeader

  ' Print Main Body

  OpenCaCustFile NumOfCaRecs, CaFile

  IF UsingIndex THEN
    GOSUB GetNameIndexML
  END IF

  FOR Cnt = 1 TO NumOfCaRecs

    IF UsingIndex THEN
      CustRecNo = array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF

    GET CaFile, CustRecNo, CaCust(1)


    help$ = "Processing Record # " + STR$(Cnt!)
     'Main Print Processing Here

    IF NOT (CaCust(1).Deleted) THEN

      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMLRptHeader
      END IF




        PRINT #RptHandle,
        PRINT #RptHandle, TAB(2); "Acct #"; CustRecNo
        PRINT #RptHandle, TAB(2); RTRIM$(CaCust(1).CustName)
        PRINT #RptHandle, TAB(2); RTRIM$(CaCust(1).Add1)
        PRINT #RptHandle, TAB(2); RTRIM$(CaCust(1).Add2)
        PRINT #RptHandle,

END IF
NEXT Cnt

  GOSUB PrintMLRptEnding

  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
'**
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
'**
  ELSE
    EntryPoint = 1
  END IF

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintMLRptHeader:
  'Print 3 alignment labels to start
   PRINT #RptHandle, TAB(2); "--- LABEL TOP ---"
   PRINT #RptHandle, TAB(2); "CUSTOMER ACCT # HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER NAME HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER ADDRESS HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER ADDRESS HERE"
   PRINT #RptHandle, TAB(2); "--- LABEL BOTTOM ---"
   PRINT #RptHandle, TAB(2); "--- LABEL TOP ---"
   PRINT #RptHandle, TAB(2); "CUSTOMER ACCT # HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER NAME HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER ADDRESS HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER ADDRESS HERE"
   PRINT #RptHandle, TAB(2); "--- LABEL BOTTOM ---"
   PRINT #RptHandle, TAB(2); "--- LABEL TOP ---"
   PRINT #RptHandle, TAB(2); "CUSTOMER ACCT # HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER NAME HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER ADDRESS HERE"
   PRINT #RptHandle, TAB(2); "CUSTOMER ADDRESS HERE"
   PRINT #RptHandle, TAB(2); "--- LABEL BOTTOM ---"
   RETURN

PrintMLRptEnding:
RETURN

SelectMLOutput:
  LibName$ = "TAX"
  ScrnName$ = "CUSTRPT1"

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

  REDIM Choice$(0 TO 2, 0 TO 2)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Summary"
  Choice$(2, 1) = "Detail"
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
   

  Form$(2, 0) = "N/A"
  Fld(2).Protected = True

  Action = 1
  ShowCursor
  ClearBack
  DisplayCaScrn ScrnName$
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%

  'printhelp help$
  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key

      IF LEFT$(Form$(1, 0), 4) = "Name" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF

      Order$ = LEFT$(Form$(1, 0), 1)

      IF LEFT$(Form$(2, 0), 1) = "S" THEN
        DetailFlag = False
      ELSE
        DetailFlag = True
      END IF
      Detail$ = Form$(2, 0)

      DevSpec$ = LEFT$(Form$(3, 0), 1)
      ExitFlag = True
      RETURN
    CASE EscKey
      AbortFlag = True
'Canceled$ = "Y"
      ExitFlag = True 'EXIT DO
'RETURN
    END SELECT
  LOOP UNTIL ExitFlag

RETURN

GetNameIndexML:
  FOR SCnt = 1 TO NumOfCaRecs
    GET CaFile, SCnt, CaCust(1)
    array(SCnt).who = UCASE$(CaCust(1).CustName)
    array(SCnt).RecNum = SCnt
    Count = NumOfCaRecs
  NEXT SCnt

'Sort Them Here
  SortT array(Start), Count, Dir, SSize, MOff, MSize
  RETURN



END SUB

SUB OpenCaCustFile (NumOfCaRecs, CaFile)
  
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #CaFile LEN = LEN(CaCust(1))
  NumOfCaRecs = LOF(CaFile) / LEN(CaCust(1))
  
END SUB

