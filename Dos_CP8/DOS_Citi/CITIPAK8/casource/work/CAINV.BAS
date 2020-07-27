DEFINT A-Z
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
  
  CustomerBalances

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF

  IF Exist("CAmenu" + Ext$) THEN
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

  

  ' Print Main Body

  OpenCaCustFile NumOfCaRecs, CaFile

  IF UsingIndex THEN
    GOSUB GetNameIndex1
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
        GOSUB PrintBalanceCustomerRptHeader
      END IF


   'Determine if Customer HAS A BALANCE & PRINT IF THEY DO!

        CustBalance# = Round#(CaCust(1).RPBStreet + CaCust(1).RPBSide + CaCust(1).RPBDrive + CaCust(1).RPBDemo + CaCust(1).RPBWater + CaCust(1).RPBSewer)
        CustBalance# = Round#(CustBalance# + CaCust(1).RIBStreet + CaCust(1).RIBSide + CaCust(1).RIBDrive + CaCust(1).RIBDemo + CaCust(1).RIBWater + CaCust(1).RIBSewer)

        IF CustBalance# <= 0 THEN GOTO CUSTSKIP

        RunPrinTotal# = 0
        RunIntTotal# = 0

        GOSUB PrintBalanceCustomerRptHeader

        PRINT #RptHandle, TAB(3); LEFT$(RTRIM$(CaCust(1).CustName), 30);
          PRINT #RptHandle, TAB(40); "    ACCT # "; CustRecNo
        PRINT #RptHandle, TAB(3); LEFT$(RTRIM$(CaCust(1).Add1), 30);
          PRINT #RptHandle, TAB(40); "LEGAL DESC "; LEFT$(RTRIM$(CaCust(1).LegalDes), 30)
        PRINT #RptHandle, TAB(3); LEFT$(RTRIM$(CaCust(1).Add2), 30);
          PRINT #RptHandle, TAB(40); "LEGAL ADDR "; LEFT$(RTRIM$(CaCust(1).PropDesc), 30)

        PRINT #RptHandle, ""
        PRINT #RptHandle, ""
        PRINT #RptHandle, ""
        PRINT #RptHandle, TAB(3); "THIS IS AN INVOICE FOR ASSESSMENTS DUE THE TOWN OF CAROLINA BEACH"
        PRINT #RptHandle, TAB(3); "THE STATUS OF THIS ACCOUNT IS SHOWN BELOW."
        PRINT #RptHandle, ""
        PRINT #RptHandle, ""
        PRINT #RptHandle, TAB(11); "- ORIG ASSESSMENT -"; TAB(31); "INTEREST"; TAB(41); "-- PRIOR PAYMENTS --"; TAB(62); "  LAST  "; TAB(70); "PRINCIPLE"
        PRINT #RptHandle, TAB(11); " AMOUNT      DATE  "; TAB(31); "  RATE  "; TAB(41); "PRINCIPLE   INTEREST"; TAB(62); "PAY DATE"; TAB(70); " BALANCE"
 
         
        IF CaCust(1).RPBStreet + CaCust(1).RIBStreet > 0 THEN
         PRINT #RptHandle, "Streets"; TAB(10); USING "#####.##"; CaCust(1).STRAmt;
          PRINT #RptHandle, TAB(22); LEFT$(Num2Date(CaCust(1).STRDate), 6) + RIGHT$(Num2Date(CaCust(1).STRDate), 2);
          PRINT #RptHandle, TAB(32); USING "##.##"; CaCust(1).STRRate;
          PRINT #RptHandle, TAB(41); USING "#####.##"; Round#(CaCust(1).STRAmt - CaCust(1).RPBStreet);
          PRINT #RptHandle, TAB(52); USING "#####.##"; Round#(CaCust(1).IStreet - CaCust(1).RIBStreet);
          IF CaCust(1).LastPayDate <> 0 THEN
           PRINT #RptHandle, TAB(62); LEFT$(Num2Date(CaCust(1).LastPayDate), 6) + RIGHT$(Num2Date(CaCust(1).LastPayDate), 2);
          END IF
          PRINT #RptHandle, TAB(71); USING "#####.##"; CaCust(1).RPBStreet
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBStreet
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBStreet
          LineCnt = LineCnt + 1
        END IF
        IF CaCust(1).RPBSide + CaCust(1).RIBSide > 0 THEN
         PRINT #RptHandle, "Sidewalk"; TAB(10); USING "#####.##"; CaCust(1).SIDEAmt;
          PRINT #RptHandle, TAB(22); LEFT$(Num2Date(CaCust(1).SIDEDate), 6) + RIGHT$(Num2Date(CaCust(1).SIDEDate), 2);
          PRINT #RptHandle, TAB(32); USING "##.##"; CaCust(1).SIDERate;
          PRINT #RptHandle, TAB(41); USING "#####.##"; Round#(CaCust(1).SIDEAmt - CaCust(1).RPBSide);
          PRINT #RptHandle, TAB(52); USING "#####.##"; Round#(CaCust(1).ISide - CaCust(1).RIBSide);
          IF CaCust(1).LastPayDate <> 0 THEN
           PRINT #RptHandle, TAB(62); LEFT$(Num2Date(CaCust(1).LastPayDate), 6) + RIGHT$(Num2Date(CaCust(1).LastPayDate), 2);
          END IF
          PRINT #RptHandle, TAB(71); USING "#####.##"; CaCust(1).RPBSide
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBSide
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBSide
          LineCnt = LineCnt + 1
        END IF
        IF CaCust(1).RPBDrive + CaCust(1).RIBDrive > 0 THEN
         PRINT #RptHandle, "Driveway"; TAB(10); USING "#####.##"; CaCust(1).DriveAmt;
          PRINT #RptHandle, TAB(22); LEFT$(Num2Date(CaCust(1).DriveDate), 6) + RIGHT$(Num2Date(CaCust(1).DriveDate), 2);
          PRINT #RptHandle, TAB(32); USING "##.##"; CaCust(1).DriveRate;
          PRINT #RptHandle, TAB(41); USING "#####.##"; Round#(CaCust(1).DriveAmt - CaCust(1).RPBDrive);
          PRINT #RptHandle, TAB(52); USING "#####.##"; Round#(CaCust(1).IDrive - CaCust(1).RIBDrive);
          IF CaCust(1).LastPayDate <> 0 THEN
           PRINT #RptHandle, TAB(62); LEFT$(Num2Date(CaCust(1).LastPayDate), 6) + RIGHT$(Num2Date(CaCust(1).LastPayDate), 2);
          END IF
          PRINT #RptHandle, TAB(71); USING "#####.##"; CaCust(1).RPBDrive
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBDrive
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBDrive
          LineCnt = LineCnt + 1
        END IF
        IF CaCust(1).RPBDemo + CaCust(1).RIBDemo > 0 THEN
         PRINT #RptHandle, "Demolit'n"; TAB(10); USING "#####.##"; CaCust(1).DemoAmt;
          PRINT #RptHandle, TAB(22); LEFT$(Num2Date(CaCust(1).DemoDate), 6) + RIGHT$(Num2Date(CaCust(1).DemoDate), 2);
          PRINT #RptHandle, TAB(32); USING "##.##"; CaCust(1).DemoRate;
          PRINT #RptHandle, TAB(41); USING "#####.##"; Round#(CaCust(1).DemoAmt - CaCust(1).RPBDemo);
          PRINT #RptHandle, TAB(52); USING "#####.##"; Round#(CaCust(1).IDemo - CaCust(1).RIBDemo);
          IF CaCust(1).LastPayDate <> 0 THEN
           PRINT #RptHandle, TAB(62); LEFT$(Num2Date(CaCust(1).LastPayDate), 6) + RIGHT$(Num2Date(CaCust(1).LastPayDate), 2);
          END IF
          PRINT #RptHandle, TAB(71); USING "#####.##"; CaCust(1).RPBDemo
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBDemo
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBDemo
          LineCnt = LineCnt + 1
        END IF
        IF CaCust(1).RPBWater + CaCust(1).RIBWater > 0 THEN
         PRINT #RptHandle, "Water"; TAB(10); USING "#####.##"; CaCust(1).WaterAmt;
          PRINT #RptHandle, TAB(22); LEFT$(Num2Date(CaCust(1).WaterDate), 6) + RIGHT$(Num2Date(CaCust(1).WaterDate), 2);
          PRINT #RptHandle, TAB(32); USING "##.##"; CaCust(1).WaterRate;
          PRINT #RptHandle, TAB(41); USING "#####.##"; Round#(CaCust(1).WaterAmt - CaCust(1).RPBWater);
          PRINT #RptHandle, TAB(52); USING "#####.##"; Round#(CaCust(1).IWater - CaCust(1).RIBWater);
          IF CaCust(1).LastPayDate <> 0 THEN
           PRINT #RptHandle, TAB(62); LEFT$(Num2Date(CaCust(1).LastPayDate), 6) + RIGHT$(Num2Date(CaCust(1).LastPayDate), 2);
          END IF
          PRINT #RptHandle, TAB(71); USING "#####.##"; CaCust(1).RPBWater
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBWater
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBWater
          LineCnt = LineCnt + 1
        END IF
        IF CaCust(1).RPBSewer + CaCust(1).RIBSewer > 0 THEN
         PRINT #RptHandle, "Sewer"; TAB(10); USING "#####.##"; CaCust(1).SewerAmt;
          PRINT #RptHandle, TAB(22); LEFT$(Num2Date(CaCust(1).SewerDate), 6) + RIGHT$(Num2Date(CaCust(1).SewerDate), 2);
          PRINT #RptHandle, TAB(32); USING "##.##"; CaCust(1).SewerRate;
          PRINT #RptHandle, TAB(41); USING "#####.##"; Round#(CaCust(1).SewerAmt - CaCust(1).RPBSewer);
          PRINT #RptHandle, TAB(52); USING "#####.##"; Round#(CaCust(1).ISewer - CaCust(1).RIBSewer);
          IF CaCust(1).LastPayDate <> 0 THEN
           PRINT #RptHandle, TAB(62); LEFT$(Num2Date(CaCust(1).LastPayDate), 6) + RIGHT$(Num2Date(CaCust(1).LastPayDate), 2);
          END IF
          PRINT #RptHandle, TAB(71); USING "#####.##"; CaCust(1).RPBSewer
          RunPrinTotal# = RunPrinTotal# + CaCust(1).RPBSewer
          RunIntTotal# = RunIntTotal# + CaCust(1).RIBSewer
          LineCnt = LineCnt + 1
        END IF

       PRINT #RptHandle, ""
       PRINT #RptHandle, ""
       PRINT #RptHandle, TAB(5); "Assessment payments may be paid in full or in 5 equal annual"
       PRINT #RptHandle, TAB(5); "installments, based on the confirmation of the assessment roll."
       PRINT #RptHandle, TAB(5); "Interest is accrued monthly (using the simple interest method)"
       PRINT #RptHandle, TAB(5); "on any outstanding principle balance."
       PRINT #RptHandle, ""
       PRINT #RptHandle, ""
       PRINT #RptHandle, "Principle Due ... "; USING "$$#####,#.##"; RunPrinTotal#
       PRINT #RptHandle, "Interest Due .... "; USING "$$#####,#.##"; RunIntTotal#
       PRINT #RptHandle, "                 -------------"
       PRINT #RptHandle, "Total Amount Due  "; USING "$$#####,#.##"; Round#(RunPrinTotal# + RunIntTotal#)
       PRINT #RptHandle, FF$
    END IF
    CustCnt = CustCnt + 1
CUSTSKIP:
    NEXT Cnt
  
  

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
  PRINT #RptHandle, TAB(29); "TOWN OF CAROLINA BEACH"
  PRINT #RptHandle, TAB(26); "FINANCE DEPT - COLLECTION DIV"
  PRINT #RptHandle, TAB(21); "1121 NORTH LAKE PARK BLD - 919-458-2999"
  PRINT #RptHandle, TAB(28); "CAROLINA BEACH, NC 28428"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(60); "DATE: "; DATE$
  PRINT #RptHandle,
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
  PRINT #RptHandle, "Demolition"; TAB(20); USING "######,#.##"; RunPrinDemo#;
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
  LOCATE 8, 33: COLOR 14: PRINT "Invoices";
  'printhelp help$
  Form$(2, 0) = "N/A"
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

SUB OpenCaCustFile (NumOfCaRecs, CaFile)
  
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #CaFile LEN = LEN(CaCust(1))
  NumOfCaRecs = LOF(CaFile) / LEN(CaCust(1))
  
END SUB

