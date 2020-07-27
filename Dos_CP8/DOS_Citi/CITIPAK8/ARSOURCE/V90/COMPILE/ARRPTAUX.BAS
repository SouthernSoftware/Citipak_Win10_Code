DEFINT A-Z
DECLARE SUB ClearBack ()
DECLARE SUB CustomerFeesListing ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SmallPause ()
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, CNumFlag%)
DECLARE SUB QuickCustomer ()
DECLARE SUB AppListing ()
DECLARE SUB CustomerBalance ()
DECLARE SUB ExpiredLicenseListing ()
DECLARE SUB LicenseListing ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB ARFixMess ()
DECLARE SUB ShowNoCode ()
DECLARE SUB OpenARCustFile (NumOfArRecs, ARFile)
DECLARE SUB TransactionJournal ()
DECLARE SUB CustomerListing ()
DECLARE SUB CatListing ()
DECLARE SUB CustomerInquiry (AccountRec&)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, Page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION PDQTimer& ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType

  CONST False = 0, True = NOT False

SUB AppListing

  SHARED Choice$()

  ReportFile$ = "ARAPPLST.PRN"  'Report File Name

  FF$ = CHR$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  GOSUB SelectAppListOutput
  IF Canceled$ = "Y" THEN EXIT SUB
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintAppListHeader
  
  ' Print Main Body
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
    NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  END IF
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      GET ARFile, ARIdxRec(1).IDXRECORD, ARCustRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      GET ARFile, Cnt!, ARCustRec(1)
      CustomerNumber = Cnt!
    END IF
    IF ARCustRec(1).Deleted <> "Y" THEN
      TCat$ = RTRIM$(Catagory$)
      IF TCat$ = RTRIM$(ARCustRec(1).BILLCAT1) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT2) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT3) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT4) OR TCat$ = RTRIM$(ARCustRec(1).BILLCAT5) OR TCat$ = "ALL" THEN
        IF LineCnt >= MaxLines THEN
          PRINT #RptHandle, FF$
          GOSUB PrintAppListHeader
        END IF
        PRINT #RptHandle, "Cust #"; CustomerNumber; TAB(15); "Bus. Name: "; ARCustRec(1).CUSTNAME
        PRINT #RptHandle, "         Applicant Name: "; ARCustRec(1).BILLNAME
        PRINT #RptHandle, "                Address: ";
        PRINT #RptHandle, TAB(26); RTRIM$(ARCustRec(1).ADDRESS1)
        PRINT #RptHandle, TAB(26); RTRIM$(ARCustRec(1).ADDRESS2)
        PRINT #RptHandle, TAB(26); RTRIM$(ARCustRec(1).CITY); ", "; ARCustRec(1).STATE; " "; ARCustRec(1).ZIPCODE
        PRINT #RptHandle, STRING$(80, "-")
        CustCnt = CustCnt + 1
        LineCnt = LineCnt + 6
      END IF
    END IF
  NEXT Cnt!
  
  GOSUB PrintAppListEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 5
  END IF
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
  
PrintAppListHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(20); "Business License Application Customer Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle,
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 3
RETURN
  
PrintAppListEnding:
  PRINT #RptHandle, "Total Customers Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
RETURN
  
  
  
SelectAppListOutput:
  LibName$ = "AR"
  ScrnName$ = "ARQCUSRP"
  
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
  
  
  'REDIM Choice$(350, 3)         ' assume maximum of 99 codes
  
  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  
  IF NumOFARCatRecs = 0 THEN
    CLOSE ARCatFile
    ShowNoCode
    EXIT SUB
  END IF
  
  NumOFARCatRecs = NumOFARCatRecs + 1
  REDIM Choice$(NumOFARCatRecs, 2)         ' assume maximum of 99 codes

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name"
  Choice$(2, 0) = "Number"
  
  
  Choice$(0, 1) = "2"
  FOR Cnt! = 1 TO (NumOFARCatRecs - 1):
    GET ARCatFile, Cnt!, ARCatCodeRec(1)
    Choice$(Cnt!, 1) = ARCatCodeRec(1).CATCODE
  NEXT Cnt!

  Choice$(NumOFARCatRecs, 1) = "ALL"
  CLOSE ARCatFile
  
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "SCREEN"
  Choice$(2, 2) = "PRINTER"
  
  Action = 1
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  Action = 1
  QPrintRC "Application Listing ", 7, 24, 14
  QPrintRC "]", 7, 44, 10


  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = Form$(1, 0)
      Catagory$ = Form$(2, 0)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      RETURN
    CASE EscKey
      Canceled$ = "Y"
      RETURN
    END SELECT
  LOOP
  
END SUB

