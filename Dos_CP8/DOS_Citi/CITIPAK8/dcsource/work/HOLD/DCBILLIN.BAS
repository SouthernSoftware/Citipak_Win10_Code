DEFINT A-Z
DECLARE SUB DCAPPLICATION ()
DECLARE SUB PrintLicenses ()
DECLARE SUB LicenseRegister ()
DECLARE SUB OpenDCCustFile (NumOfArRecs%, ARFile%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB ARAPPLICATION ()
DECLARE SUB ShowNoCode ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB ShowCursor ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB WaitForAction ()
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, action, Cnf AS Config)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
   
   DIM SHARED DCCustRec(1) AS DCCustRecType
   DIM SHARED DCCustIdxRec(1) AS DCCustIdxRecType

   TYPE Struct
      who AS STRING * 14
      RecNum AS INTEGER
   END TYPE


   
 STACK 8000

   '--Dim the choice array to the number of menu items
          DCAPPLICATION
   RUN "DCmenu"

SUB DCAPPLICATION
  
  SHARED Choice$()



  LibName$ = "DC"
  ScrnName$ = "DCAPPRN"
  help$ = "Print Applications!!!"
  LOCATE 1, 1, 0
  
  REDIM DCCatCodeRec(1) AS DCCatCodeRecType
  DCCatCodeRecLen = LEN(DCCatCodeRec(1))
  DCCatFile = FREEFILE
  OPEN "DCCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCCatFile LEN = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatCodeRecLen
  
  IF NumOFDCCatRecs = 0 THEN
   CLOSE DCCatFile
   ShowNoCode
   EXIT SUB
  END IF
  
  NumOFDCCatRecs = NumOFDCCatRecs + 1


   REDIM Choice$(0 TO NumOFDCCatRecs, 0 TO 1)

   Choice$(0, 0) = "1"
  FOR Cnt! = 1 TO (NumOFDCCatRecs - 1):
  GET DCCatFile, Cnt!, DCCatCodeRec(1)
   
   Choice$(Cnt!, 0) = DCCatCodeRec(1).CATCODE
  NEXT Cnt!
   Choice$(NumOFDCCatRecs, 0) = "ALL"
  CLOSE DCCatFile


   Choice$(0, 1) = "3"
   Choice$(1, 1) = "Standard"
   Choice$(2, 1) = "Custom"
   

  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  printhelp help$
   

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

  action = 1

  REM check for code file
   
  DO

MainBody:
    EditForm Form$(), Fld(), Frm(1), Cnf, action


    
    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      IF ASC(LEFT$(Form$(1, 0), 1)) = 32 THEN
       Done = False
       COLOR 12
       LOCATE 9, 34: PRINT "Catagory Not Set": PRINT CHR$(7);
       COLOR 15
       SLEEP 3: LOCATE 9, 34
       PRINT STRING$(23, 32)
       action = 1: Frm(1).FldNo = 1: GOTO MainBody
      END IF
      IF ASC(LEFT$(Form$(2, 0), 1)) = 32 THEN
       Done = False
       COLOR 12
       LOCATE 10, 34: PRINT "Year Not Set": PRINT CHR$(7);
       SLEEP 3: LOCATE 10, 34
       COLOR 15
       PRINT STRING$(23, 32)
       action = 1
       Frm(1).FldNo = 2
       GOTO MainBody
      END IF
      IF ASC(LEFT$(Form$(3, 0), 1)) = 32 THEN
       Done = False
       COLOR 12
       LOCATE 11, 34
       PRINT "Form Type Not Set": PRINT CHR$(7);
       SLEEP 3: LOCATE 11, 34
       COLOR 15
       PRINT STRING$(23, 32)
       action = 1
       Frm(1).FldNo = 3
       GOTO MainBody
      END IF
       Code$ = LEFT$(Form$(1, 0), 3)
       IF LEFT$(Form$(3, 0), 6) = "Custom" THEN GOSUB PrintCustom
       IF LEFT$(Form$(3, 0), 8) = "Standard" THEN GOSUB PrintStandard
       
    CASE EscKey
      EXIT SUB
    END SELECT

  LOOP


PrintCustom:

  ReportFile$ = "CUSTAPPS.RPT"
  RptHandle = FREEFILE

  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  CustRecLen = LEN(DCCustRec(1))
  TrHandle = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen

  'REDIM DCCustIdxRec(1) AS DCCustIdxRecType     ' open customer file

  IdxCustRecLen = LEN(DCCustIdxRec(1))
  IdxTrHandle = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS IdxTrHandle LEN = IdxCustRecLen
  IdxTrNumRecs = LOF(IdxTrHandle) \ IdxCustRecLen

  FOR Cnt = 1 TO IdxTrNumRecs
   GET IdxTrHandle, Cnt, DCCustIdxRec(1)
    GET TrHandle, DCCustIdxRec(1).IDXRECORD, DCCustRec(1)
IF DCCustRec(1).Deleted <> "Y" THEN
   IF Code$ = "ALL" THEN
     GOSUB PrintCustomForm
     ELSE
     IF LEFT$(DCCustRec(1).BILLCAT, 3) = Code$ THEN
     GOSUB PrintCustomForm
     END IF
   END IF
    
END IF
  NEXT Cnt

    CLOSE                       'Close all open files now
    EntryPoint = 1
    Header$ = "Applications"
    PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
    KILL ReportFile$
    EXIT SUB

PrintCustomForm:
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(31); "TOWN OF DUMFRIES"
   PRINT #RptHandle, TAB(24); "BUSINESS LICENSE APPLICATION"
   PRINT #RptHandle, TAB(31); "For Year: "; Form$(2, 0)
   PRINT #RptHandle, ""
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "Business Name: "; DCCustRec(1).CustName
   PRINT #RptHandle, TAB(5); "              -----------------------------------------------------------"
   PRINT #RptHandle, TAB(5); "Street Address of Business: "
   PRINT #RptHandle, TAB(5); "                           ----------------------------------------------"
   PRINT #RptHandle, TAB(5); "Zoning of Business Location: "
   PRINT #RptHandle, TAB(5); "                            ---------------------------------------------"
   PRINT #RptHandle, TAB(5); "Telephone Number: "
   PRINT #RptHandle, TAB(5); "                 --------------------------------------------------------"
   PRINT #RptHandle, TAB(5); "_________________________________________________________________________"
   PRINT #RptHandle, TAB(5); "Applicant's Name: "; DCCustRec(1).BILLNAME
   PRINT #RptHandle, TAB(5); "                 --------------------------------------------------------"
   PRINT #RptHandle, TAB(5); "Applicant's Address: "; DCCustRec(1).ADDRESS1
   PRINT #RptHandle, TAB(5); "                    -----------------------------------------------------"
   PRINT #RptHandle, TAB(5); "Telephone Number: "; DCCustRec(1).HPHONE
   PRINT #RptHandle, TAB(5); "                 --------------------------------------------------------"
   REM 22 lines printed here
   PRINT #RptHandle, TAB(5); "TYPE OF BUSINESS LICENSE APPLYING FOR:"
   PRINT #RptHandle, TAB(5); ""
   PRINT #RptHandle, TAB(5); "_______ Contracting or Construction $30.00 or 10 cents per $100.00 gross"
   PRINT #RptHandle, TAB(5); "           receipts whichever is greater."
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "_______ Retail Sales $30.00 or 12 cents per $100.00 whichever is greater."
   PRINT #RptHandle, TAB(5); ""
   PRINT #RptHandle, TAB(5); "_______ Financial, Real Estate or Professional Service $30.00 or "
   PRINT #RptHandle, TAB(5); "           40 cents per $100.00 whichever is greater."
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "_______ Repair, Personal or Business Service $30.00 or 20 cents per $100.00"
   PRINT #RptHandle, TAB(5); "            whichever is greater."
   PRINT #RptHandle, TAB(5); ""
   PRINT #RptHandle, TAB(5); "_______ Other (Specify) ______________________________________________"
   PRINT #RptHandle, TAB(5); ""
   PRINT #RptHandle, TAB(5); "Estimate of ______________ gross receipts or preceding year's gross "
   PRINT #RptHandle, TAB(5); "receipts ______________________"
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "AMOUNT OF LICENSE TAX FOR YEAR JAN. 1, THROUGH DEC. 31, 19___ IS:$_______"
   PRINT #RptHandle, TAB(5); "ANY SPECIAL CONDITIONS OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED "
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "ACTIVITY SHALL BE CONDUCTED: ____________________________________________"
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "_________________________________________________________________________"
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "_________________________________________________________________________"
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "I certify that the statements and figures set forth on this application"
   PRINT #RptHandle, TAB(5); "DCe true to the best of my knowledge."
   PRINT #RptHandle, TAB(5); "                                      ___________________________________"
   PRINT #RptHandle, TAB(5); "                                            Signature of Applicant"
   PRINT #RptHandle, TAB(5); ""
   PRINT #RptHandle, TAB(5); "To Avoid Late Penalty Charge of 10%, License Renewals Must Be Made Before "
   PRINT #RptHandle, TAB(5); "February 1."
   PRINT #RptHandle, TAB(5);
   PRINT #RptHandle, TAB(5); "Return Application and Fee to:"
   PRINT #RptHandle, TAB(5); Form$(4, 0)
   PRINT #RptHandle, TAB(5); Form$(5, 0)
   PRINT #RptHandle, TAB(5); Form$(6, 0)
   PRINT #RptHandle, CHR$(12);
   TotalCust = TotalCust + 1
   RETURN

   
RETURN
PrintStandard:
RETURN


END SUB

SUB OpenDCCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile

  DCCustRecLen = LEN(DCCustRec(1))
  ARFile = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = DCCustRecLen
  NumOfArRecs = LOF(ARFile) \ DCCustRecLen

END SUB

SUB ShowNoCode
  LibName$ = "DC"
  ScrnName$ = "DCNOCODE"
  help$ = "NEW A/R Customer Entry"
  LOCATE 1, 1, 0


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


  PRINT CHR$(7);
  ShowCursor
  LibFile2Scrn "DC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  action = 1
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, action

    SELECT CASE Frm(1).KeyCode
    CASE EscKey
     EXIT SUB
    END SELECT

  LOOP

END SUB

