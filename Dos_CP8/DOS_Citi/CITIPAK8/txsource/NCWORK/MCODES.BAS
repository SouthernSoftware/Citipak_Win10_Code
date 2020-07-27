DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)

DEFINT A-Z
SUB MortCodeReport
  SHARED Choice$()
  DIM MCode$(250), MDesc$(250)

  REDIM MortCodeRec(1) AS MortCodeRecType
  MortCodeRecLen = LEN(MortCodeRec(1))
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxMORT.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  MaxLines = 56
  LineCnt = 0
  CustCnt = 0

  LibName$ = "TAX"
  ScrnName$ = "LATERPT"
  

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
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  Choice$(0, 2) = "3"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  Form$(2, 0) = "Detail"
  Fld(2).Protected = -1
  Action = 1
  ClearBack

  ShowCursor

  DisplayTaxScrn ScrnName$
  LOCATE 8, 24: COLOR 14: PRINT "Mortgage Code Report"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      DetailFlag$ = LEFT$(Form$(2, 0), 1)
      DevSpec$ = LEFT$(Form$(3, 0), 1)
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag

  IF AbortFlag THEN EXIT SUB

  RptHandle = FREEFILE

  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  GOSUB PrintMCHeader

  GOSUB MCodes
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndexMC
  END IF

  ClearBack
  ShowProcessingScrn "Mortgage Code Report"


  FOR MCodes = 1 TO NumOfCatRecs
  LOCATE 24, 1: PRINT STRING$(79, " ");
  LOCATE 24, 1: COLOR 11
  'QPrintRC
  PRINT "Processing Mortgage Code ";
  COLOR 15
  PRINT MCode$(MCodes); "  - "; RTRIM$(MDesc$(MCodes));
  
  PRINT #RptHandle, "Mortgage Code "; MCode$(MCodes); "  - "; MDesc$(MCodes)
  LineCnt = LineCnt + 1
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF

    GET TaxFile, CustRecNo, TaxCustRec(1)

    IF NOT TaxCustRec(1).Deleted THEN

      ' Check Line on Page and Form Feed if Necessary
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintMCHeader
        PRINT #RptHandle, "Mortgage Code "; MCode$(MCodes); "  - "; MDesc$(MCodes)
      END IF

      nme$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      nme$ = QPTrim$(nme$)      'this one cleans up those with only last name

      'First Show Property Records

      IF TaxCustRec(1).FirstPropRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPropRec
        WHILE PropertyRecord! <> 0
          GET #PropTaxFile, PropertyRecord!, PropertyRec(1)
          IF QPTrim$(PropertyRec(1).MortCode) = MCode$(MCodes) THEN
            PRINT #RptHandle, "Acct # "; USING "######"; CustRecNo;
            PRINT #RptHandle, TAB(20); "PIN # "; PropertyRec(1).RealPin
            PRINT #RptHandle, nme$; TAB(60); " MC: "; PropertyRec(1).MortCode
            PRINT #RptHandle, TaxCustRec(1).Addr1; TAB(60); "VAL:"; USING "$$#######,#"; PropertyRec(1).PROPVALU
            PRINT #RptHandle, TaxCustRec(1).Addr2
            PRINT #RptHandle, RTRIM$(TaxCustRec(1).City); " ";
            PRINT #RptHandle, RTRIM$(TaxCustRec(1).State); " ";
            PRINT #RptHandle, RTRIM$(TaxCustRec(1).Zip)
            PRINT #RptHandle, STRING$(79, "-")
            
            LineCnt = LineCnt + 6
            CustCnt = CustCnt + 1
          END IF
          PropertyRecord! = PropertyRec(1).NextRec
        WEND
      END IF
    
    END IF
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT Cnt
  NEXT MCodes

  GOSUB PrintMCEnding

  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF

  ERASE Array, Frm, Form$, Fld, TaxCustRec

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB

PrintMCHeader:
  Page = Page + 1
  PRINT #RptHandle, TAB(24); "Tax Customer Mortgage Code Report"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; Page
  PRINT #RptHandle, Dash80$
  LineCnt = 5
  RETURN

PrintMCEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "        Total Listings Printed: "; USING "#####"; CustCnt
  PRINT #RptHandle, FF$
  RETURN

GetNameIndexMC:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT

  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN


  
MCodes:
  MortFile = FREEFILE
  OPEN "TAXMORT.DAT" FOR RANDOM AS MortFile LEN = MortCodeRecLen
  NumOfCatRecs = LOF(MortFile) \ MortCodeRecLen
  IF NumOfCatRecs > 0 THEN
    GOSUB SortMortCodes
    REDIM MChoice$(1 TO NumOfCatRecs)
    FOR Cnt = 1 TO NumOfCatRecs
      GET MortFile, Array(Cnt).RecNum, MortCodeRec(1)
      MCode$(Cnt) = RTRIM$(MortCodeRec(1).MortCode)
      MDesc$(Cnt) = RTRIM$(MortCodeRec(1).BName)
  
    NEXT Cnt
  END IF


SortMortCodes:
  REDIM Array(1 TO NumOfCatRecs) AS Struct
  FOR Cnt = 1 TO NumOfCatRecs
    GET MortFile, Cnt, MortCodeRec(1)
    Array(Cnt).who = MortCodeRec(1).MortCode + STRING$(11, " ")
    Array(Cnt).RecNum = Cnt
  NEXT Cnt
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN

END SUB

