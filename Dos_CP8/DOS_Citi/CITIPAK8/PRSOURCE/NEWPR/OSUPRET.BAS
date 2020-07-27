DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB oSupRetReport
  
  SHARED Choice$(), PrnDef$()
  
  REDIM TempScrn(1)
  
  REDIM Choice$(12, 1)          '
  REDIM Text$(1 TO 5)

  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  
  REDIM Unit(1)              AS UnitFileRecType
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  Choice$(0, 1) = "3,4"
  
  FOR Cnt = 1 TO 12
    Choice$(Cnt, 1) = QPTrim$(DedCodes(Cnt).DCDESC1)
  NEXT
  
  
  FirstTime = True
  REDIM Frm(1) AS FormInfo
  
  FormName$ = "SRETIRPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  BlockClear
  DisplayMiscScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      Form$(2, 0) = DATE$
      Form$(1, 0) = LEFT$(DATE$, 3) + "01-" + RIGHT$(DATE$, 4)
      Action = 2
    END IF
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      LowDate = Date2Num(Form$(1, 0))
      HighDate = Date2Num(Form$(2, 0))
      GPct# = VAL(Form$(5, 0))
      LPct# = VAL(Form$(6, 0))
      TempDed$ = QPTrim$(Form$(3, 0))
      FOR Cnt = 1 TO 12
        IF TempDed$ = QPTrim$(DedCodes(Cnt).DCDESC1) THEN
          VCodeNum = Cnt
          EXIT FOR
        END IF
      NEXT
      TempDed$ = QPTrim$(Form$(4, 0))
      FOR Cnt = 1 TO 12
        IF TempDed$ = QPTrim$(DedCodes(Cnt).DCDESC1) THEN
          LCodeNum = Cnt
          EXIT FOR
        END IF
      NEXT
      'IF (HighDate >= LowDate) AND LCodeNum > 0 AND VCodeNum > 0 AND GPct# > 0 THEN
      IF (HighDate >= LowDate) THEN 'AND GPct# > 0 THEN
        OKFlag = True
      ELSE
        CursorOff
        SaveScrn TempScrn()
        DisplayMiscScrn ReportParmError
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 2
        OKFlag = False
      END IF
    CASE EscKey
      OKFlag = True
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag             'proper key not set
  
  CursorOff
  
  ERASE Form$, Fld, TempScrn, Frm
  
  IF ExitFlag THEN EXIT SUB
  
  RptName$ = "PRRPTS\401K.RPT"
  
  VCalcAmt# = 0: GVTotal# = 0
  LCalcAmt# = 0: GLTotal# = 0
  GCalcAmt# = 0:
  
  'OPEN "dodo.lst" FOR OUTPUT AS #10

  FF$ = CHR$(12)
  
  EPRTotal# = 0
  
  Image1$ = ",#####.##"
  Image2$ = "###,###.##"
  
  REDIM TransHRec(1)     AS TransRecType
  REDIM Emp2Rec(1)  AS EmpData2Type
  REDIM K401Rec(1)       AS K401RptType
  
  EmpRecSize = LEN(Emp2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  
  IdxRecLen = 2
  
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  
  RptTitle$ = "Supplemental Retirement Report."
  ShowProcessingScrn RptTitle$
  
  'got input here

  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 12, RHandle
  FClose RHandle
  
  OPEN RptName$ FOR APPEND AS #1
  
  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  GOSUB K401Header
  
  FOR RecNo = 1 TO NumOfRecs
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    GCalcAmt# = 0

    FGetRTA DHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize

    IF Emp2Rec(1).EMPTDATE > 0 THEN
      IF LowDate > Emp2Rec(1).EMPTDATE THEN
        GOTO SkipEm
      END IF
    END IF
    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipEm
    END IF
    
    TransRecNum& = CLNG(Emp2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate

        IF VCodeNum > 0 THEN
          IF TransHRec(1).DAmt(VCodeNum) <> 0 THEN
            VCalcAmt# = Round(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          END IF
        END IF
        IF LCodeNum > 0 THEN
          IF TransHRec(1).DAmt(LCodeNum) > 0 THEN
            LCalcAmt# = Round(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          END IF
        END IF
        EmpRType$ = UCASE$(LEFT$(LTRIM$(Emp2Rec(1).EMPRETTP), 1))
        IF EmpRType$ = "L" OR EmpRType$ = "G" THEN

          GCalcAmt# = Round(GCalcAmt# + TransHRec(1).GrossPay)
          UsingThisOne = True
          'PRINT #10, TransHRec(1).TDist(DCnt).DAcct; "  ";
          'FOR DCnt = 1 TO 8
          '  PRINT #10, TransHRec(1).TDist(DCnt).DAcct; "  ";
          '  PRINT #10, TransHRec(1).TDist(DCnt).DPct; "  ";
          'NEXT
          'PRINT #10,
        END IF
      CASE ELSE
      END SELECT
      
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          IF EmpRType$ = "L" THEN
            EPct# = LPct#
          ELSE
            EPct# = GPct#
          END IF
          GOSUB PrintThisOne
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF

    LOOP
    
SkipEm:
    ShowPctComp RecNo, NumOfRecs
  NEXT

  'CLOSE 10
  GOSUB PrintGrandTots
  FClose DHandle
  FClose THandle
  PRINT #1, PrnDef$(1);
  CLOSE
  
  ERASE TransHRec, Emp2Rec, K401Rec, IdxBuff, DedCodes
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
PrintThisOne:
  'EPrinted = EPrinted + 1
IF EPct# > 0 OR VCalcAmt# > 0 OR LCalcAmt# > 0 THEN
  EPrinted = EPrinted + 1
  REDIM K401Rec(1)       AS K401RptType
  TotalVAmt# = Round(TotalVAmt# + VCalcAmt#)
  TotalLAmt# = Round(TotalLAmt# + LCalcAmt#)
  TMatchAmt# = Round((GCalcAmt# * EPct#) * .01)
  TotalMatchAmt# = Round(TotalMatchAmt# + TMatchAmt#)
  
  K401Rec(1).EmpName = LTRIM$(RTRIM$(Emp2Rec(1).EMPFNAME)) + " " + LTRIM$(RTRIM$(Emp2Rec(1).EMPLNAME))
  K401Rec(1).SSN = LEFT$(Emp2Rec(1).EMPSSN, 3) + "-" + MID$(Emp2Rec(1).EMPSSN, 4, 2) + "-" + MID$(Emp2Rec(1).EMPSSN, 6, 4)
  RSET K401Rec(1).VAmt = LTRIM$(RTRIM$(FUsing(STR$(VCalcAmt#), Image1$)))
  IF VCalcAmt# = 0 THEN
    RSET K401Rec(1).VAmt = ".00"
  END IF

  RSET K401Rec(1).LAmt = LTRIM$(RTRIM$(FUsing(STR$(LCalcAmt#), Image1$)))
  IF LCalcAmt# = 0 THEN
    RSET K401Rec(1).LAmt = ".00"
  END IF

  RSET K401Rec(1).MAmt = LTRIM$(RTRIM$(FUsing(STR$(TMatchAmt#), Image1$)))
  IF TMatchAmt# = 0 THEN
    RSET K401Rec(1).MAmt = ".00"
  END IF


  RSET K401Rec(1).Batch = QPTrim$(Unit(1).BBTBATCH)

  TempDate$ = Num2Date(HighDate)
  K401Rec(1).HDate = LEFT$(TempDate$, 6) + RIGHT$(TempDate$, 2)
  PRINT #1, K401Rec(1).EmpName; K401Rec(1).SSN; "   "; K401Rec(1).VAmt; " "; K401Rec(1).LAmt;
  'PRINT #1, "  "; K401Rec(1).MAmt; K401Rec(1).Batch; "  "; K401Rec(1).HDate
  PRINT #1, "  "; K401Rec(1).MAmt; "    "; K401Rec(1).HDate

END IF
'STOP
RETURN
  
K401Header:
  TempDate$ = DATE$
  TempDate$ = LEFT$(TempDate$, 6) + RIGHT$(TempDate$, 2)
  Text$(1) = "Supplemental Retirement Income Plan of NC"
  Text$(2) = "BB&T Administrator"
  Text$(3) = "Payroll Center Name: " + QPTrim$(Unit(1).UFEMPR)
  Text$(4) = "Payroll Center Number: " + QPTrim$(Unit(1).BBTCNTNO)

  Text1Len = LEN(Text$(1))
  Text2Len = LEN(Text$(2))
  Text3Len = LEN(Text$(3))
  Text4Len = LEN(Text$(4))

  PRINT #1, SPACE$(48 - (Text1Len \ 2)); Text$(1); TAB(81); "Date: "; TempDate$
  PRINT #1, SPACE$(48 - (Text2Len \ 2)); Text$(2); TAB(81); "Page: 1 "
  PRINT #1, SPACE$(48 - (Text3Len \ 2)); Text$(3)
  PRINT #1, SPACE$(48 - (Text4Len \ 2)); Text$(4)
  PRINT #1,
  PRINT #1, "                                                 Pre-Tax    Post-Tax     Employer"
'  PRINT #1, "Employee Name                       SS#          Vol Cont   Loan Pmts    Contrib'n      Batch     Date"
  PRINT #1, "Employee Name                       SS#          Vol Cont   Loan Pmts    Contrib'n      Date"
  PRINT #1, ""
  RETURN
  
PrintGrandTots:
  Totals$ = SPACE$(95)

  RSET K401Rec(1).VAmt = LTRIM$(RTRIM$(FUsing(STR$(TotalVAmt#), Image1$)))
  RSET K401Rec(1).LAmt = LTRIM$(RTRIM$(FUsing(STR$(TotalLAmt#), Image1$)))
  RSET K401Rec(1).MAmt = LTRIM$(RTRIM$(FUsing(STR$(TotalMatchAmt#), Image2$)))
  
  IF TotalVAmt# = 0 THEN
    RSET K401Rec(1).VAmt = ".00"
  END IF
  IF TotalLAmt# = 0 THEN
    RSET K401Rec(1).LAmt = ".00"
  END IF
  IF TotalMatchAmt# = 0 THEN
    RSET K401Rec(1).MAmt = ".00"
  END IF
  

  LSET Totals$ = "Totals"
  MID$(Totals$, 37) = STR$(EPrinted)
  MID$(Totals$, 47) = K401Rec(1).VAmt
  MID$(Totals$, 59) = K401Rec(1).LAmt
  MID$(Totals$, 72) = K401Rec(1).MAmt

  MID$(Totals$, 87) = LTRIM$(RTRIM$(FUsing(STR$(Round#(TotalVAmt# + TotalLAmt# + TotalMatchAmt#)), Image2$)))

  PRINT #1,
  PRINT #1, Totals$
  PRINT #1, FF$
  
  RETURN
  
END SUB

