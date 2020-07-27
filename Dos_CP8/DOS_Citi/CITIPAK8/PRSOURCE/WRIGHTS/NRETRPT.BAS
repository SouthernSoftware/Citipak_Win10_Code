DEFINT A-Z

SUB nRetirementRpt
  
  SHARED PrnDef$()
  
  FirstTime = True
  REDIM TempScrn(1)
  REDIM Frm(1) AS FormInfo
  FormName$ = "ERETIRPT"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  'MaxLines = 60
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
      IF LowDate <= HighDate THEN
        MonthNum = QPValI(LEFT$(Form$(2, 0), 2))
        YearNum = QPValI(RIGHT$(Form$(2, 0), 4))
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
  
  CursorOff
  
  RptName$ = "PRRPTS\RETIRE.RPT"
  
  REDIM TransHRec(1)   AS TransRecType
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM Pg(1)          AS STRING * 3
  REDIM Dash(1)        AS STRING * 91
  
  Dash(1) = STRING$(91, "-")
  
  EmpRecSize = LEN(Emp2Rec(1))
  TRecSize = LEN(TransHRec(1))
  
  EmpFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM AS EmpFile LEN = EmpRecSize
  NumOfRec = LOF(EmpFile) / EmpRecSize
  REDIM EmpLBuff(1 TO NumOfRec) AS EmpSortType
  
  FOR Cnt = 1 TO NumOfRec
    GET #EmpFile, Cnt, Emp2Rec(1)
    EmpRet$ = UCASE$(LEFT$(LTRIM$(Emp2Rec(1).EMPRETTP), 1))
    SELECT CASE EmpRet$
    CASE "L"
      LCnt = LCnt + 1
      REDIM PRESERVE EmpLBuff(1 TO LCnt) AS EmpSortType
      EmpLBuff(LCnt).EmpNO = Emp2Rec(1).EmpNO
      EmpLBuff(LCnt).RecNo = Cnt
    CASE "G"
      GCnt = GCnt + 1
      REDIM PRESERVE EmpGBuff(1 TO GCnt) AS EmpSortType
      EmpGBuff(GCnt).EmpNO = Emp2Rec(1).EmpNO
      EmpGBuff(GCnt).RecNo = Cnt
    END SELECT
  NEXT
  CLOSE EmpFile
  
  SortT EmpLBuff(1), LCnt, 0, 12, 0, 10
  SortT EmpGBuff(1), GCnt, 0, 12, 0, 10
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  MaxLines = 50
  
  '-=-=-=-=-=-=-=-=-=-=-=
  
  RptTitle$ = "Employee Retirement Report."
  ShowProcessingScrn RptTitle$
  
  TCol = 40 - (LEN(RptTitle$) \ 2) + 1
  PctRow = 11
  
  QPrintRC "Pass 1 of 2", PctRow, TCol, 112
  
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 11, RHandle
  
  FClose RHandle
  
  OPEN RptName$ FOR APPEND AS #1
  
  FOpenS TransHistFileName, THandle
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  GOSUB LRetRptHeader
  
  FOR RecNo = 1 TO LCnt
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    MatchCnt = 0
    FGetRTA DHandle, Emp2Rec(1), CLNG(EmpLBuff(RecNo).RecNo), EmpRecSize
    
    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipNotL
    END IF
    
    TransRecNum& = CLNG(Emp2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        RetWage# = Round(RetWage# + TransHRec(1).RetGrossPay)   ' was .GrossPay
        RETAMT# = Round(RETAMT# + TransHRec(1).RetireAmt)
        'MatAmt# = Round(MatAmt# + TransHRec(1).MatchRetAmt)
        EMatch# = Round#(TransHRec(1).RetGrossPay * .0447)
        MatAmt# = Round(MatAmt# + EMatch#)
        MatchCnt = MatchCnt + 1
        UsingThisOne = True
      END SELECT
      
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          GOSUB PrintEmpRetLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB LRetRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
      
    LOOP
    
SkipNotL:
    ShowPctComp RecNo, LCnt
  NEXT
  
  GOSUB RetLTotals
  
  'start of PASS 2
  
  QPrintRC "Pass 2 of 2", PctRow, TCol, 112
  
  GTotal# = 0
  RTotal# = 0
  MTotal# = 0
  GOSUB GRetRptHeader
  
  FOR RecNo = 1 TO GCnt
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    FGetRTA DHandle, Emp2Rec(1), CLNG(EmpGBuff(RecNo).RecNo), EmpRecSize
    
    IF Emp2Rec(1).LastTransRec <= 0 THEN
      GOTO SkipNotG
    END IF
    
    TransRecNum& = CLNG(Emp2Rec(1).LastTransRec)
    
    DO
      FGetRTA THandle, TransHRec(1), TransRecNum&, TRecSize
      SELECT CASE TransHRec(1).CheckDate
      CASE LowDate TO HighDate
        IF UCASE$(LEFT$(LTRIM$(Emp2Rec(1).EMPRETTP), 1)) = "G" THEN
          RetWage# = Round(RetWage# + TransHRec(1).RetGrossPay) 'was GrossPay
          RETAMT# = Round(RETAMT# + TransHRec(1).RetireAmt)
          'MatAmt# = Round(MatAmt# + TransHRec(1).MatchRetAmt)
          EMatch# = Round#(TransHRec(1).RetGrossPay * .0494)
          MatAmt# = Round(MatAmt# + EMatch#)
          UsingThisOne = True
        END IF
      CASE ELSE
      END SELECT
      
      IF TransHRec(1).PrevTransRec <= 0 THEN
        IF UsingThisOne THEN
          GOSUB PrintEmpRetLine
          IF LineCnt >= MaxLines THEN
            PRINT #1, FF$
            GOSUB GRetRptHeader
          END IF
        END IF
        EXIT DO
      ELSE
        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
      END IF
    LOOP
    
    
SkipNotG:
    ShowPctComp RecNo, GCnt
  NEXT
  
  GOSUB RetLTotals
  PRINT #1, PrnDef$(1);
  
  FClose THandle
  FClose DHandle
  
  CLOSE
  
  ERASE TransHRec, Emp2Rec, Unit, Pg, Dash
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
  
PrintEmpRetLine:
  PRINT #1, LEFT$(Emp2Rec(1).EMPSSN, 3); "-"; MID$(Emp2Rec(1).EMPSSN, 4, 2); "-"; MID$(Emp2Rec(1).EMPSSN, 6, 4);
  PRINT #1, TAB(14); LEFT$(Emp2Rec(1).EMPRETNO, 12);
  PRINT #1, TAB(28); QPTrim$(Emp2Rec(1).EMPLNAME); ", "; QPTrim$(Emp2Rec(1).EMPFNAME);
  PRINT #1, TAB(55); USING "#####,.##"; RetWage#; TAB(70); RETAMT#;
  PRINT #1, TAB(83); USING "#####,.##"; MatAmt#
  LineCnt = LineCnt + 1         'employeesprinted = employeesprinted + 1
  GTotal# = Round(GTotal# + RetWage#)
  RTotal# = Round(RTotal# + RETAMT#)
  MTotal# = Round(MTotal# + MatAmt#)
  RETURN
  
RetLTotals:
  PRINT #1, Dash(1)
  PRINT #1, TAB(28); "Totals:";
  PRINT #1, TAB(53); USING "#######,.##"; GTotal#; TAB(68); RTotal#;
  PRINT #1, TAB(81); USING "#######,.##"; MTotal#
  PRINT #1, FF$
  RETURN
  
LRetRptHeader:
  LPage = LPage + 1
  PRINT #1, "Department of State Treasurer"; TAB(70); "     Page:"; LPage
  PRINT #1, "North Carolina Law Enforcement Officers Benefit and Retirement Fund"
  PRINT #1, "Monthly Report of Subject Wages And Retirement Contributions"
  PRINT #1,
  PRINT #1, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); TAB(70); "Unit Code: "; Unit(1).UFRETIDL
  PRINT #1, "Month: "; MonthName$(MonthNum); " "; YearNum
  PRINT #1,
  PRINT #1, "                                                   Wages Subject    Retirement     Employer"
  PRINT #1, "Soc Sec #    Ret #         Employee Name           To Retirem't     Deductions        Match"
  PRINT #1, Dash(1)
  LineCnt = 10
  RETURN
  
GRetRptHeader:
  GPage = GPage + 1
  PRINT #1, "Department of State Treasurer"; TAB(70); "     Page:"; GPage
  PRINT #1, "North Carolina Local Government Retirement System."
  PRINT #1, "Monthly Report of Subject Wages And Retirement Contributions"
  PRINT #1,
  PRINT #1, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); TAB(70); "Unit Code: "; Unit(1).UFRETID
  PRINT #1, "Month: "; MonthName$(MonthNum); " "; YearNum
  PRINT #1,
  PRINT #1, "                                                   Wages Subject    Retirement     Employer"
  PRINT #1, "Soc Sec #    Ret #         Employee Name           To Retirem't     Deductions        Match"
  PRINT #1, Dash(1)
  LineCnt = 10
  RETURN
  
  
  
END SUB

