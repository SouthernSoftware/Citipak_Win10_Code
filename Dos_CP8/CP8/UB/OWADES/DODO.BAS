DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB MasterBalanceListing
  
  SHARED Choice$()
  REDIM Choice$(0 TO 15, 0 TO 3)

  Choice$(0, 3) = "8"
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  TownName$ = UBSetUpRec(1).UTILNAME
  IF INSTR(TownName$, "MOWAS") > 0 THEN
    MowasaFlag = True
  END IF

  RCnt = RCnt + 1
  Choice$(RCnt, 3) = "  ALL Revenues"

  REDIM Fmt$(1 TO 3)
  Fmt$(1) = "####,#.##"
  Fmt$(2) = "#####"
  Fmt$(3) = "######,#.##"

  FOR Cnt = 1 TO 15
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      EXIT FOR
    ELSE
      RCnt = RCnt + 1
      Choice$(RCnt, 3) = QPTrim$(STR$(RCnt - 1)) + ") " + TempRev$
    END IF
  NEXT
  ERASE UBSetUpRec

  Choice$(0, 0) = "3"
  Choice$(1, 0) = "Past Due Balance"
  Choice$(2, 0) = "Current Balance"
  Choice$(3, 0) = "Total Balance"
  
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "  ALL"
  Choice$(2, 1) = "Active"
  Choice$(3, 1) = "Inactive"
  Choice$(4, 1) = "Balance"
  
  Choice$(0, 2) = "6"
  Choice$(1, 2) = "Customer Name Order"
  Choice$(2, 2) = "Account Number Order"
  Choice$(3, 2) = "Location Number Order"
  
  LibName$ = "UB"
  ScrnName$ = "UBMBLIST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  Fld(1).LowRange = 0
  Fld(2).LowRange = 0

  Action = 1
  FirstTime = True
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      LSET Form$(1, 0) = "0"
      LSET Form$(2, 0) = "99"
      LSET Form$(4, 0) = "0"
      LSET Form$(7, 0) = "Y"
      IF MowasaFlag = False THEN
        Fld(8).Protected = True
        QPrintRC SPACE$(16), 18, 19, -1
      ELSE
        LSET Form$(8, 0) = Choice$(1, 3)
      END IF
      FirstTime = False
      Action = 1
    END IF
    
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB CheckBalInfo
      
    CASE EscKey
      ExitFlag = True
      EXIT SUB
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  IF ExitFlag THEN GOTO ExitBalRpt
  
  'Main Body Start
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  
  TownName$ = QPTrim$(UBSetup(1).UTILNAME)
  TabStop = 40 - (LEN(TownName$) / 2)
  IF TabStop < 1 THEN TabStop = 1
  TownName$ = SPACE$(TabStop) + TownName$
  
  REDIM RevenueName(1 TO 15) AS STRING * 10
  FOR RCnt = 1 TO 15
    TRevName$ = QPTrim$(UBSetup(1).Revenues(RCnt).REVNAME)
    IF LEN(TRevName$) > 0 THEN
      AndPos = INSTR(TRevName$, "&")
      IF AndPos THEN
        MID$(TRevName$, AndPos) = " "
      END IF
      RevenueName(RCnt) = TRevName$
    ELSE
      MaxRevenue = RCnt - 1
      EXIT FOR
    END IF
  NEXT
  REDIM RevTotals(1 TO MaxRevenue) AS DOUBLE
  
  IF UsingName OR UsingBook THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  MaxLines = 55
  Dash80$ = STRING$(80, "-")
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBALIST.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Master Balance Listing."
  
  GOSUB DoCustRptHeader
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      AcctNo& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNo& = Cnt
    END IF
    
    GET UBCust, AcctNo&, UBCustRec(1)
    'IF AcctNo& = 34 THEN STOP

    RealBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

'110598 Old code could skew totals, A credit in one revenue and a debit
'       in another revenue, that totaled to zero wouldn't show!!!

    'IF RealBalance# = 0 THEN
    '  GOTO BSkipEM
    'END IF

'110598 Code to fix the above bug
    'If all there rev sources are "0" then skipem
    OKToSkip = True
    FOR RevChk = 1 TO 15
      IF Round#(UBCustRec(1).CurrRevAmts(RevChk)) <> 0 THEN
        OKToSkip = False
        EXIT FOR
      END IF
    NEXT
    IF OKToSkip THEN
      GOTO BSkipEM
    END IF
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoCustRptHeader
    END IF
    
    IF UseStatus THEN           'if they care about the cust status, or want all.
      CStatus$ = LEFT$(QPTrim$(UBCustRec(1).Status), 1)
      IF CStatus$ <> RStatus$ THEN
        GOTO BSkipEM
      END IF
    END IF
    
    Book$ = QPTrim$(UBCustRec(1).Book)
    SEQNUMB$ = QPTrim$(UBCustRec(1).SEQNUMB)
    IF LEN(Book$) = 0 THEN
      Book$ = "  "
    END IF
    Book = Value#(Book$, ECode)
    IF Book < BegRoute OR Book > EndRoute THEN
      GOTO BSkipEM
    END IF
    
    SELECT CASE BalType$
    CASE "P"
      ChkBalance# = UBCustRec(1).PrevBalance
    CASE "C"
      ChkBalance# = UBCustRec(1).CurrBalance
    CASE "T"
      ChkBalance# = RealBalance#
    END SELECT
    
    'if there (balance to check) is more than or equal to minimum, or
    'RealBalance# < 0 then printem
    IF (ChkBalance# >= MinBal#) OR (RealBalance# < 0) THEN
      IF RevSource > 0 THEN
        IF UBCustRec(1).CurrRevAmts(RevSource) <> 0 THEN
          PRINT #UBRpt, USING Fmt$(2); AcctNo&;
          PRINT #UBRpt, TAB(8); Book$; "-"; SEQNUMB$;
          PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CUSTNAME, 31);
          PRINT #UBRpt, TAB(60); USING Fmt$(1); Round#(UBCustRec(1).CurrRevAmts(RevSource))
          LineCnt = LineCnt + 1
          TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrRevAmts(RevSource))
          CustCnt = CustCnt + 1
        END IF
      ELSE
        PRINT #UBRpt, USING Fmt$(2); AcctNo&;
        PRINT #UBRpt, TAB(8); Book$; "-"; SEQNUMB$;
        PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CUSTNAME, 31);
        PRINT #UBRpt, TAB(50); USING Fmt$(1); Round#(UBCustRec(1).CurrBalance);
        PRINT #UBRpt, TAB(61); USING Fmt$(1); Round#(UBCustRec(1).PrevBalance);
        PRINT #UBRpt, TAB(72); USING Fmt$(1); RealBalance#
        LineCnt = LineCnt + 1
        TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrBalance)
        TPrevBalance# = Round#(TPrevBalance# + UBCustRec(1).PrevBalance)
        CustCnt = CustCnt + 1
      END IF
      
      IF Detail THEN
        GOSUB PrintDetail
      END IF
      
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    
BSkipEM:
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoCustRptFooter
  
  CLOSE UBCust, UBRpt
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Balance Listing Report.", "UBBALIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'Main Body Exit
ExitBalRpt:
  
  EXIT SUB
  
DoCustRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, TAB(26); "Customer Balance Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, "Acct #"; TAB(9); "Location"; TAB(18); "Customer Name";
  IF RevSource > 0 THEN
    PRINT #UBRpt, TAB(63); "Amount"
  ELSE
    PRINT #UBRpt, TAB(52); "Cur Bal"; TAB(62); "Past Due"; TAB(73); "Acct Bal"
  END IF
  PRINT #UBRpt, Dash80$
  LineCnt = 5
RETURN
  
DoCustRptFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt;
  IF RevSource = 0 THEN
    PRINT #UBRpt, TAB(49); USING Fmt$(3); TCurrBalance#;
    PRINT #UBRpt, TAB(60); USING Fmt$(3); TPrevBalance#;
    PRINT #UBRpt, TAB(71); USING Fmt$(3); (TCurrBalance# + TPrevBalance#)
    IF Detail THEN
      FOR Cnt = 1 TO MaxRevenue
        Detail$ = SPACE$(28)
        LSET Detail$ = RevenueName(Cnt)
        MID$(Detail$, 15) = Fmt$(3)
        PRINT #UBRpt, USING Detail$; RevTotals(Cnt)
      NEXT
    ELSE
      PRINT #UBRpt,
    END IF
  ELSE
    PRINT #UBRpt, TAB(58); USING Fmt$(3); TCurrBalance#
  END IF

  PRINT #UBRpt,
  PRINT #UBRpt, "Report Parameters"
  PRINT #UBRpt, "                   Routes:"; STR$(BegRoute); " to"; STR$(EndRoute)
  PRINT #UBRpt, USING "          Minimum Balance:$$#####.##"; MinBal#;
  PRINT #UBRpt, "               Customer Status:"; Stat$
  PRINT #UBRpt, "           Printing Order:"; POrder$;
  PRINT #UBRpt, "              Balance Type:"; Bal$;
  IF RevSource > 0 THEN
    PRINT #UBRpt, TAB(40); "  Revenue Source: "; Choice$(RevSource + 1, 3);
  END IF
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  LineCnt = LineCnt + 4
  
  PRINT #UBRpt, FF$
RETURN
  
PrintDetail:
  TCnt = 0
  Detail$ = SPACE$(18)
  First = 1
  Last = MaxRevenue
  IF MaxRevenue < Last THEN
    Last = MaxRevenue
  END IF
  FOR RCnt = First TO Last
    TCnt = TCnt + 1
    TabStop = (TCnt * 21) - 20
    IF TabStop > 81 THEN
      LineCnt = LineCnt + 1
      TCnt = 1
      TabStop = (TCnt * 21) - 20
    END IF
    LSET Detail$ = RevenueName(RCnt)
    MID$(Detail$, 10) = "#####.##"
    RevTotals(RCnt) = Round#(RevTotals(RCnt) + UBCustRec(1).CurrRevAmts(RCnt))
    PRINT #UBRpt, TAB(TabStop); USING Detail$; UBCustRec(1).CurrRevAmts(RCnt);
    
  NEXT
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  LineCnt = LineCnt + 2
RETURN
  
  
CheckBalInfo:
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  BegRoute = QPValI(Form$(1, 0))
  EndRoute = QPValI(Form$(2, 0))
  BalType$ = LEFT$(Form$(3, 0), 1)
  Order$ = LEFT$(QPTrim$(Form$(6, 0)), 1)
  MinBal# = Value#(Form$(4, 0), ECode)
  RStatus$ = LEFT$(Form$(5, 0), 1)
  Detail = LEFT$(Form$(7, 0), 1) = "Y" 'Evaluates to TRUE if the field = "Y"
  Rev$ = QPTrim$(LEFT$(Form$(8, 0), 1))
  RevSource = Value#(Rev$, ECode)
  IF RevSource > 0 THEN
    Detail = False
  END IF
  
  IF BegRoute > EndRoute THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Route Order", 10, 30, -1
    QPrintRC "Please Correct and Try Again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
    GOTO ParmErrorRet
  END IF
  SELECT CASE BalType$
  CASE "P"
    Bal$ = " PAST DUE"
  CASE "C"
    Bal$ = " CURRENT"
  CASE "T"
    Bal$ = " TOTAL BALANCE"
  CASE ELSE
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Balance Type", 10, 29, -1
    QPrintRC "Please Correct and Try Again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 3
    Action = 2
    GOTO ParmErrorRet
  END SELECT
  
  SELECT CASE Order$
  CASE "C"
    IndexName$ = NameIndexFile
    UsingName = True
    POrder$ = " CUSTOMER NAME"
  CASE "A"
    POrder$ = " ACCOUNT NUMBER"
    IndexName$ = ""
    UsingAcct = True
  CASE "L"
    POrder$ = " LOCATION NUMBER"
    IndexName$ = BookIndexFile
    UsingBook = True
  CASE ELSE
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Printing Order", 10, 28, -1
    QPrintRC "Please Correct and Try Again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 6
    Action = 2
    GOTO ParmErrorRet
  END SELECT
  SELECT CASE RStatus$
  CASE "A"
    UseStatus = True
    Stat$ = " ACTIVE"
  CASE "I"
    UseStatus = True
    Stat$ = " INACTIVE"
  CASE "B"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  CASE ELSE
    Stat$ = " ALL"
    UseStatus = False
  END SELECT
  
  OKFlag = True
  
ParmErrorRet:
  
RETURN

END SUB

