DEFINT A-Z
SUB CustomerCutOffListing
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 0 TO 1)
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Thirty Days or Older"
  Choice$(2, 0) = "Sixty Days or Older"
  Choice$(3, 0) = "Ninety Days or Older"
  
  Choice$(0, 1) = "3"
  Choice$(1, 1) = " Customer Name Order   "
  Choice$(2, 1) = " Account Number Order  "
  Choice$(3, 1) = " Location Number Order "
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  
  AbortFlag = False
  PageNo = 0

  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetUpRecLen

  TownName$ = UBSetUp(1).UTILNAME

  IF INSTR(TownName$, "SPENCER") THEN
    SpenFlag = True
  END IF

  IF INSTR(TownName$, "JOHNSONVILLE") THEN
    JVILLEFlag = True
  END IF

  IF INSTR(TownName$, "ANSONVILLE") THEN
    AnsonFlag = True
  END IF

  LibName$ = "UB"
  ScrnName$ = "UBCOLIST"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Set Defaults as Anticpated Response
  Action = 1
  
  FirstTime = True
  
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF SpenFlag OR AnsonFlag OR JVILLEFlag AND FirstTime THEN
      QPrintRC SPACE$(19), 11, 18, -1
      Fld(2).Protected = True
    END IF

    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = "0"
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      MinAmt# = Value#(Form$(1, 0), ECode)
      Days$ = Form$(2, 0)
      Order$ = LEFT$(QPTrim$(Form$(3, 0)), 1)
      SELECT CASE Order$
      CASE "C"
        IndexName$ = NameIndexFile
        UsingName = True
        OKFlag = True
      CASE "A"
        IndexName$ = ""
        UsingAcct = True
        OKFlag = True
      CASE "L"
        IndexName$ = BookIndexFile
        UsingBook = True
        OKFlag = True
      CASE ELSE
        GOSUB ShowErrScrn
      END SELECT
    CASE ESC
      GOTO ExitCutOffListing
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL OKFlag
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUp(1))
  
  AgeDate = Date2Num%(DATE$)
  
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
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  UBRpt = FREEFILE
  OPEN "UBCOLIST.RPT" FOR OUTPUT AS UBRpt
  
  UBSetUp = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetUp LEN = UBSetUpRecLen
  IF LOF(UBSetUp) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetUp, 1, UBSetUp(1)
    TownName$ = UBSetUp(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetUp
  
  
  BlockClear
  ShowProcessingScrn "Customer Cut-Off Report."
  
  GOSUB DoCutOffRptHeader
  
  LongView = INSTR(TownName$, "LONG VIEW")
  Landis = INSTR(TownName$, "LANDIS")
  Tucka = INSTR(TownName$, "TUCK")
  Ellen = INSTR(TownName$, "ELLEN")

  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      AcctNo& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNo& = Cnt
    END IF

    GET UBCust, AcctNo&, UBCustRec(1)

    IF SpenFlag THEN
      AcctBalance# = UBCustRec(1).PrevBalance
      TotBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
      IF AcctBalance# > 0 THEN
        IF UBCustRec(1).CutOFFYN = "Y" THEN
          IF UBCustRec(1).Status = "A" THEN
            IF AcctBalance# >= MinAmt# THEN
              GOSUB printline
            END IF
          END IF
        END IF
      END IF
      GOTO skip2here
    END IF


'special mod for ellenboro
    IF Ellen OR JVILLEFlag OR AnsonFlag THEN
      AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
      TotBalance# = AcctBalance#
      IF AcctBalance# > 0 THEN
        IF UBCustRec(1).CutOFFYN = "Y" THEN
          IF UBCustRec(1).Status = "A" THEN
            IF AcctBalance# >= MinAmt# THEN
              GOSUB printline
            END IF
          END IF
        END IF
      END IF
      GOTO skip2here
    END IF
'end of ellenboro mod

    IF UBCustRec(1).Status = "A" THEN
      IF UBCustRec(1).CutOFFYN = "Y" THEN
        ' Calc Balance
'        IF LongView > 0 GOTO Skip1
        GOSUB CutOffCalcBal

  'not sure about this
        IF Landis THEN
          AcctBalance# = UBCustRec(1).PrevBalance
        ELSE
          AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
        END IF

        TotBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance

        'Check if Town of LongView if so remove taps rev's 5 & 6
      
        IF LongView > 0 THEN
          'AcctBalance# = UBCustRec(1).PrevBalance
          AcctBalance# = AcctBalance# - (UBCustRec(1).CurrRevAmts(5) + UBCustRec(1).CurrRevAmts(6))
        END IF
        Days$ = LEFT$(Days$, 1)
      
        IF Days$ = "T" AND (Amt2# + Amt3# + Amt4#) > MinAmt# AND AcctBalance# > 0 THEN
          GOSUB printline
        END IF
        IF Days$ = "S" AND (Amt3# + Amt4#) > MinAmt# AND AcctBalance# > 0 THEN
          GOSUB printline
        END IF
        IF Days$ = "N" AND Amt4# > MinAmt# AND AcctBalance# > 0 THEN
          GOSUB printline
        END IF

        IF LongView > 0 THEN
          IF AcctBalance# > MinAmt# THEN
            GOSUB printline
          END IF
        END IF
skip2here:
        IF AskAbandonPrint% THEN
          AbortFlag = True
          EXIT FOR
        END IF
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  GOSUB DoCutOffRptFooter:
  
  CLOSE
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Customer Cut-Off Report.", "UBCOLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBCOLIST.RPT"
  
ExitCutOffListing:
  
  EXIT SUB
  
DoCutOffRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(26); "Customer Cut Off Listing Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$;
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Location  ActNo. Customer Name"; TAB(43); "Service Address"; TAB(74); "Balance"
  IF NOT LongView > 0 THEN
    PRINT #UBRpt, "Meter Number"; TAB(17); "Last Read"; TAB(30); "Reading at Cut-Off"
  END IF
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DoCutOffRptFooter:
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Total Customers to Cut Off: "; USING "#####,#"; CustomerCnt
  PRINT #UBRpt, "     Total Cut Off Balance: "; USING "#####.##"; TBalance#
  RETURN
  
ShowErrScrn:
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  CursorOff
  WazzWind 13, 27, 19, 63, 78, 1, True
  QPrintRC "Invalid Printing Order!", 15, 34, 79
  QPrintRC "Press Any Key To Continue.", 17, 33, 79
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 2
  Action = 1
  RETURN
  
CutOffCalcBal:
  'Age1#=1-30 day  Age2#=31-60 day  Age3#=61-90 day   Age4#=Over 90
  Age1# = 0
  Age2# = 0
  Age3# = 0
  Age4# = 0
  Amt1# = 0
  Amt2# = 0
  Amt3# = 0
  Amt4# = 0
  
  TotalPayments# = 0
  Trans& = UBCustRec(1).LastTrans
  
  DO WHILE Trans& > 0
    GET UBTrans, Trans&, UBTransRec(1)
    IF UBTransRec(1).TransDate <= AgeDate THEN
      IF AgeDate - UBTransRec(1).TransDate > 120 THEN
        IF UBTransRec(1).RunBalance = 0 THEN EXIT DO
      END IF
      GOSUB CutOffLineCalc
    END IF
    Trans& = UBTransRec(1).PrevTrans
  LOOP
  
  GOSUB CutOffFinishCalc
  RETURN
  
CutOffFinishCalc:
  IF TotalPayments# > Amt4# THEN
    TotalPayments# = TotalPayments# - Amt4#
    Amt4# = 0
  ELSE
    Amt4# = Amt4# - TotalPayments#
    TotalPayments# = 0
  END IF
  IF TotalPayments# > Amt3# THEN
    TotalPayments# = TotalPayments# - Amt3#
    Amt3# = 0
  ELSE
    Amt3# = Amt3# - TotalPayments#
    TotalPayments# = 0
  END IF
  IF TotalPayments# > Amt2# THEN
    TotalPayments# = TotalPayments# - Amt2#
    Amt2# = 0
  ELSE
    Amt2# = Amt2# - TotalPayments#
    TotalPayments# = 0
  END IF
  Amt1# = Amt1# - TotalPayments#
  LineTotal# = Amt1# + Amt2# + Amt3# + Amt4#
  RETURN
  
  
CutOffLineCalc:
  
  IF UBTransRec(1).TransType = 107 OR UBTransRec(1).TransType = 7 THEN
    RETURN
  END IF
  IF UBTransRec(1).TransType = 4 OR UBTransRec(1).TransType = 5 OR UBTransRec(1).TransType = 8 OR UBTransRec(1).TransType = 12 OR UBTransRec(1).TransType = 104 OR UBTransRec(1).TransType = 105 OR UBTransRec(1).TransType = 109 THEN
    TotalPayments# = TotalPayments# + UBTransRec(1).TransAmt
    TotalPayments# = Round#(TotalPayments#)
    RETURN
  END IF
  
  SELECT CASE UBTransRec(1).TransType
  CASE 1, 2, 3, 6, 11, 10, 101, 102, 103, 112
    Days& = AgeDate - UBTransRec(1).TransDate
    IF Days& <= 30 THEN
      Amt1# = Amt1# + UBTransRec(1).TransAmt
      Amt1# = Round#(Amt1#)
      'RETURN
    END IF
    IF Days& > 30 AND Days& <= 60 THEN
      Amt2# = Amt2# + UBTransRec(1).TransAmt
      Amt2# = Round#(Amt2#)
      'RETURN
    END IF
    IF Days& > 60 AND Days& <= 90 THEN
      Amt3# = Amt3# + UBTransRec(1).TransAmt
      Amt3# = Round#(Amt3#)
      'RETURN
    END IF
    IF Days& > 90 THEN
      Amt4# = Amt4# + UBTransRec(1).TransAmt
      Amt4# = Round#(Amt4#)
      'RETURN
    END IF
  END SELECT
  'STOP
  
  RETURN

printline:
  TBalance# = Round#(TBalance# + AcctBalance#)
  IF LineCnt > MaxLines THEN
    PRINT #UBRpt, CHR$(12)
    GOSUB DoCutOffRptHeader
  END IF
  '*************************************
  '   Main body of Printing goes here
  PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; USING "######"; AcctNo&;
  PRINT #UBRpt, TAB(18); LEFT$(UBCustRec(1).CustName, 25); LEFT$(UBCustRec(1).ServAddr, 22);
  PRINT #UBRpt, TAB(70); USING "$$####,#.##"; TotBalance#
  LineCnt = LineCnt + 1
  IF LongView > 0 THEN
    GOTO Skip1
  END IF
  FOR MrtCnt = 1 TO 7
    IF UBCustRec(1).LocMeters(MrtCnt).MtrType <> " " THEN
      PRINT #UBRpt, UBCustRec(1).LocMeters(MrtCnt).MtrNum;
      PRINT #UBRpt, TAB(17); UBCustRec(1).LocMeters(MrtCnt).CurRead;
      PRINT #UBRpt, TAB(30); "_________________"
      LineCnt = LineCnt + 1
    END IF
  NEXT
  PRINT #UBRpt, STRING$(79, "-"): LineCnt = LineCnt + 1
Skip1:
  IF Landis THEN
    PRINT #UBRpt,
    LineCnt = LineCnt + 2
  END IF
  CustomerCnt = CustomerCnt + 1
  '*************************************
  RETURN
  
END SUB

