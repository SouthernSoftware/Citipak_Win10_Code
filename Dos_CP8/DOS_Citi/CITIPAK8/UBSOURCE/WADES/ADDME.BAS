DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB CustomerFlatRates

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 0 TO 1)

  Choice$(0, 1) = "2"
  Choice$(1, 1) = " Customer Name Order   "
  Choice$(2, 1) = " Account Number Order  "
  Choice$(3, 1) = " Location Number Order "

  UsingBook = False
  UsingAcct = False
  UsingName = False

  AbortFlag = False
  PageNo = 0

  LibName$ = "UB"
  ScrnName$ = "UBFLCUST"

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
    
    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = "1"
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      RevSource = Value#(Form$(1, 0), ECode)
      Days$ = Form$(2, 0)
      Order$ = LEFT$(QPTrim$(Form$(2, 0)), 1)
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
        'GOSUB ShowErrScrn
      END SELECT
    CASE ESC
      GOTO ExitCustomerFlat
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL OKFlag
  
  'IF ExitFlag THEN GOTO ExitCustomerFlat

  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  '  REDIM UBLocaRec(1) AS UBLocationRecType
  '  UBLocaRecLen = LEN(UBLocaRec(1))

  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))

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

  '  UBLoca = FREEFILE
  '  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBLoca LEN = UBLocaRecLen

  UBRpt = FREEFILE
  OPEN "UBFRLIST.RPT" FOR OUTPUT AS UBRpt

  UBSetup = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetup LEN = UBSetUpRecLen
  IF LOF(UBSetup) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetup, 1, UBSetup(1)
    TownName$ = UBSetup(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetup


  BlockClear
  ShowProcessingScrn "Flat Rate Listing."

  GOSUB DoMastFlatHeader

  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook THEN
      GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
      acctnumber = IdxBuff(Cnt).RecNum
    ELSE
      GET UBCust, Cnt, UBCustRec(1)
      acctnumber = Cnt
    END IF

    'LocationNumber$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
    'LocationNumber$ = QPTrim$(LocationNumber$)

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoMastFlatHeader
    END IF

    '*************************************
    '   Main body of Printing goes here
    IF UBCustRec(1).DelFlag <> -1 AND UBCustRec(1).Status = "A" THEN
      PrnCnt = 0
      DidOne = False
      FOR FRCnt = 1 TO 4
        IF UBCustRec(1).FlatRates(FRCnt).RevSrc = RevSource THEN
          DidOne = True
          PRINT #UBRpt, USING "Acct:#####"; acctnumber;
          PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).CustName, 30);
          CustFlat# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAmt * UBCustRec(1).FlatRates(FRCnt).NumMin)
          PRINT #UBRpt, TAB(45); QPTrim$(UBCustRec(1).FlatRates(FRCnt).FRDesc); TAB(65); USING "#####.##"; CustFlat#
          TotalFlat# = Round#(TotalFlat# + CustFlat#)
          PrnCnt = PrnCnt + 1
         EXIT FOR
        END IF
      NEXT

      IF DidOne THEN
        PrnCnt = PrnCnt + 1
        PRINT #UBRpt, Dash80$
        LineCnt = LineCnt + PrnCnt
        CustCnt = CustCnt + 1
      END IF
      '*************************************
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT

  GOSUB DoMastFlatFooter:

  CLOSE UBCust, UBRpt

  ERASE IdxBuff, UBCustRec

  IF NOT AbortFlag THEN
    PrintRptFile "Master Customer Report.", "UBFRLIST.RPT", 1, RetCode, EntryPoint
  END IF

  KillFile "UBFRLIST.RPT"

ExitCustomerFlat:

  EXIT SUB

DoMastFlatHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(28); "Customer Flat Rate Listing"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, Dash80$
  LineCnt = 4
  RETURN

DoMastFlatFooter:
  PRINT #UBRpt, "Totals:"; TAB(10); "Customers: "; USING "#####,#"; CustCnt
  PRINT #UBRpt, TAB(10); "   Amount:"; USING "######,#.##"; TotalFlat#
  RETURN



END SUB

