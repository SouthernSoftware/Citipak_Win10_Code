DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB PrintCustByRate

  SHARED Choice$()

  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

'*****************
  LibName$ = "UB"
  ScrnName$ = "CUSBYRAT"

  '--Get the total number of fields from all pages
  NumFlds = NumFlds + LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  '--Increment StartEl to next

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs

  'add show a message
  IF NumOfRateRecs <= 0 THEN
    GOTO ExitCustByRate
  END IF

  REDIM Choice$(0 TO NumOfRateRecs, 0 TO 1)
  Choice$(0, 0) = "1"
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    Choice$(Cnt, 0) = SPACE$(30)
    LSET Choice$(Cnt, 0) = UBRateTblRec(1).RATECODE
    MID$(Choice$(Cnt, 0), 6, 24) = UBRateTblRec(1).RATEDESC
  NEXT
  CLOSE
  'SortStr2 VARPTR(Choice$(1)), Cnt - 1, 0

  RateRec = 0
  PrnRate = False

  Action = 1
  FirstTime = True

  BlockClear

  DisplayUBScrn "CUSBYRAT"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = "N"
      LSET Form$(3, 0) = "Y"
      Action = 2
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      WhatRate$ = QPTrim$(LEFT$(Form$(1, 0), 4))
      IF LEN(WhatRate$) > 0 THEN
        IF QPTrim$(Form$(2, 0)) = "Y" THEN
          PrnRate = True
          FOR Cnt = 1 TO NumOfRateRecs
            ThisRate$ = QPTrim$(LEFT$(Choice$(Cnt, 0), 4))
            IF ThisRate$ = WhatRate$ THEN
              RateRec = Cnt
              EXIT FOR
            END IF
          NEXT
        END IF
        SELECT CASE Form$(3, 0)
        CASE "Y"
          ActOnly = True
        END SELECT
        OkFlag = True

      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT
    
    '--Check screen page
  LOOP UNTIL ExitFlag OR OkFlag

  IF ExitFlag THEN
    GOTO ExitCustByRate
  END IF
  
'****************

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBCSBYRT.RPT" FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn "Customer Listing By Rate Code."

  GOSUB CustByRateHeader

  FOR Cnt = 1 TO IdxNumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF ActOnly THEN
        IF UBCustRec(1).Status <> "A" THEN
          GOTO SkipCustRate
        END IF
      END IF
      FOR SCnt = 1 TO 15
        ThisRate$ = QPTrim$(UBCustRec(1).Serv(SCnt).RATECODE)
        IF WhatRate$ = ThisRate$ THEN
          PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " "; FUsing$(STR$(IdxBuff(Cnt).RecNum), "#####"); "  "; LEFT$(UBCustRec(1).CUSTNAME, 25); "  "; LEFT$(UBCustRec(1).SERVADDR, 30); " "; UBCustRec(1).Status
          CustCnt = CustCnt + 1
          LineCnt = LineCnt + 1
          SELECT CASE UBCustRec(1).Status
          CASE "A"
            Active = Active + 1
          CASE "F"
            Final = Final + 1
          CASE "I"
            InActive = InActive + 1
          CASE "B"
            Balance = Balance + 1
          CASE ELSE
            UnKnown = UnKnown + 1
          END SELECT
          EXIT FOR
        END IF
      NEXT

      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB CustByRateHeader
      END IF
    END IF
SkipCustRate:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT

  GOSUB CustByRateTotals

  ERASE IdxBuff, UBCustRec   'free up memory

  IF PrnRate THEN
    GOSUB PrintRateCode
  END IF

  CLOSE UBCust, UBRpt

  IF NOT AbortFlag THEN
    PrintRptFile "Customer Listing by Rate Code.", "UBCSBYRT.RPT", 1, RetCode, EntryPoint
  END IF

ExitCustByRate:
  EXIT SUB

CustByRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing By Rate Code      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "RATE CODE: "; WhatRate$
  PRINT #UBRpt, "Location   Acct.  Customer Name             Service Address             Status"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN

CustByRateTotals:
  'PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: ####"; Active
  PRINT #UBRpt, USING "   Final: ####"; Final
  PRINT #UBRpt, USING "Inactive: ####"; InActive
  PRINT #UBRpt, USING " Balance: ####"; Balance
  PRINT #UBRpt, USING " Unknown: ####"; UnKnown
  PRINT #UBRpt, USING " Deleted: ####"; DeletedCnt
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: ####"; CustCnt
  PRINT #UBRpt, CHR$(12)
RETURN

PrintRateCode:

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))

  NumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen

  IF NumRateRecs = 0 THEN
    GOTO PrintRateExit
  END IF

  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  GET UBFile, RateRec, UBRateTblRec(1)
  CLOSE UBFile

  REDIM StepText(1 TO 10) AS STRING * 40

  GOSUB PrintRateHeader
    PRINT #UBRpt, "       Rate Code:  "; UBRateTblRec(1).RATECODE
    PRINT #UBRpt, "     Description:  "; UBRateTblRec(1).RATEDESC
    PRINT #UBRpt, "  Minimum Charge:"; FUsing$(STR$(UBRateTblRec(1).MINAMT), "#######.##")
    PRINT #UBRpt, "   Minimum Units:"; FUsing$(STR$(UBRateTblRec(1).MINUNITS), "##########")
    PRINT #UBRpt, "      Max Amount:"; FUsing$(STR$(UBRateTblRec(1).MaxAmt), "######.##")
    PRINT #UBRpt, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
    FOR RCnt = 1 TO 10
      LSET StepText$(RCnt) = ""
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITS >= 0 THEN
        MID$(StepText$(RCnt), 8) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITS), "########")
      END IF
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITAMT >= 0 THEN
        MID$(StepText$(RCnt), 25) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT), "####.######")
      END IF
    NEXT
    PRINT #UBRpt, "     First Break:"; StepText$(1)
    PRINT #UBRpt, "    Second Break:"; StepText$(2)
    PRINT #UBRpt, "     Third Break:"; StepText$(3)
    PRINT #UBRpt, "    Fourth Break:"; StepText$(4)
    PRINT #UBRpt, "     Fifth Break:"; StepText$(5)
    PRINT #UBRpt, "     Sixth Break:"; StepText$(6)
    PRINT #UBRpt, "   Seventh Break:"; StepText$(7)
    PRINT #UBRpt, "    Eighth Break:"; StepText$(8)
    PRINT #UBRpt, "     Ninth Break:"; StepText$(9)
    PRINT #UBRpt, "        All Over:"; StepText$(10)
    PRINT #UBRpt,
    PRINT #UBRpt, Dash80$
    PRINT #UBRpt, CHR$(12)

  ERASE UBRateTblRec, StepText

  GOTO PrintRateExit

PrintRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Rate Table Listing."
  PRINT #UBRpt, "RATE CODE: "; WhatRate$; TAB(70); "Page:"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, Dash80$
  NumPrinted = 0
RETURN

PrintRateExit:

RETURN

END SUB

