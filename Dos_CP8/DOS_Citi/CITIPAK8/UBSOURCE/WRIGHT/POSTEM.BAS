DECLARE SUB nPostBillTrans ()
DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION ErrorScrn (WhatError%, Acct&)
DECLARE FUNCTION Exist% (filename$)
DECLARE FUNCTION FileSize& (filename$)
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION GetBillBook% ()
DECLARE FUNCTION GetBillCycle% ()
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (filename$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE SUB GetPreBillOrder (Choice%, ExitFlag%, SeqFlag$)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (filename$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#, ThisCycle%, ThisBook%)
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB PostBillTrans ()
DECLARE SUB PreBillReport ()
DECLARE SUB PrintHighLowReport ()
DECLARE SUB PrintRptFile (RptTitle$, filename$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RateCodeErrScrn (RATECODE$)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBLog (Text$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()

  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'

  STACK 8000

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

  IF INSTR(UBSetUpRec(1).DEFCITY, "MCCOR") > 0 THEN
    MccorFlag = True
  END IF

  '033198 Added Conway special handling

  DIM SHARED ConwayFlag AS INTEGER

      nPostBillTrans
      HideCursor
      ClearScrn
      END


SUB nPostBillTrans
  UBLog "IN: Bill Posting."

  IF NOT Exist(UBBillsFile) THEN
    UBLog "ERROR: UBBILLS.DAT Calculation file NOT FOUND!"
    CursorOff
    BlockClear
    DisplayUBScrn "NON2POST"
    WaitForAction
    GOTO ExitBillPost:
  END IF

  IF NOT Exist("UBBILLS.PRN") THEN
    UBLog "ERROR: UBBILLS.PRN Print File NOT FOUND!"
    CursorOff
    BlockClear
    DisplayUBScrn "NOTPRNTD"
    WaitForAction
    GOTO ExitBillPost:
  END IF

  LibName$ = "UB"
  ScrnName$ = "PSTBILLS"

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file

  TownName$ = UBSetUpRec(1).UTILNAME

  'Section to check for customer modifications
  'Town of Lilesville Special Discount Situation

  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF

  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0

  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True

  '--Set screen number to one and display screen

  BlockClear

  DisplayUBScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      OKFlag = True
    CASE EscKey
      ExitFlag = True
    END SELECT

  LOOP UNTIL OKFlag OR ExitFlag

  IF ExitFlag THEN
    BlockClear
    CursorOff
    DisplayUBScrn "POSTCAN"
    WaitForAction
    UBLog "ABORTED:  Bill Posting"
    GOTO ExitBillPost
  END IF

  UBLog "START: Posting Transactions."

  REDIM UBBillRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType

  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBBill = FREEFILE
  OPEN UBBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBBillRecLen

  NumOfTranRecs& = LOF(UBTran) \ UBBillRecLen
  NumOfBillRecs = LOF(UBBill) \ UBBillRecLen

  ShowProcessingScrn "Posting Billing Transactions"
  FOR 1295 = 1 TO NumOfBillRecs
    GET UBBill, BillCnt, UBBillRec(1)
    IF (UBBillRec(1).ActiveFlag AND UBBillRec(1).TransAmt > 0) OR (UBBillRec(1).NONProfit = "Y") THEN
      PostedCnt& = PostedCnt& + 1
      NumOfTranRecs& = NumOfTranRecs& + 1       'point to next trans to write
      GET UBCust, BillCnt, UBCustRec(1)
      FOR MRCnt = 1 TO 2
        WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
        IF UBCustRec(1).Monthly(MRCnt).PayAmt > 0 AND WhatService > 0 THEN
          TestAmt# = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + UBCustRec(1).Monthly(MRCnt).PayAmt)
          IF TestAmt# > UBCustRec(1).Monthly(MRCnt).AMTOWED THEN
            HowMuch# = Round#(UBCustRec(1).Monthly(MRCnt).AMTOWED - UBCustRec(1).Monthly(MRCnt).TotAmtPD)
          ELSE
            HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
          END IF
          UBCustRec(1).Monthly(MRCnt).TotAmtPD = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + HowMuch#)
        END IF
      NEXT
      '062597 added removal of nonrecurring flat rates
      FRFlag = False
      FOR FRCnt = 1 TO 4        'Remove non-recurring flat rates
        IF UBCustRec(1).FlatRates(FRCnt).FRFREQ = "N" THEN
          UBCustRec(1).FlatRates(FRCnt).FRDESC = ""
          UBCustRec(1).FlatRates(FRCnt).FRAMT = 0
          UBCustRec(1).FlatRates(FRCnt).FRFREQ = ""
          UBCustRec(1).FlatRates(FRCnt).REVSRC = 0
          UBCustRec(1).FlatRates(FRCnt).NumMin = 0
          FRFlag = True
        END IF
      NEXT
      IF FRFlag THEN
        UBLog "BILL POST: Removed Flat Rate. Acct:" + STR$(BillCnt)
      END IF
      UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBCustRec(1).CurrBalance = UBBillRec(1).TransAmt
      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      FOR RevCnt = 1 TO MaxRevsCnt
        '01-20-98 Added previous revenue array for Autodistribute.
        'Shift-F9 in payment entry
        UBCustRec(1).PrevRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBCustRec(1).PrevRevAmts(RevCnt))
        UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      NEXT
      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
      UBBillRec(1).TransDesc = "Utility Billing"
      UBBillRec(1).TransDate = UBBillRec(1).BillDate
      FOR MtrCnt = 1 TO 7
        CubMtr = False
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead >= 0 THEN
          IF UBCustRec(1).LocMeters(MtrCnt).MTRUnit = "C" THEN
            CubMtr = True
          END IF
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          IF ReadAmt& < 0 THEN  'Meter rolled over or, been misread
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          END IF
          IF CubMtr THEN
            ReadAmt& = ReadAmt& * 7.481
          END IF
          IF ReadAmt& < 1 THEN
            ReadAmt& = 1
          END IF
          IF UBCustRec(1).LocMeters(MtrCnt).AvgUse < 1 THEN
            UBCustRec(1).LocMeters(MtrCnt).AvgUse = 1
          END IF
          IF UBCustRec(1).LocMeters(MtrCnt).UseCnt < 1 THEN
            UBCustRec(1).LocMeters(MtrCnt).UseCnt = 1
          END IF
          TUse# = ReadAmt& + (UBCustRec(1).LocMeters(MtrCnt).AvgUse * UBCustRec(1).LocMeters(MtrCnt).UseCnt)
          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
          UBCustRec(1).LocMeters(MtrCnt).AvgUse = TUse# / UBCustRec(1).LocMeters(MtrCnt).UseCnt
          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
        END IF
      NEXT
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      UBCustRec(1).LastTrans = NumOfTranRecs&

      IF IndianFlag THEN
        UBCustRec(1).USERCODE1 = ""
      END IF

      PUT UBCust, BillCnt, UBCustRec(1)
      PUT UBTran, NumOfTranRecs&, UBBillRec(1)
    END IF
    ShowPctComp BillCnt, NumOfBillRecs
  NEXT
  CLOSE
  UBLog "  DONE: Posting Transactions."
  UBLog "POSTED:" + STR$(PostedCnt&) + " New BILL Transactions."

  KillFile UBBillsFile
  KillFile "UBBILLS.PRN"

  UBLog "KILLED: UBBILLS.DAT & UBBILLS.PRN"
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitBillPost:
  UBLog "OUT: Bill Posting." + CrLf$

END SUB

