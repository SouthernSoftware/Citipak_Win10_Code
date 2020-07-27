DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SendDist2GL ()
DECLARE SUB ShowCursor ()

'$INCLUDE: 'DefCnf.BI'

DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB WaitForAction ()
DECLARE SUB CursorOff ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPStrL$ (Number&)
DECLARE FUNCTION QPStrI$ (Number%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB QPrintRC (x$, Row, Col, Colr)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION ubGetAcctTitle$ (RecordNumber)

  CONST False = 0, True = NOT False

  TYPE TranRecInfoType
    TranDate  AS INTEGER
    TranRecNo AS LONG
  END TYPE

  '$INCLUDE: 'GL.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'

  ShowCursor
  SendDist2GL

  IF INSTR(COMMAND$, "TEST") THEN
    'ClearScrn
    HideCursor
  ELSE
    RUN "ubmenu"
  END IF

SUB SendDist2GL
  
  Today$ = DATE$
  Ref$ = "UB" + LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)

  REDIM TranInfo(1) AS TranRecInfoType

  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)

  CursorOff
  BlockClear

  REDIM GJRec(1 TO 2) AS TrEditRecType
  GJRecLen = LEN(GJRec(1))
  GJFile = FREEFILE
  OPEN "GLGJED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  CLOSE
  
  IF NumEdTrans > 0 THEN
    Ok = MsgBox%("UBSENDGL", "JEDNZERO")
    GOTO SendExit
  END IF
  
  Ok = MsgBox%("UBSENDGL", "MUSTEXT2")
  SELECT CASE Ok
  CASE 2
    GOTO SendExit
  END SELECT
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  AcctMeth$ = QPTrim$(UBSetUpRec(1).MethAcct)
  IF (LEN(AcctMeth$) = 0) THEN
    BlockClear
    Ok = MsgBox%("UBSENDGL", "NOMETHOD")
    GOTO SendExit
  END IF

  SELECT CASE AcctMeth$
  CASE "C"
    InterfaceMethod = 1
  CASE "A"
    InterfaceMethod = 2
  CASE ELSE
    GOTO SendExit
  END SELECT

  RPTFile = FREEFILE
  OPEN "UBNOTFND.RPT" FOR OUTPUT AS RPTFile
  GOSUB NotFoundHeader

  ShowProcessingScrn "Verifying GL Transfer Accounts"
  
  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    ELSE
      REDIM PRESERVE GJInfo(1 TO RevCnt) AS GJXferRecType
      GJInfo(RevCnt).RevText = TempRev$
      GJInfo(RevCnt).BAcctInfo.DAcctNo = UBSetUpRec(1).BillAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).BAcctInfo.CAcctNo = UBSetUpRec(1).BillAcct(RevCnt).CreditAcct
      GJInfo(RevCnt).PAcctInfo.DAcctNo = UBSetUpRec(1).PayAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).PAcctInfo.CAcctNo = UBSetUpRec(1).PayAcct(RevCnt).CreditAcct
      IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
        GJInfo(RevCnt).DAcctInfo.DAcctNo = UBSetUpRec(1).DepAcct(RevCnt).DebitAcct
        GJInfo(RevCnt).DAcctInfo.CAcctNo = UBSetUpRec(1).DepAcct(RevCnt).CreditAcct
      END IF
    END IF
  NEXT

  'check to see if they are valid GL accounts
  GOSUB ValidateGLAccounts

  IF BadAcct THEN
    GOTO SendExit
  END IF
  
  DisplayUBScrn "PRORPT"
  QPrintRC "Searching Utility Transactions.", 9, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  
  REDIM UBTransRec(1)  AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTransRecLen

  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #UBTran, TCnt&, UBTransRec(1)
    IF LEN(QPTrim$(UBTransRec(1).Posted2GL)) = 0 THEN
      'Store trans rec numbers and dates in array
      FoundCnt = FoundCnt + 1
      REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
      TranInfo(FoundCnt).TranDate = UBTransRec(1).TransDate
      TranInfo(FoundCnt).TranRecNo = TCnt&
    ELSE
      NGCnt = NGCnt + 1
    END IF
    RSET P2S$ = QPStrI$(FoundCnt)
    QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    IF NGCnt >= 5 THEN EXIT FOR
  NEXT

  IF FoundCnt = 0 THEN
    CLOSE
    BlockClear
    Ok = MsgBox%("UBSENDGL", "NONEFOND")
    GOTO SendExit
  END IF

  SortT TranInfo(1), FoundCnt, 0, 6, 0, -1  'sort'em by date. oldest first
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize

  BlockClear
  ShowProcessingScrn "Transferring Utility Transactions"

  GJFile = FREEFILE
  OPEN "GLGJED.DAT" FOR RANDOM SHARED AS GJFile LEN = GJRecLen

  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  FOR Cnt = 1 TO FoundCnt
    IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      LastTran = Cnt - 1
      GOSUB ProcessThisBunch
      FirstTran = Cnt
    END IF
  NEXT
  LastTran = FoundCnt
  GOSUB ProcessThisBunch

  'transactions as interfaced
  BlockClear
  ShowProcessingScrn "Updating Utility Transactions"
  FOR Cnt = 1 TO FoundCnt
    GET #UBTran, TranInfo(Cnt).TranRecNo, UBTransRec(1)
    UBTransRec(1).Posted2GL = "Y"
    'PUT #UBTran, TranInfo(Cnt).TranRecNo, UBTransRec(1)
    'SmallPause
    ShowPctComp Cnt, FoundCnt
  NEXT
  CLOSE
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
SendExit:

EXIT SUB

NotFoundHeader:
  PageNo = PageNo + 1
  PRINT #RPTFile, "Utility Billing GL Transfer Invalid Account Listing."; TAB(70); "Page:"; PageNo
  PRINT #RPTFile, TownName$
  PRINT #RPTFile, "Report Date: "; DATE$
  PRINT #RPTFile, "Revenue           Acct. Type              Debit Acct.          Credit Acct."
  PRINT #RPTFile, Dash80$
  NumPrinted = 0
RETURN

PrintBadAcct:
    IF LEN(QPTrim$(BadCAcct$)) = 0 THEN
      BadCAcct$ = "Undefined"
    END IF

    PRINT #RPTFile, GJInfo(RevCnt).RevText;

    SELECT CASE ActT
    CASE 1
      ActPg$ = "Billing"
    CASE 2
      ActPg$ = "Payment"
    CASE 3
      ActPg$ = "Deposit"
    END SELECT
    PRINT #RPTFile, TAB(22); ActPg$;
    PRINT #RPTFile, TAB(43); BadDAcct$; TAB(64); BadCAcct$
RETURN

ProcessThisBunch:
  FOR RevCnt = 1 TO NumOfRevs
    GJInfo(RevCnt).BAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).BAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).PAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).PAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).DAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).DAcctInfo.DebitAmt = 0
  NEXT

  FOR PCnt = FirstTran TO LastTran
    IF PCnt = FirstTran THEN
      WorkDate = TranInfo(PCnt).TranDate
    END IF
    GET #UBTran, TranInfo(PCnt).TranRecNo, UBTransRec(1)

    SELECT CASE InterfaceMethod
    CASE 1  'Cash Central
      SELECT CASE UBTransRec(1).TransType
        CASE TranUtilityBill          ' 1=Utility bill
          'no action
        CASE TranLateCharge           ' 2=late charge
          'no action
        CASE TranReconnectFee         ' 3=reconnect fee
          'no action
        CASE TranBillPayment          ' 4=Bill Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).PAcctInfo.DebitAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
          NEXT
        CASE TranAppliedDeposit       ' 5=Applied Deposit
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).PAcctInfo.DebitAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
          NEXT
        CASE TranPenaltyCharge        ' 6=Penalty Charge
          'no action
        CASE TranDepositPayment       ' 7=Deposit Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).DAcctInfo.DebitAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
          NEXT
        CASE TranDraftPayment         ' 8=Draft Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).PAcctInfo.DebitAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
          NEXT
        CASE TranRefundDeposit        ' 9=Refund Deposit
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).DAcctInfo.DebitAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
          NEXT
        CASE TranBeginBalance         '10=Beginning Balance
          'no action
        CASE TranUpwardAdjustment     '11=Upward Adjustments
          'no action
        CASE TranDownwardAdjustment   '12=Downward Adjustments
          'no action
      END SELECT

    CASE 2      'Accural
      SELECT CASE UBTransRec(1).TransType
        CASE TranUtilityBill          ' 1=Utility bill
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
          NEXT
        CASE TranLateCharge           ' 2=late charge
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
          NEXT
        CASE TranReconnectFee         ' 3=reconnect fee
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
          NEXT
        CASE TranBillPayment          ' 4=Bill Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).PAcctInfo.DebitAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
          NEXT
        CASE TranAppliedDeposit       ' 5=Applied Deposit
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).PAcctInfo.DebitAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
          NEXT
        CASE TranPenaltyCharge        ' 6=Penalty Charge
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
          NEXT
        CASE TranDepositPayment       ' 7=Deposit Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).DAcctInfo.DebitAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
          NEXT
        CASE TranDraftPayment         ' 8=Draft Payment
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).PAcctInfo.DebitAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
          NEXT
        CASE TranRefundDeposit        ' 9=Refund Deposit
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).DAcctInfo.DebitAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
          NEXT
        CASE TranBeginBalance         '10=Beginning Balance
          'no action
        CASE TranUpwardAdjustment     '11=Upward Adjustments
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).BAcctInfo.DebitAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
          NEXT
        CASE TranDownwardAdjustment   '12=Downward Adjustments
          FOR RevCnt = 1 TO NumOfRevs
            GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
            GJInfo(RevCnt).PAcctInfo.DebitAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
          NEXT
       END SELECT
    END SELECT
    'SmallPause
    ShowPctComp PCnt, FoundCnt
  NEXT


'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).BAcctInfo.CreditAmt <> 0 THEN
      GJRec(1).AcctRec = GJInfo(RevCnt).BAcctInfo.CRecNo
      GJRec(1).AcctNum = GJInfo(RevCnt).BAcctInfo.CAcctNo
      GJRec(1).AcctName = GJInfo(RevCnt).BAcctInfo.CTitle
      GJRec(1).TrDate = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).CrAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
      GJRec(1).EType = "C"
      GJRec(1).Desc = "FROM UTILITIES"
      GJRec(1).Src = "UB"
      PUT #GJFile, , GJRec(1)
    END IF
    IF GJInfo(RevCnt).BAcctInfo.DebitAmt <> 0 THEN
      GJRec(2).AcctRec = GJInfo(RevCnt).BAcctInfo.DRecNo
      GJRec(2).AcctNum = GJInfo(RevCnt).BAcctInfo.DAcctNo
      GJRec(2).AcctName = GJInfo(RevCnt).BAcctInfo.DTitle
      GJRec(2).TrDate = WorkDate
      GJRec(2).Ref = Ref$
      GJRec(2).DrAmt = GJInfo(RevCnt).BAcctInfo.DebitAmt
      GJRec(2).EType = "D"
      GJRec(2).Desc = "FROM UTILITIES"
      GJRec(2).Src = "UB"
      PUT #GJFile, , GJRec(2)
    END IF
  NEXT

  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).PAcctInfo.CreditAmt <> 0 THEN
      GJRec(1).AcctRec = GJInfo(RevCnt).PAcctInfo.CRecNo
      GJRec(1).AcctNum = GJInfo(RevCnt).PAcctInfo.CAcctNo
      GJRec(1).AcctName = GJInfo(RevCnt).PAcctInfo.CTitle
      GJRec(1).TrDate = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).CrAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
      GJRec(1).EType = "C"
      GJRec(1).Desc = "FROM UTILITIES"
      GJRec(1).Src = "UB"
      PUT #GJFile, , GJRec(1)
    END IF
    IF GJInfo(RevCnt).PAcctInfo.DebitAmt <> 0 THEN
      GJRec(2).AcctRec = GJInfo(RevCnt).PAcctInfo.DRecNo
      GJRec(2).AcctNum = GJInfo(RevCnt).PAcctInfo.DAcctNo
      GJRec(2).AcctName = GJInfo(RevCnt).PAcctInfo.DTitle
      GJRec(2).TrDate = WorkDate
      GJRec(2).Ref = Ref$
      GJRec(2).DrAmt = GJInfo(RevCnt).PAcctInfo.DebitAmt
      GJRec(2).EType = "D"
      GJRec(2).Desc = "FROM UTILITIES"
      GJRec(2).Src = "UB"
      PUT #GJFile, , GJRec(2)
    END IF
  NEXT

  FOR RevCnt = 1 TO NumOfRevs
    REDIM GJRec(1 TO 2) AS TrEditRecType
    IF GJInfo(RevCnt).DAcctInfo.CreditAmt <> 0 THEN
      GJRec(1).AcctRec = GJInfo(RevCnt).DAcctInfo.CRecNo
      GJRec(1).AcctNum = GJInfo(RevCnt).DAcctInfo.CAcctNo
      GJRec(1).AcctName = GJInfo(RevCnt).DAcctInfo.CTitle
      GJRec(1).TrDate = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).CrAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
      GJRec(1).EType = "C"
      GJRec(1).Desc = "FROM UTILITIES"
      GJRec(1).Src = "UB"
      PUT #GJFile, , GJRec(1)
    END IF
    IF GJInfo(RevCnt).DAcctInfo.DebitAmt <> 0 THEN
      GJRec(2).AcctRec = GJInfo(RevCnt).DAcctInfo.DRecNo
      GJRec(2).AcctNum = GJInfo(RevCnt).DAcctInfo.DAcctNo
      GJRec(2).AcctName = GJInfo(RevCnt).DAcctInfo.DTitle
      GJRec(2).TrDate = WorkDate
      GJRec(2).Ref = Ref$
      GJRec(2).DrAmt = GJInfo(RevCnt).DAcctInfo.DebitAmt
      GJRec(2).EType = "D"
      GJRec(2).Desc = "FROM UTILITIES"
      GJRec(2).Src = "UB"
      PUT #GJFile, , GJRec(2)
    END IF
  NEXT

BunchReturn:
RETURN

ValidateGLAccounts:
  BadAcct = False
  FOR RevCnt = 1 TO NumOfRevs
    'Billing Accounts
    IF InterfaceMethod = 2 THEN
'NOTE: We Only check billing accounts if Accural method
      ActT = 1
      ThisAcct = ubFindAcct(GJInfo(RevCnt).BAcctInfo.DAcctNo)
      IF ThisAcct <= 0 THEN
        BadDAcct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).BAcctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.DTitle = ubGetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      END IF
      ThisAcct = ubFindAcct(GJInfo(RevCnt).BAcctInfo.CAcctNo)
      IF ThisAcct <= 0 THEN
        BadCAcct$ = GJInfo(RevCnt).BAcctInfo.CAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).BAcctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.CTitle = ubGetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      END IF
      GOSUB PrintBadAcct
    END IF

    'Payment Accounts
    ActT = 2
    ThisAcct = ubFindAcct(GJInfo(RevCnt).PAcctInfo.DAcctNo)
    IF ThisAcct <= 0 THEN
      BadDAcct$ = GJInfo(RevCnt).PAcctInfo.DAcctNo
      BadAcct = True
    ELSE
      GJInfo(RevCnt).PAcctInfo.DRecNo = ThisAcct
      GJInfo(RevCnt).PAcctInfo.DTitle = ubGetAcctTitle$(ThisAcct)
      BadDAcct$ = "     OK"
    END IF
    ThisAcct = ubFindAcct(GJInfo(RevCnt).PAcctInfo.CAcctNo)
    IF ThisAcct <= 0 THEN
      BadCAcct$ = GJInfo(RevCnt).PAcctInfo.CAcctNo
      BadAcct = True
    ELSE
      GJInfo(RevCnt).PAcctInfo.CRecNo = ThisAcct
      GJInfo(RevCnt).PAcctInfo.CTitle = ubGetAcctTitle$(ThisAcct)
      BadCAcct$ = "     OK"
    END IF
    GOSUB PrintBadAcct

    'Deposit Accounts
    ActT = 3
    IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
      ThisAcct = ubFindAcct(GJInfo(RevCnt).DAcctInfo.DAcctNo)
      IF ThisAcct <= 0 THEN
        BadDAcct$ = GJInfo(RevCnt).DAcctInfo.DAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).DAcctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).DAcctInfo.DTitle = ubGetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      END IF
    ELSE
      BadDAcct$ = "    N/A"
    END IF
    IF UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" THEN
      ThisAcct = ubFindAcct(GJInfo(RevCnt).DAcctInfo.CAcctNo)
      IF ThisAcct <= 0 THEN
        BadCAcct$ = GJInfo(RevCnt).DAcctInfo.CAcctNo
        BadAcct = True
      ELSE
        GJInfo(RevCnt).DAcctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).DAcctInfo.CTitle = ubGetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      END IF
    ELSE
      BadCAcct$ = "    N/A"
    END IF
    GOSUB PrintBadAcct
    ShowPctComp RevCnt, NumOfRevs
  NEXT
  CLOSE RPTFile

  IF BadAcct THEN
    BlockClear
    Ok = MsgBox%("UBSENDGL", "UBADACTS")
    BlockClear
    EntryPoint = 0
    PrintRptFile "GL Transfer Invalid Account List.", "UBNOTFND.RPT", 1, RetCode, EntryPoint
  END IF
  KILL "UBNOTFND.RPT"
RETURN

END SUB

SUB SmallPause
St! = TIMER
St! = St! + .1
DO WHILE St! > TIMER
LOOP
END SUB

