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
DECLARE FUNCTION QPStrL$ (number&)
DECLARE FUNCTION QPStrI$ (number%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB QPrintRC (x$, Row, Col, Colr)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION ubGetAcctTitle$ (RecordNumber)

  CONST False = 0, True = NOT False

  TYPE TranRecInfoType
    TranDate  AS INTEGER
    TranRecNo AS LONG
  END TYPE

  '$INCLUDE: 'GL.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'cmfiles.bi'

  ShowCursor

  REDIM SHARED MCodeRec(1) AS MiscCodeRecType
  REDIM SHARED GLSetup(1) AS GLSetupRecType
  

  SendDist2GL

  IF INSTR(COMMAND$, "TEST") THEN
    'ClearScrn
    HideCursor
  ELSE
    RUN "cmmenu"
  END IF

SUB SendDist2GL
  
  Today$ = DATE$
  Ref$ = "CM" + LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)

  REDIM TranInfo(1) AS TranRecInfoType
  DIM MiscRec#(500), MiscAmt#(500), Fund$(100), FundAmt#(100)



  Dash80$ = STRING$(80, "-")
  P2S$ = SPACE$(4)

  CursorOff
  BlockClear
  REDIM GJRec(1) AS TrEditRecType
  GJRecLen = LEN(GJRec(1))
  GJFile = FREEFILE
  OPEN "GLGJED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJFile LEN = GJRecLen
  NumEdTrans = LOF(GJFile) \ GJRecLen
  CLOSE
  
  IF NumEdTrans > 0 THEN
    Ok = MsgBox%("UBSENDGL", "JEDNZERO")
    GOTO SendExit
  END IF
  
  Ok = MsgBox%("UBSENDGL", "MUSTEXT3")
  SELECT CASE Ok
  CASE 2
    GOTO SendExit
  END SELECT
  

  
  RPTFile = FREEFILE
  OPEN "UBNOTFND.RPT" FOR OUTPUT AS RPTFile
  GOSUB NotFoundHeader


  DisplayUBScrn "PRORPT"
  QPrintRC "Searching Cash Transactions.", 9, 26, 126
  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  
  REDIM CMTransRec(1)  AS CMTransRecType
  CMTransRecLen = LEN(CMTransRec(1))

  CMTran = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMTran LEN = CMTransRecLen
  NumOfTRecs& = LOF(CMTran) \ CMTransRecLen
  LOCK #CMTran

  FOR TCnt& = NumOfTRecs& TO 1 STEP -1
    GET #CMTran, TCnt&, CMTransRec(1)
    IF LEN(QPTrim$(CMTransRec(1).Trans2GL)) = 0 OR QPTrim$(CMTransRec(1).Trans2GL) = "N" THEN
      'Store trans rec numbers and dates in array
      FoundCnt = FoundCnt + 1
      REDIM PRESERVE TranInfo(FoundCnt) AS TranRecInfoType
      TranInfo(FoundCnt).TranDate = CMTransRec(1).TransDate
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
  ShowProcessingScrn "Transferring Misc Cash Transactions"

  GJFile = FREEFILE
  OPEN "GLGJED.DAT" FOR RANDOM SHARED AS GJFile LEN = GJRecLen

  MCFile = FREEFILE
  OPEN "CMMISCCD.DAT" FOR RANDOM SHARED AS MCFile LEN = LEN(MCodeRec(1))

  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  WorkDate = ThisDate

  FOR Cnt = 1 TO FoundCnt
   IF ThisDate <> TranInfo(Cnt).TranDate THEN
      ThisDate = TranInfo(Cnt).TranDate
      GOSUB ProcessThisBunch: DayCount = 0: WorkDate = ThisDate
   END IF

   GET #CMTran, TranInfo(Cnt).TranRecNo, CMTransRec(1)

   IF DayCount = 0 THEN
      
            FOR MCnt = 1 TO 5
            MiscRevAmt# = (CMTransRec(1).TransRevAmt(MCnt))
            MiscRevAmt# = Round#(MiscRevAmt#)
            IF MiscRevAmt# <> 0 THEN
               'If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
              IF CMTransRec(1).TransRevAmt(MCnt + 5) >= 1 THEN
                DayCount = DayCount + 1
                MiscRec#(DayCount) = CMTransRec(1).TransRevAmt(MCnt + 5)
                MiscAmt#(DayCount) = MiscRevAmt#
               END IF
             END IF
           NEXT MCnt
     
    ELSE
      FOR MCnt = 1 TO 5
       MiscRevAmt# = (CMTransRec(1).TransRevAmt(MCnt))
       MiscRevAmt# = Round#(MiscRevAmt#)
      DO WHILE MiscRevAmt# <> 0
            FOR FindCount = 1 TO DayCount
             IF MiscRec#(FindCount) = CMTransRec(1).TransRevAmt(MCnt + 5) THEN
              MiscAmt#(FindCount) = MiscAmt#(FindCount) + MiscRevAmt#
              MiscRevAmt# = 0
              EXIT DO
             END IF
            NEXT FindCount
            DayCount = DayCount + 1
            MiscRec#(DayCount) = CMTransRec(1).TransRevAmt(MCnt + 5)
            MiscAmt#(DayCount) = MiscRevAmt#
            MiscRevAmt# = 0
      LOOP
      NEXT MCnt

      END IF
  NEXT Cnt
     GOSUB ProcessThisBunch

  'transactions as interfaced
  BlockClear
  ShowProcessingScrn "Updating Cash Transactions"
  FOR Cnt = 1 TO FoundCnt
    GET #CMTran, TranInfo(Cnt).TranRecNo, CMTransRec(1)
    CMTransRec(1).Trans2GL = "Y"
    PUT #CMTran, TranInfo(Cnt).TranRecNo, CMTransRec(1)
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
  PRINT #RPTFile, TOWNNAME$
  PRINT #RPTFile, "Report Date: "; DATE$
  PRINT #RPTFile, "Revenue           Acct. Type              Debit Acct.          Credit Acct."
  PRINT #RPTFile, Dash80$
  NumPrinted = 0
RETURN

PrintBadAcct:
    IF LEN(QPTrim$(BadCAcct$)) = 0 THEN
      BadCAcct$ = "Undefined"
    END IF
    
    PRINT #RPTFile, TAB(43); BadDAcct$; TAB(64); BadCAcct$
    RETURN

ProcessThisBunch:
' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund

IF DayCount <= 0 THEN RETURN

FundCnt = 0                     ' Set Funds Used to Zero

FOR Process = 1 TO DayCount
  GET #MCFile, MiscRec#(Process), MCodeRec(1)
      Acct$ = MCodeRec(1).GLAcctNumb
      CreditFlag = 1            'Must Send Flag to Tell System to Format Acct #
      GOSUB GetAcctName

      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(Acct$)
      GJRec(1).AcctName = AcctName$
      GJRec(1).TrDate = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).CrAmt = MiscAmt#(Process)
      GJRec(1).DrAmt = 0
      GJRec(1).EType = "C"
      GJRec(1).Desc = "FROM CASH MGMT"
      GJRec(1).Src = "CR"
      PUT #GJFile, , GJRec(1)

  'Add Up Fund Total Here for Cash Credit Entry
     IF FundCnt = 0 THEN
        FundCnt = 1
        Fund$(FundCnt) = LEFT$(Acct$, Fund)
        FundAmt#(FundCnt) = MiscAmt#(Process)
     ELSE
        FoundFund = 0
          FOR FCnt = 1 TO FundCnt
            IF Fund$(FCnt) = LEFT$(Acct$, Fund) THEN
              FoundFund = 1
              FundAmt#(FCnt) = FundAmt#(FCnt) + MiscAmt#(Process)
            END IF
          NEXT FCnt
        IF FoundFund = 0 THEN
          FundCnt = FundCnt + 1
          Fund$(FundCnt) = LEFT$(Acct$, Fund)
          FundAmt#(FundCnt) = MiscAmt#(Process)
        END IF
      END IF

  NEXT Process

   'Now Make Matching Debit Entries to Cash Account

   FOR Cash = 1 TO FundCnt
       Acct$ = Fund$(Cash) + GLSetup(1).CRCashAcct
       Acct$ = LEFT$(Acct$, TotAcctLen + 2)
       CreditFlag = 0
       GOSUB GetAcctName

       GJRec(1).AcctRec = 0
       GJRec(1).AcctNum = Acct$
       GJRec(1).AcctName = AcctName$
       GJRec(1).TrDate = WorkDate
       GJRec(1).Ref = Ref$
       GJRec(1).DrAmt = FundAmt#(Cash)
       GJRec(1).CrAmt = 0
       GJRec(1).EType = "D"
       GJRec(1).Desc = "FROM CASH MGMT"
       GJRec(1).Src = "CR"
       PUT #GJFile, , GJRec(1)
  NEXT Cash
BunchReturn:
  RETURN


ValidateGLAccounts:
      'BadAcct = False
      'ThisAcct = ubFindAcct(GJInfo(RevCnt).BAcctInfo.DAcctNo)
     ' IF ThisAcct <= 0 THEN
     ' BadDAcct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
     '   BadAcct = True
     ' ELSE
     '   GJInfo(RevCnt).BAcctInfo.DRecNo = ThisAcct
     '   GJInfo(RevCnt).BAcctInfo.DTitle = ubGetAcctTitle$(ThisAcct)
     '   BadDAcct$ = "     OK"
     ' END IF
    'GOSUB PrintBadAcct
    '
  'CLOSE RPTFile

 ' IF BadAcct THEN
 '   BlockClear
 '   Ok = MsgBox%("UBSENDGL", "UBADACTS")
 '   BlockClear
 '   EntryPoint = 0
 '   PrintRptFile "GL Transfer Invalid Account List.", "UBNOTFND.RPT", 1, RetCode, EntryPoint
 ' END IF
 ' KILL "UBNOTFND.RPT"
 ' RETURN

GetAcctName:
     CLOSE 13
     OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = LEN(GLSetup(1))
     GET 13, 1, GLSetup(1)
     Fund = GLSetup(1).FundLen
     Acct = GLSetup(1).AcctLen
     Obj = GLSetup(1).DetLen
     TotAcctLen = GLSetup(1).TotAcctLen
     CLOSE 13
     OPEN "glacct.idx" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = 16
     FIELD 13, 14 AS AcctNumber$, 2 AS RecordNumber$
     
     IF CreditFlag = 1 THEN
       Acct$ = LEFT$(Acct$, Fund) + "-" + MID$(Acct$, Fund + 1, Acct) + "-" + MID$(Acct$, Fund + Acct + 1, Obj)
       Acct$ = LEFT$(Acct$, TotAcctLen + 2)
     END IF
     FOR Cnt! = 1 TO LOF(13) / 16: GET 13, Cnt!
       IF RTRIM$(AcctNumber$) = RTRIM$(Acct$) THEN
        Record = CVI(RecordNumber$): ng = 1:
        OPEN "glacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 178
        FIELD 12, 16 AS nu$, 30 AS AName$
        GET 12, Record
        AcctName$ = AName$
        CLOSE 12: CLOSE 13
        RETURN
       END IF
      NEXT Cnt!
      ng = 0
      CLOSE 12: CLOSE 13
      RETURN



END SUB

SUB SmallPause
St! = TIMER
St! = St! + .1
DO WHILE St! > TIMER
LOOP
END SUB

