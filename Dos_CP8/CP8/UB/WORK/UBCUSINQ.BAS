DEFINT A-Z
DECLARE SUB CursorOff ()
DECLARE SUB RestScrn (array())
DECLARE SUB SaveScrn (array())
DECLARE SUB WaitForAction ()
DECLARE FUNCTION CustHasMsg% (RecNo&)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Round# (N#)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB BlockClear ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, PColor%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  ''$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubowner.BI'

  CONST False = 0, True = NOT False

  TYPE CycleType
     CustCnt AS LONG
     PendCnt AS LONG
  END TYPE

SUB CustomerInquiry (RecNo&)

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  IF RecNo& = 0 THEN
    BlockClear
    LookUp RecNo&, "Customer Inquiry", 2, True, False
    IF RecNo& = 0 THEN
      GOTO ExitInquiry
    END IF
  ELSE
    InfoFlag = True
  END IF

'*********Setup the owner control arrays
  OwnNumFlds = LibNumberOfFields("UB", "UBOWNER")
  REDIM OwnFrm(1) AS FormInfo
  REDIM OwnForm$(OwnNumFlds, 2)
  REDIM OwnFld(OwnNumFlds) AS FieldInfo

  OwnStartEL = 0
  LibGetFldDef "UB", "UBOWNER", OwnStartEL, OwnFld(), OwnForm$(), ErrCode
  OwnForm$(0, 0) = SPACE$(UBOwnerRecLen)
'***************
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
MainBody:
  'RecNo& = RecNo&
  
  LibName$ = "UB"
  ScrnName$ = "UBCUSINQ"
  
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
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1

  DisplayUBScrn ScrnName$
  
  IF InfoFlag THEN
    QPrintRC "Information", 2, 17, 14
    QPrintRC " ]", 2, 28, -1
  END IF

  FirstTime = True

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadCustInfo
      GOSUB CheckCustMsg
      IF GotMsgFlag THEN
        QPrintRC "F7", 22, 46, 254
      ELSE
        QPrintRC "F7", 22, 46, 126
      END IF
      QPrintRC QPTrim$(UBCustRec(1).ServAddr), 8, 20, -1
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
      
    CASE F4KEY  'Customer History
      IF RecNo& > 0 THEN
        ShowCustHistory RecNo&
        Action = 2
      END IF
    CASE F6KEY
      ShowCustConsHist RecNo&
      Action = 2
    CASE F3Key
      ShowWrkOrdHistory RecNo&
      Action = 1
    CASE F7KEY
      CustMessageSystem RecNo&
      FirstTime = True
      Action = 1
      GOTO MainBody
    CASE F8Key
      GOSUB ShowOwnerInfo
      Action = 1
    CASE F10Key
      Done = True
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
    
  LOOP UNTIL Done
  
ExitInquiry:
  EXIT SUB
  '***************
  
LoadCustInfo:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE
  
  LSET Form$(1, 0) = STR$(RecNo&)
  LSET Form$(2, 0) = UBCustRec(1).Book
  LSET Form$(3, 0) = UBCustRec(1).SeqNumb
  LSET Form$(4, 0) = UBCustRec(1).Status
  LSET Form$(5, 0) = UBCustRec(1).CustName
  LSET Form$(6, 0) = UBCustRec(1).ADDR1
  LSET Form$(7, 0) = UBCustRec(1).ADDR2
  LSET Form$(8, 0) = UBCustRec(1).CITY
  LSET Form$(9, 0) = UBCustRec(1).STATE
  LSET Form$(10, 0) = UBCustRec(1).ZIPCODE
  LSET Form$(11, 0) = UBCustRec(1).HPHONE
  LSET Form$(12, 0) = Num2Date$(UBCustRec(1).OPENDATE)
  LSET Form$(13, 0) = UBCustRec(1).CUTOFFYN
  Soc$ = UBCustRec(1).SOSEC
  LSET Form$(14, 0) = LEFT$(Soc$, 3) + "-" + MID$(Soc$, 4, 2) + "-" + MID$(Soc$, 6)
  LSET Form$(15, 0) = UBCustRec(1).DRVLIC
  LSET Form$(16, 0) = UBCustRec(1).WPHONE
  LSET Form$(17, 0) = UBCustRec(1).TAXEXPT
  LSET Form$(18, 0) = UBCustRec(1).CUSTTYPE
  LSET Form$(19, 0) = STR$(UBCustRec(1).BILLCOPY)

  LSET Form$(20, 0) = UBCustRec(1).Zone
  LSET Form$(21, 0) = UBCustRec(1).USEDRAFT
  LSET Form$(22, 0) = UBCustRec(1).BANKNAME
  LSET Form$(23, 0) = UBCustRec(1).BANKLOC
  LSET Form$(24, 0) = UBCustRec(1).TRANSIT
  LSET Form$(25, 0) = UBCustRec(1).BANKACCT
  LSET Form$(26, 0) = STR$((Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)))
  LSET Form$(27, 0) = STR$(Round#(UBCustRec(1).CurrBalance))
  LSET Form$(28, 0) = STR$(Round#(UBCustRec(1).PrevBalance))
  LSET Form$(29, 0) = STR$(Round#(UBCustRec(1).DepositAmt))
  'UBCustRec(1).DepositAmt = 100
RETURN

CheckCustMsg:
  IF CustHasMsg(RecNo&) THEN
    GotMsgFlag = True
  ELSE
    GotMsgFlag = False
  END IF
RETURN

ShowOwnerInfo:

  OwnFrm(1).FldNo = 1
  OExitFlag = False

  REDIM TempScr(0)
  SaveScrn TempScr()

  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen
  GET UBFile, RecNo&, UBOwnerRec(1)
  CLOSE

  BCopy VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), UBOwnerRecLen, 0
  UnPackBuffer 0, 0, OwnForm$(), OwnFld()

  OAction = 1

  MPaintBox 2, 5, 22, 75, 8
  DisplayUBScrn "UBOWNER"

  QPrintRC "Press any Key to Continue.", 19, 29, 15

  EditForm OwnForm$(), OwnFld(), OwnFrm(1), Cnf, OAction

  CursorOff

  WaitForAction

  RestScrn TempScr()
  ERASE TempScr

RETURN


END SUB

SUB CycleCount

  Title$ = "Cycle Count Summary Report"
  Dash80$ = STRING$(80, "-")

  REDIM Cycles(-1 TO 99) AS CycleType

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpRecLen

  TownName$ = QPTrim$(UBSetUp(1).UTILNAME)
  TabStop = 40 - (LEN(TownName$) / 2)
  IF TabStop < 1 THEN TabStop = 1
  TownName$ = SPACE$(TabStop) + TownName$
  
  UBRpt = FREEFILE
  OPEN "UBCYCLE.RPT" FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn Title$

  GOSUB DoLocaRptHeader

'***************
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumCust& = LOF(UBFile) \ UBCustRecLen
  
  FOR CCnt& = 1 TO NumCust&
    GET UBFile, CCnt&, UBCustRec(1)
    IF UBCustRec(1).BILLCYCL >= 0 AND (NOT UBCustRec(1).DelFlag) THEN
      Cycles(UBCustRec(1).BILLCYCL).CustCnt = Cycles(UBCustRec(1).BILLCYCL).CustCnt + 1
      IF UBCustRec(1).Status = "P" THEN
        Cycles(UBCustRec(1).BILLCYCL).PendCnt = Cycles(UBCustRec(1).BILLCYCL).PendCnt + 1
      END IF
    ELSE
      IF NOT UBCustRec(1).DelFlag THEN
        'LPRINT CCnt&
        Cycles(-1).CustCnt = Cycles(-1).CustCnt + 1
        IF UBCustRec(1).Status = "P" THEN
          Cycles(-1).PendCnt = Cycles(-1).PendCnt + 1
        END IF
      END IF
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      GOTO CycleAbortExit
    END IF

    ShowPctCompL CCnt&, NumCust&
  NEXT

  FOR zz = 0 TO 99
    IF Cycles(zz).CustCnt > 0 THEN
      PRINT #UBRpt, USING "######"; zz; TAB(30); Cycles(zz).CustCnt; TAB(48); Cycles(zz).PendCnt
    END IF
  NEXT

  IF Cycles(-1).CustCnt > 0 THEN
    PRINT #UBRpt, "    ??";
    PRINT #UBRpt, TAB(30); USING "######"; Cycles(-1).CustCnt; TAB(48); Cycles(-1).PendCnt
  END IF
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, CHR$(12)

CycleAbortExit:

  CLOSE

  ERASE UBCustRec

  IF NOT AbortFlag THEN
    PrintRptFile Title$, "UBCYCLE.RPT", 1, RetCode, EntryPoint
  END IF


EXIT SUB

DoLocaRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TownName$
  PRINT #UBRpt, "                           Cycle Count Summary Report"
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt,
  PRINT #UBRpt, "Bill Cycle            Customer Count            Pending"
  PRINT #UBRpt, Dash80$
RETURN

END SUB

