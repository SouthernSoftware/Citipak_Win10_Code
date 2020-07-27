DEFINT A-Z
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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  ''$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  ''$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'

  CONST False = 0, True = NOT False

SUB CustomerInquiry (RecNo&)
  
  IF RecNo& = 0 THEN
    BlockClear
    LookUp RecNo&, "Customer Inquiry", 2, True, False
    IF RecNo& = 0 THEN
      GOTO ExitInquiry
    END IF
  ELSE
    InfoFlag = True
  END IF
  
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
  LSET Form$(2, 0) = UBCustRec(1).BOOK
  LSET Form$(3, 0) = UBCustRec(1).SEQNUMB
  LSET Form$(4, 0) = UBCustRec(1).Status
  LSET Form$(5, 0) = UBCustRec(1).CUSTNAME
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
  LSET Form$(20, 0) = UBCustRec(1).USEDRAFT
  LSET Form$(21, 0) = UBCustRec(1).BANKNAME
  LSET Form$(22, 0) = UBCustRec(1).BANKLOC
  LSET Form$(23, 0) = UBCustRec(1).TRANSIT
  LSET Form$(24, 0) = UBCustRec(1).BANKACCT
  LSET Form$(25, 0) = STR$((Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)))
  LSET Form$(26, 0) = STR$(Round#(UBCustRec(1).CurrBalance))
  LSET Form$(27, 0) = STR$(Round#(UBCustRec(1).PrevBalance))
  LSET Form$(28, 0) = STR$(Round#(UBCustRec(1).DepositAmt))
RETURN

CheckCustMsg:
  IF CustHasMsg(RecNo&) THEN
    GotMsgFlag = True
  ELSE
    GotMsgFlag = False
  END IF
RETURN

END SUB

