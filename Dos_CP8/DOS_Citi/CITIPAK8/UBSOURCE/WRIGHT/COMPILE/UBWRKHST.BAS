DEFINT A-Z

DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()


  TYPE FLen2
    V AS STRING * 64
  END TYPE

  '$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
  '$INCLUDE: 'FORMEDIT.bi'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'setcnf.bi'

  CONST False = 0, True = NOT False

SUB ShowWrkOrdHistory (CustRec&)
  
  REDIM TScrn1(0)
  SaveScrn TScrn1()
  
  DisplayUBScrn "UBCUHIST"
  
  REDIM MChoice(1 TO 1) AS FLen2
  
  REDIM WorkOrderRec(1) AS WorkOrderRecType
  WorkOrderRecLen = LEN(WorkOrderRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, CustRec&, UBCustRec(1)
  CLOSE UBFile
  
  UBWrkOrd = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
  
  PrevTranRec& = UBCustRec(1).WOLastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
      GET UBWrkOrd, PrevTranRec&, WorkOrderRec(1)
      LSET MChoice(DCnt).V = FUsing(STR$(PrevTranRec&), "#####")
      MID$(MChoice(DCnt).V, 15) = Num2Date$(WorkOrderRec(1).ENTRYDATE)
      MID$(MChoice(DCnt).V, 32) = Num2Date$(WorkOrderRec(1).CompleteByDate)
      IF WorkOrderRec(1).CompletedDate <= 0 THEN
        MID$(MChoice(DCnt).V, 50) = "Open"
      ELSE
        MID$(MChoice(DCnt).V, 50) = Num2Date$(WorkOrderRec(1).CompletedDate)
      END IF
      MID$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      PrevTranRec& = WorkOrderRec(1).PrevTransRec
    LOOP
    
    CLOSE UBWrkOrd
    
    RestScrn TScrn1()
    MPaintBox 3, 5, 22, 75, 8
    REDIM TScrn2(2000)
    SaveScrn TScrn2()
    
WOTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    Balance$ = Title$
    LSET Balance$ = " Work Order       Entry           Complete         Completed    "
    LSET Title$ = "   Number          Date           By Date            Date       "
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    TitleBox BoxBot + 3, Col, MaxLen + 3, "       Press <ESC> to continue.", Cnf
    QPrintRC "Transactions:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
    QPrintRC Balance$, Row - 2, Col, 112
    QPrintRC Title$, Row - 1, Col, 112
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    
    DO
      LOCATE Row, Col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TScrn1()
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        RestScrn TScrn2()
        GOTO WOShowDetail
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE UBWrkOrd
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
  END IF
  
  RestScrn TScrn1()
  ERASE MChoice, TScrn1
  ERASE UBCustRec, WorkOrderRec
  
  EXIT SUB
  
WOShowDetail:
  CursorOff
  TransRecNum& = CVL(RIGHT$(MChoice(Choice).V, 4))
  UBWrkOrd = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
  GET UBWrkOrd, TransRecNum&, WorkOrderRec(1)
  CLOSE UBWrkOrd
  
  LibFile2Scrn "UBWRKORD", "WODETAIL", MonoCode, Attribute%, ErrCode
  
  QPrintRC STR$(TransRecNum&), 4, 23, 15
  IF WorkOrderRec(1).CompletedDate <= 0 THEN
    TDate$ = "Open"
  ELSE
    TDate$ = Num2Date$(WorkOrderRec(1).CompletedDate)
  END IF
  QPrintRC Num2Date$(WorkOrderRec(1).ENTRYDATE), 5, 23, 15
  QPrintRC TDate$, 4, 63, 15
  QPrintRC Num2Date$(WorkOrderRec(1).CompleteByDate), 5, 63, 15
  
  FOR Cnt = 1 TO 6
    QPrintRC WorkOrderRec(1).OrdersText.Text(Cnt), Cnt + 7, 8, 15
    QPrintRC WorkOrderRec(1).RepliesText.Text(Cnt), Cnt + 14, 8, 15
  NEXT
  
  WaitForAction
  RestScrn TScrn2()
  GOTO WOTop
  
END SUB

