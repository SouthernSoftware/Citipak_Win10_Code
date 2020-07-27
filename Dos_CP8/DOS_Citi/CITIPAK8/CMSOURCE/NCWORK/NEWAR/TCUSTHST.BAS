DECLARE SUB DisplayTaxScrn (ScrnName$)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB ShowTCustHist (CustRec&)
  
  IF CustRec& < 0 THEN
    AdjShadow = True
    CustRec& = ABS(CustRec&)
  ELSE
    AdjShadow = False
  END IF

  u$ = CHR$(24)
  d$ = CHR$(25)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  DisplayTaxScrn "LOADHIST"
  
  REDIM MChoice(1 TO 1) AS FLen2

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  REDIM TaxCustRec(1) AS TaxCustType
  
  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTran(1))
  
  TaxFile = FREEFILE
  OPEN "TaxCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  GET TaxFile, CustRec&, TaxCustRec(1)
  CLOSE TaxFile
  
  'CurBal# = TaxCustRec(1).CurrBalance
  'PreBal# = TaxCustRec(1).PrevBalance
  
Top:
  
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  
  PrevTranRec& = TaxCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
      GET TaxTran, PrevTranRec&, TaxTran(1)
      LSET MChoice(DCnt).V = Num2Date(TaxTran(1).TransDate)
      GOSUB GetTransType
      MID$(MChoice(DCnt).V, 13) = TType$
      MID$(MChoice(DCnt).V, 41) = FUsing(STR$(TaxTran(1).Amount), "#####.##")
      'this will show the actual trans number in the list
      MID$(MChoice(DCnt).V, 52) = FUsing(STR$(PrevTranRec&), "######")
      'MID$(MChoice(DCnt).V, 52) = FUsing(STR$(TaxTran(1).RunBalance), "#####.##")
      MID$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      PrevTranRec& = TaxTran(1).LastTrans
    LOOP
    
    CLOSE TaxTran
    
    RestScrn TempScrn()

    IF AdjShadow THEN
      MPaintBox 6, 4, 18, 76, 8
    ELSE
      MPaintBox 3, 5, 22, 74, 8
    END IF
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "   Date            Description              Amount             "
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    WazzWind BoxBot + 2, Col, BoxBot + 5, MaxLen + 3 + Col, 10, 4, True
    QPrintRC "  Use:  " + u$ + "-" + d$ + " to select.", BoxBot + 3, Col + 3, 15
    QPrintRC u$, BoxBot + 3, Col + 11, 14
    QPrintRC d$, BoxBot + 3, Col + 13, 14
    
    QPrintRC "Total: " + STR$(DCnt), BoxBot + 4, Col + 3, 15
    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, Col + 33, 15
    QPrintRC "        [ENTER] for detail.", BoxBot + 4, Col + 33, 15
    QPrintRC "ESC", BoxBot + 3, Col + 43, 14
    QPrintRC "ENTER", BoxBot + 4, Col + 42, 14
    
    QPrintRC Title$, Row - 1, Col, 112

    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    
    DO
      LOCATE Row, Col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        'RestScrn TempScrn()
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        RestScrn TempScrn()
        GOTO ShowTransDetail
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE TaxTran
    Ok = MsgBox%("Tax.QSL", "NOCTRANS")
    'RestScrn TempScrn()
  END IF
  
  RestScrn TempScrn()
  ERASE MChoice
  ERASE TempScrn, TaxTran, TaxCustRec
  
EXIT SUB
  
ShowTransDetail:
  CursorOff
  TransRecNum& = CVL(RIGHT$(MChoice(Choice).V, 4))
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET TaxTran, TransRecNum&, TaxTran(1)
  GOSUB GetTransType
  CLOSE TaxTran  'NOTE: Close must be after GetTransType

  DisplayTaxScrn "TRDETAIL"
  QPrintRC Num2Date(TaxTran(1).TransDate), 7, 28, -1
  QPrintRC TaxTran(1).Description, 8, 28, -1
  QPrintRC FUsing$(STR$(TaxTran(1).Amount), "#####.##"), 9, 28, -1
  QPrintRC BillType$, 10, 28, -1
  QPrintRC TaxYear$, 11, 60, -1
  QPrintRC Post2GL$, 11, 28, -1
  QPrintRC FUsing$(STR$(Principle#), "#######.##"), 13, 28, -1
  QPrintRC FUsing$(STR$(Interest#), "#######.##"), 14, 28, -1
  QPrintRC FUsing$(STR$(Penalty#), "#######.##"), 15, 28, -1
  QPrintRC FUsing$(STR$(AdCost#), "#######.##"), 16, 28, -1

  WaitForAction
  RestScrn TempScrn2()
  GOTO HistTop
  
GetTransType:

  Principle# = 0
  Interest# = 0
  Penalty# = 0
  AdCost# = 0

  TType$ = TaxTran(1).Description
  BillType$ = ""
  TaxYear$ = "N/A"
  Post2GL$ = "N"
  IF TaxTran(1).Posted2GL = "Y" THEN
    Post2GL$ = "Y"
  END IF

  SELECT CASE TaxTran(1).TranType
  CASE 1
    SELECT CASE TaxTran(1).BillType
    CASE "R"
      BillType$ = "Real-Estate"
    CASE "P"
      BillType$ = "Personal Property"
    CASE "C"
      BillType$ = "Combined"
    CASE "M"
      BillType$ = "Manual"
    END SELECT
    TaxYear$ = QPTrim$(STR$(TaxTran(1).TaxYear))
    Principle# = Round#(TaxTran(1).Revenue.Principle1 + TaxTran(1).Revenue.Principle2 + TaxTran(1).Revenue.Principle3)
    Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4 + TaxTran(1).Revenue.Principle5)
  CASE 2
    BillType$ = "Payment"
    Principle# = Round#(TaxTran(1).Revenue.Principle1PD + TaxTran(1).Revenue.Principle2PD + TaxTran(1).Revenue.Principle3PD)
    Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4PD + TaxTran(1).Revenue.Principle5PD)
  CASE 3
    BillType$ = "Release"
  CASE 4
    BillType$ = "Interest"
    Interest# = TaxTran(1).Revenue.Interest#
    IF TaxTran(1).BelongTo > 0 THEN
      GET TaxTran, TaxTran(1).BelongTo, TaxTran(2)
      TaxYear$ = QPTrim$(STR$(TaxTran(2).TaxYear))
    END IF
  CASE 5
    BillType$ = "Collection/Ad Cost"
  CASE 7
    BillType$ = "Adjustment"
    Principle# = Round#(TaxTran(1).Revenue.Principle1PD + TaxTran(1).Revenue.Principle2PD + TaxTran(1).Revenue.Principle3PD)
    Principle# = Round#(Principle# + TaxTran(1).Revenue.Principle4PD + TaxTran(1).Revenue.Principle5PD)
    Interest# = TaxTran(1).Revenue.InterestPD#
  CASE ELSE
    BillType$ = "?????"
  END SELECT

RETURN
  

END SUB

