DECLARE SUB DisplayCRScrn (ScrnName$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag%, Cnf AS ANY)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB SelectChks2Recon
  
  IF WEnvTest THEN
    MaxLoad = 2499
  ELSE
    FreMem& = FRE(-1)
    FreMem& = FreMem& - 20000
    MaxLoad = FreMem& \ 66
    IF MaxLoad > 4999 THEN
      MaxLoad = 4999
    END IF
  END IF
  
  
  ShowProcessingScrn "Reading Check Information File"
  
  REDIM OSChk(1) AS OSChkRecType
  ChkRecLen = LEN(OSChk(1))
  
  'CRHandle = FREEFILE
  'OPEN "crchk.dat" FOR RANDOM SHARED AS CRHandle LEN = ChkRecLen
  
  NumChkRecs = LOF(CRHandle) \ ChkRecLen
  
  IF NumChkRecs = 0 THEN
    'CLOSE
    DisplayCRScrn "MENUBAK"
    OK = MsgBox%("CHECKREC.QSL", "NOCTRANS")
    GOTO AbortExit
  END IF
  
  FOR CCnt = 1 TO NumChkRecs
    GET CRHandle, CCnt, OSChk(1)
    DCnt = DCnt + 1
    Show$ = "Checks Found:" + FUsing(STR$(DCnt), "######")
    QPrintRC Show$, 11, 26, 112
    
    REDIM PRESERVE MTChoice(1 TO DCnt) AS FLen2
    REDIM PRESERVE Picked(1 TO DCnt)
    
    LSET MTChoice(DCnt).V = FUsing(STR$(OSChk(1).ChkNum), "########")
    MID$(MTChoice(DCnt).V, 11) = OSChk(1).ChkDate
    MID$(MTChoice(DCnt).V, 21) = LEFT$(OSChk(1).Desc, 28)
    
    SELECT CASE OSChk(1).Src
    CASE 0
      T$ = "A"
    CASE 1
      T$ = "P"
    END SELECT
    MID$(MTChoice(DCnt).V, 51) = T$
    
    'IF OSChk(1).Amt > 120000 THEN STOP
    
    MID$(MTChoice(DCnt).V, 52) = FUsing(STR$(OSChk(1).Amt), "#######.##")
    
    IF OSChk(1).Cleared > 0 THEN
      PCnt = PCnt + 1
      Picked(PCnt) = DCnt
    END IF
    
    ShowPctComp CCnt, NumChkRecs
    
    IF DCnt > MaxLoad THEN EXIT FOR
    
  NEXT
  
  'CLOSE
  BlockClear
  
HistTop:
  Help$ = SPACE$(80)
  
  MaxLen = 59   'Set menu width to zero
  Action = 0    '0 means stay in the menu until they select something
  
  IF Choice < 1 THEN
    Choice = 1  'Pre-load choice to highlight
  END IF
  
  Title$ = SPACE$(68)
  LSET Title$ = "   Chk No.  Chk Date      Description              Src   Amount TAG"
  
  '--Find max menu width
  '--Center Menu within Screen
  
  LSET Help$ = "Budgetary Accounting"
  
  
  Col = ((80 - 60) \ 2)
  
  Row = 3
  
  DisplayCRScrn "PickOpti"
  QPrintRC Title$, Row - 1, Col - 4, 112
  PickList MTChoice(), Picked(), DCnt, Row, UpDateFlag, Cnf
  
  'OK = MsgBox%("UB.QSL", "NOCTRANS")
  
  IF UpDateFlag THEN
    ShowProcessingScrn "Processing Selections"
    
    FOR CCnt = 1 TO NumChkRecs
      GET CRHandle, CCnt, OSChk(1)
      OSChk(1).Cleared = 0
      PUT CRHandle, CCnt, OSChk(1)
      ShowPctComp CCnt, NumChkRecs
    NEXT
    
    FOR PCnt = 1 TO DCnt
      GET CRHandle, Picked(PCnt), OSChk(1)
      OSChk(1).Cleared = 1
      PUT CRHandle, Picked(PCnt), OSChk(1)
      ShowPctComp PCnt, DCnt
    NEXT
    
    'close and reopen the checks file! should
    'force windows to commit all to disk
    CLOSE CRHandle
    IF NOT OpenLockCRFile(CRHandle) THEN
      BlockClear
      DisplayCRScrn "LOCKERR"
      WaitForAction
      CLS
      HideCursor
      END
    END IF
  END IF
  
AbortExit:
  
END SUB

