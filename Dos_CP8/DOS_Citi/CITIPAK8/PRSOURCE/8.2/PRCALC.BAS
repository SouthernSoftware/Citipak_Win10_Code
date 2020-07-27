DEFINT A-Z
DECLARE SUB PCPickEmpList (TransRecNo%, EmpRecNo%)
DECLARE SUB Post2BA (FileName$, BadTrans%)
  
DECLARE SUB PCPrintPayRegisterS ()
DECLARE SUB MakeTransInActive ()
DECLARE FUNCTION PromptPeriodWasActive% ()
  
DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FldNum (FldName$, Fld() AS ANY)
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION PCDelFromPay% ()
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION Value# (e$, ErCode%)
DECLARE SUB AccruLeave (SaveScrnFlag%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB CalcFields (StartOfForm, FldNo, Form$(), Fld() AS ANY)
DECLARE SUB CalcPay (TransRec AS ANY, TransRecNo%, ReCalcFlag)
DECLARE SUB CreateEmpTransRecs (RecNo%)
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode, Attribute, ErrorCode)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB MakeDefaultTransActs ()
DECLARE SUB ManualMenu ()
DECLARE SUB PCEdScrnCalc (TransRecNo%, EmpName$, EmpRecNo%)
DECLARE SUB PCEnterEdTrans (TransRecNo%, EmpRecNo%)
DECLARE SUB PCGetEmp2Rec (EmpRecNo%, ErrorCode%)
DECLARE SUB PCGetEmpNum (EmpNumber%)
DECLARE SUB PCLoadPayFreqs ()
DECLARE SUB PCLoadSystemFiles ()
DECLARE SUB PCLookUpEmp (EmpNum$, TRecNum%, ERecNum%)
DECLARE SUB PCPrintManRegister ()
DECLARE SUB PCPrintPayRegister ()
DECLARE SUB PCSetPeriodDefault ()
DECLARE SUB PRCheckMenu ()
DECLARE SUB ParseHourly2Trans ()
DECLARE SUB ParseManual2Trans (ManTrans AS ANY)
DECLARE SUB ParseSalary2Trans ()
DECLARE SUB ParseScrnCalc2Trans ()
DECLARE SUB ParseTrans2Hourly ()
DECLARE SUB ParseTrans2Manual (ManTrans AS ANY)
DECLARE SUB ParseTrans2Salary ()
DECLARE SUB ParseTrans2ScrnCalc ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB PostTransactions ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintArray (FirstFld%, LastFld%, Form$(), Fld() AS ANY)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ReplaceString (Work$, Old$, New$)
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB ShowCursor ()
DECLARE SUB UnPackBuffer (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB UpDateTransFile (TransRecNo%)
DECLARE SUB WaitForAction ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION CheckFldEmpty% (FldData$)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB WazzWind (BYVAL TopRow, BYVAL LeftCol, BYVAL BotRow, BYVAL RghtCol, BYVAL FrameColor, BYVAL FrameType, BYVAL Shadow)
DECLARE FUNCTION DOSError ()
'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'ErnCodes.Bi'
'$INCLUDE: 'Leave.Bi'
'$INCLUDE: 'StateTax.Bi'
'$INCLUDE: 'FedTax.Bi'
'$INCLUDE: 'PREIC.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRPPDEF.Bi'
'$INCLUDE: 'PRTRANS.Bi'
''$INCLUDE: 'oPRTRANS.Bi'
'$INCLUDE: 'PRTIMEIN.Bi'
'$INCLUDE: 'ScrCalc1.Bi'
'$INCLUDE: 'PRCHECK.Bi'
'$INCLUDE: 'PROSCHK.bi'
'$INCLUDE: 'prif.bi'
  
  CONST False = 0
  CONST True = NOT False
  
  CONST manual = 2
  CONST Normal = 1
  
  
  REDIM SHARED DedCodes(0)       AS DedCodeRecType
  REDIM SHARED ErnCodes(0)       AS ErnCodeRecType
  REDIM SHARED RetireRec(0)      AS RetireRecType
  
  REDIM SHARED StateTax(0)       AS StateTaxRecType
  REDIM SHARED FedTax(0)         AS FederalTaxRecType
  REDIM SHARED EICRec(0)         AS EICRecType
  REDIM SHARED PrdDefRec(1)      AS PeriodDefaultRecType
  
  'DIM SHARED EmpRec1   AS EmpData1Type
  'DIM SHARED EmpRec3   AS EmpData3Type
  
  DIM SHARED EMP2Rec(1)   AS EmpData2Type
  
  DIM SHARED SalInput(1) AS SalaryInputType
  DIM SHARED HourInput(1) AS HourlyInputType
  
  REDIM SHARED ScrnCalc(1)      AS ScrnCalcType
  
  REDIM SHARED TransRec(1)      AS TransRecType
  
  'DIM SHARED ScrnCalc     AS ScrnCalcType
  'DIM SHARED UndoScrnCalc AS ScrnCalcType
  
  DIM SHARED THandle AS INTEGER, TRecSize AS INTEGER, EntryType AS INTEGER
  
  DIM SHARED EmpName$
  DIM SHARED PayPFreq$(1 TO 7)
  
  'REDIM SHARED Items(0) AS FLen
  
  'PRINT LEN(EmpRec1), LEN(EmpRec3), LEN(EMP2Rec(1))
  
  'ON TIMER(1) GOSUB ShowStack
  'TIMER ON
  
  'ShowStack:
  '  QPrintRC STR$(FRE(-2)), 25, 1, 31
  '  RETURN
  
  '  DIM Emp3 AS EmpData3Type
  '  FOpenS EmpData3Name, Handle3
  '  RecLen = LEN(Emp3)
  '  NumOfRec = FileSize(EmpData3Name) \ RecLen
  '  FOR Cnt = 1 TO NumOfRec
  '    FGetRTA Handle3, Emp3, CLNG(Cnt), RecLen
  '    Emp3.Data1RecNum = Cnt
  '    FPutRTA Handle3, Emp3, CLNG(Cnt), RecLen
  '  NEXT
  '  FClose Handle3
  'END
  
  '  DIM Emp2 AS EmpData2Type
  '  RecLen = LEN(Emp2)
  '
  '  OPEN EmpData2Name FOR RANDOM AS #1 LEN = RecLen
  '
  '  NumOfRec = FileSize(EmpData2Name) \ RecLen
  '  FOR Cnt = 1 TO NumOfRec
  '    GET #1, Cnt, Emp2
  '     'IF Emp2.EmpPin = 171 THEN PRINT Emp2.EmpLName; Emp2.EmpFName
  '     IF UCASE$(QPTrim$(Emp2.EmpLName)) = "WILLIS" THEN STOP
  ''    PRINT Emp2.EmpPin
  '    z! = TIMER
  '    DO: LOOP WHILE z! + .1 > TIMER
  '    'z$ = INPUT$(1)
  '    IF INSTR(UCASE$(Emp2.EmpPType), "HOUR") THEN LPRINT Emp2.EmpNo
  '
  '  NEXT
  '  CLOSE
  ''STOP
  'END
  '
'  TDate = Date2Num("03-29-1996")
'  DIM TR AS TransRecType
'  RecLen = LEN(TR)
  '
  'KILL "prdata\newtranh.dat"
  
'  NumOfRec = FileSize(TransHistFileName) \ RecLen
'  OPEN TransHistFileName FOR RANDOM AS #1 LEN = RecLen
  
'  OPEN "prdata\newtranh.dat" FOR RANDOM AS #2 LEN = RecLen
  '
  'RecCnt = 1
  
'  FOR Cnt = 1 TO NumOfRec
'    GET #1, Cnt, TR
    '
'    IF TR.EmpPin = 47 THEN      'AND TR.CheckDate = TDate THEN
      
      'IF TR.CheckNum = 20651 OR TR.CheckNum = 17 THEN
'      Removed = Removed + 1
'    ELSE
'      PUT #2, , TR
'    END IF
    'ELSE
    '  PUT #2, , TR
    'END IF
    '     PRINT TR.CheckNum, TR.TotRegWage, TR.GrossWage, TR.NETPAY
    'STOP
    '      IF TR.CheckNum > 0 THEN
    '        'PUT #2, RecCnt, TR
    '        RecCnt = RecCnt + 1
    '        Removed = Removed + 1
    '      END IF
    '    ELSE
    '      'PUT #2, RecCnt, TR
    '      RecCnt = RecCnt + 1
    '    END IF
'  NEXT
'  CLOSE
'  PRINT "Total removed:"; Removed
'  END
  '
  'PRINT TotalTra, Net#
  
  '  DIM TR AS TransRecType
  '  LastEmpPin = -100
  
  '  RecLen = LEN(TR)
  '  Date1 = Date2Num("06-07-1995")
  'Date2 = Date2Num("08-31-1995")
  
  '  OPEN TransHistFileName FOR RANDOM AS #1 LEN = RecLen
  '  OPEN "prdata\Temphist.dat" FOR RANDOM AS #2 LEN = RecLen
  
  '  NumOfRec = FileSize(TransHistFileName) \ RecLen
  '  FOR Cnt = 1 TO NumOfRec
  '    GET #1, Cnt, TR
  '    IF TR.RetireAmt = 0 THEN '    LastEmpPin <> TR.EmpPin THEN  'TR.CheckDate <= Date2 THEN
  '      TR.MatchRetAmt = 0
  '      DidCnt = DidCnt + 1
  '      PUT #1, Cnt, TR
  '    END IF
  
  '      PRINT TR.GrossPay
  '      STOP
  '      EXIT FOR
  '    END IF
  
  '     IF TR.DAMT(2) > 0 THEN
  '       TGrossPay# = Round(TR.GROSSPAY - TR.DAMT(2))
  '     ELSE
  '       TGrossPay# = TR.GROSSPAY
  '     END IF
  '    TR.FedGrossPay = TGrossPay#
  '    TR.StaGrossPay = TGrossPay#
  '    TR.SocGrossPay = TR.GROSSPAY
  '    TR.MedGrossPay = TR.GROSSPAY
  '    TR.RetGrossPay = TR.GROSSPAY
  '    PUT #1, Cnt, TR
  '  NEXT
  '  CLOSE
  '  CLS
  '  PRINT "Number of Trans:"; NumOfRec; " Changed:"; DidCnt

SUB CalcPay (TransRec AS TransRecType, TransRecNo%, ReCalcFlag)
  
  SHARED TaxText$()
  
  REDIM DedAmts#(1 TO 12)
  REDIM ErnAmts#(1 TO 3)
  
  ErnAmts#(1) = TransRec.EAmt(1)
  ErnAmts#(2) = TransRec.EAmt(2)
  ErnAmts#(3) = TransRec.EAmt(3)
  
  
  '*************************************************************
  'Hourly processing
  
  SELECT CASE TransRec.PayType
  CASE "H"
    TransRec.TotRegWage = Round#(TransRec.BaseRate * TransRec.RegHrsPaid)
    TransRec.TotOTWage = Round#(TransRec.OTRate * TransRec.OTHrsPaid)
    TransRec.GrossWage = Round#(TransRec.TotRegWage + TransRec.TotOTWage)
    
    IF TransRec.GrossWage > 0 THEN
      FOR Cnt = 1 TO 8          'split out wage distrubtions
        TransRec.TDist(Cnt).DRWage = Round#(TransRec.TDist(Cnt).DRHrs * TransRec.BaseRate)
        TransRec.TDist(Cnt).DOWage = Round#(TransRec.TDist(Cnt).DOHrs * TransRec.OTRate)
        IF TransRec.TDist(Cnt).DRWage > 0 THEN LastRActiveDist = Cnt
        IF TransRec.TDist(Cnt).DOWage > 0 THEN LastOActiveDist = Cnt
        TransRec.TDist(Cnt).DPct = Round#(100 * (TransRec.TDist(Cnt).DRWage + TransRec.TDist(Cnt).DOWage) / TransRec.GrossWage)
      NEXT
    ELSE        'deal with zero gross wage, but has an addtional earning dist
      'here
      FOR Cnt = 1 TO 8          'No Gross All from Add earnings
        TransRec.TDist(Cnt).DRWage = 0
        TransRec.TDist(Cnt).DOWage = 0
        '    TransRec.TDist(Cnt).DPct = 0
      NEXT
      TransRec.TDist(1).DPct = 100
    END IF
    'calc and adjust regular wage distributions
    IF LastRActiveDist > 0 THEN
      DO
        TotalRegDist# = 0
        FOR Cnt = 1 TO 8
          TotalRegDist# = Round#(TotalRegDist# + TransRec.TDist(Cnt).DRWage)
        NEXT
        IF TotalRegDist# <> TransRec.TotRegWage THEN
          IF TotalRegDist# > TransRec.TotRegWage THEN
            DistDif# = Round#(TotalRegDist# - TransRec.TotRegWage)
            TransRec.TDist(LastRActiveDist).DRWage = TransRec.TDist(LastRActiveDist).DRWage - DistDif#
          ELSEIF TotalRegDist# < TransRec.TotRegWage THEN
            DistDif# = Round#(TransRec.TotRegWage - TotalRegDist#)
            TransRec.TDist(LastRActiveDist).DRWage = TransRec.TDist(LastRActiveDist).DRWage + DistDif#
          END IF
        END IF
      LOOP UNTIL TotalRegDist# = TransRec.TotRegWage
    END IF
    'calc and adjust Overtime wage distributions
    DO
      TotalOTDist# = 0
      FOR Cnt = 1 TO 8
        TotalOTDist# = Round#(TotalOTDist# + TransRec.TDist(Cnt).DOWage)
      NEXT
      IF TotalOTDist# <> TransRec.TotOTWage THEN
        IF TotalOTDist# > TransRec.TotOTWage THEN
          DistDif# = Round#(TotalOTDist# - TransRec.TotOTWage)
          TransRec.TDist(LastOActiveDist).DOWage = TransRec.TDist(LastOActiveDist).DOWage - DistDif#
        ELSEIF TotalOTDist# < TransRec.TotOTWage THEN
          DistDif# = Round#(TransRec.TotOTWage - TotalOTDist#)
          TransRec.TDist(LastOActiveDist).DOWage = TransRec.TDist(LastOActiveDist).DOWage + DistDif#
        END IF
      END IF
    LOOP UNTIL TotalOTDist# = TransRec.TotOTWage
    
    'maybe? put code here to deal with no gross and alt earns
    '     if TransRec.GrossWage=0 and TransRec.TotAddEarn > 0
    
    '**********************************************************************
    'Salaried processing
  CASE "S"
    IF LEN(QPTrim$(TransRec.PaySFlag)) = 0 THEN
      TransRec.PaySFlag = "Y"
    END IF
    IF TransRec.PaySFlag = "Y" THEN
      IF NOT ReCalcFlag THEN
        TransRec.GrossWage = Round#(TransRec.BaseRate)
        TransRec.TotRegWage = Round#(TransRec.BaseRate)         '******
      ELSE
        TransRec.GrossWage = TransRec.TotRegWage
      END IF
      FOR Cnt = 1 TO 8
        
        TransRec.TDist(Cnt).DRWage = Round#((TransRec.TDist(Cnt).DPct * .01) * TransRec.GrossWage)
        'TransRec.TDist(Cnt).DPct = Round#((TransRec.TDist(Cnt).DRWage / TransRec.GrossWage) * 100)
        
        IF TransRec.TDist(Cnt).DRWage > 0 THEN LastActiveDist = Cnt
      NEXT
      
      IF LastActiveDist = 0 THEN
        TransRec.TActive = False
        EXIT SUB
      END IF
      DO
        TotalWageDist# = 0
        FOR Cnt = 1 TO 8
          TotalWageDist# = Round#(TotalWageDist# + TransRec.TDist(Cnt).DRWage)
        NEXT
        IF TotalWageDist# <> TransRec.GrossWage THEN
          IF TotalWageDist# > TransRec.GrossWage THEN
            DistDif# = Round#(TotalWageDist# - TransRec.GrossWage)
            TransRec.TDist(LastActiveDist).DRWage = TransRec.TDist(LastActiveDist).DRWage - DistDif#
          ELSEIF TotalWageDist# < TransRec.GrossWage THEN
            DistDif# = Round#(TransRec.GrossWage - TotalWageDist#)
            TransRec.TDist(LastActiveDist).DRWage = TransRec.TDist(LastActiveDist).DRWage + DistDif#
          END IF
        END IF
      LOOP UNTIL TotalWageDist# = TransRec.GrossWage
    ELSE        'deal with zero gross wage, but has an addtional earning dist
      'here
      
      FOR Cnt = 1 TO 8
        TransRec.TDist(Cnt).DRWage = 0
        TransRec.TDist(Cnt).DPct = 0
      NEXT
      'testing VVV
      TransRec.TDist(1).DPct = 100
      TransRec.GrossWage = 0
      TransRec.TotRegWage = 0
      TransRec.BaseRate = 0
      
      'TransRec.GrossPay = 0
      'TransRec.NetPay = 0
    END IF
  END SELECT
  '
  '***********************************************************************
  
  TransRec.GrossPay = Round#(TransRec.GrossWage + TransRec.TotAdditEarn)
  TransRec.FedGrossPay = TransRec.GrossPay
  TransRec.StaGrossPay = TransRec.GrossPay
  TransRec.SocGrossPay = TransRec.GrossPay
  TransRec.MedGrossPay = TransRec.GrossPay
  TransRec.RetGrossPay = TransRec.GrossPay
  
  'RetGrossPay# = TransRec.GrossPay
  
  '****************************************************************
  'Deduction processing
  '02/27/94 added for-next structure to "Deduction calc" below.
  'Reduced from "219" to "21" lines of code
  
  IF NOT ReCalcFlag THEN
    TransRec.TotDedAmt = 0
    FOR Cnt = 1 TO 12           '12 total deductions
      IF ASCII(PrdDefRec(1).UseDed(Cnt)) = UseY THEN
        IF EMP2Rec(1).EmpDed(Cnt).DAmt > 0 THEN
          SELECT CASE ASCII(EMP2Rec(1).EmpDed(Cnt).DPct)
          CASE 65               'Amount
            DedAmts#(Cnt) = EMP2Rec(1).EmpDed(Cnt).DAmt
          CASE 80               'Percent
            SELECT CASE ASCII(EMP2Rec(1).EmpDed(Cnt).DOTI)
            CASE UseN           'no
              DedAmts#(Cnt) = Round#(TransRec.TotRegWage * (EMP2Rec(1).EmpDed(Cnt).DAmt * .01))
            CASE 32, UseY       'Yes, or "space (blank)"
              DedAmts#(Cnt) = Round#(TransRec.GrossPay * (EMP2Rec(1).EmpDed(Cnt).DAmt * .01))
            END SELECT
          END SELECT
          TransRec.DAmt(Cnt) = DedAmts#(Cnt)
        ELSE
          TransRec.DAmt(Cnt) = 0
        END IF
      END IF
      TransRec.TotDedAmt = Round#(TransRec.TotDedAmt + TransRec.DAmt(Cnt))
    NEXT
  ELSE          'recalc flag is true
    '10-02-94 *** Fixed recalc for deduction that are Percentage based.
    TransRec.TotDedAmt = 0
    FOR Cnt = 1 TO 12
      IF TransRec.DAmt(Cnt) > 0 THEN
        DedAmts#(Cnt) = TransRec.DAmt(Cnt)
        SELECT CASE ASCII(EMP2Rec(1).EmpDed(Cnt).DPct)
        CASE 65 'Amount
          DedAmts#(Cnt) = TransRec.DAmt(Cnt)
        CASE 80 'Percent
          SELECT CASE ASCII(EMP2Rec(1).EmpDed(Cnt).DOTI)
          CASE UseN             'no
            DedAmts#(Cnt) = Round#(TransRec.TotRegWage * (EMP2Rec(1).EmpDed(Cnt).DAmt * .01))
          CASE 32, UseY         'Yes, or "space (blank)"
            DedAmts#(Cnt) = Round#(TransRec.GrossPay * (EMP2Rec(1).EmpDed(Cnt).DAmt * .01))
          END SELECT
        END SELECT
      END IF
      TransRec.DAmt(Cnt) = DedAmts#(Cnt)
      TransRec.TotDedAmt = Round#(TransRec.TotDedAmt + TransRec.DAmt(Cnt))
    NEXT
  END IF
  '****************************************************************
  
  'calc taxable amounts "DEDUCTIONS"
  '  IF NOT ReCalcFlag THEN
  FOR Cnt = 1 TO 12
    IF DedCodes(Cnt).DCFWT1 = "Y" THEN
      IF DedAmts#(Cnt) > 0 THEN
        TransRec.FedGrossPay = Round#(TransRec.FedGrossPay - DedAmts#(Cnt))
      END IF
    END IF
    IF DedCodes(Cnt).DCSWT1 = "Y" THEN
      IF DedAmts#(Cnt) > 0 THEN
        TransRec.StaGrossPay = Round#(TransRec.StaGrossPay - DedAmts#(Cnt))
      END IF
    END IF
    IF DedCodes(Cnt).DCSOC1 = "Y" THEN
      IF DedAmts#(Cnt) > 0 THEN
        TransRec.SocGrossPay = Round#(TransRec.SocGrossPay - DedAmts#(Cnt))
      END IF
    END IF
    IF DedCodes(Cnt).DCMED1 = "Y" THEN
      IF DedAmts#(Cnt) > 0 THEN
        TransRec.MedGrossPay = Round#(TransRec.MedGrossPay - DedAmts#(Cnt))
      END IF
    END IF
  NEXT
  
  '   if retamt > 0 and not tax exempt then adjust fedgrosspay
  
  'calc taxable amounts "EARNINGS"
  FOR Cnt = 1 TO 3
    
    IF ErnCodes(Cnt).ERNFWT1 = "N" THEN
      TransRec.FedGrossPay = Round#(TransRec.FedGrossPay - ErnAmts#(Cnt))
    END IF
    IF ErnCodes(Cnt).ERNSWT1 = "N" THEN
      TransRec.StaGrossPay = Round#(TransRec.StaGrossPay - ErnAmts#(Cnt))
    END IF
    IF ErnCodes(Cnt).ERNSOC1 = "N" THEN
      TransRec.SocGrossPay = Round#(TransRec.SocGrossPay - ErnAmts#(Cnt))
    END IF
    IF ErnCodes(Cnt).ERNMED1 = "N" THEN
      TransRec.MedGrossPay = Round#(TransRec.MedGrossPay - ErnAmts#(Cnt))
    END IF
    IF ErnCodes(Cnt).ERNRET1 = "N" THEN
      TransRec.RetGrossPay = Round#(TransRec.RetGrossPay - ErnAmts#(Cnt))
    END IF
  NEXT
  
  'got soc med gross here
  'calc retirement W/H
  
  IF LEN(QPTrim$(EMP2Rec(1).EMPRETNO)) THEN
    FOR Cnt = 1 TO 6
      IF UCASE$(QPTrim$(RetireRec(Cnt).TYPEDES1)) = UCASE$(QPTrim$(EMP2Rec(1).EMPRETTP)) THEN
        EXIT FOR
      END IF
    NEXT
    
    SELECT CASE Cnt
    CASE 1 TO 6
      SELECT CASE RetireRec(Cnt).TYPEOT1
      CASE "Y"  'include overtime in retirement calc!
        TransRec.RetireAmt = Round#(TransRec.RetGrossPay * (RetireRec(Cnt).TYPEWH1 * .01))
        TransRec.MatchRetAmt = Round#(TransRec.RetGrossPay * (RetireRec(Cnt).TYPEM1 * .01))
      CASE "N"  'nope don't include OT
        '04-04-95 ***  fixed bug RetGrossPay
        'If OT not included in calc of retamt
        'OT wages allready added to RetGross must subtract ot wage
        TransRec.RetGrossPay = Round#(TransRec.RetGrossPay - TransRec.TotOTWage)
        TransRec.RetireAmt = Round#(TransRec.RetGrossPay * (RetireRec(Cnt).TYPEWH1 * .01))
        TransRec.MatchRetAmt = Round#(TransRec.RetGrossPay * (RetireRec(Cnt).TYPEM1 * .01))
      END SELECT
    CASE ELSE
      TransRec.RetGrossPay = 0
      TransRec.RetireAmt = 0
      TransRec.MatchRetAmt = 0
    END SELECT
  ELSE
    TransRec.RetGrossPay = 0
    TransRec.RetireAmt = 0
    TransRec.MatchRetAmt = 0
  END IF
  
  '08-26-94 added code here to check if retirement is tax deferred.
  'adjust taxable amts  after retirement calc
  IF Cnt < 7 THEN
    IF QPTrim$(RetireRec(Cnt).TYPETD1) = "Y" THEN

      TransRec.FedGrossPay = Round#(TransRec.FedGrossPay - TransRec.RetireAmt)

      TransRec.StaGrossPay = Round#(TransRec.StaGrossPay - TransRec.RetireAmt)
    END IF
  END IF
  'add retirement to total deductions
  TransRec.TotDedAmt = Round#(TransRec.TotDedAmt + TransRec.RetireAmt)
  
  SocExempt$ = QPTrim$(EMP2Rec(1).EMPSOCX)
  MedExempt$ = QPTrim$(EMP2Rec(1).EMPMEDX)
  
  'calc social security W/H
  SELECT CASE SocExempt$
  CASE "N", ""
    TransRec.SocTaxAmt = Round#((TransRec.SocGrossPay * FedTax(1).FTMEMPSS) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FedTax(1).FTMEMPSS) * .01)
      TransRec.SocTaxAmt = Round#(TransRec.SocTaxAmt + TaxFring#)
    END IF
    TransRec.MatchSocAmt = Round#((TransRec.SocGrossPay * FedTax(1).FTMEMRSS) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FedTax(1).FTMEMRSS) * .01)
      TransRec.MatchSocAmt = Round#(TransRec.MatchSocAmt + TaxFring#)
    END IF
  CASE "Y"
    TransRec.SocTaxAmt = 0
    TransRec.SocGrossPay = 0
    TransRec.MatchSocAmt = 0
  END SELECT
  
  'calc medicare W/H
  SELECT CASE MedExempt$
  CASE "N", ""
    TransRec.MedTaxAmt = Round#((TransRec.MedGrossPay * FedTax(1).FTMEMPM) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FedTax(1).FTMEMPM) * .01)
      TransRec.MedTaxAmt = Round#(TransRec.MedTaxAmt + TaxFring#)
    END IF
    TransRec.MatchMedAmt = Round#((TransRec.MedGrossPay * FedTax(1).FTMEMRM) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FedTax(1).FTMEMRM) * .01)
      TransRec.MatchMedAmt = Round#(TransRec.MatchMedAmt + TaxFring#)
    END IF
  CASE "Y"
    TransRec.MedTaxAmt = 0
    TransRec.MedGrossPay = 0
    TransRec.MatchMedAmt = 0
  END SELECT
  
  '--------------------------------------------------------------------
  'Start of State and federal tax calc
  
  FedExempt$ = QPTrim$(EMP2Rec(1).EMPFEDX)
  StaExempt$ = QPTrim$(EMP2Rec(1).EMPSTAX)
  
  FOR Cnt = 1 TO 7
    IF EMP2Rec(1).EMPPFREQ = PayPFreq$(Cnt) THEN
      EXIT FOR
    END IF
  NEXT
  
  SELECT CASE Cnt
  CASE 1
    PayFreq = 52
  CASE 2
    PayFreq = 26
  CASE 3
    PayFreq = 24
  CASE 4
    PayFreq = 12
  CASE 5
    PayFreq = 4
  CASE 6
    PayFreq = 2
  CASE 7
    PayFreq = 1
  END SELECT
  
  SELECT CASE FedExempt$        'are they fed exempt?
  CASE "N", ""  'no they aren't
    SELECT CASE QPTrim$(EMP2Rec(1).EMPFEDO2)    'using fixed amount or percent?
    CASE ""     'no
      AnnualizedFedGross# = Round#(TransRec.FedGrossPay * PayFreq)
      GOSUB CalcFedTax
      '12-13-94 Added code to process taxable fringe field
      IF TransRec.TaxFring > 0 THEN
        PriorFedTax# = TransRec.FedTaxAmt
        'this variable has to remain this name
        'because it is used in the gosub to calcfedtax
        AnnualizedFedGross# = TransRec.TaxFring
        GOSUB CalcFedTax
        TransRec.FedTaxAmt = Round#(TransRec.FedTaxAmt + PriorFedTax#)
      END IF
    CASE "P"    'using fixed percent
      TransRec.FedTaxAmt = Round#(TransRec.FedGrossPay * (EMP2Rec(1).EMPFEDO1 * .01))
    CASE "A"    'using fixed amount
      TransRec.FedTaxAmt = EMP2Rec(1).EMPFEDO1
    END SELECT
  CASE "Y"      'yes they are fed tax exempt
    TransRec.FedTaxAmt = 0
    AnnualizedFedGross# = 0
  END SELECT
  
  'AnnualizedFedGross# or TransRec.FedTaxAmt
  
  SELECT CASE StaExempt$        'are they state exempt?
  CASE "N", ""  'no they aren't
    SELECT CASE QPTrim$(EMP2Rec(1).EMPSTAO2)    'using fixed amount or percent?
    CASE ""     'no
      AnnualizedStaGross# = Round#(TransRec.StaGrossPay * PayFreq)
      IF TransRec.TaxFring > 0 THEN
        AnnualizedStaGross# = Round#(AnnualizedStaGross# + TransRec.TaxFring)
      END IF
      GOSUB CalcStaTax
      '12-13-94 Added code to process taxable fringe field
      '          IF TransRec.TaxFring > 0 THEN
      '
      '            PriorStaTax# = TransRec.StaTaxAmt
      '            'this variable has to remain this name
      '            'because it is used in the gosub to calcStatax
      'JIM
      '
      '            AnnualizedStaGross# = Round#(TransRec.TaxFring * PayFreq)
      '
      ''            AnnualizedStaGross# = TransRec.TaxFring
      '
      '            GOSUB CalcStaTax
      '            TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + PriorStaTax#)
      '          END IF
    CASE "P"    'using fixed percent
      TransRec.StaTaxAmt = Round#(TransRec.StaGrossPay * (EMP2Rec(1).EMPSTAO1 * .01))
    CASE "A"    'using fixed amount
      TransRec.StaTaxAmt = EMP2Rec(1).EMPSTAO1
    END SELECT
  CASE "Y"      'yes they are fed tax exempt
    TransRec.StaTaxAmt = 0
    AnnualizedStaGross# = 0
  END SELECT
  
  TransRec.TotTaxAmt = Round#(TransRec.StaTaxAmt + TransRec.FedTaxAmt)
  TransRec.TotTaxAmt = Round#(TransRec.TotTaxAmt + TransRec.MedTaxAmt + TransRec.SocTaxAmt)
  TransRec.NetPay = Round#((TransRec.GrossPay - TransRec.TotTaxAmt) - TransRec.TotDedAmt)
  
  IF TransRec.NetPay < 0 THEN
    TransRec.TActive = False
  END IF
  
  'added EIC calc 06/06/1994
  'TransRec.NetPay
  EICGross# = Round(TransRec.FedGrossPay * PayFreq)
  
  SELECT CASE QPTrim$(EMP2Rec(1).EMPEIC)
  CASE "0", ""
    TransRec.EICAmt = 0
  CASE "1"
    SELECT CASE EICGross#
    CASE IS < EICRec(1).EIC(1).EIC1NVR0
      TransRec.EICAmt = Round((TransRec.FedGrossPay * EICRec(1).EIC(1).EIC1AMT0) * .01)
    CASE EICRec(1).EIC(1).EIC1OVR1 + 1 TO EICRec(1).EIC(1).EIC1NVR1
      TransRec.EICAmt = Round(EICRec(1).EIC(1).EIC1AMT1 / PayFreq)
    CASE IS > EICRec(1).EIC(1).EIC1OVR2
      WageDiff# = Round(EICGross# - EICRec(1).EIC(1).EIC1EXES)
      EXSDiff# = Round((WageDiff# * EICRec(1).EIC(1).EIC1LESS) * .01)
      TransRec.EICAmt = Round((EICRec(1).EIC(1).EIC1AMT1 - EXSDiff#) / PayFreq)
      IF TransRec.EICAmt < 1 THEN
        TransRec.EICAmt = 0
      END IF
    END SELECT
  CASE "2"
    SELECT CASE EICGross#
    CASE IS < EICRec(1).EIC(2).EIC1NVR0
      TransRec.EICAmt = Round((TransRec.FedGrossPay * EICRec(1).EIC(2).EIC1AMT0) * .01)
    CASE EICRec(1).EIC(2).EIC1OVR1 + 1 TO EICRec(1).EIC(2).EIC1NVR1
      TransRec.EICAmt = Round(EICRec(1).EIC(2).EIC1AMT1 / PayFreq)
    CASE IS > EICRec(1).EIC(2).EIC1OVR2
      WageDiff# = Round(EICGross# - EICRec(1).EIC(2).EIC1EXES)
      EXSDiff# = Round((WageDiff# * EICRec(1).EIC(2).EIC1LESS) * .01)
      TransRec.EICAmt = Round((EICRec(1).EIC(2).EIC1AMT1 - EXSDiff#) / PayFreq)
      IF TransRec.EICAmt < 1 THEN
        TransRec.EICAmt = 0
      END IF
    END SELECT
  END SELECT
  
  TransRec.NetPay = Round(TransRec.EICAmt + TransRec.NetPay)
  
  '12-10-94  this was to make sure that no gross can be less than 0
  '          corrects a bug in no reg wages alt-earnings non-taxable
  IF TransRec.FedGrossPay < 0 THEN TransRec.FedGrossPay = 0
  IF TransRec.StaGrossPay < 0 THEN TransRec.StaGrossPay = 0
  IF TransRec.SocGrossPay < 0 THEN TransRec.SocGrossPay = 0
  IF TransRec.MedGrossPay < 0 THEN TransRec.MedGrossPay = 0
  IF TransRec.RetGrossPay < 0 THEN TransRec.RetGrossPay = 0
  
  
  EXIT SUB
  
CalcFedTax:
  
  SELECT CASE QPTrim$(EMP2Rec(1).EMPFEDS)       'are they single or married?
  CASE "M"
    IF EMP2Rec(1).EMPFEDA < 0 THEN EMP2Rec(1).EMPFEDA = 0
    TaxableAmtFed# = Round#(AnnualizedFedGross# - (EMP2Rec(1).EMPFEDA * FedTax(1).FTMSDAA))
    IF TaxableAmtFed# < FedTax(1).FTM(3, 1) THEN
      TransRec.FedTaxAmt = 0
    ELSE
      FOR Cnt = 1 TO 10
        IF FedTax(1).FTM(3, Cnt) > TaxableAmtFed# THEN EXIT FOR
      NEXT
      Cnt = Cnt - 1
      TaxableAmtFed# = Round#(TaxableAmtFed# - FedTax(1).FTM(3, Cnt))
      TransRec.FedTaxAmt = Round#(Round#(FedTax(1).FTM(1, Cnt) + (TaxableAmtFed# * (FedTax(1).FTM(2, Cnt) * .01))) / PayFreq)
    END IF
  CASE "S", ""
    IF EMP2Rec(1).EMPFEDA < 0 THEN EMP2Rec(1).EMPFEDA = 0
    TaxableAmtFed# = Round#(AnnualizedFedGross# - (EMP2Rec(1).EMPFEDA * FedTax(1).FTSSDAA))
    IF TaxableAmtFed# < FedTax(1).FTS(3, 1) THEN
      TransRec.FedTaxAmt = 0
    ELSE
      FOR Cnt = 1 TO 10
        IF FedTax(1).FTS(3, Cnt) > TaxableAmtFed# THEN EXIT FOR
      NEXT
      Cnt = Cnt - 1
      TaxableAmtFed# = Round#(TaxableAmtFed# - FedTax(1).FTS(3, Cnt))
      TransRec.FedTaxAmt = Round#(Round#(FedTax(1).FTS(1, Cnt) + (TaxableAmtFed# * (FedTax(1).FTS(2, Cnt) * .01))) / PayFreq)
    END IF
  END SELECT
  
  '***02-29-96  Fixed bug that caused a doubling of the additional
  '             taxes withheld, if taxable fringe was included.
  '             Also fixed in state calc of taxes.
  IF EMP2Rec(1).EMPFEDAA > 0 AND PriorFedTax# = 0 THEN
    TransRec.FedTaxAmt = Round#(TransRec.FedTaxAmt + EMP2Rec(1).EMPFEDAA)
  END IF
  
  RETURN
  
CalcStaTax:
  
  SELECT CASE TaxText$(0)
  CASE "OK"
    '$INCLUDE: 'OKstatax.bi'
    
  CASE "NC", "VA", "AL"
    '$INCLUDE: 'ncstatax.bi'
    
  CASE "SC"
    '$INCLUDE: 'SCstatax.bi'
    
  END SELECT
  
  
  
  RETURN
  
END SUB

SUB CreateEmpTransRecs (RecNo)
  
  'don't make a transaction for a terminated employee
  IF EMP2Rec(1).EMPTDATE > 0 OR EMP2Rec(1).Deleted = True THEN
    EXIT SUB
  END IF
  
  REDIM TempTransRec(1) AS TransRecType
  
  TRecSize = LEN(TempTransRec(1))
  
  TempTransRec(1).TActive = True
  
  PayType$ = UCASE$(QPTrim$(EMP2Rec(1).EMPPTYPE))
  PayFreq$ = UCASE$(QPTrim$(EMP2Rec(1).EMPPFREQ))
  
  TempTransRec(1).EmpPin = EMP2Rec(1).EmpPin
  TempTransRec(1).PayPdStart = PrdDefRec(1).PERBEG
  TempTransRec(1).PayPdEnd = PrdDefRec(1).PEREND
  
  'adjust and copy base rate
  IF EMP2Rec(1).EMPPRATE < 0 THEN
    TempTransRec(1).BaseRate = 0
  ELSE
    TempTransRec(1).BaseRate = EMP2Rec(1).EMPPRATE
  END IF
  
  'adjust and copy overtime rate
  IF EMP2Rec(1).EMPORATE < 0 THEN
    TempTransRec(1).OTRate = 0
  ELSE
    TempTransRec(1).OTRate = EMP2Rec(1).EMPORATE
  END IF
  
  IF PrdDefRec(1).USEAE1 = "Y" THEN
    IF EMP2Rec(1).EMPEAMT1 < 0 THEN TempTransRec(1).EAmt(1) = 0 ELSE TempTransRec(1).EAmt(1) = EMP2Rec(1).EMPEAMT1
    TempTransRec(1).EDist(1).EAmt = TempTransRec(1).EAmt(1)
    TempTransRec(1).EDist(1).EAcct = EMP2Rec(1).EMPEACT1
    'ELSE
    'help, I'm a duck
    
  END IF
  IF PrdDefRec(1).USEAE2 = "Y" THEN
    IF EMP2Rec(1).EMPEAMT2 < 0 THEN TempTransRec(1).EAmt(2) = 0 ELSE TempTransRec(1).EAmt(2) = EMP2Rec(1).EMPEAMT2
    TempTransRec(1).EDist(2).EAmt = TempTransRec(1).EAmt(2)
    TempTransRec(1).EDist(2).EAcct = EMP2Rec(1).EMPEACT2
  END IF
  IF PrdDefRec(1).USEAE3 = "Y" THEN
    IF EMP2Rec(1).EMPEAMT3 < 0 THEN TempTransRec(1).EAmt(3) = 0 ELSE TempTransRec(1).EAmt(3) = EMP2Rec(1).EMPEAMT3
    TempTransRec(1).EDist(3).EAmt = TempTransRec(1).EAmt(3)
    TempTransRec(1).EDist(3).EAcct = EMP2Rec(1).EMPEACT3
  END IF
  
  
  TempTransRec(1).TotAdditEarn = TempTransRec(1).EAmt(1) + TempTransRec(1).EAmt(2) + TempTransRec(1).EAmt(3)
  
  SELECT CASE PayType$
  CASE "HOURLY"
    TempTransRec(1).PayType = "H"
    FOR Cnt = 1 TO 8
      IF EMP2Rec(1).EDist(Cnt).DAmt < 0 THEN
        TempTransRec(1).TDist(Cnt).DRHrs = 0
      ELSE
        TempTransRec(1).TDist(Cnt).DRHrs = EMP2Rec(1).EDist(Cnt).DAmt
        TempTransRec(1).RegHrsWork = Round#(TempTransRec(1).RegHrsWork + TempTransRec(1).TDist(Cnt).DRHrs)
      END IF
      TempTransRec(1).TDist(Cnt).DAcct = EMP2Rec(1).EDist(Cnt).DAcct
    NEXT
    TempTransRec(1).RegHrsPaid = TempTransRec(1).RegHrsWork
  CASE "SALARIED"
    TempTransRec(1).PayType = "S"
    FOR Cnt = 1 TO 8
      IF EMP2Rec(1).EDist(Cnt).DAmt < 0 THEN
        TempTransRec(1).TDist(Cnt).DPct = 0
      ELSE
        TempTransRec(1).TDist(Cnt).DPct = EMP2Rec(1).EDist(Cnt).DAmt
      END IF
      TempTransRec(1).TDist(Cnt).DAcct = EMP2Rec(1).EDist(Cnt).DAcct
    NEXT
  END SELECT
  
  CalcPay TempTransRec(1), RecNo, False
  
  FPutRTA THandle, TempTransRec(1), CLNG(RecNo), TRecSize
  
  
END SUB

SUB MakeDefaultTransActs
  
  DisplayMiscScrn CreatingPayRoll
  EmpRecSize = LEN(EMP2Rec(1))
  
  NumOfRecs = FileSize(EmpData2Name) \ EmpRecSize
  
  FOpenS TransWorkFileName, THandle             'open it
  FOpenS EmpData2Name, DHandle  'open employee data file
  
  FOR RecNo = 1 TO NumOfRecs
    FGetRTA DHandle, EMP2Rec(1), CLNG(RecNo), EmpRecSize
    
    PayType$ = UCASE$(QPTrim$(EMP2Rec(1).EMPPFREQ))
    
    SELECT CASE PayType$
    CASE "WEEKLY"
      IF PrdDefRec(1).PAYWK = "Y" THEN MakeTrans = True
    CASE "BI-WEEKLY"
      IF PrdDefRec(1).PAYBIWK = "Y" THEN MakeTrans = True
    CASE "SEMI-MONTHLY"
      IF PrdDefRec(1).PAYSEMIM = "Y" THEN MakeTrans = True
    CASE "MONTHLY"
      IF PrdDefRec(1).PAYMO = "Y" THEN MakeTrans = True
    CASE "QUARTERLY"
      IF PrdDefRec(1).PAYQTR = "Y" THEN MakeTrans = True
    CASE "SEMI-ANNUALLY"
      IF PrdDefRec(1).PAYSEMIA = "Y" THEN MakeTrans = True
    CASE "ANNUALLY"
      IF PrdDefRec(1).PAYANNL = "Y" THEN MakeTrans = True
    END SELECT
    IF MakeTrans THEN
      CreateEmpTransRecs RecNo
      MakeTrans = False
    END IF
    PctC$ = STR$(INT((RecNo / NumOfRecs) * 100))
    QPrintRC PctC$, 12, 34, 112
  NEXT
  
  FClose THandle
  FClose DHandle
  
END SUB

SUB MakeTransInActive
  
  REDIM TTran(1) AS TransRecType
  TTranSize = LEN(TTran(1))
  NumOfRecs = FileSize(TransWorkFileName) \ TTranSize
  FOpenS TransWorkFileName, TTHandle            'open it
  
  FOR Cnt = 1 TO NumOfRecs
    RecNo& = Cnt
    FGetRTA TTHandle, TTran(1), RecNo&, TTranSize
    TTran(1).TActive = False
    FPutRTA TTHandle, TTran(1), RecNo&, TTranSize
  NEXT
  
  FClose TTHandle
END SUB

SUB ManualMenu

  PCLoadPayFreqs
  PCLoadSystemFiles
  

  DO
    PayMenu 4, Choice, 3
    
    SELECT CASE Choice
    CASE 1
      EntryType = manual
      PCGetEmpNum LastEmpNum
    CASE 2
      EntryType = manual
      'PCPrintManRegister             'unrem
    CASE 3
      EntryType = manual
      PostTransactions
    END SELECT
  LOOP UNTIL Choice = EscKey
  Choice = 0
  
END SUB

SUB ParseHourly2Trans
  
  FOR Cnt = 1 TO 8
    TransRec(1).TDist(Cnt).DAcct = HourInput(1).HDist(Cnt).DAcct
    TransRec(1).TDist(Cnt).DRHrs = HourInput(1).HDist(Cnt).DRHrs
    TransRec(1).TDist(Cnt).DOHrs = HourInput(1).HDist(Cnt).DOHrs
  NEXT
  
  TransRec(1).RegHrsWork = HourInput(1).WORKHRS
  TransRec(1).VacUsed = HourInput(1).VACHRS
  TransRec(1).SickUsed = HourInput(1).SICKHRS
  TransRec(1).HOLHOURS = HourInput(1).HOLHRS
  
  TransRec(1).CompUsed = HourInput(1).COMPHRS
  
  TransRec(1).RegHrsPaid = HourInput(1).TOTHRSPD
  
  TransRec(1).OTHours = HourInput(1).OTWORKED
  TransRec(1).OTHrsPaid = HourInput(1).OTHRSPD
  TransRec(1).OT2COMP = HourInput(1).OT2COMP
  
  
  TransRec(1).EAmt(1) = HourInput(1).ALTEARN1
  TransRec(1).EAmt(2) = HourInput(1).ALTEARN2
  TransRec(1).EAmt(3) = HourInput(1).ALTEARN3
  TransRec(1).TotAdditEarn = HourInput(1).TOTAERN               'this is sum of (alt earn) fields
  
  TransRec(1).EDist(1).EAcct = HourInput(1).AERNDST1
  TransRec(1).EDist(2).EAcct = HourInput(1).AERNDST2
  TransRec(1).EDist(3).EAcct = HourInput(1).AERNDST3
  TransRec(1).EDist(4).EAcct = HourInput(1).AERNDST4
  TransRec(1).EDist(5).EAcct = HourInput(1).AERNDST5
  TransRec(1).EDist(6).EAcct = HourInput(1).AERNDST6
  
  TransRec(1).EDist(1).EAmt = HourInput(1).AERNAMT1
  TransRec(1).EDist(2).EAmt = HourInput(1).AERNAMT2
  TransRec(1).EDist(3).EAmt = HourInput(1).AERNAMT3
  TransRec(1).EDist(4).EAmt = HourInput(1).AERNAMT4
  TransRec(1).EDist(5).EAmt = HourInput(1).AERNAMT5
  TransRec(1).EDist(6).EAmt = HourInput(1).AERNAMT6
  
  TransRec(1).TaxFring = HourInput(1).TaxFring
END SUB

SUB ParseManual2Trans (ManTrans AS ManualTransRecType)
  'STOP
  TransRec(1).PayPdStart = ManTrans.PDSTART
  TransRec(1).PayPdEnd = ManTrans.PDEND
  
  TransRec(1).CheckDate = ManTrans.ChkDate
  
  TransRec(1).CheckNum = ManTrans.CheckNum
  'STOP
  
  TransRec(1).RegHrsWork = ManTrans.REGHRS
  TransRec(1).RegHrsPaid = TransRec(1).RegHrsWork
  
  TransRec(1).SickUsed = ManTrans.SICKHRS
  TransRec(1).VacUsed = ManTrans.VACHRS
  TransRec(1).CompUsed = ManTrans.COMPHRS
  TransRec(1).HOLHOURS = ManTrans.HOLHOURS
  
  TransRec(1).OTHrsPaid = ManTrans.OTHRSPD
  
  TransRec(1).TotRegWage = ManTrans.REGWAGE
  TransRec(1).TotOTWage = ManTrans.OTWAGE
  
  TransRec(1).TDist(1).DAcct = ManTrans.DISTACT1
  TransRec(1).TDist(1).DRWage = ManTrans.WAGEAMT1
  
  TransRec(1).TDist(2).DAcct = ManTrans.DISTACT2
  TransRec(1).TDist(2).DRWage = ManTrans.WAGEAMT2
  
  TransRec(1).TDist(3).DAcct = ManTrans.DISTACT3
  TransRec(1).TDist(3).DRWage = ManTrans.WAGEAMT3
  
  TransRec(1).TDist(4).DAcct = ManTrans.DISTACT4
  TransRec(1).TDist(4).DRWage = ManTrans.WAGEAMT4
  
  TransRec(1).GrossPay = ManTrans.GrossPay
  
  TransRec(1).FedTaxAmt = ManTrans.FedTax
  TransRec(1).StaTaxAmt = ManTrans.StaTax
  TransRec(1).SocTaxAmt = ManTrans.SocTax
  TransRec(1).MedTaxAmt = ManTrans.MedTax
  TransRec(1).RetireAmt = ManTrans.RETAMT
  TransRec(1).TotTaxAmt = ManTrans.TOTTAX
  
  FOR Cnt = 1 TO 12
    TransRec(1).DAmt(Cnt) = ManTrans.DAmt(Cnt)
  NEXT
  
  TransRec(1).TotDedAmt = ManTrans.TOTDED
  TransRec(1).EICAmt = ManTrans.EIC
  
  TransRec(1).NetPay = ManTrans.NetPay
  
  TransRec(1).FedGrossPay = ManTrans.FEDGROSS
  TransRec(1).StaGrossPay = ManTrans.STAGROSS
  TransRec(1).SocGrossPay = ManTrans.SOCGROSS
  TransRec(1).MedGrossPay = ManTrans.MEDGROSS
  TransRec(1).RetGrossPay = ManTrans.RETGROSS
  
  'IF ManTrans.RETAMT < 0 THEN
  '  TransRec(1).RetGrossPay=
END SUB

SUB ParseSalary2Trans
  '1
  TransRec(1).PaySFlag = SalInput(1).PAYSAL
  
  TransRec(1).VacUsed = SalInput(1).VACHRS
  TransRec(1).SickUsed = SalInput(1).SICKHRS
  
  TransRec(1).HOLHOURS = SalInput(1).HOLHRS
  
  FOR Cnt = 1 TO 8
    TransRec(1).TDist(Cnt).DAcct = SalInput(1).SDist(Cnt).DAcct
    TransRec(1).TDist(Cnt).DPct = SalInput(1).SDist(Cnt).DPct
  NEXT
  
  TransRec(1).EAmt(1) = SalInput(1).ALTEARN1
  TransRec(1).EAmt(2) = SalInput(1).ALTEARN2
  TransRec(1).EAmt(3) = SalInput(1).ALTEARN3
  
  TransRec(1).EDist(1).EAmt = SalInput(1).AERNAMT1
  TransRec(1).EDist(2).EAmt = SalInput(1).AERNAMT2
  TransRec(1).EDist(3).EAmt = SalInput(1).AERNAMT3
  TransRec(1).EDist(4).EAmt = SalInput(1).AERNAMT4
  TransRec(1).EDist(5).EAmt = SalInput(1).AERNAMT5
  TransRec(1).EDist(6).EAmt = SalInput(1).AERNAMT6
  
  TransRec(1).EDist(1).EAcct = SalInput(1).AERNDST1
  TransRec(1).EDist(2).EAcct = SalInput(1).AERNDST2
  TransRec(1).EDist(3).EAcct = SalInput(1).AERNDST3
  TransRec(1).EDist(4).EAcct = SalInput(1).AERNDST4
  TransRec(1).EDist(5).EAcct = SalInput(1).AERNDST5
  TransRec(1).EDist(6).EAcct = SalInput(1).AERNDST6
  
  TransRec(1).TotAdditEarn = SalInput(1).TOTAERN
  TransRec(1).TaxFring = SalInput(1).TaxFring
  
END SUB

SUB ParseScrnCalc2Trans
  
  
  TransRec(1).TotRegWage = ScrnCalc(1).REGEARN
  TransRec(1).TotOTWage = ScrnCalc(1).OTEARN
  
  TransRec(1).EAmt(1) = ScrnCalc(1).ALTEARN1
  TransRec(1).EAmt(2) = ScrnCalc(1).ALTEARN2
  TransRec(1).EAmt(3) = ScrnCalc(1).ALTEARN3
  
  TransRec(1).SocTaxAmt = ScrnCalc(1).SocTax
  TransRec(1).MedTaxAmt = ScrnCalc(1).MedTax
  TransRec(1).FedTaxAmt = ScrnCalc(1).FedTax
  TransRec(1).StaTaxAmt = ScrnCalc(1).StaTax
  
  TransRec(1).RetireAmt = ScrnCalc(1).RETIRE
  
  TransRec(1).DAmt(1) = ScrnCalc(1).DED1
  TransRec(1).DAmt(2) = ScrnCalc(1).DED2
  TransRec(1).DAmt(3) = ScrnCalc(1).DED3
  TransRec(1).DAmt(4) = ScrnCalc(1).DED4
  TransRec(1).DAmt(5) = ScrnCalc(1).DED5
  
  TransRec(1).DAmt(6) = ScrnCalc(1).DED6
  TransRec(1).DAmt(7) = ScrnCalc(1).DED7
  TransRec(1).DAmt(8) = ScrnCalc(1).DED8
  TransRec(1).DAmt(9) = ScrnCalc(1).DED9
  TransRec(1).DAmt(10) = ScrnCalc(1).DED10
  TransRec(1).DAmt(11) = ScrnCalc(1).DED11
  TransRec(1).DAmt(12) = ScrnCalc(1).DED12
  
  TransRec(1).GrossPay = ScrnCalc(1).GrossPay
  
  TransRec(1).TotDedAmt = 0
  FOR Cnt = 1 TO 12
    TransRec(1).TotDedAmt = Round#(TransRec(1).TotDedAmt + TransRec(1).DAmt(Cnt))
  NEXT
  
  'fix from region-d
  TransRec(1).TotDedAmt = Round#(TransRec(1).TotDedAmt + TransRec(1).RetireAmt)
  
  TransRec(1).EICAmt = ScrnCalc(1).EIC
  TransRec(1).NetPay = ScrnCalc(1).NetPay
  
  
END SUB

SUB ParseTrans2Hourly
  
  FOR Cnt = 1 TO 8
    HourInput(1).HDist(Cnt).DAcct = TransRec(1).TDist(Cnt).DAcct
    HourInput(1).HDist(Cnt).DRHrs = TransRec(1).TDist(Cnt).DRHrs
    HourInput(1).HDist(Cnt).DOHrs = TransRec(1).TDist(Cnt).DOHrs
  NEXT
  
  HourInput(1).WORKHRS = TransRec(1).RegHrsWork
  HourInput(1).VACHRS = TransRec(1).VacUsed
  HourInput(1).SICKHRS = TransRec(1).SickUsed
  HourInput(1).HOLHRS = TransRec(1).HOLHOURS
  HourInput(1).COMPHRS = TransRec(1).CompUsed
  HourInput(1).TOTHRSPD = TransRec(1).RegHrsPaid
  
  
  HourInput(1).OTWORKED = TransRec(1).OTHours
  HourInput(1).OTHRSPD = TransRec(1).OTHrsPaid
  HourInput(1).OT2COMP = TransRec(1).OT2COMP
  
  HourInput(1).ALTEARN1 = TransRec(1).EAmt(1)   '*
  HourInput(1).ALTEARN2 = TransRec(1).EAmt(2)   'actual earning amounts
  HourInput(1).ALTEARN3 = TransRec(1).EAmt(3)   '*
  
  HourInput(1).AERNDST1 = TransRec(1).EDist(1).EAcct            '*
  HourInput(1).AERNDST2 = TransRec(1).EDist(2).EAcct            '
  HourInput(1).AERNDST3 = TransRec(1).EDist(3).EAcct            'Earnings distribution accts.
  HourInput(1).AERNDST4 = TransRec(1).EDist(4).EAcct            '
  HourInput(1).AERNDST5 = TransRec(1).EDist(5).EAcct            '
  HourInput(1).AERNDST6 = TransRec(1).EDist(6).EAcct            '*
  
  HourInput(1).AERNAMT1 = TransRec(1).EDist(1).EAmt             '*
  HourInput(1).AERNAMT2 = TransRec(1).EDist(2).EAmt             '
  HourInput(1).AERNAMT3 = TransRec(1).EDist(3).EAmt             'earnings amounts as distributed
  HourInput(1).AERNAMT4 = TransRec(1).EDist(4).EAmt             'to accts.
  HourInput(1).AERNAMT5 = TransRec(1).EDist(5).EAmt             '
  HourInput(1).AERNAMT6 = TransRec(1).EDist(6).EAmt             '*
  
  HourInput(1).TOTAERN = TransRec(1).TotAdditEarn
  HourInput(1).TaxFring = TransRec(1).TaxFring
  
END SUB

SUB ParseTrans2Manual (ManTrans AS ManualTransRecType)
  
  ManTrans.PDSTART = TransRec(1).PayPdStart
  ManTrans.PDEND = TransRec(1).PayPdEnd
  ManTrans.ChkDate = TransRec(1).CheckDate
  ManTrans.CheckNum = TransRec(1).CheckNum
  
  ManTrans.REGHRS = TransRec(1).RegHrsWork
  ManTrans.SICKHRS = TransRec(1).SickUsed
  ManTrans.VACHRS = TransRec(1).VacUsed
  ManTrans.COMPHRS = TransRec(1).CompUsed
  ManTrans.HOLHOURS = TransRec(1).HOLHOURS
  
  ManTrans.OTHRSPD = TransRec(1).OTHrsPaid
  
  ManTrans.REGWAGE = TransRec(1).TotRegWage
  ManTrans.OTWAGE = TransRec(1).TotOTWage
  
  ManTrans.DISTACT1 = TransRec(1).TDist(1).DAcct
  ManTrans.WAGEAMT1 = TransRec(1).TDist(1).DRWage
  
  ManTrans.DISTACT2 = TransRec(1).TDist(2).DAcct
  ManTrans.WAGEAMT2 = TransRec(1).TDist(2).DRWage
  
  ManTrans.DISTACT3 = TransRec(1).TDist(3).DAcct
  ManTrans.WAGEAMT3 = TransRec(1).TDist(3).DRWage
  ManTrans.DISTACT4 = TransRec(1).TDist(4).DAcct
  ManTrans.WAGEAMT4 = TransRec(1).TDist(4).DRWage
  
  ManTrans.GrossPay = TransRec(1).GrossPay
  
  ManTrans.FedTax = TransRec(1).FedTaxAmt
  ManTrans.StaTax = TransRec(1).StaTaxAmt
  ManTrans.SocTax = TransRec(1).SocTaxAmt
  ManTrans.MedTax = TransRec(1).MedTaxAmt
  ManTrans.RETAMT = TransRec(1).RetireAmt
  ManTrans.TOTTAX = TransRec(1).TotTaxAmt
  
  FOR Cnt = 1 TO 12
    ManTrans.DAmt(Cnt) = TransRec(1).DAmt(Cnt)
  NEXT
  ManTrans.TOTDED = TransRec(1).TotDedAmt
  ManTrans.EIC = TransRec(1).EICAmt
  ManTrans.NetPay = TransRec(1).NetPay
  
  ManTrans.FEDGROSS = TransRec(1).FedGrossPay
  ManTrans.STAGROSS = TransRec(1).StaGrossPay
  ManTrans.SOCGROSS = TransRec(1).SocGrossPay
  ManTrans.MEDGROSS = TransRec(1).MedGrossPay
  ManTrans.RETGROSS = TransRec(1).RetGrossPay
  
END SUB

SUB ParseTrans2Salary
  
  IF LEN(QPTrim$(TransRec(1).PaySFlag)) = 0 THEN
    SalInput(1).PAYSAL = "Y"
  ELSE
    SalInput(1).PAYSAL = TransRec(1).PaySFlag
  END IF
  
  SalInput(1).VACHRS = TransRec(1).VacUsed
  SalInput(1).SICKHRS = TransRec(1).SickUsed
  SalInput(1).HOLHRS = TransRec(1).HOLHOURS
  
  FOR Cnt = 1 TO 8
    SalInput(1).SDist(Cnt).DAcct = TransRec(1).TDist(Cnt).DAcct
    SalInput(1).SDist(Cnt).DPct = TransRec(1).TDist(Cnt).DPct
  NEXT
  
  SalInput(1).ALTEARN1 = TransRec(1).EAmt(1)
  SalInput(1).ALTEARN2 = TransRec(1).EAmt(2)
  SalInput(1).ALTEARN3 = TransRec(1).EAmt(3)
  
  SalInput(1).AERNAMT1 = TransRec(1).EDist(1).EAmt
  SalInput(1).AERNAMT2 = TransRec(1).EDist(2).EAmt
  SalInput(1).AERNAMT3 = TransRec(1).EDist(3).EAmt
  SalInput(1).AERNAMT4 = TransRec(1).EDist(4).EAmt
  SalInput(1).AERNAMT5 = TransRec(1).EDist(5).EAmt
  SalInput(1).AERNAMT6 = TransRec(1).EDist(6).EAmt
  
  SalInput(1).AERNDST1 = TransRec(1).EDist(1).EAcct
  SalInput(1).AERNDST2 = TransRec(1).EDist(2).EAcct
  SalInput(1).AERNDST3 = TransRec(1).EDist(3).EAcct
  SalInput(1).AERNDST4 = TransRec(1).EDist(4).EAcct
  SalInput(1).AERNDST5 = TransRec(1).EDist(5).EAcct
  SalInput(1).AERNDST6 = TransRec(1).EDist(6).EAcct
  
  SalInput(1).TOTAERN = TransRec(1).TotAdditEarn
  SalInput(1).TaxFring = TransRec(1).TaxFring
  
END SUB

SUB ParseTrans2ScrnCalc
  'STOP
  ScrnCalc(1).REGEARN = TransRec(1).TotRegWage
  ScrnCalc(1).OTEARN = TransRec(1).TotOTWage
  
  ScrnCalc(1).ALTEARN1 = TransRec(1).EAmt(1)
  ScrnCalc(1).ALTEARN2 = TransRec(1).EAmt(2)
  ScrnCalc(1).ALTEARN3 = TransRec(1).EAmt(3)
  
  ScrnCalc(1).SocTax = TransRec(1).SocTaxAmt
  ScrnCalc(1).MedTax = TransRec(1).MedTaxAmt
  ScrnCalc(1).FedTax = TransRec(1).FedTaxAmt
  ScrnCalc(1).StaTax = TransRec(1).StaTaxAmt
  ScrnCalc(1).RETIRE = TransRec(1).RetireAmt
  
  ScrnCalc(1).DED1 = TransRec(1).DAmt(1)
  ScrnCalc(1).DED2 = TransRec(1).DAmt(2)
  ScrnCalc(1).DED3 = TransRec(1).DAmt(3)
  ScrnCalc(1).DED4 = TransRec(1).DAmt(4)
  ScrnCalc(1).DED5 = TransRec(1).DAmt(5)
  ScrnCalc(1).DED6 = TransRec(1).DAmt(6)
  ScrnCalc(1).DED7 = TransRec(1).DAmt(7)
  ScrnCalc(1).DED8 = TransRec(1).DAmt(8)
  ScrnCalc(1).DED9 = TransRec(1).DAmt(9)
  ScrnCalc(1).DED10 = TransRec(1).DAmt(10)
  ScrnCalc(1).DED11 = TransRec(1).DAmt(11)
  ScrnCalc(1).DED12 = TransRec(1).DAmt(12)
  
  ScrnCalc(1).GrossPay = TransRec(1).GrossPay
  
  ScrnCalc(1).TOTDED = Round#(TransRec(1).TotDedAmt + TransRec(1).TotTaxAmt)
  
  ScrnCalc(1).EIC = TransRec(1).EICAmt
  ScrnCalc(1).NetPay = TransRec(1).NetPay
  
  
END SUB

FUNCTION PCDelFromPay
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  DisplayMiscScrn DelFromPay
  DO
    DO
      Ky$ = INKEY$
      ButtonPress 1, 0, MouseButton, XPos, YPos
    LOOP UNTIL LEN(Ky$) OR MouseButton
    IF MouseButton THEN
      MRow = (YPos \ 8) + 1
      MCol = (XPos \ 8)         '14 30
      SELECT CASE MRow
      CASE 16
        SELECT CASE MCol
        CASE 30 TO 40
          PressButton F3Key, 16, 30, 40
        CASE 42 TO 53
          PressButton EscKey, 16, 42, 53
        CASE ELSE
        END SELECT
        Ky$ = INKEY$
      CASE ELSE
      END SELECT
    END IF
    
    SELECT CASE Ky$
    CASE CHR$(0) + "="
      DelPayFlag = -1
      ExitFlag = -1
    CASE CHR$(27)
      DelPayFlag = 0
      ExitFlag = -1
    END SELECT
    
  LOOP UNTIL ExitFlag
  
  PCDelFromPay = DelPayFlag
  
  RestScrn TempScrn()
  ERASE TempScrn
  
END FUNCTION

SUB PCEdScrnCalc (TransRecNo, EmpName$, EmpRecNo)
  
  REDIM TempScrn(0)
  REDIM FrmESC(1) AS FormInfo
  REDIM ScrnCalc(1)      AS ScrnCalcType
  REDIM OldFldVal(6 TO 17) AS DOUBLE
  REDIM EarnDes$(1 TO 3)
  REDIM DedDes$(1 TO 12)
  
  ScrnRetFld = 22
  
  TransRecLen = LEN(TransRec(1))
  
  FormName$ = "SCRNCALC"
  
  NumFlds = LibNumberOfFields(CalcQLib, FormName$)
  
  REDIM FormESC$(NumFlds, 2)    'DIM the form data array
  REDIM FldESC(NumFlds) AS FieldInfo            'DIM the field information array
  
  StartEl = 0   'Load first form at array start
  LibGetFldDef CalcQLib, FormName$, StartEl, FldESC(), FormESC$(), ErrCode
  
  PCGetEmp2Rec EmpRecNo, ErrorCode
  
  IF TransRec(1).TActive = False THEN
    CalcPay TransRec(1), TransRecNo, False
  ELSE
    CalcPay TransRec(1), TransRecNo, True
  END IF
  
  FormESC$(0, 0) = SPACE$(LEN(ScrnCalc(1)))
  
  ParseTrans2ScrnCalc
  
  BCopy VARSEG(ScrnCalc(1)), VARPTR(ScrnCalc(1)), SSEG(FormESC$(0, 0)), SADD(FormESC$(0, 0)), LEN(ScrnCalc(1)), 0
  
  UnPackBuffer 0, 0, FormESC$(), FldESC()
  
  FOR Cnt = 6 TO 17
    OldFldVal(Cnt) = Value#(FormESC$(Cnt, 0), ErrCode)
  NEXT
  
  OldRetAmt# = Value#(FormESC$(ScrnRetFld, 0), ErrCode)
  
  IF TransRec(1).PayType = "H" THEN
    FldESC(1).Protected = True
  END IF
  
  FOR Cnt = 1 TO 3
    EarnDes$(Cnt) = QPTrim$(ErnCodes(Cnt).ERNCODE1)
  NEXT
  FOR Cnt = 1 TO 12
    DedDes$(Cnt) = QPTrim$(DedCodes(Cnt).DCDESC1)
  NEXT
  
  DO
    LibFile2Scrn CalcQLib, FormName$, MonoCode, -1, ErrCode
    GOSUB DisplayInfoSC
    DO          '"Pole" the editing procedure
      EditForm FormESC$(), FldESC(), FrmESC(1), Cnf, Action
      IF FrmESC(1).FldEdited AND FrmESC(1).PrevFld = 1 AND FrmESC(1).FldNo > 1 THEN
        GOSUB ReCalcPay
        OldRetAmt# = Value#(FormESC$(ScrnRetFld, 0), ErrCode)
      END IF
      
      SELECT CASE FrmESC(1).PrevFld
      CASE 6 TO 17
        IF FrmESC(1).FldNo <> FrmESC(1).PrevFld AND OldFldVal(FrmESC(1).PrevFld) <> Value#(FormESC$(FrmESC(1).PrevFld, 0), ErrCode) THEN
          GOSUB ReCalcPay
          OldFldVal(FrmESC(1).PrevFld) = Value#(FormESC$(FrmESC(1).PrevFld, 0), ErrCode)
        END IF
      END SELECT
      
      IF FrmESC(1).Presses AND FrmESC(1).MRow = 25 THEN
        SELECT CASE FrmESC(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 26 TO 37           'F3
          PressButton F3Key, 25, 26, 37
        CASE 53 TO 66           'F10
          PressButton F0Key, 25, 53, 66
        END SELECT
      END IF
      
      IF FrmESC(1).KeyCode THEN
        SELECT CASE FldESC(FrmESC(1).PrevFld).FType
        CASE 3, 4, 5, 10, 18
          IF QPTrim$(FormESC$(FrmESC(1).PrevFld, 0)) = "" THEN LSET FormESC$(FrmESC(1).FldNo, 0) = "0"
        END SELECT
      END IF
      
      SELECT CASE FrmESC(1).KeyCode
      CASE F0Key
        IF FrmESC(1).KeyCode = EscKey AND FrmESC(1).Edited THEN
          SaveFlag = PromptSaveData
        ELSEIF FrmESC(1).KeyCode = F0Key THEN
          IF Value#(FormESC$(NumFlds, 0), ErrCode) < 0 THEN
            SaveScrn TempScrn()
            DisplayMiscScrn NegativePay
            WaitForAction
            RestScrn TempScrn()
            SaveFlag = 1
          ELSE
            SaveFlag = True
          END IF
        END IF
        SELECT CASE SaveFlag
        CASE True               'yep save data to disk
          ExitFlag = True
          BCopy SSEG(FormESC$(0, 0)), SADD(FormESC$(0, 0)), VARSEG(ScrnCalc(1)), VARPTR(ScrnCalc(1)), LEN(ScrnCalc(1)), 0
          ParseScrnCalc2Trans   'TransRec', ScrnCalc(1)
          '10-01-94 03 refixed   Added fix to adjust fed gross and state gross
          'if ret amt was edited
          IF OldRetAmt# <> Value#(FormESC$(ScrnRetFld, 0), ErrCode) THEN
            DifRetAmt# = OldRetAmt# - Value#(FormESC$(ScrnRetFld, 0), ErrCode)
            TransRec(1).FedGrossPay = Round(TransRec(1).FedGrossPay + DifRetAmt#)
            TransRec(1).StaGrossPay = Round(TransRec(1).StaGrossPay + DifRetAmt#)
          END IF
          TransRec(1).TActive = True
          UpDateTransFile TransRecNo
        CASE False              'nope don't save abandon edit
          ExitFlag = True
        CASE 1  'oops continue editing
          Action = 2
          ExitFlag = False
        END SELECT
      CASE F3Key
        IF PCDelFromPay THEN
          TransRec(1).TActive = 0
          UpDateTransFile TransRecNo
          ExitFlag = True
        END IF
      CASE ELSE
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
  ERASE TempScrn, FormESC$, FldESC
  
  EXIT SUB
  
DisplayInfoSC:
  
  TempE$ = SPACE$(10)
  HideCursor
  QPrintRC EmpName$, 1, 38, -1
  FOR Cnt = 1 TO 3
    RSET TempE$ = EarnDes$(Cnt)
    QPrintRC TempE$, 7, 21 + (Cnt * 15), -1
  NEXT
  OffSet = 4
  FOR Cnt = 1 TO 6
    RSET TempE$ = DedDes$(Cnt)
    QPrintRC TempE$, 10, OffSet, -1
    OffSet = OffSet + 12
  NEXT
  OffSet = 4
  DCod = 7
  FOR Cnt = 1 TO 6
    RSET TempE$ = DedDes$(DCod)
    QPrintRC TempE$, 12, OffSet, -1
    OffSet = OffSet + 12
    DCod = DCod + 1
  NEXT
  ShowCursor
  
  RETURN
  
ReCalcPay:
  'make sure all fields in the form are calculated
  CalcFields 0, 0, FormESC$(), FldESC()
  'copy the all fields to scrn calc type variable
  BCopy SSEG(FormESC$(0, 0)), SADD(FormESC$(0, 0)), VARSEG(ScrnCalc(1)), VARPTR(ScrnCalc(1)), LEN(ScrnCalc(1)), 0
  'copy scrn calcs fields to the transaction fields
  ParseScrnCalc2Trans
  'recalc tax's, net pay, deductions, etc.
  CalcPay TransRec(1), TransRecNo, True
  'copy transaction data back to scrn calc type variable
  ParseTrans2ScrnCalc
  'copy scrn calc type back to scrn calc form
  BCopy VARSEG(ScrnCalc(1)), VARPTR(ScrnCalc(1)), SSEG(FormESC$(0, 0)), SADD(FormESC$(0, 0)), LEN(ScrnCalc(1)), 0
  'unpack form$(0,0) to fields
  UnPackBuffer 0, 0, FormESC$(), FldESC()
  'update the scrn display
  PrintArray 0, 0, FormESC$(), FldESC()
  RETURN
  
END SUB

SUB PCEnterEdTrans (TransRecNo, EmpRecNo)
  
  REDIM TransRec(1) AS TransRecType
  REDIM FrmEET(1) AS FormInfo
  REDIM EarnDes$(1 TO 3)
  REDIM TempScrn(0)
  REDIM ManTrans(1) AS ManualTransRecType
  
  ManRecLen = LEN(ManTrans(1))
  TransRecLen = LEN(TransRec(1))
  EmpRecLen = LEN(EMP2Rec(1))
  
  'FGetAH PPDefaultFileName, PrdDefRec(1), LEN(PrdDefRec(1)), 1
  FOR Cnt = 1 TO 3
    EarnDes$(Cnt) = QPTrim$(ErnCodes(Cnt).ERNCODE1)
  NEXT
  
  FOpenS EmpData2Name, Handle
  FGetRTA Handle, EMP2Rec(1), CLNG(EmpRecNo), EmpRecLen
  FClose Handle
  
  FOpenS TransWorkFileName, THandle
  FGetRTA THandle, TransRec(1), CLNG(TransRecNo), TransRecLen
  
  OTActive = TransRec(1).TActive
  ' ^^^ use this to not kill a mistaken entry for allready active
  '     employee transaction
  
  IF NOT OTActive THEN
    IF EMP2Rec(1).EMPTDATE > -32767 THEN
      SaveScrn TempScrn()
      DisplayMiscScrn EMPNotFound
      WaitForAction
      FClose THandle
      GOTO NoGood
    ELSE
      CreateEmpTransRecs TransRecNo
      FGetRTA THandle, TransRec(1), CLNG(TransRecNo), TransRecLen
    END IF
  END IF
  
  FClose THandle
  
  EmpName$ = QPTrim$(EMP2Rec(1).EmpNo) + "   " + QPTrim$(EMP2Rec(1).EMPFNAME) + " " + QPTrim$(EMP2Rec(1).EMPLNAME)
  
  'fix for manual
  IF EntryType = Normal THEN
    SELECT CASE TransRec(1).PayType
    CASE "H"
      FormName$ = "HTIMEHRS"
    CASE "S"
      
      FormName$ = "SALRYNPT"
      
    END SELECT
  ELSE
    FormName$ = "MANTRAN"
  END IF
  '-=-=
  
  NumFlds = LibNumberOfFields(CalcQLib, FormName$)
  REDIM FormEET$(NumFlds, 2)    'DIM the form data array
  REDIM FldEET(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CalcQLib, FormName$, StartEl, FldEET(), FormEET$(), ErrCode
  
  'fix for MANUAL
  IF EntryType = Normal THEN
    FOR Cnt = 1 TO 3
      IF LEN(EarnDes$(Cnt)) = 0 THEN
        FldEET(FldNum("ALTEARN" + LTRIM$(STR$(Cnt)), FldEET())).Protected = True
      END IF
    NEXT
  END IF
  
  FrmEET(1).FldNo = 1           'Start editing on field #1
  FrmEET(1).InsStat = False     'Set the insert state (-1 = Insert on)
  FrmEET(1).StartEl = 0         'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  BlockClear
  
  'fix for MANUAL
  
  IF EntryType = Normal THEN
    SELECT CASE TransRec(1).PayType
    CASE "H"
      TOTHrs = FldNum("TOTHRSPD", FldEET())
      DstHrs = FldNum("DTREGHRS", FldEET())
      OTHRSPD = FldNum("OTHRSPD", FldEET())
      DstOTHrs = FldNum("DTOTHRS", FldEET())
      TOTAERN = FldNum("TOTAERN", FldEET())
      DstAErn = FldNum("TOTEADST", FldEET())
      TimeRecLen = LEN(HourInput(1))
      FormEET$(0, 0) = SPACE$(TimeRecLen)
      BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(HourInput(1)), VARPTR(HourInput(1)), TimeRecLen, 0
      ParseTrans2Hourly
      BCopy VARSEG(HourInput(1)), VARPTR(HourInput(1)), SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), TimeRecLen, 0
      UnPackBuffer 0, 0, FormEET$(), FldEET()
    CASE "S"    'the name of this field in the
      TotPct = FldNum("TOTOTHRS", FldEET())     'salary form is wrong
      TOTAERN = FldNum("TOTAERN", FldEET())
      DstAErn = FldNum("TOTEADST", FldEET())
      TimeRecLen = LEN(SalInput(1))
      FormEET$(0, 0) = SPACE$(TimeRecLen)
      BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(SalInput(1)), VARPTR(SalInput(1)), TimeRecLen, 0
      ParseTrans2Salary
      
      BCopy VARSEG(SalInput(1)), VARPTR(SalInput(1)), SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), TimeRecLen, 0
      UnPackBuffer 0, 0, FormEET$(), FldEET()
    END SELECT
    
  ELSE          'MANUAL do manual transaction setup here
    FormEET$(0, 0) = SPACE$(ManRecLen)
    IF OTActive THEN
      ParseTrans2Manual ManTrans(1)
    ELSE
      ManTrans(1).ChkDate = -32767
      ManTrans(1).PDSTART = -32767
      ManTrans(1).PDEND = -32767
      ManTrans(1).CheckNum = 0
      ManTrans(1).DISTACT1 = EMP2Rec(1).EDist(1).DAcct
      ManTrans(1).DISTACT2 = EMP2Rec(1).EDist(2).DAcct
      ManTrans(1).DISTACT3 = EMP2Rec(1).EDist(3).DAcct
      ManTrans(1).DISTACT4 = EMP2Rec(1).EDist(4).DAcct
      
    END IF
    BCopy VARSEG(ManTrans(1)), VARPTR(ManTrans(1)), SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), ManRecLen, 0
    UnPackBuffer 0, 0, FormEET$(), FldEET()
  END IF
  '-=-=
  
  DO            'Pole the editing procedure
    LibFile2Scrn CalcQLib, FormName$, MonoCode, -1, ErrCode
    
    GOSUB DisplayInfo
    
    DO
      EditForm FormEET$(), FldEET(), FrmEET(1), Cnf, Action
      'add check to make sure there is an account number for
      'every distribution and visa versa
      
      IF FrmEET(1).Presses AND FrmEET(1).MRow = 25 THEN
        
        SELECT CASE FrmEET(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 26 TO 37           'F3
          PressButton F3Key, 25, 26, 37
        CASE 53 TO 66           'F10
          PressButton F0Key, 25, 53, 66
        END SELECT
      END IF
      
      '09-30-94 *** Added check for mandatory fields
      IF NOT EntryType = Normal THEN
        IF (FrmEET(1).KeyCode <> 0 AND FrmEET(1).PrevFld <> FrmEET(1).FldNo) OR (FrmEET(1).KeyCode = F0Key) THEN
          SELECT CASE FrmEET(1).PrevFld
          CASE 1, 2, 3
            IF Date2Num(FormEET$(FrmEET(1).PrevFld, 0)) < 0 THEN
              FrmEET(1).FldNo = FrmEET(1).PrevFld
              FrmEET(1).KeyCode = 0
              FrmEET(1).StartEl = 0
              SaveScrn TempScrn()
              DisplayMiscScrn NoBlank
              WaitForAction
              RestScrn TempScrn()
            END IF
          END SELECT
        END IF
      END IF
      
      IF FrmEET(1).KeyCode THEN
        IF EntryType = Normal THEN
          SELECT CASE FldEET(FrmEET(1).PrevFld).FType
          CASE 3, 4, 5, 10, 18
            IF QPTrim$(FormEET$(FrmEET(1).PrevFld, 0)) = "" THEN LSET FormEET$(FrmEET(1).FldNo, 0) = "0"
          END SELECT
        END IF
      END IF
      
      SELECT CASE FrmEET(1).KeyCode
        
      CASE F0Key
        IF FrmEET(1).KeyCode = F0Key THEN
          'fix for MANUAL
          IF EntryType = Normal THEN
            SELECT CASE TransRec(1).PayType
            CASE "H"
              IF Value(FormEET$(TOTHrs, 0), Oops) <> Value(FormEET$(DstHrs, 0), Oops) THEN DstError = True
              IF Value(FormEET$(OTHRSPD, 0), Oops) <> Value(FormEET$(DstOTHrs, 0), Oops) THEN DstError = True
              IF Value(FormEET$(TOTAERN, 0), Oops) <> Value(FormEET$(DstAErn, 0), Oops) THEN DstError = True
            CASE "S"
              'IF FormEET$(1, 0) = "Y" THEN
              IF Value(FormEET$(TotPct, 0), Oops) <> 100 THEN DstError = True
              IF Value(FormEET$(TOTAERN, 0), Oops) <> Value(FormEET$(DstAErn, 0), Oops) THEN DstError = True
              'END IF
            END SELECT
            
            'make sure there are accounts for every amount
            
            '12-12-94  Adjusted field offsets for taxable fringe
            '         6    20
            FOR Cnt = 6 TO 20 STEP 2
              IF Value(FormEET$(Cnt, 0), Oops) > 0 AND LEN(QPTrim$(FormEET$(Cnt - 1, 0))) = 0 THEN
                FrmEET(1).FldNo = Cnt - 1
                AcctError = True
              END IF
            NEXT
            '01-15-96  Fixed offset bug in alt-earnings (Amts with no acct.)
            '         27    37
            FOR Cnt = 27 TO 37 STEP 2
              IF Value(FormEET$(Cnt, 0), Oops) > 0 AND LEN(QPTrim$(FormEET$(Cnt - 1, 0))) = 0 THEN
                FrmEET(1).FldNo = Cnt - 1
                AcctError = True
              END IF
            NEXT
          ELSE
            
            'MANUAL   do manual transaction distribution stuff here
          END IF
        END IF
        
        IF DstError OR AcctError THEN
          SaveScrn TempScrn()
          IF DstError THEN
            DisplayMiscScrn DistributErr
          ELSE
            DisplayMiscScrn AccountErr
          END IF
          WaitForAction
          FrmEET(1).KeyCode = 0
          SaveFlag = 1
          RestScrn TempScrn()
          ERASE TempScrn
          DstError = False
          AcctError = False
        END IF
        '***********************************
        'fix for normal not to save if esc key pressed   6/09/94
        IF FrmEET(1).KeyCode = EscKey AND EntryType = Normal THEN               'AND NOT OTActive THEN
          SaveFlag = True       'PromptSaveData
        ELSEIF FrmEET(1).KeyCode = EscKey AND EntryType = manual AND OTActive = False THEN
          SaveFlag = False
          TransRec(1).TActive = 0
          UpDateTransFile TransRecNo
        ELSEIF FrmEET(1).KeyCode = F0Key THEN
          'TransRec(1).TActive = True
          SaveFlag = True
        END IF
        '************************************
        SELECT CASE SaveFlag
        CASE True               'yep save data to disk
          ExitFlag = True
          'fix for MANUAL
          SELECT CASE EntryType
          CASE Normal
            SELECT CASE TransRec(1).PayType
            CASE "H"
              BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(HourInput(1)), VARPTR(HourInput(1)), TimeRecLen, 0
              ParseHourly2Trans
            CASE "S"
              BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(SalInput(1)), VARPTR(SalInput(1)), TimeRecLen, 0
              ParseSalary2Trans
            END SELECT
          CASE manual
            TransRec(1).TActive = True
            BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(ManTrans(1)), VARPTR(ManTrans(1)), ManRecLen, 0
            ParseManual2Trans ManTrans(1)
            'fix for manual  6/09/94
            IF PrdDefRec(1).MACTIVE = 0 THEN
              PrdDefRec(1).MACTIVE = True
              FPutAH PPDefaultFileName, PrdDefRec(1), LEN(PrdDefRec(1)), 1
            END IF
          END SELECT
          UpDateTransFile TransRecNo
          'do normal scrn calc screen if not a manual transaction
          IF FrmEET(1).KeyCode = F0Key AND EntryType = Normal THEN
            PCEdScrnCalc TransRecNo, EmpName$, EmpRecNo
          END IF
        CASE False              'nope don't save abandon edit
          ExitFlag = True
        CASE 1  'oops continue editing
          Action = 2
          ExitFlag = False
        END SELECT
      CASE F3Key
        IF PCDelFromPay THEN
          TransRec(1).TActive = 0
          UpDateTransFile TransRecNo
          ExitFlag = True
        END IF
      CASE ELSE
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
NoGood:
  ERASE TempScrn, FormEET$, FldEET
  
  EXIT SUB
  
DisplayInfo:
  
  Temp$ = SPACE$(10)
  HideCursor
  QPrintRC EmpName$, 1, 38, -1
  'fix for MANUAL
  IF EntryType = Normal THEN
    OffSet = 14
    FOR Cnt = 1 TO 3
      RSET Temp$ = EarnDes$(Cnt)
      QPrintRC Temp$, OffSet + Cnt, 22, -1
    NEXT
    
    '12-13-94  Adjusted field offsets for taxable fringe
    SELECT CASE TransRec(1).PayType
    CASE "H"
      CalcFields 0, 11, FormEET$(), FldEET()
      CalcFields 0, 12, FormEET$(), FldEET()
      'CalcFields 0, 36, FormEET$(), FldEET()
      CalcFields 0, 41, FormEET$(), FldEET()
    CASE "S"
      CalcFields 0, 6, FormEET$(), FldEET()
      CalcFields 0, 22, FormEET$(), FldEET()
      CalcFields 0, 27, FormEET$(), FldEET()
    END SELECT
  ELSE          'fix for MANUAL
    FOR Cnt = 1 TO 12
      QPrintRC QPTrim$(DedCodes(Cnt).DCDESC1), Cnt + 4, 44, -1
    NEXT
  END IF
  ShowCursor
  
  RETURN
  
END SUB

SUB PCGetEmp2Rec (EmpRecNo%, ErrorCode)
  
  FOpenS EmpData2Name, Handle   'open employee data file
  FGetRTA Handle, EMP2Rec(1), CLNG(EmpRecNo), LEN(EMP2Rec(1))
  
  'IF EMP2Rec(1).IsLocked THEN
  '  ErrorCode = True
  'ELSE
  '  EMP2Rec(1).IsLocked = True
  '  FPutRTA Handle, EMP2Rec(1), CLNG(EmpRecNo), LEN(EMP2Rec(1))
  '  ErrorCode = False
  'END IF
  
  FClose Handle
  
END SUB

SUB PCGetEmpNum (EmpNumber)
  
  REDIM FrmGEN(1) AS FormInfo
  
  REDIM TempScrn(0)
  
  FormName$ = "TRANSTRT"
  NumFlds = LibNumberOfFields(CalcQLib, FormName$)
  REDIM FormGEN$(NumFlds, 2)    'DIM the form data array
  REDIM FldGEN(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CalcQLib, FormName$, StartEl, FldGEN(), FormGEN$(), ErrCode
  
  '----- Setup TYPE for setting and reading form editing information.
  FrmGEN(1).FldNo = 1           'Start editing on field #1
  FrmGEN(1).InsStat = False     'Set the insert state (-1 = Insert on)
  FrmGEN(1).StartEl = 0         'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  BlockClear
  
  DO            'Pole the editing procedure
    'look
    LSET FormGEN$(1, 0) = ""
    LibFile2Scrn CalcQLib, FormName$, MonoCode, -1, ErrCode
    GOSUB DisplayName
    
    DO
      EditForm FormGEN$(), FldGEN(), FrmGEN(1), Cnf, Action
      
      IF FrmGEN(1).Presses AND FrmGEN(1).MRow = 25 THEN
        SELECT CASE FrmGEN(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 47 TO 59           'F7
          PressButton F7KEY, 25, 47, 59
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        END SELECT
      END IF
      
      SELECT CASE FrmGEN(1).KeyCode
      CASE EnterKey
        WorkNum$ = LTRIM$(RTRIM$(FormGEN$(1, 0)))
        IF LEN(WorkNum$) THEN
          PCLookUpEmp WorkNum$, TransRecNo, EmpRecNo
          IF TransRecNo > 0 THEN
            PCEnterEdTrans TransRecNo, EmpRecNo
            Action = 1
            EXIT DO
          ELSE
            SaveScrn TempScrn()
            DisplayMiscScrn EMPNotFound
            WaitForAction
            LSET FormGEN$(1, 0) = ""
            RestScrn TempScrn()
            PrintArray 1, 1, FormGEN$(), FldGEN()
          END IF
        ELSE
          ExitFlag = True
        END IF
'04-07-96 Added F7 pick trans list for Mr. Dew
      CASE F7KEY
        'PCPickEmpList TransRecNo, EmpRecNo
        Action = 1
        EXIT DO
'-=-=-=-=
      CASE EscKey
        ExitFlag = True
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  CursorOff
  
  ERASE TempScrn, FormGEN$, FldGEN, FrmGEN
  
  EXIT SUB
  
DisplayName:
  HideCursor
  IF EntryType = Normal THEN
    QPrintRC "Payroll", 1, 3, 112
  ELSE
    QPrintRC "Manual", 1, 4, 112
  END IF
  IF LEN(EmpName$) > 0 THEN QPrintRC EmpName$, 1, 45, 112
  ShowCursor
  RETURN
  
END SUB

SUB PCLoadPayFreqs
  PayPFreq$(1) = "Weekly          "
  PayPFreq$(2) = "Bi-Weekly       "
  PayPFreq$(3) = "Semi-Monthly    "
  PayPFreq$(4) = "Monthly         "
  PayPFreq$(5) = "Quarterly       "
  PayPFreq$(6) = "Semi-Annually   "
  PayPFreq$(7) = "Annually        "
END SUB

SUB PCLoadSystemFiles
  REDIM DedCodes(1 TO 12)   AS DedCodeRecType
  REDIM ErnCodes(1 TO 3)    AS ErnCodeRecType
  REDIM RetireRec(1 TO 6)   AS RetireRecType
  REDIM StateTax(1)         AS StateTaxRecType
  REDIM FedTax(1)           AS FederalTaxRecType
  REDIM EICRec(1)           AS EICRecType
  REDIM PrdDefRec(1)        AS PeriodDefaultRecType
  
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  FGetAH ErnCodeFileName, ErnCodes(1), LEN(ErnCodes(1)), 3
  FGetAH RetireFileName, RetireRec(1), LEN(RetireRec(1)), 6
  FGetAH FederalTaxFileName, FedTax(1), LEN(FedTax(1)), 1
  FGetAH StateTaxFileName, StateTax(1), LEN(StateTax(1)), 1
  FGetAH EICFileName, EICRec(1), LEN(EICRec(1)), 2
  FGetAH PPDefaultFileName, PrdDefRec(1), LEN(PrdDefRec(1)), 1
  
END SUB

SUB PCLookUpEmp (EmpNum$, TRecNum, ERecNum)
  
  REDIM EmpRec1(1)     AS EmpData1Type
  REDIM LUEmpRec2(1)   AS EmpData2Type
  
  TRecNum = 0
  ERecNum = 0
  
  RecLen = LEN(EmpRec1(1))
  RecLen2 = LEN(LUEmpRec2(1))
  
  NumOfRecs = FileSize(EmpData1Name) \ RecLen
  
  FOpenS EmpData1Name, Handle
  FOpenS EmpData2Name, Handle2  'open employee data file
  
  FOR Cnt = 1 TO NumOfRecs
    FGetT Handle, EmpRec1(1), RecLen
    TestNum$ = QPTrim$(EmpRec1(1).EmpNo)
    IF TestNum$ = EmpNum$ THEN
      FGetRTA Handle2, LUEmpRec2(1), CLNG(Cnt), RecLen2
      IF LUEmpRec2(1).EMPTDATE <= -32767 AND NOT LUEmpRec2(1).Deleted = True THEN
        TRecNum = Cnt           'EmpRec1(1).TransRecNum
        ERecNum = Cnt           'EmpRec1(1).Data1RecNum
      END IF
      EXIT FOR
    END IF
  NEXT
  
  FClose Handle
  FClose Handle2
  ERASE EmpRec1, LUEmpRec2
  
END SUB

SUB PCSetPeriodDefault
  
  ' NO EDITING IN HERE UNLESS ok to delete all PERIOD TRANSACTION
  ' OR, OK IF NO PERIOD TRANSACTIONS ARE PENDING
  
  REDIM FrmPPD(1) AS FormInfo
  
  REDIM TempScrn(0)
  
  FormName$ = "PRPPDEF"
  NumFlds = LibNumberOfFields(CalcQLib, FormName$)
  REDIM FormPPD$(NumFlds, 2)    'DIM the form data array
  REDIM FldPPD(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CalcQLib, FormName$, StartEl, FldPPD(), FormPPD$(), ErrCode
  
  PeriodDefaultRecLen = LEN(PrdDefRec(1))
  
  FormPPD$(0, 0) = SPACE$(PeriodDefaultRecLen)
  
  FldOffSet = 11
  FOR Cnt = 1 TO 12
    IF QPTrim$(DedCodes(Cnt).DCDESC1) = "" THEN
      LSET FormPPD$((FldOffSet + Cnt), 0) = ""
      FldPPD(FldOffSet + Cnt).Protected = True
    END IF
  NEXT
  
  FGetAH PPDefaultFileName, PrdDefRec(1), PeriodDefaultRecLen, 1
  BCopy VARSEG(PrdDefRec(1)), VARPTR(PrdDefRec(1)), SSEG(FormPPD$(0, 0)), SADD(FormPPD$(0, 0)), PeriodDefaultRecLen, 0
  UnPackBuffer 0, 0, FormPPD$(), FldPPD()
  
  WasActive = PrdDefRec(1).PACTIVE
  
  UseDefaultF = FldNum("USEDEF", FldPPD())
  
  '----- Setup TYPE for setting and reading form editing information.
  FrmPPD(1).FldNo = 1           'Start editing on field #1
  FrmPPD(1).InsStat = False     'Set the insert state (-1 = Insert on)
  FrmPPD(1).StartEl = 0         'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '  BlockClear
  
  'ASK JIM
  '  IF PCCheckChecksActive THEN
  '    GOTO DefaultErrExit
  '  END IF
  
  BlockClear
  
  DO            'Poll the editing procedure
    LibFile2Scrn CalcQLib, FormName$, MonoCode, -1, ErrCode
    
    FOR Cnt = 1 TO 12
      SELECT CASE Cnt
      CASE 1 TO 4
        QPrintRC DedCodes(Cnt).DCDESC1, 14 + Cnt, 12, Cnf.NonMen
      CASE 5 TO 8
        QPrintRC DedCodes(Cnt).DCDESC1, 10 + Cnt, 35, Cnf.NonMen
      CASE 9 TO 12
        QPrintRC DedCodes(Cnt).DCDESC1, 6 + Cnt, 59, Cnf.NonMen
      END SELECT
      FldOffSet = FldOffSet + 2
    NEXT
    
    zz = 12
    FOR Cnt = 1 TO 2
      QPrintRC ErnCodes(Cnt).ERNCODE1, 21, zz, Cnf.NonMen
      zz = zz + 23
    NEXT
    QPrintRC ErnCodes(Cnt).ERNCODE1, 21, zz + 1, Cnf.NonMen
    
    'PRINT FRE(-1), FRE(-2), FRE("  "), FRE(CurDrive$)
    'WaitForAction
    
    
    DO
      EditForm FormPPD$(), FldPPD(), FrmPPD(1), Cnf, Action
      
      IF NOT FirstFlag THEN
        FirstFlag = True
        SaveScrn TempScrn()
      END IF
      
      IF FrmPPD(1).Presses AND FrmPPD(1).MRow = 25 THEN
        SELECT CASE FrmPPD(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        END SELECT
      END IF
      
      SELECT CASE FrmPPD(1).KeyCode
      CASE F0Key, EscKey
        SaveScrn TempScrn()
        SELECT CASE FrmPPD(1).KeyCode
        CASE EscKey
          IF FrmPPD(1).Edited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF FrmPPD(1).Edited THEN
            SaveFlag = True
          ELSEIF PrdDefRec(1).PACTIVE = False THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          'ADD routine like PromptSaveData here to
          'Make warn transaction must be overwritten here
          'DisplayMiscScrn PayrollInProg
          IF WasActive THEN
            SaveFlag = PromptPeriodWasActive
            IF SaveFlag THEN
              DisplayMiscScrn UpdatingDisk
              MakeTransInActive
            END IF
          END IF
          IF SaveFlag THEN
            BCopy SSEG(FormPPD$(0, 0)), SADD(FormPPD$(0, 0)), VARSEG(PrdDefRec(1)), VARPTR(PrdDefRec(1)), PeriodDefaultRecLen, 0
            PrdDefRec(1).PACTIVE = True
            FPutAH PPDefaultFileName, PrdDefRec(1), PeriodDefaultRecLen, 1
          END IF
          ExitFlag = True
        CASE False
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
  IF NOT SaveFlag THEN EXIT SUB
  
  IF FormPPD$(UseDefaultF, 0) = "Y" THEN
    EntryType = Normal
    RestScrn TempScrn()
    MakeDefaultTransActs
  END IF
  
DefaultErrExit:
  
  ERASE TempScrn, FormPPD$, FldPPD
  
END SUB

SUB PostTransactions
  
  FirstFlag = True
  BlockClear
  
  DisplayMiscScrn ReadyToPost
  
  DO
    DO
      Ky$ = INKEY$
      ButtonPress 1, 0, MouseButton, XPos, YPos
    LOOP UNTIL LEN(Ky$) OR MouseButton
    
    IF MouseButton THEN
      MRow = (YPos \ 8) + 1
      MCol = (XPos \ 8)         '14 30
      SELECT CASE MRow
      CASE 16
        SELECT CASE MCol
        CASE 30 TO 41
          PressButton EscKey, 16, 30, 41
        CASE 43 TO 53
          PressButton F0Key, 16, 43, 53
        CASE ELSE
        END SELECT
        Ky$ = INKEY$
      CASE ELSE
      END SELECT
    END IF
    
    SELECT CASE Ky$
    CASE CHR$(0) + "D"
      PostOk = -1
      ExitFlag = -1
    CASE CHR$(27)
      PostOk = 0
      ExitFlag = -1
    END SELECT
    
  LOOP UNTIL ExitFlag
  
  IF PostOk = 0 THEN EXIT SUB
  
  BlockClear
  
  DisplayMiscScrn Posting
  
  PostDate = Date2Num(DATE$)
  
  REDIM TransRec(1)   AS TransRecType
  REDIM EmpRec2(1)    AS EmpData2Type
  REDIM EmpRec3(1)    AS EmpData3Type
  REDIM Check(1)      AS PRCheckRecType
  
  REDIM PSysRec(1)    AS RegDSysFileRecType
  
  REDIM OSIFRec(1)    AS OSChkRecType
  
  REDIM PDR(1)        AS PeriodDefaultRecType
  
  TransRecLen = LEN(TransRec(1))
  Emp2RecLen = LEN(EmpRec2(1))
  Emp3RecLen = LEN(EmpRec3(1))
  CheckRecLen = LEN(Check(1))
  OSChkRecLen = LEN(OSIFRec(1))
  
  FGetAH SysFileName, PSysRec(1), LEN(PSysRec(1)), 1
  
  '05/10/94   Fix to set false after posting
  '    -vvv-  is used as a flag to menu entry points
  '06/09/94   Fix for manual trans
  PDRLen = LEN(PDR(1))
  FGetAH PPDefaultFileName, PDR(1), PDRLen, 1   'get pay period defaults
  SELECT CASE EntryType
  CASE Normal
    PrdDefRec(1).PACTIVE = False
    PDR(1).PACTIVE = False      'set active flag = false
  CASE manual
    PrdDefRec(1).MACTIVE = False
    PDR(1).MACTIVE = False      'set active flag = false
  END SELECT
  
  FPutAH PPDefaultFileName, PDR(1), PDRLen, 1   'update period file
  
  NumOfRecs& = FileSize(ChecksFileName) \ CheckRecLen
  
  NextHistRec& = (FileSize(TransHistFileName) \ TransRecLen) + 1
  
  FOpenS ChecksFileName, CHandle                'open checks data file
  FOpenS EmpData2Name, EHandle2 'open employee data 2 file   HRS
  FOpenS EmpData3Name, EHandle3 'open employee data 3 file   YTD
  FOpenS TransWorkFileName, THandle             'open transaction work file
  FOpenS TransHistFileName, HHandle             'open transaction history file
  
  OSChkFile$ = QPTrim$(PSysRec(1).CITIDIR)

  IF LEN(OSChkFile$) > 0 THEN
    DoOSChkFlag = True
  END IF
  IF DoOSChkFlag THEN
    IF RIGHT$(OSChkFile$, 1) <> "\" THEN
      OSChkFile$ = OSChkFile$ + "\CRCHK.DAT"
    ELSE
      OSChkFile$ = OSChkFile$ + "CRCHK.DAT"
    END IF
    IF NOT Exist(OSChkFile$) THEN
      FCreate OSChkFile$
    END IF
    FOpenS OSChkFile$, OHandle  'open checks data file
    IF OHandle > 0 THEN
      OSNumRec = (FileSize(OSChkFile$) \ OSChkRecLen) + 1
    ELSE
      DoOSChkFlag = False
    END IF
  END IF
  
  FOR Cnt& = 1 TO NumOfRecs&
    FGetRTA CHandle, Check(1), Cnt&, CheckRecLen                'get the check data rec
    FGetRTA THandle, TransRec(1), Cnt&, TransRecLen
    
    'IF Check(1).CActive = -1 OR TransRec(1).TActive = -1 THEN
    IF TransRec(1).TActive = -1 THEN
      'if this is an active check or transaction
      IF FirstFlag THEN
        FirstFlag = False
        GLIFDate$ = Num2Date(Check(1).CheckDate)
        ReplaceString GLIFDate$, "-", "/"
        ReplaceString GLIFDate$, "1994", "94"
        ReplaceString GLIFDate$, "1995", "95"
        ReplaceString GLIFDate$, "1996", "96"
        ReplaceString GLIFDate$, "1997", "97"
        ReplaceString GLIFDate$, "1998", "98"
        ReplaceString GLIFDate$, "1999", "99"
        GLIFDate$ = QPTrim$(GLIFDate$)
      END IF
      
      
      '** Get all data from files that need to be updated
      FGetRTA EHandle2, EmpRec2(1), Cnt&, Emp2RecLen
      FGetRTA EHandle3, EmpRec3(1), Cnt&, Emp3RecLen
      
      '** Update employee 2 file and adjust previous transaction pointer
      IF EmpRec2(1).EMPVACE < 0 THEN EmpRec2(1).EMPVACE = 0
      IF EmpRec2(1).EMPSLE < 0 THEN EmpRec2(1).EMPSLE = 0
      IF EmpRec2(1).EMPCTE < 0 THEN EmpRec2(1).EMPCTE = 0
      
      IF EmpRec2(1).EMPVUSED < 0 THEN EmpRec2(1).EMPVUSED = 0
      IF EmpRec2(1).EMPSLUSE < 0 THEN EmpRec2(1).EMPSLUSE = 0
      IF EmpRec2(1).EMPCTUSE < 0 THEN EmpRec2(1).EMPCTUSE = 0
      
      EmpRec2(1).EMPVUSED = Round(EmpRec2(1).EMPVUSED + TransRec(1).VacUsed)
      EmpRec2(1).EMPVBAL = Round(EmpRec2(1).EMPVACE - EmpRec2(1).EMPVUSED)
      
      EmpRec2(1).EMPSLUSE = Round(EmpRec2(1).EMPSLUSE + TransRec(1).SickUsed)
      EmpRec2(1).EMPSLBAL = Round(EmpRec2(1).EMPSLE - EmpRec2(1).EMPSLUSE)
      
      'fix for comp earned 05/04/94
      EmpRec2(1).EMPCTUSE = Round(EmpRec2(1).EMPCTUSE + TransRec(1).CompUsed)
      IF EntryType = Normal THEN
        EmpRec2(1).EMPCTBAL = Check(1).CompBal
        EmpRec2(1).EMPCTE = Check(1).CompEarn
      ELSE
        EmpRec2(1).EMPCTBAL = Round(EmpRec2(1).EMPCTE - EmpRec2(1).EMPCTUSE)
      END IF
      
      
      'adjust and update (last - previous) transaction pointers
      IF EmpRec2(1).LastTransRec >= 0 THEN
        TransRec(1).PrevTransRec = EmpRec2(1).LastTransRec
      ELSE
        TransRec(1).PrevTransRec = 0
      END IF
      
      EmpRec2(1).LastTransRec = CINT(NextHistRec&)
      
      '** Update employee 3 file
      'EmpRec3(1).YTDGrossPay = Round(EmpRec3(1).YTDGrossPay + Check(1).GrossPay)
      '-=-=man
      EmpRec3(1).YTDGrossPay = Round(EmpRec3(1).YTDGrossPay + TransRec(1).GrossPay)
      EmpRec3(1).YTDFedGrossPay = Round(EmpRec3(1).YTDFedGrossPay + TransRec(1).FedGrossPay)
      EmpRec3(1).YTDStaGrossPay = Round(EmpRec3(1).YTDStaGrossPay + TransRec(1).StaGrossPay)
      EmpRec3(1).YTDSocGrossPay = Round(EmpRec3(1).YTDSocGrossPay + TransRec(1).SocGrossPay)
      EmpRec3(1).YTDMedGrossPay = Round(EmpRec3(1).YTDMedGrossPay + TransRec(1).MedGrossPay)
      
      EmpRec3(1).YTDRegPay = Round(EmpRec3(1).YTDRegPay + TransRec(1).TotRegWage)
      EmpRec3(1).YTDOTPay = Round(EmpRec3(1).YTDOTPay + TransRec(1).TotOTWage)
      
      EmpRec3(1).YTDNet = Round(EmpRec3(1).YTDNet + TransRec(1).NetPay)
      
      EmpRec3(1).YTDFederal = Round(EmpRec3(1).YTDFederal + TransRec(1).FedTaxAmt)
      EmpRec3(1).YTDState = Round(EmpRec3(1).YTDState + TransRec(1).StaTaxAmt)
      EmpRec3(1).YTDSocial = Round(EmpRec3(1).YTDSocial + TransRec(1).SocTaxAmt)
      EmpRec3(1).YTDMedicare = Round(EmpRec3(1).YTDMedicare + TransRec(1).MedTaxAmt)
      EmpRec3(1).YTDRetire = Round(EmpRec3(1).YTDRetire + TransRec(1).RetireAmt)
      
      'year to date totals on deductions
      FOR Cnt = 1 TO 12
        EmpRec3(1).YTDDAmt(Cnt) = Round(EmpRec3(1).YTDDAmt(Cnt) + TransRec(1).DAmt(Cnt))
        EmpRec3(1).YTDDAmtT = Round(EmpRec3(1).YTDDAmtT + TransRec(1).DAmt(Cnt))
      NEXT
      
      'year to date totals on alt earnings
      EmpRec3(1).YTDEarn1 = Round(EmpRec3(1).YTDEarn1 + TransRec(1).EAmt(1))
      EmpRec3(1).YTDEarn2 = Round(EmpRec3(1).YTDEarn2 + TransRec(1).EAmt(2))
      EmpRec3(1).YTDEarn3 = Round(EmpRec3(1).YTDEarn3 + TransRec(1).EAmt(3))
      EmpRec3(1).YTDEarnT = Round(EmpRec3(1).YTDEarn1 + EmpRec3(1).YTDEarn2 + EmpRec3(1).YTDEarn3)
      
      '** Update MISC transaction history data
      'added fix for manual transaction entry 5/13/94  friday the 13th
      
      IF EntryType = Normal THEN
        TransRec(1).CheckNum = Check(1).CheckNum
        TransRec(1).CheckDate = Check(1).CheckDate
        TransRec(1).PostDate = PostDate
      END IF
      
      '** Added Update EIC year to date.  6/06/94
      IF TransRec(1).EICAmt > 0 THEN
        EmpRec3(1).YTDEIC = Round(EmpRec3(1).YTDEIC + TransRec(1).EICAmt)
      END IF
      
      '** Update active flags
      TransRec(1).TActive = False
      Check(1).CActive = False
      '** Update OSChk File
      IF DoOSChkFlag THEN
        OSIFRec(1).ChkNum = Check(1).CheckNum
        OSIFRec(1).ChkDate = GLIFDate$
        OSIFRec(1).Desc = Check(1).EmpName
        OSIFRec(1).Amt = Check(1).NetPay
        OSIFRec(1).Src = 1
        FPutRTA OHandle, OSIFRec(1), CLNG(OSNumRec), OSChkRecLen
        OSNumRec = OSNumRec + 1
      END IF
      '** Update DISK files
      FPutRTA CHandle, Check(1), Cnt&, CheckRecLen
      FPutRTA THandle, TransRec(1), Cnt&, TransRecLen
      FPutRTA HHandle, TransRec(1), NextHistRec&, TransRecLen
      FPutRTA EHandle2, EmpRec2(1), Cnt&, Emp2RecLen
      FPutRTA EHandle3, EmpRec3(1), Cnt&, Emp3RecLen
      NextHistRec& = NextHistRec& + 1
    END IF
    GOSUB PostPctComplete
  NEXT
  
  FClose EHandle2
  FClose EHandle3
  FClose CHandle
  FClose THandle
  FClose HHandle
  IF DoOSChkFlag THEN
    FClose OHandle
  END IF
  
  REDIM GLIFRec79(1)    AS GLIFDataType79
  REDIM GLIFRec10(1)    AS GLIFDataType10
  REDIM GLIFRec11(1)    AS GLIFDataType11
  REDIM GLIFRec12(1)    AS GLIFDataType12
  
  GLIFRecLen = LEN(GLIFRec12(1))
  GLIFRecLen79 = LEN(GLIFRec79(1))
  GLIFRecLen10 = LEN(GLIFRec10(1))
  GLIFRecLen11 = LEN(GLIFRec11(1))
  
  'process gl transfer file
  IF Exist("TEMPIF.DAT") THEN
    PRIFDir$ = QPTrim$(PSysRec(1).CITIDIR)
    IF RIGHT$(PRIFDir$, 1) <> "\" THEN
      PRIFDir$ = PRIFDir$ + "\PRIF.DAT"
    ELSE
      PRIFDir$ = PRIFDir$ + "PRIF.DAT"
    END IF
    NextRec& = (FileSize(PRIFDir$) \ GLIFRecLen) + 1
    NumOfRecs = FileSize("TEMPIF.DAT") \ GLIFRecLen
    
    FOpenS "TEMPIF.DAT", IHandle
    IF NOT Exist(PRIFDir$) THEN
      FCreate PRIFDir$
    END IF
    
    FOpenS PRIFDir$, OHandle
    
    FOR Cnt = 1 TO NumOfRecs
      FGetRTA IHandle, GLIFRec12(1), CLNG(Cnt), GLIFRecLen
      GLIFRec12(1).TranDate = GLIFDate$
      FPutRTA IHandle, GLIFRec12(1), CLNG(Cnt), GLIFRecLen

      SELECT CASE PSysRec(1).GLActLen
      CASE 7, 9
        GLIFRec79(1).TranAcct = QPTrim$(GLIFRec12(1).TranAcct)
        GLIFRec79(1).TranDate = GLIFRec12(1).TranDate
        GLIFRec79(1).TranDesc = GLIFRec12(1).TranDesc
        GLIFRec79(1).CrAmt = GLIFRec12(1).CrAmt
        GLIFRec79(1).DrAmt = GLIFRec12(1).DrAmt
        GLIFRec79(1).Detail = GLIFRec12(1).Detail
        GLIFRec79(1).Source = GLIFRec12(1).Source
        GLIFRec79(1).FromFlag = GLIFRec12(1).FromFlag
        FPutRTA OHandle, GLIFRec79(1), NextRec&, GLIFRecLen79
      CASE 10
        GLIFRec10(1).TranAcct = QPTrim$(GLIFRec12(1).TranAcct)
        GLIFRec10(1).TranDate = GLIFRec12(1).TranDate
        GLIFRec10(1).TranDesc = GLIFRec12(1).TranDesc
        GLIFRec10(1).CrAmt = GLIFRec12(1).CrAmt
        GLIFRec10(1).DrAmt = GLIFRec12(1).DrAmt
        GLIFRec10(1).Detail = GLIFRec12(1).Detail
        GLIFRec10(1).Source = GLIFRec12(1).Source
        GLIFRec10(1).FromFlag = GLIFRec12(1).FromFlag
        FPutRTA OHandle, GLIFRec10(1), NextRec&, GLIFRecLen10
      CASE 11
        GLIFRec11(1).TranAcct = QPTrim$(GLIFRec12(1).TranAcct)
        GLIFRec11(1).TranDate = GLIFRec12(1).TranDate
        GLIFRec11(1).TranDesc = GLIFRec12(1).TranDesc
        GLIFRec11(1).CrAmt = GLIFRec12(1).CrAmt
        GLIFRec11(1).DrAmt = GLIFRec12(1).DrAmt
        GLIFRec11(1).Detail = GLIFRec12(1).Detail
        GLIFRec11(1).Source = GLIFRec12(1).Source
        GLIFRec11(1).FromFlag = GLIFRec12(1).FromFlag
        FPutRTA OHandle, GLIFRec11(1), NextRec&, GLIFRecLen11
      CASE 12
        FPutRTA OHandle, GLIFRec12(1), NextRec&, GLIFRecLen
      END SELECT
      NextRec& = NextRec& + 1
    NEXT
    
    FClose IHandle
    FClose OHandle

'This is for nicks autopost
        SELECT CASE PSysRec(1).GLActLen
          CASE 7, 9, 12
            'Post2BA PSysRec(1).CITIDIR, BadAccts
            IF BadAccts = -1 THEN
              'there was a posting to gl error
            ELSE
              KillFile QPTrim$(PSysRec(1).CITIDIR) + "\PRIF.DAT"
            END IF
        END SELECT
    
    IF Exist("TEMPIF.DAT") THEN
      KillFile "LASTIF.DAT"
      NAME "TEMPIF.DAT" AS "LASTIF.DAT"
    END IF
  END IF
  
  
  'done with gl transfer file
  
  ERASE PSysRec, TransRec, EmpRec2, EmpRec3, Check, PDR
  ERASE GLIFRec79, GLIFRec11, GLIFRec12, OSIFRec
  
  BlockClear
  
  DisplayMiscScrn UpdatedOK
  WaitForAction
  
  EXIT SUB
  
PostPctComplete:
  CALL QPrintRC(STR$(INT((Cnt& / NumOfRecs&) * 100)), 12, 35, 112)
  RETURN
  
END SUB

SUB PRCalcMenu
  
  PCLoadPayFreqs
  PCLoadSystemFiles
  
  'SPLIT
  REDIM SysRec(1) AS RegDSysFileRecType
  FGetAH SysFileName, SysRec(1), LEN(SysRec(1)), 1
  IF SysRec(1).SplitFlag = "Y" THEN
    SplitFlag = True
  ELSE
    SplitFlag = False
  END IF
  ERASE SysRec
  
  
  DO
    PayMenu PayProcessMenu, Choice, 7
    '-=-=-=-=-=-=-=-=
    SELECT CASE Choice
    CASE 2
      IF NOT PrdDefRec(1).MACTIVE THEN
        PCSetPeriodDefault
      ELSE
        BlockClear
        DisplayMiscScrn NoNormalNow
        WaitForAction
      END IF
    CASE 1, 3 TO 9
      IF PrdDefRec(1).PACTIVE OR Choice = 1 OR Choice = 7 THEN
        SELECT CASE Choice
        CASE 1
          'AccruLeave False       'unrem
        CASE 3  'payroll transaction entry
          EntryType = Normal
          PCGetEmpNum LastEmpNum
        CASE 4
          IF SplitFlag THEN
            'PCPrintPayRegisterS     'unrem me
          ELSE
            'PCPrintPayRegister     'unrem me
          END IF
        CASE 5
          'PRCheckMenu            'unrem
        CASE 6
          EntryType = Normal
          PostTransactions
        CASE 7  'manual transaction entry
          IF PrdDefRec(1).PACTIVE THEN
            'No manual transaction durning payroll active error
            BlockClear
            DisplayMiscScrn NoManualNow
            WaitForAction
          ELSE
            ManualMenu
          END IF
        CASE 8
          'VoidTransaction               'unrem
          
        END SELECT
      ELSE
        IF NOT PrdDefRec(1).MACTIVE THEN
          BlockClear
          DisplayMiscScrn NoPeriodDefaults
          WaitForAction
        ELSE
          BlockClear
          DisplayMiscScrn NoNormalNow
          WaitForAction
        END IF
        Choice = 1
      END IF
    END SELECT
  LOOP UNTIL Choice = EscKey
  
END SUB

FUNCTION PromptPeriodWasActive
  
  REDIM Frm(1) AS FormInfo
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  SaveFlag = 2
  
  FormName$ = PayrollInProg
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  DisplayMiscScrn PayrollInProg
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      SaveFlag = True
    CASE EscKey
      SaveFlag = False
    END SELECT
    
  LOOP WHILE SaveFlag = 2       'proper key not set
  
  PromptPeriodWasActive = SaveFlag
  CursorOff
  
  RestScrn TempScrn()
  
  ERASE TempScrn, Form$, Fld, Frm
  
  
END FUNCTION

SUB UpDateTransFile (TransRecNo)
  FOpenS TransWorkFileName, Handle              'open it
  FPutRTA Handle, TransRec(1), CLNG(TransRecNo), LEN(TransRec(1))
  FClose Handle
END SUB

