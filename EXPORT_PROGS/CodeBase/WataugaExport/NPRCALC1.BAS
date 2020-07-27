DEFINT A-Z
DECLARE SUB PRCalcMenu ()
DECLARE SUB PCPickEmpList (TransRecNo%, EmpRecNo%)
DECLARE SUB Post2GL (FileName$, PSysRec() AS ANY, BadTrans%)
'DECLARE SUB Post2BA (FileName$, BadTrans%)
DECLARE FUNCTION FmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
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
DECLARE SUB PRDraftMenu ()
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
DECLARE SUB PCEdScrnCalc (TransRec() AS ANY, TransRecNo%, EmpName$, EmpRecNo%)
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
DECLARE SUB ParseManual2Trans (TransRec() AS ANY, ManTrans AS ANY)
DECLARE SUB ParseTrans2Manual (TransRec() AS ANY, ManTrans AS ANY)
DECLARE SUB ParseHourly2Trans (TransRec() AS ANY, HourInput() AS ANY)
DECLARE SUB ParseTrans2Hourly (TransRec() AS ANY, HourInput() AS ANY)
DECLARE SUB ParseSalary2Trans (TransRec() AS ANY, SalaryInput() AS ANY)
DECLARE SUB ParseTrans2Salary (TransRec() AS ANY, SalaryInput() AS ANY)
DECLARE SUB ParseScrnCalc2Trans (TransRec() AS ANY, ScrnCalc() AS ANY)
DECLARE SUB ParseTrans2ScrnCalc (TransRec() AS ANY, ScrnCalc() AS ANY)
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB PostTransactions (EntryType)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintArray (FirstFld%, LastFld%, Form$(), Fld() AS ANY)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ReplaceString (Work$, Old$, New$)
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB ShowCursor ()
DECLARE SUB UnPackBuffer (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB UpDateTransFile (TransRec() AS ANY, TransRecNo%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION CheckFldEmpty% (FldData$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB WazzWind (BYVAL TopRow, BYVAL LeftCol, BYVAL BotRow, BYVAL RghtCol, BYVAL FrameColor, BYVAL FrameType, BYVAL Shadow)
DECLARE FUNCTION DOSError ()
DECLARE FUNCTION Monitor% ()
  
  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
  '$INCLUDE: 'FieldInf.bi'                '        "        field information
  '$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'Retire.Bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'ErnCodes.Bi'
  '$INCLUDE: 'Leave.Bi'
  '$INCLUDE: 'StateTax.Bi'
  '$INCLUDE: 'FedTax.Bi'
  '$INCLUDE: 'PREIC.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'PRTIMEIN.Bi'
  '$INCLUDE: 'ScrCalc1.Bi'
  '$INCLUDE: 'PRCHECK.Bi'
  '$INCLUDE: 'PROSCHK.bi'
  '$INCLUDE: 'prif.bi'
  '$INCLUDE: 'GL.bi'
  '$INCLUDE: 'setcnf.bi'
  

  CONST False = 0
  CONST True = NOT False
  
  CONST Manual = 2
  CONST Normal = 1
  
  REDIM SHARED DedCodes(0)       AS DedCodeRecType
  REDIM SHARED ErnCodes(0)       AS ErnCodeRecType
  REDIM SHARED RetireRec(0)      AS RetireRecType
  REDIM SHARED StateTax(0)       AS StateTaxRecType
  REDIM SHARED FEDTAX(0)         AS FederalTaxRecType
  REDIM SHARED EICRec(0)         AS EICRecType
  REDIM SHARED PrdDefRec(0)      AS PeriodDefaultRecType
  
  DIM SHARED EMP2Rec(1)   AS EmpData2Type
  DIM SHARED THandle AS INTEGER, TRecSize AS INTEGER, EntryType AS INTEGER
  DIM SHARED EmpName$
  DIM SHARED PayPFreq$(1 TO 7)

  PRCalcMenu

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
    ELSE    'deal with zero gross wage, but has an addtional earning dist here
      
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
      TransRec.RetireAmt = 0
      TransRec.MatchRetAmt = 0
    END SELECT
  ELSE
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
    TransRec.SocTaxAmt = Round#((TransRec.SocGrossPay * FEDTAX(1).FTMEMPSS) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FEDTAX(1).FTMEMPSS) * .01)
      TransRec.SocTaxAmt = Round#(TransRec.SocTaxAmt + TaxFring#)
    END IF
    TransRec.MatchSocAmt = Round#((TransRec.SocGrossPay * FEDTAX(1).FTMEMRSS) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FEDTAX(1).FTMEMRSS) * .01)
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
    TransRec.MedTaxAmt = Round#((TransRec.MedGrossPay * FEDTAX(1).FTMEMPM) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FEDTAX(1).FTMEMPM) * .01)
      TransRec.MedTaxAmt = Round#(TransRec.MedTaxAmt + TaxFring#)
    END IF
    TransRec.MatchMedAmt = Round#((TransRec.MedGrossPay * FEDTAX(1).FTMEMRM) * .01)
    IF TransRec.TaxFring > 0 THEN
      TaxFring# = Round#((TransRec.TaxFring * FEDTAX(1).FTMEMRM) * .01)
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
      GOSUB CalcStaTax
      '12-13-94 Added code to process taxable fringe field
      IF TransRec.TaxFring > 0 THEN
        PriorStaTax# = TransRec.StaTaxAmt
        'this variable has to remain this name
        'because it is used in the gosub to calcStatax
        AnnualizedStaGross# = TransRec.TaxFring
        GOSUB CalcStaTax
        TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + PriorStaTax#)
      END IF
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
  
  IF TransRec.GrossPay <= 0 THEN
    TransRec.TActive = False
  END IF
  
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
    TaxableAmtFed# = Round#(AnnualizedFedGross# - (EMP2Rec(1).EMPFEDA * FEDTAX(1).FTMSDAA))
    IF TaxableAmtFed# < FEDTAX(1).FTM(3, 1) THEN
      TransRec.FedTaxAmt = 0
    ELSE
      FOR Cnt = 1 TO 10
        IF FEDTAX(1).FTM(3, Cnt) > TaxableAmtFed# THEN EXIT FOR
      NEXT
      Cnt = Cnt - 1
      TaxableAmtFed# = Round#(TaxableAmtFed# - FEDTAX(1).FTM(3, Cnt))
      TransRec.FedTaxAmt = Round#(Round#(FEDTAX(1).FTM(1, Cnt) + (TaxableAmtFed# * (FEDTAX(1).FTM(2, Cnt) * .01))) / PayFreq)
    END IF
  CASE "S", ""
    IF EMP2Rec(1).EMPFEDA < 0 THEN EMP2Rec(1).EMPFEDA = 0
    TaxableAmtFed# = Round#(AnnualizedFedGross# - (EMP2Rec(1).EMPFEDA * FEDTAX(1).FTSSDAA))
    IF TaxableAmtFed# < FEDTAX(1).FTS(3, 1) THEN
      TransRec.FedTaxAmt = 0
    ELSE
      FOR Cnt = 1 TO 10
        IF FEDTAX(1).FTS(3, Cnt) > TaxableAmtFed# THEN EXIT FOR
      NEXT
      Cnt = Cnt - 1
      TaxableAmtFed# = Round#(TaxableAmtFed# - FEDTAX(1).FTS(3, Cnt))
      TransRec.FedTaxAmt = Round#(Round#(FEDTAX(1).FTS(1, Cnt) + (TaxableAmtFed# * (FEDTAX(1).FTS(2, Cnt) * .01))) / PayFreq)
    END IF
  END SELECT
  IF EMP2Rec(1).EMPFEDAA > 0 THEN
    TransRec.FedTaxAmt = Round#(TransRec.FedTaxAmt + EMP2Rec(1).EMPFEDAA)
  END IF
  
  RETURN
  
CalcStaTax:
  
  SELECT CASE TaxText$(0)
  CASE "OK"
    ''$INCLUDE: 'OKstatax.bi'
    
  CASE "NC", "VA"
    '$INCLUDE: 'ncstatax.bi'
    
  CASE "SC"
    ''$INCLUDE: 'SCstatax.bi'
    
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

SUB ManualMenu
  
  DO
    PayMenu 4, Choice, 3
    
    SELECT CASE Choice
    CASE 1
      EntryType = Manual
      PCGetEmpNum LastEmpNum
    CASE 2
      EntryType = Manual
      'PCPrintManRegister             'unrem
    CASE 3
      EntryType = Manual  'Manuauaual=2
      PostTransactions 2
    END SELECT
  LOOP UNTIL Choice = EscKey
  Choice = 0
  
END SUB

SUB ParseManual2Trans (TransRec() AS TransRecType, ManTrans AS ManualTransRecType)
  'STOP
  TransRec(1).PayPdStart = ManTrans.PDSTART
  TransRec(1).PayPdEnd = ManTrans.PDEND
  
  TransRec(1).CheckDate = ManTrans.CHKDATE
  
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
  
  TransRec(1).FedTaxAmt = ManTrans.FEDTAX
  TransRec(1).StaTaxAmt = ManTrans.STATAX
  TransRec(1).SocTaxAmt = ManTrans.SOCTAX
  TransRec(1).MedTaxAmt = ManTrans.MEDTAX
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

SUB ParseTrans2Manual (TransRec() AS TransRecType, ManTrans AS ManualTransRecType)
  
  ManTrans.PDSTART = TransRec(1).PayPdStart
  ManTrans.PDEND = TransRec(1).PayPdEnd
  ManTrans.CHKDATE = TransRec(1).CheckDate
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
  
  ManTrans.FEDTAX = TransRec(1).FedTaxAmt
  ManTrans.STATAX = TransRec(1).StaTaxAmt
  ManTrans.SOCTAX = TransRec(1).SocTaxAmt
  ManTrans.MEDTAX = TransRec(1).MedTaxAmt
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

SUB PCEdScrnCalc (TransRec() AS TransRecType, TransRecNo, EmpName$, EmpRecNo)
  
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
  
  ParseTrans2ScrnCalc TransRec(), ScrnCalc()
  
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
          'ParseScrnCalc2Trans           'TransRec', ScrnCalc(1)
          ParseScrnCalc2Trans TransRec(), ScrnCalc()
          '10-01-94 03 refixed   Added fix to adjust fed gross and state gross
          'if ret amt was edited
          IF OldRetAmt# <> Value#(FormESC$(ScrnRetFld, 0), ErrCode) THEN
            DifRetAmt# = OldRetAmt# - Value#(FormESC$(ScrnRetFld, 0), ErrCode)
            TransRec(1).FedGrossPay = Round(TransRec(1).FedGrossPay + DifRetAmt#)
            TransRec(1).StaGrossPay = Round(TransRec(1).StaGrossPay + DifRetAmt#)
          END IF
          TransRec(1).TActive = True
          UpDateTransFile TransRec(), TransRecNo
        CASE False              'nope don't save abandon edit
          ExitFlag = True
        CASE 1  'oops continue editing
          Action = 2
          ExitFlag = False
        END SELECT
      CASE F3Key
        IF PCDelFromPay THEN
          TransRec(1).TActive = 0
          UpDateTransFile TransRec(), TransRecNo
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
  ParseScrnCalc2Trans TransRec(), ScrnCalc()
  'recalc tax's, net pay, deductions, etc.
  CalcPay TransRec(1), TransRecNo, True
  'copy transaction data back to scrn calc type variable
  'ParseTrans2ScrnCalc
  ParseTrans2ScrnCalc TransRec(), ScrnCalc()
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
  REDIM SalInput(1)  AS SalaryInputType
  REDIM HourInput(1) AS HourlyInputType
  
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
      ParseTrans2Hourly TransRec(), HourInput()
      BCopy VARSEG(HourInput(1)), VARPTR(HourInput(1)), SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), TimeRecLen, 0
      UnPackBuffer 0, 0, FormEET$(), FldEET()
    CASE "S"    'the name of this field in the
      TotPct = FldNum("TOTOTHRS", FldEET())     'salary form is wrong
      TOTAERN = FldNum("TOTAERN", FldEET())
      DstAErn = FldNum("TOTEADST", FldEET())
      TimeRecLen = LEN(SalInput(1))
      FormEET$(0, 0) = SPACE$(TimeRecLen)
      BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(SalInput(1)), VARPTR(SalInput(1)), TimeRecLen, 0
      ParseTrans2Salary TransRec(), SalInput()
      BCopy VARSEG(SalInput(1)), VARPTR(SalInput(1)), SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), TimeRecLen, 0
      UnPackBuffer 0, 0, FormEET$(), FldEET()
    END SELECT
    
  ELSE          'MANUAL do manual transaction setup here
    FormEET$(0, 0) = SPACE$(ManRecLen)
    IF OTActive THEN
      ParseTrans2Manual TransRec(), ManTrans(1)
    ELSE
      ManTrans(1).CHKDATE = -32767
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
              GOSUB ShowBlankError
              FrmEET(1).FldNo = FrmEET(1).PrevFld
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
        
      CASE F10Key
        
        IF FrmEET(1).KeyCode = F0Key THEN
          'fix for MANUAL
          IF EntryType = Normal THEN
            '*************
            'make sure there are accounts for every amount
            '12-12-94  Adjusted field offsets for taxable fringe
            '01-15-96  Fixed offset bug in alt-earnings (Amts with no acct.)
            '09-19-96  Fixed again Fixes for Both Salary & Hourly
            SELECT CASE TransRec(1).PayType
            CASE "H"
              IF Value(FormEET$(TOTHrs, 0), Oops) <> Value(FormEET$(DstHrs, 0), Oops) THEN DstError = True
              IF Value(FormEET$(OTHRSPD, 0), Oops) <> Value(FormEET$(DstOTHrs, 0), Oops) THEN DstError = True
              IF Value(FormEET$(TOTAERN, 0), Oops) <> Value(FormEET$(DstAErn, 0), Oops) THEN DstError = True
              'check regwage dists
              FOR Cnt = 11 TO 32 STEP 3
                'IF Value(FormEET$(Cnt, 0), Oops) > 0 AND LEN(QPTrim$(FormEET$(Cnt - 1, 0))) = 0 THEN
                IF Value(FormEET$(Cnt + 1, 0), Oops) > 0 OR Value(FormEET$(Cnt + 2, 0), Oops) > 0 THEN
                  IF LEN(QPTrim$(FormEET$(Cnt, 0))) = 0 THEN
                    FrmEET(1).FldNo = Cnt
                    AcctError = True
                  END IF
                END IF
              NEXT
              'check alt-earn dists
              FOR Cnt = 41 TO 51 STEP 2
                IF Value(FormEET$(Cnt, 0), Oops) > 0 AND LEN(QPTrim$(FormEET$(Cnt - 1, 0))) = 0 THEN
                  FrmEET(1).FldNo = Cnt - 1
                  AcctError = True
                END IF
              NEXT
              
            CASE "S"
              'IF FormEET$(1, 0) = "Y" THEN
              IF Value(FormEET$(TotPct, 0), Oops) <> 100 THEN DstError = True
              IF Value(FormEET$(TOTAERN, 0), Oops) <> Value(FormEET$(DstAErn, 0), Oops) THEN DstError = True
              
              FOR Cnt = 6 TO 20 STEP 2
                IF Value(FormEET$(Cnt, 0), Oops) > 0 AND LEN(QPTrim$(FormEET$(Cnt - 1, 0))) = 0 THEN
                  FrmEET(1).FldNo = Cnt - 1
                  AcctError = True
                END IF
              NEXT
              FOR Cnt = 27 TO 37 STEP 2
                IF Value(FormEET$(Cnt, 0), Oops) > 0 AND LEN(QPTrim$(FormEET$(Cnt - 1, 0))) = 0 THEN
                  FrmEET(1).FldNo = Cnt - 1
                  AcctError = True
                END IF
              NEXT
            END SELECT
          ELSE     'Manual entry field checks
'042898 Added more error checking for invalid entries
            'this will check the dates
            FOR ZZCnt = 1 TO 3
              IF Date2Num(FormEET$(ZZCnt, 0)) < 0 THEN
                GOSUB ShowBlankError
                GOTO ExitManCheck
              END IF
            NEXT
            'this will check the hours

            FOR ZZCnt = 4 TO 11
              ThisFld$ = QPTrim$(FormEET$(ZZCnt, 0))
              IF LEN(ThisFld$) = 0 THEN
                GOSUB ShowBlankError
                GOTO ExitManCheck
              END IF
            NEXT
            FOR ZZCnt = 12 TO 18 STEP 2
              ThisFld$ = QPTrim$(FormEET$(ZZCnt, 0))
              IF LEN(ThisFld$) = 0 THEN
                GOSUB ShowBlankError
                GOTO ExitManCheck
              END IF
            NEXT
            FOR ZZCnt = 19 TO NumFlds
              ThisFld$ = QPTrim$(FormEET$(ZZCnt, 0))
              IF LEN(ThisFld$) = 0 THEN
                GOSUB ShowBlankError
                EXIT FOR
              END IF
            NEXT
ExitManCheck:
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
        ELSEIF FrmEET(1).KeyCode = EscKey AND EntryType = Manual AND OTActive = False THEN
          SaveFlag = False
          TransRec(1).TActive = 0
          UpDateTransFile TransRec(), TransRecNo
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
              ParseHourly2Trans TransRec(), HourInput()
            CASE "S"
              BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(SalInput(1)), VARPTR(SalInput(1)), TimeRecLen, 0
              ParseSalary2Trans TransRec(), SalInput()
            END SELECT
          CASE Manual
            TransRec(1).TActive = True
            BCopy SSEG(FormEET$(0, 0)), SADD(FormEET$(0, 0)), VARSEG(ManTrans(1)), VARPTR(ManTrans(1)), ManRecLen, 0
            ParseManual2Trans TransRec(), ManTrans(1)
            'fix for manual  6/09/94
            IF PrdDefRec(1).MACTIVE = 0 THEN
              PrdDefRec(1).MACTIVE = True
              FPutAH PPDefaultFileName, PrdDefRec(1), LEN(PrdDefRec(1)), 1
            END IF
          END SELECT
          UpDateTransFile TransRec(), TransRecNo
          'do normal scrn calc screen if not a manual transaction
          IF FrmEET(1).KeyCode = F0Key AND EntryType = Normal THEN
            PCEdScrnCalc TransRec(), TransRecNo, EmpName$, EmpRecNo
          END IF
        CASE False              'nope don't save abandon edit
          ExitFlag = True
        CASE 1                  'Continue editing
          Action = 2
          ExitFlag = False
        END SELECT
      CASE F3Key
        IF PCDelFromPay THEN
          TransRec(1).TActive = 0
          UpDateTransFile TransRec(), TransRecNo
          ExitFlag = True
        END IF
      CASE ELSE
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
NoGood:
  ERASE TempScrn, FormEET$, FldEET
  ERASE TransRec, FrmEET, EarnDes$
  ERASE TempScrn, ManTrans, SalInput
  ERASE HourInput

  
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
      CalcFields 0, 1, FormEET$(), FldEET()
      CalcFields 0, 12, FormEET$(), FldEET()
      CalcFields 0, 13, FormEET$(), FldEET()
      CalcFields 0, 42, FormEET$(), FldEET()
    CASE "S"
      CalcFields 0, 7, FormEET$(), FldEET()
      CalcFields 0, 23, FormEET$(), FldEET()
      CalcFields 0, 28, FormEET$(), FldEET()
    END SELECT
  ELSE          'fix for MANUAL
    FOR Cnt = 1 TO 12
      QPrintRC QPTrim$(DedCodes(Cnt).DCDESC1), Cnt + 4, 44, Cnf.ActivCh
    NEXT
  END IF
  ShowCursor
RETURN

ShowBlankError:
  FrmEET(1).FldNo = ZZCnt
  FrmEET(1).KeyCode = 0
  FrmEET(1).StartEl = 0
  SaveFlag = 1
  SaveScrn TempScrn()
  DisplayMiscScrn NoBlank
  WaitForAction
  RestScrn TempScrn()
RETURN

END SUB

SUB PCGetEmp2Rec (EmpRecNo%, ErrorCode)
  
  FOpenS EmpData2Name, Handle   'open employee data file
  FGetRTA Handle, EMP2Rec(1), CLNG(EmpRecNo), LEN(EMP2Rec(1))
  
  FClose Handle
  
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
        QPrintRC DedCodes(Cnt).DCDESC1, 14 + Cnt, 12, -1
      CASE 5 TO 8
        QPrintRC DedCodes(Cnt).DCDESC1, 10 + Cnt, 35, -1
      CASE 9 TO 12
        QPrintRC DedCodes(Cnt).DCDESC1, 6 + Cnt, 59, -1
      END SELECT
      FldOffSet = FldOffSet + 2
    NEXT
    
    ZZ = 12
    FOR Cnt = 1 TO 2
      QPrintRC ErnCodes(Cnt).ERNCODE1, 21, ZZ, -1
      ZZ = ZZ + 23
    NEXT
    QPrintRC ErnCodes(Cnt).ERNCODE1, 21, ZZ + 1, -1
    
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

SUB PostTransactions (EntryType)
  
  
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
  'REDIM GLSetUpRec(1)    AS GLSetupRecType
  
  TransRecLen = LEN(TransRec(1))
  Emp2RecLen = LEN(EmpRec2(1))
  Emp3RecLen = LEN(EmpRec3(1))
  CheckRecLen = LEN(Check(1))
  OSChkRecLen = LEN(OSIFRec(1))
  
  FGetAH SysFileName, PSysRec(1), LEN(PSysRec(1)), 1
  
  REDIM GLSetUpRec(1)    AS GLSetupRecType
  GLSetupRecLen = LEN(GLSetUpRec(1))
  
  GLSetUpName$ = QPTrim$(PSysRec(1).CITIDIR) + "\GLSETUP.DAT"
  IF Exist(GLSetUpName$) THEN
    FGetAH GLSetUpName$, GLSetUpRec(1), LEN(GLSetUpRec(1)), 1
  END IF
  
  '05/10/94   Fix to set false after posting
  '    -vvv-  is used as a flag to menu entry points
  '06/09/94   Fix for manual trans
  
  PDRLen = LEN(PDR(1))
  FGetAH PPDefaultFileName, PDR(1), PDRLen, 1   'get pay period defaults
  SELECT CASE EntryType
  CASE Normal
    PrdDefRec(1).PACTIVE = False
    PDR(1).PACTIVE = False      'set active flag = false
  CASE Manual
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
    OSNumRec = (FileSize(OSChkFile$) \ OSChkRecLen) + 1
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
        GLDate = Check(1).CheckDate
        ReplaceString GLIFDate$, "-", "/"
        ReplaceString GLIFDate$, "1999", "99"
        ReplaceString GLIFDate$, "2000", "00"
        FOR ZZ = 1 TO 9
          ReplaceString GLIFDate$, "200" + QPTrim$(STR$(ZZ)), "0" + QPTrim$(STR$(ZZ))
        NEXT
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
      
      '022498
      EmpRec2(1).HolUsed = Round(EmpRec2(1).HolUsed + TransRec(1).HOLHOURS)
      EmpRec2(1).HOLBAL = Round(EmpRec2(1).HOLERN - EmpRec2(1).HolUsed)
      
      EmpRec2(1).PerUsed = Round(EmpRec2(1).PerUsed + TransRec(1).PerHours)
      EmpRec2(1).PERBAL = Round(EmpRec2(1).PERERN - EmpRec2(1).PerUsed)
      '***************
      
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
        OSIFRec(1).CHKDATE = GLIFDate$
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
  
  REDIM GLRec(1)        AS GLTransRecType
  REDIM GLIFRec14(1)    AS GLIFDataType14
  
  GLIFRecLen = LEN(GLIFRec14(1))
  GLRecLen = LEN(GLRec(1))
  
  'process gl transfer file
  
  PRIF$ = "PRIF.DAT"
  NextRec& = 1
  NumOfRecs = FileSize("TEMPIF.DAT") \ GLIFRecLen
  
  FCreate PRIF$
  FOpenS "TEMPIF.DAT", IHandle
  FOpenS PRIF$, OHandle
  'fix me
  FOR Cnt = 1 TO NumOfRecs
    FGetRTA IHandle, GLIFRec14(1), CLNG(Cnt), GLIFRecLen
    'GLRec(1).AcctNum = FmtAcct$(GLIFRec14(1).TranAcct, GLSetUpRec(1).FundLen, GLSetUpRec(1).AcctLen, GLSetUpRec(1).DetLen)
    'unrem
    GLRec(1).TrDate = GLDate
    GLRec(1).Desc = GLIFRec14(1).TranDesc
    GLRec(1).CrAmt = GLIFRec14(1).CrAmt
    GLRec(1).DrAmt = GLIFRec14(1).DrAmt
    GLRec(1).Src = GLIFRec14(1).Source
    FPutRTA OHandle, GLRec(1), NextRec&, GLRecLen
    NextRec& = NextRec& + 1
  NEXT
  
  FClose IHandle
  FClose OHandle
  
  'Post2GL PRIF$, PSysRec(), BadAccts   'unrem
  
  IF BadAccts <> 0 THEN
    'there was a posting to gl error
  ELSE
    KillFile "PRIF.DAT"
  END IF
  
  IF Exist("TEMPIF.DAT") THEN
    KillFile "LASTIF.DAT"
    NAME "TEMPIF.DAT" AS "LASTIF.DAT"
  END IF
  
  
  'done with gl transfer file
  
  ERASE PSysRec, TransRec, EmpRec2, EmpRec3, Check, PDR
  ERASE GLIFRec14, OSIFRec
  
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

  REDIM DedCodes(1 TO 12)   AS DedCodeRecType
  REDIM ErnCodes(1 TO 3)    AS ErnCodeRecType
  REDIM RetireRec(1 TO 6)   AS RetireRecType
  REDIM StateTax(1)         AS StateTaxRecType
  REDIM FEDTAX(1)           AS FederalTaxRecType
  REDIM EICRec(1)           AS EICRecType
  REDIM PrdDefRec(1)        AS PeriodDefaultRecType

  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  FGetAH ErnCodeFileName, ErnCodes(1), LEN(ErnCodes(1)), 3
  FGetAH RetireFileName, RetireRec(1), LEN(RetireRec(1)), 6
  FGetAH FederalTaxFileName, FEDTAX(1), LEN(FEDTAX(1)), 1
  FGetAH StateTaxFileName, StateTax(1), LEN(StateTax(1)), 1
  FGetAH EICFileName, EICRec(1), LEN(EICRec(1)), 2
  FGetAH PPDefaultFileName, PrdDefRec(1), LEN(PrdDefRec(1)), 1
  
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
    PayMenu PayProcessMenu, Choice, 9
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
      IF PrdDefRec(1).PACTIVE OR Choice = 1 OR Choice = 9 THEN
        SELECT CASE Choice
        CASE 1
          'AccruLeave False       'unrem
        CASE 3  'payroll transaction entry
          EntryType = Normal
          PCGetEmpNum LastEmpNum
        CASE 4
          'IF SplitFlag THEN
            'PCPrintPayRegisterS     'unrem me
          'ELSE
          '  PCPrintPayRegister      'unrem me
          'END IF
        CASE 5
          FCreate PayrollStatusFile
          RUN "prcheck"
        CASE 6
          'PRDraftMenu
        CASE 7
          EntryType = Normal
          PostTransactions Normal
        CASE 8
          RUN "PRVOIDCK"

        CASE 9  'manual transaction entry
'          IF PrdDefRec(1).PACTIVE THEN
'            'No manual transaction durning payroll active error
'            BlockClear
'            DisplayMiscScrn NoManualNow
'            WaitForAction
'          ELSE
            ManualMenu
'          END IF
          
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

