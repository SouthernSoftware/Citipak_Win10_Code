DEFINT A-Z
DECLARE SUB RPTSetupPRN (RPTNum%, FHandle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB PCPrintManRegister ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOFRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION QPStrL$ (LongNum&)

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

''$INCLUDE: 'FieldInf.bi'    '        "          field information
'$INCLUDE: 'FormEdit.bi'    '        "          form editing information

'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'ErnCodes.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'PRSYSCTR.Bi'
'$INCLUDE: 'PRPPDEF.Bi'
'$INCLUDE: 'PRTRANS.Bi'
'$INCLUDE: 'prif.bi'

CONST False = 0, True = NOT False

SUB PCPrintManRegister

  RptTitle$ = "Manual Register Report"
  ShowProcessingScrn RptTitle$

  REDIM TransRec(1) AS TransRecType
  REDIM EmpRec1(1)  AS EmpData1Type
  REDIM Unit(1)     AS UnitFileRecType

  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  'REDIM ErnCodes(1 TO 3) AS ErnCodeRecType

  REDIM TotDeds(1 TO 12) AS DOUBLE
  'REDIM TotErns(1 TO 3) AS DOUBLE

  REDIM EDRHrs(1)   AS STRING * 11
  REDIM EDOHrs(1)   AS STRING * 11
  REDIM EDRPay(1)   AS STRING * 11
  REDIM EDOPay(1)   AS STRING * 11
  REDIM EDEarn(1)   AS STRING * 11
  REDIM EDGroP(1)   AS STRING * 11

  REDIM EDSAmt(1)   AS STRING * 11
  REDIM EDMAmt(1)   AS STRING * 11
  REDIM EDRAmt(1)   AS STRING * 11

  'REDIM ENum(1)     AS STRING * 14
  REDIM ENum(1)     AS STRING * 13
  REDIM EName(1)    AS STRING * 32

  REDIM MTDates(1 TO 3)  AS STRING * 22
  REDIM ChkNo(1)         AS STRING * 10
  REDIM DAcct(1 TO 4)    AS STRING * 22
  REDIM DAmts(1 TO 4)    AS STRING * 22

  REDIM FedGro(1)   AS STRING * 11
  REDIM StaGro(1)   AS STRING * 11
  REDIM SocGro(1)   AS STRING * 11
  REDIM MedGro(1)   AS STRING * 11
  REDIM RetGro(1)   AS STRING * 11

  REDIM BRat(1)     AS STRING * 11
  REDIM ORat(1)     AS STRING * 11
  REDIM Fill11(1)   AS STRING * 11

  REDIM SCnt(1)     AS STRING * 11
  REDIM HCnt(1)     AS STRING * 11

  REDIM RHrs(1)     AS STRING * 11
  REDIM VHrs(1)     AS STRING * 11
  REDIM SHrs(1)     AS STRING * 11
  REDIM HHrs(1)     AS STRING * 11
  REDIM CHrs(1)     AS STRING * 11
  REDIM THrs(1)     AS STRING * 11
  REDIM OTHrs(1)    AS STRING * 11
  REDIM OTPaid(1)   AS STRING * 11
  REDIM OTComp(1)   AS STRING * 11

  REDIM RErnP(1)    AS STRING * 11
  REDIM OErnP(1)    AS STRING * 11

  REDIM GPayP(1)    AS STRING * 11
  REDIM SSTaxP(1)   AS STRING * 11
  REDIM MTaxP(1)    AS STRING * 11
  REDIM FTaxP(1)    AS STRING * 11
  REDIM STaxP(1)    AS STRING * 11
  REDIM RetirP(1)   AS STRING * 11
  REDIM NetPayP(1)  AS STRING * 11
  REDIM Ded(1)      AS STRING * 11

  REDIM EEicP(1)   AS STRING * 11
  REDIM Ern(1)      AS STRING * 11
  REDIM Pg(1)       AS STRING * 5
  REDIM EMPLine(1) AS STRING * 132

  REDIM Dash(1) AS STRING * 132

  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  'FGetAH ErnCodeFileName, ErnCodes(1), LEN(ErnCodes(1)), 3

  Image0$ = "####"
  Image$ = ",#####.##"
  Image3$ = ",######.##"
  Image5$ = ",#######.##"

  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GRetGross# = 0

  LSET Dash(1) = STRING$(132, "-")
  LSET Fill11(1) = ""

  TransRecLen = LEN(TransRec(1))
  Emp1RecLen = LEN(EmpRec1(1))

  IdxRecLen = 2  'LEN(TempIdxNRec)
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOFRecs = IdxFileSize& \ IdxRecLen

  SalCnt = 0
  HrlCnt = 0

  LineCnt = 0
  MaxLines = 45
  Page = 1

  REDIM IdxBuff(1 TO NumOFRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOFRecs

  DTitle$ = ""
  FOR Cnt = 1 TO 12
    TDed$ = QPTrim$(DedCodes(Cnt).DCDESC1)
    IF LEN(TDed$) > 0 THEN
      LastDed = LastDed + 1
      RSET Ded(1) = TDed$
      DTitle$ = DTitle$ + Ded(1)
    ELSE
      EXIT FOR
    END IF
  NEXT

  IF LastDed < 12 THEN
    DTitle$ = SPACE$(11 * (12 - LastDed)) + DTitle$
  END IF


'---------------------------------------------
  ETitle$ = "   Reg Earn   O/T Earn                        Gross Pay        EIC    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"
  SumHeader2$ = ETitle$
            
'------------------------------------------------------------------

  FCreate ManRegisterRptName

  FOpenS ManRegisterRptName, RHandle

  RPTSetupPRN 12, RHandle

  FOpenS EmpData1Name, NHandle   'open it
  FOpenS TransWorkFileName, THandle   'open it

  GOSUB PrintManualHeader

  FOR Cnt = 1 TO NumOFRecs
    FGetRTA THandle, TransRec(1), CLNG(IdxBuff(Cnt)), TransRecLen
    IF TransRec(1).TActive = True THEN
      FGetRTA NHandle, EmpRec1(1), CLNG(IdxBuff(Cnt)), Emp1RecLen
      GOSUB SumAndPrintTime
      LineCnt = LineCnt + 7
      IF LineCnt >= MaxLines AND Cnt < NumOFRecs THEN
        LineCnt = 0
        FPut RHandle, FF$
        GOSUB PrintManualHeader
      END IF
    END IF
    ShowPctComp Cnt, NumOFRecs
  NEXT

  GOSUB PrintSumTotal
  RPTSetupPRN 0, RHandle
  FClose THandle
  FClose NHandle
  FClose RHandle

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

  ERASE TransRec, EmpRec1, Unit
  ERASE DedCodes, TotDeds, EDRHrs, EDOHrs
  ERASE EDRPay, EDOPay, EDEarn, EDGroP, EDSAmt, EDMAmt, EDRAmt, ENum, EName
  ERASE BRat, ORat, Fill11, SCnt, HCnt, RHrs, VHrs, SHrs, HHrs, CHrs, THrs
  ERASE OTHrs, OTPaid, OTComp, RErnP, OErnP, GPayP, SSTaxP, MTaxP, FTaxP, STaxP
  ERASE RetirP, NetPayP, Ded, Ern, Pg, Dash
  ERASE IdxBuff
'----------------------------------------------------------------------------

  PrintRptFile RptTitle$, ManRegisterRptName, 1, RetCode, 1       'unrem

EXIT SUB


PrintManualHeader:
  RSET Pg(1) = STR$(Page)
  FPut RHandle, Unit(1).UFEMPR + SPACE$(87) + "Page:" + Pg(1) + CrLf$
  FPut RHandle, "Manual Transaction Register" + CrLf$
  FPut RHandle, "Report Date: " + DATE$ + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "Employee No   Name                             Beg Date              End Date              Chk Date                         Check No" + CrLf$
  FPut RHandle, "  Base Rate   O/T Rate               Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid   O/T Comp" + CrLf$
  FPut RHandle, ETitle$ + CrLf$
  FPut RHandle, DTitle$ + CrLf$
  FPut RHandle, "                Dist 1                Dist 2                Dist 3                Dist 4  Fed Gross  Sta Gross  Med Gross  Soc Gross" + CrLf$
  FPut RHandle, "                   Amt                   Amt                   Amt                   Amt                                   Ret Gross" + CrLf$
  FPut RHandle, Dash(1) + CrLf$
  LincCnt = LineCnt + 11
  Page = Page + 1
RETURN

SumAndPrintTime:

 REGHRS# = Round#(REGHRS# + TransRec(1).RegHrsWork)
 VACHRS# = Round#(VACHRS# + TransRec(1).VacUsed)
 SICKHRS# = Round#(SICKHRS# + TransRec(1).SickUsed)
 HOLHRS# = Round#(HOLHRS# + TransRec(1).HOLHOURS)
 COMPHRS# = Round#(COMPHRS# + TransRec(1).CompUsed)

 TotalHrs# = Round(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)

 TotEIC# = Round#(TotEIC# + TransRec(1).EICAmt)

'-=-=-=-=-=-=-=
 TOTHrs# = Round#(TOTHrs# + TransRec(1).OTHours)

 IF TransRec(1).OTHrsPaid <> 0 THEN
   TOTPaid# = Round#(TOTPaid# + TransRec(1).OTHrsPaid)
 END IF

 TOTComp# = Round#(TOTComp# + TransRec(1).OT2Comp)

 TRegWage# = Round#(TRegWage# + TransRec(1).TotRegWage)

 IF TransRec(1).TotOTWage > 0 THEN
   TOTWage# = Round#(TOTWage# + TransRec(1).TotOTWage)
 END IF

 GPay# = Round#(GPay# + TransRec(1).GROSSPAY)
 SSTax# = Round#(SSTax# + TransRec(1).SocTaxAmt)
 MTax# = Round#(MTax# + TransRec(1).MedTaxAmt)
 FTax# = Round#(FTax# + TransRec(1).FedTaxAmt)
 STax# = Round#(STax# + TransRec(1).StaTaxAmt)

 IF TransRec(1).RetireAmt <> 0 THEN
   RETTOT# = Round(RETTOT# + TransRec(1).RetireAmt)
 END IF
 'TOTDED# = Round(TOTDED# + TransRec(1).TotDedAmt)

 TNetPay# = Round#(TNetPay# + TransRec(1).NETPAY)

 GFedGross# = Round#(GFedGross# + TransRec(1).FedGrossPay)
 GStaGross# = Round#(GStaGross# + TransRec(1).StaGrossPay)
 GSocGross# = Round#(GSocGross# + TransRec(1).SocGrossPay)
 GMedGross# = Round#(GMedGross# + TransRec(1).MedGrossPay)
 GRetGross# = Round#(GRetGross# + TransRec(1).RetGrossPay)

 RSET FedGro(1) = FUsing(STR$(TransRec(1).FedGrossPay), Image3$)
 RSET StaGro(1) = FUsing(STR$(TransRec(1).StaGrossPay), Image3$)
 RSET MedGro(1) = FUsing(STR$(TransRec(1).MedGrossPay), Image3$)
 RSET SocGro(1) = FUsing(STR$(TransRec(1).SocGrossPay), Image3$)
 RSET RetGro(1) = FUsing(STR$(TransRec(1).RetGrossPay), Image3$)

 LSET ENum(1) = LTRIM$(EmpRec1(1).EMPNO)
 LSET EName(1) = QPTrim$(EmpRec1(1).EMPLNAME) + ", " + QPTrim$(EmpRec1(1).EMPFNAME)
 RSET BRat(1) = FUsing(STR$(TransRec(1).BaseRate), Image3$)
 RSET ORat(1) = FUsing(STR$(TransRec(1).OTRate), Image3$)
'look here

 RSET RHrs(1) = FUsing(STR$(TransRec(1).RegHrsWork), Image$)

 RSET VHrs(1) = FUsing(STR$(TransRec(1).VacUsed), Image$)
 RSET SHrs(1) = FUsing(STR$(TransRec(1).SickUsed), Image$)
 RSET HHrs(1) = FUsing(STR$(TransRec(1).HOLHOURS), Image$)
 RSET CHrs(1) = FUsing(STR$(TransRec(1).CompUsed), Image$)

 TotHrsPaid# = TransRec(1).RegHrsPaid + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed

 RSET THrs(1) = FUsing(STR$(TotHrsPaid#), Image$)

 RSET OTHrs(1) = FUsing(STR$(TransRec(1).OTHours), Image$)
 RSET OTPaid(1) = FUsing(STR$(TransRec(1).OTHrsPaid), Image$)
 RSET OTComp(1) = FUsing(STR$(TransRec(1).OT2Comp), Image$)

 RSET EEicP(1) = FUsing(STR$(TransRec(1).EICAmt), Image3$)

 SELECT CASE TransRec(1).PayType
   CASE "S"
     RSET RHrs(1) = "Salaried"
     SalCnt = SalCnt + 1
   CASE ELSE
     HrlCnt = HrlCnt + 1
 END SELECT

'=======
 RSET RErnP(1) = FUsing(STR$(TransRec(1).TotRegWage), Image3$)
 RSET OErnP(1) = FUsing(STR$(TransRec(1).TotOTWage), Image3$)

 RSET GPayP(1) = FUsing(STR$(TransRec(1).GROSSPAY), Image3$)
 RSET SSTaxP(1) = FUsing(STR$(TransRec(1).SocTaxAmt), Image3$)
 RSET MTaxP(1) = FUsing(STR$(TransRec(1).MedTaxAmt), Image3$)
 RSET FTaxP(1) = FUsing(STR$(TransRec(1).FedTaxAmt), Image3$)
 RSET STaxP(1) = FUsing(STR$(TransRec(1).StaTaxAmt), Image3$)

 RSET RetirP(1) = FUsing(STR$(TransRec(1).RetireAmt), Image3$)

 RSET NetPayP(1) = FUsing(STR$(TransRec(1).NETPAY), Image3$)

 SumDed$ = ""
 FOR Cnt2 = 1 TO LastDed
   IF TransRec(1).DAMT(Cnt2) <> 0 THEN
     TotDeds(Cnt2) = Round#(TotDeds(Cnt2) + TransRec(1).DAMT(Cnt2))
     RSET Ded(1) = FUsing(STR$(TransRec(1).DAMT(Cnt2)), Image3$)
   ELSE
     RSET Ded(1) = FUsing(" 0", Image3$)
   END IF
   SumDed$ = SumDed$ + Ded(1)
 NEXT
 IF LastDed < 12 THEN
   SumDed$ = SPACE$(11 * (12 - LastDed)) + SumDed$
 END IF

'----------------------------------------------
  SumErn$ = SPACE$(22)

  LSET MTDates(1) = Num2Date$(TransRec(1).PayPdStart)
  LSET MTDates(2) = Num2Date$(TransRec(1).PayPdEnd)
  LSET MTDates(3) = Num2Date$(TransRec(1).CheckDate)

  RSET ChkNo(1) = QPStrL(TransRec(1).CHECKNUM)

 FOR Cnt2 = 1 TO 4
   RSET DAcct(Cnt2) = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
   RSET DAmts(Cnt2) = FUsing(STR$(TransRec(1).TDist(Cnt2).DRWage), Image3$)
 NEXT
'-------------------------------------------------------
 'RSET EMPLine$(1) = EEicP(1)
 'MID$(EMPLine$(1), 1) = ENum(1) + EName(1) + MTDates(1) + MTDates(2) + MTDates(3) + ChkNo(1)
 'RSET EMPLine$(1) = EEicP(1)
 'EMPLine$ = ENum(1) + EName(1) + MTDates(1) + MTDates(2) + MTDates(3) + ChkNo(1)


 FPut RHandle, ENum(1) + EName(1) + MTDates(1) + MTDates(2) + MTDates(3) + Fill11(1) + ChkNo(1) + CrLf$
 FPut RHandle, BRat(1) + ORat(1) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
 FPut RHandle, CHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + OTComp(1) + CrLf$
 FPut RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + EEicP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1)
 FPut RHandle, RetirP(1) + NetPayP(1) + CrLf$
 'FPut RHandle, FedGro(1) + StaGro(1) + MedGro(1) + SocGro(1) + CrLf$
 FPut RHandle, SumDed$ + CrLf$
 FPut RHandle, DAcct(1) + DAcct(2) + DAcct(3) + DAcct(4) + FedGro(1) + StaGro(1) + SocGro(1) + MedGro(1) + CrLf$
 FPut RHandle, DAmts(1) + DAmts(2) + DAmts(3) + DAmts(4) + Fill11(1) + Fill11(1) + Fill11(1) + RetGro(1) + CrLf$
 FPut RHandle, CrLf$

RETURN

PrintSumTotal:
  RSET SCnt(1) = FUsing(STR$(SalCnt), Image0$)
  RSET HCnt(1) = FUsing(STR$(HrlCnt), Image0$)

  RSET THrs(1) = FUsing(STR$(TotalHrs#), Image3$)
  RSET RHrs(1) = FUsing(STR$(REGHRS#), Image3$)
  RSET VHrs(1) = FUsing(STR$(VACHRS#), Image3$)
  RSET SHrs(1) = FUsing(STR$(SICKHRS#), Image3$)
  RSET HHrs(1) = FUsing(STR$(HOLHRS#), Image3$)
  RSET CHrs(1) = FUsing(STR$(COMPHRS#), Image3$)
  RSET OTHrs(1) = FUsing(STR$(TOTHrs#), Image3$)
  RSET OTPaid(1) = FUsing(STR$(TOTPaid#), Image3$)
  RSET OTComp(1) = FUsing(STR$(TOTComp#), Image3$)

  RSET RErnP(1) = FUsing(STR$(TRegWage#), Image3$)
  RSET OErnP(1) = FUsing(STR$(TOTWage#), Image3$)

  RSET GPayP(1) = FUsing(STR$(GPay#), Image3$)
  RSET SSTaxP(1) = FUsing(STR$(SSTax#), Image3$)
  RSET MTaxP(1) = FUsing(STR$(MTax#), Image3$)
  RSET FTaxP(1) = FUsing(STR$(FTax#), Image3$)
  RSET STaxP(1) = FUsing(STR$(STax#), Image3$)

  RSET RetirP(1) = FUsing(STR$(RETTOT#), Image3$)
  RSET NetPayP(1) = FUsing(STR$(TNetPay#), Image3$)

  SumDed$ = ""
  FOR Cnt2 = 1 TO LastDed
    RSET Ded(1) = FUsing(STR$(TotDeds(Cnt2)), Image3$)
    SumDed$ = SumDed$ + Ded(1)
  NEXT
  IF LastDed < 12 THEN
    SumDed$ = SPACE$(11 * (12 - LastDed)) + SumDed$
  END IF
'---------------------------------------------------------
'  SumErn$ = ""
'  FOR Cnt2 = 1 TO LastErn
'    RSET Ern(1) = FUsing(STR$(TotErns(Cnt2)), Image3$)
'    SumErn$ = SumErn$ + Ern(1)
'  NEXT
'  IF LastErn < 3 THEN
'    SumErn$ = SPACE$(11 * (3 - LastErn)) + SumErn$
    SumErn$ = SPACE$(22)
'  END IF

'---------------------------------------------
' Page = Page + 1
  RSET EEicP(1) = FUsing(STR$(TotEIC#), Image5$)
  RSET Pg(1) = STR$(Page)
  FPut RHandle, FF$
  FPut RHandle, Unit(1).UFEMPR + SPACE$(87) + "Page:" + Pg(1) + CrLf$
  FPut RHandle, "Manual Transaction Summary" + CrLf$
  FPut RHandle, "Report Date: " + DATE$ + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, Dash(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "   Salaried     Hourly               Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid   O/T Comp" + CrLf$
  FPut RHandle, SCnt(1) + HCnt(1) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
  FPut RHandle, CHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + OTComp(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, SumHeader2$ + CrLf$ ' "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay" + CrLf$
  FPut RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + EEicP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1)
  FPut RHandle, STaxP(1) + RetirP(1) + NetPayP(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, DTitle$ + CrLf$
  FPut RHandle, SumDed$ + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "  Fed Gross  Sta Gross  Med Gross  Soc Gross  Ret Gross" + CrLf$

  RSET FTaxP(1) = FUsing(STR$(GFedGross#), Image5$)
  RSET STaxP(1) = FUsing(STR$(GStaGross#), Image5$)
  RSET MTaxP(1) = FUsing(STR$(GMedGross#), Image5$)
  RSET SSTaxP(1) = FUsing(STR$(GSocGross#), Image5$)
  RSET RetGro(1) = FUsing(STR$(GRetGross#), Image5$)

  FPut RHandle, FTaxP(1) + STaxP(1) + MTaxP(1) + SSTaxP(1) + RetGro(1) + CrLf$

  FPut RHandle, CrLf$
  FPut RHandle, Dash(1) + CrLf$
  FPut RHandle, FF$

RETURN

END SUB

