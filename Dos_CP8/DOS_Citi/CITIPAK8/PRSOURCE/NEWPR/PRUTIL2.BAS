DECLARE SUB edittrans ()
DECLARE SUB ResetTrPointers ()
DECLARE SUB ResetPR ()
DEFINT A-Z
'DECLARE SUB PCPickEmpList (TransRecNo%, EmpRecNo%)
'DECLARE SUB Post2BA (FileName$, BadTrans%)
'DECLARE SUB PCPrintPayRegisterS ()
'DECLARE SUB MakeTransInActive ()
'DECLARE FUNCTION PromptPeriodWasActive% ()
'DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION Date2Num (DateIn$)
'DECLARE FUNCTION Exist% (FileName$)
'DECLARE FUNCTION FileSize& (FileName$)
'DECLARE FUNCTION FldNum (FldName$, Fld() AS ANY)
'DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION Num2Date$ (Number%)
'DECLARE FUNCTION PCDelFromPay% ()
'DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (DblNum#)
'DECLARE FUNCTION Value# (e$, ErCode%)
'DECLARE SUB AccruLeave (SaveScrnFlag%)
'DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
'DECLARE SUB BlockClear ()
'DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
'DECLARE SUB CalcFields (StartOfForm, FldNo, Form$(), Fld() AS ANY)
'DECLARE SUB CalcPay (TransRec AS ANY, TransRecNo%, ReCalcFlag)
'DECLARE SUB CreateEmpTransRecs (RecNo%)
'DECLARE SUB CursorOff ()
'DECLARE SUB DisplayMiscScrn (ScrnName$)
'DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
'DECLARE SUB FClose (Handle%)
'DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
'DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
'DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
'DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
'DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
'DECLARE SUB HideCursor ()
'DECLARE SUB KillFile (FileName$)
'DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode, Attribute, ErrorCode)
'DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
'DECLARE SUB MakeDefaultTransActs ()
'DECLARE SUB ManualMenu ()
'DECLARE SUB PCEdScrnCalc (TransRecNo%, EmpName$, EmpRecNo%)
'DECLARE SUB PCEnterEdTrans (TransRecNo%, EmpRecNo%)
'DECLARE SUB PCGetEmp2Rec (EmpRecNo%, ErrorCode%)
'DECLARE SUB PCGetEmpNum (EmpNumber%)
'DECLARE SUB PCLoadPayFreqs ()
'DECLARE SUB PCLoadSystemFiles ()
'DECLARE SUB PCLookUpEmp (EmpNum$, TRecNum%, ERecNum%)
'DECLARE SUB PCPrintManRegister ()
'DECLARE SUB PCPrintPayRegister ()
'DECLARE SUB PCSetPeriodDefault ()
'DECLARE SUB PRCheckMenu ()
'DECLARE SUB ParseHourly2Trans ()
'DECLARE SUB ParseManual2Trans (ManTrans AS ANY)
'DECLARE SUB ParseSalary2Trans ()
'DECLARE SUB ParseScrnCalc2Trans ()
'DECLARE SUB ParseTrans2Hourly ()
'DECLARE SUB ParseTrans2Manual (ManTrans AS ANY)
'DECLARE SUB ParseTrans2Salary ()
'DECLARE SUB ParseTrans2ScrnCalc ()
'DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
'DECLARE SUB PostTransactions ()
'DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
'DECLARE SUB PrintArray (FirstFld%, LastFld%, Form$(), Fld() AS ANY)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
'DECLARE SUB ReplaceString (Work$, Old$, New$)
'DECLARE SUB RestScrn (ScrnArray%())
'DECLARE SUB SaveScrn (ScrnArray%())
'DECLARE SUB ShowCursor ()
'DECLARE SUB UnPackBuffer (FirstFld, LastFld, Form$(), Fld() AS ANY)
'DECLARE SUB UpDateTransFile (TransRecNo%)
'DECLARE SUB WaitForAction ()
'DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
'DECLARE FUNCTION CheckFldEmpty% (FldData$)
'DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
'DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
'DECLARE SUB WazzWind (BYVAL TopRow, BYVAL LeftCol, BYVAL BotRow, BYVAL RghtCol, BYVAL FrameColor, BYVAL FrameType, BYVAL Shadow)
'DECLARE FUNCTION DOSError ()

DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION FUsing$ (Number$, Image$)

'--procedures in this module
DECLARE SUB PrnEmpMatchTrans ()
DECLARE SUB formloadsub ()
DECLARE SUB MarkTrans (Low, High)
DECLARE SUB PurgeDelTrans ()
DECLARE SUB Mark4Del ()
DECLARE SUB PrintTrans ()
DECLARE SUB ResetLastPR ()
DECLARE SUB ChangeTrDate ()
DECLARE SUB ResetLeaveHrs ()

'$INCLUDE: 'DefCnf.bi'
'$INCLUDE: 'FieldInf.bi'
'$INCLUDE: 'FormEdit.bi'
'$INCLUDE: 'PRFiles.bi'
''$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PREmpold.bi'
'$INCLUDE: 'PageInfo.bi'
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
'$INCLUDE: 'PRTIMEIN.Bi'
'$INCLUDE: 'ScrCalc1.Bi'
'$INCLUDE: 'PRCHECK.Bi'
'$INCLUDE: 'PROSCHK.bi'
'$INCLUDE: 'prif.bi'

'$INCLUDE: 'setcnf.bi'

CONST False = 0
CONST True = NOT False

'edittrans
'ResetLeaveHrs
'ResetLastPR
'END

CLS
PRINT " Payroll Utilities"
PRINT
PRINT " 1.  Mark Transactions for Deletion."
PRINT " 2.  Print Transaction File"
PRINT " 3.  Print Transactions by Date by Employee."
PRINT " 4.  Delete Marked Transactions."
PRINT " 5.  Change Transaction Date"
PRINT " 6.  Re-Edit Posted Payroll"
PRINT " 7.  Zero Emp Trans Pointers"
PRINT
PRINT " Esc=Quit"

K$ = INPUT$(1)
SELECT CASE K$
  CASE "1"
    Mark4Del
  CASE "2"
    PrintTrans
  CASE "3"
    PrnEmpMatchTrans
  CASE "4"
    PurgeDelTrans
  CASE "5"
    ChangeTrDate
  CASE "6"
    ResetPR
  CASE "7"
    ResetTrPointers
  CASE CHR$(27)
    CLS : END
END SELECT

SUB ChangeTrDate

  CLS
  LOCATE 1, 1, 0
  QPrintRC "Change Payroll Transaction Date", 1, 1, -1

  LOCATE 2, 1
  INPUT "Enter the Date to Replace (MMDDYY): ", ODate$
  
  LOCATE 3, 1
  INPUT "Enter the New Date (MMDDYY): ", NDate$
  
  NewDate = Date2Num(NDate$)
  OldDate = Date2Num(ODate$)
  
  ND$ = Num2Date(NewDate)
  OD$ = Num2Date(OldDate)

  LOCATE 4, 1
  PRINT "Ready to replace " + OD$ + " with " + ND$ + "."
  PRINT "Enter=Go  Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  
  DIM Tr AS TransRecType
  TRecLen = LEN(Tr)
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = TRecLen
  NumTrans = LOF(1) \ TRecLen

  RptFileName$ = "PRRPTS\PRTRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS #2 LEN = 80
  LPTPort = 1

  ToPrint$ = SPACE$(80)
  RptTitle$ = "PR Trans List"

  FOR T = 1 TO NumTrans
    GET #1, T, Tr

    TrNum$ = STR$(T)
    QPrintRC "Processing:" + TrNum$, 6, 1, -1

    IF Tr.CheckDate = OldDate THEN
      Tr.CheckDate = NewDate
      PUT #1, T, Tr
      LSET ToPrint$ = ""
      MID$(ToPrint$, 2) = TrNum$
      MID$(ToPrint$, 10) = Num2Date(Tr.CheckDate)
      MID$(ToPrint$, 25) = STR$(Tr.CHECKNUM)
      MID$(ToPrint$, 32) = FUsing(STR$(Tr.NETPAY), "######,.##")
      MID$(ToPrint$, 45) = STR$(Tr.TActive)
      PRINT #2, ToPrint$
      TotNet# = TotNet# + Tr.NETPAY
    END IF
    
  NEXT

  LSET ToPrint$ = ""
  MID$(ToPrint$, 2) = "Total Net Pay"
  MID$(ToPrint$, 32) = FUsing(STR$(TotNet#), "######,.##")
  PRINT #2, ToPrint$

  CLOSE

  PrintRptFile RptTitle$, RptFileName$, LPTPort, RetCode, EntryPoint
  CLS


END SUB

SUB edittrans

  DIM Tr AS TransRecType
  TRecLen = LEN(Tr)
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = TRecLen
  NumTrans = LOF(1) \ TRecLen

  T = 2740

  GET #1, T, Tr
  '  Tr.SocTaxAmt = -11.24
  Tr.MedTaxAmt = -11.24
  PUT #1, T, Tr

  CLOSE : END

'Tr.TActive
'Tr.PrevTransRec
'Tr.EmpPin
'Tr.PaySFlag
'Tr.CHECKNUM        'Payroll Check Number
'Tr.PayPdStart      'Start of Pay Period
'Tr.PayPdEnd        'End of Pay Period
'Tr.CheckDate       'Date checks written
'Tr.PostDate        'Date Transaction are posted
'Tr.PayType         'Salaried or Hourly
'Tr.BaseRate        'Base Rate or Salary Amt
'Tr.OTRate          'Overtime Rate
'Tr.RegHrsWork      'Hours worked this period
'Tr.VacUsed         'vacation used this period
'Tr.SickUsed        'Sick hours used this period
'Tr.CompUsed        'comp hours used this period
'Tr.HOLHOURS        'holiday hours used this period
'Tr.RegHrsPaid      'sum of reg hours paid
'Tr.OTHours         'OT hours this period
'Tr.OTHrsPaid       'OT hours paid this period
'Tr.OT2Comp         'Hours to comp time

'1 to 8 wage distributions
'AS TransWageDistType
'Tr.TDist(w).DAcct
'Tr.TDist(w).DRHrs
'Tr.TDist(w).DOHrs
'Tr.TDist(w).DPct
'Tr.TDist(w).DRWage
'Tr.TDist(w).DOWage

'Tr.TotRegWage      'Total Reg Wage distributions
'Tr.TotOTWage       'Total OT Wage distributions
'Tr.GrossWage       'Reg Wage + OT Wage

'1 TO 6 Add Earnings distribitions
'Tr.EDist(a).EAcct
'Tr.EDist(a).EAmt

'Tr.TotAdditEarn    'Total Additional Earnings
'Tr.GROSSPAY        'Add Earnings + GrossWage
'Tr.SocGrossPay     'Social Security Gross
'Tr.MedGrossPay     'Medicare Gross
'Tr.FedGrossPay     'Federal Gross
'Tr.StaGrossPay     'State Gross
'Tr.SocTaxAmt       'Social Security Tax W/H
'Tr.MedTaxAmt       'Medicare Tax W/H
'Tr.FedTaxAmt       'Fed Tax W/H
'Tr.StaTaxAmt       'State Tax W/H
'Tr.TotTaxAmt       'Total Taxes W/H
'Tr.RetireAmt       'Retirement W/H
'Tr.DAMT(1)         '1 to 12 Voluntary Deduction amounts / pcts
'Tr.TotDedAmt       'Total Voluntary Deductions
'Tr.EICAmt
'Tr.NETPAY
'Tr.PeriodHistRec   'AS INTEGER
'Tr.MatchRetAmt     'Employer's Retirement Match
'Tr.MatchSocAmt     'Employer's Social Secity Match
'Tr.MatchMedAmt     'Employer's Medicare Match
'Tr.RetGrossPay     'Retirement Gross
'Tr.TaxFring        'Taxable Fringe
'Tr.Pads(1 TO 2)    'future

END SUB

SUB formloadsub
'
'  LibName$ = "PRQL\PREMPLOY"
'  NumOfScrns = 3
'  ScrnName$ = "EMPMA"
'  REDIM Frm(1)    AS FormInfo
'  REDIM FormName$(1 TO NumOfScrns)
'  FOR Scr = 1 TO NumOfScrns
'    FormName$(Scr) = ScrnName$ + LTRIM$(STR$(Scr))
'  NEXT
'
'  '-- Get total number of fields
'  NumFlds = True
'  FOR Scr = 1 TO NumOfScrns
'    NumFlds = NumFlds + LibNumberOfFields(LibName$, FormName$(Scr)) + 1
'  NEXT
'
'  REDIM Form$(NumFlds, 2)
'  REDIM Fld(NumFlds) AS FieldInfo
'  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
'
'  'for each page, get first and last fields
'  StartEl = 0   'Load first form at array start
'  FOR Scr = 1 TO NumOfScrns     'Get field defs
'    LibGetFldDef LibName$, FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
'    PageInfo(Scr).PageNo = Scr
'    PageInfo(Scr).FirstFld = StartEl + 1
'    StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
'    PageInfo(Scr).LastFld = StartEl - 1
'  NEXT
'
'  Scr = 1
'  LibFile2Scrn LibName$, FormName$(Scr), MonoCode, -1, ErrCode
'  Action = 1
'  DO
'    EditForm Form$(), Fld(), Frm(1), Cnf, Action
'  LOOP UNTIL Frm(1).KeyCode = 27
'
END SUB

SUB Mark4Del

  DO
    CLS
    INPUT "First record to mark for deletion (Q=Quit): ", L$
    IF L$ = "Q" OR L$ = "q" THEN
      cancel = True
      EXIT DO
    END IF

    INPUT "Last record to mark for deletion (Q=Quit): ", H$
    IF H$ = "Q" OR H$ = "q" THEN
      cancel = True
      EXIT DO
    END IF

    Low = VAL(L$)
    High = VAL(H$)
    IF Low > High GOTO tryagain


    msg$ = "Ready to mark records " + L$ + " thru " + H$ + "?"
    Ok$ = "Y=Go, any other key to cancel."
    QPrintRC msg$, 4, 1, -1
    QPrintRC Ok$, 5, 1, -1
    K$ = INPUT$(1)
    IF K$ = "Y" THEN
      EXIT DO
       cancel = 0
    ELSE
      EXIT DO
      cancel = -1
    END IF

tryagain:
  LOOP

  IF cancel THEN
    CLS
    EXIT SUB
  END IF

  MarkTrans Low, High


END SUB

SUB MarkTrans (Low, High)
  
  msg$ = "Marking Records...."
  QPrintRC msg$, 6, 1, -1

  DIM Tr1 AS TransRecType
  TRecLen = LEN(Tr1)
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = TRecLen
  
  FOR T = Low TO High
    GET #1, T, Tr1
    Tr1.TActive = NOT Tr1.TActive
    PUT #1, T, Tr1
  NEXT

  CLOSE
  QPrintRC msg$ + "Done.", 6, 1, -1

END SUB

SUB PrintTrans

  CLS

  DIM Tr AS TransRecType
  TRecLen = LEN(Tr)
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = TRecLen
  NumTrans = LOF(1) \ TRecLen

  RptFileName$ = "PRRPTS\PRTRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS #2 LEN = 80
  LPTPort = 1

  RptTitle$ = "PR Trans List"

  FOR T = 1 TO NumTrans
    GET #1, T, Tr
    TrNum$ = STR$(T)
    QPrintRC "Processing:" + TrNum$, 1, 1, -1
    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 2) = TrNum$
    MID$(ToPrint$, 10) = Num2Date(Tr.CheckDate)
    MID$(ToPrint$, 25) = STR$(Tr.CHECKNUM)
    MID$(ToPrint$, 32) = FUsing(STR$(Tr.NETPAY), "######,.##")
    MID$(ToPrint$, 45) = STR$(Tr.TActive)
    PRINT #2, ToPrint$
  NEXT

  CLOSE

  PrintRptFile RptTitle$, RptFileName$, LPTPort, RetCode, EntryPoint
  CLS

END SUB

SUB PrnEmpMatchTrans

  CLS
  LOCATE 1, 1, 0
  QPrintRC "Print Matching Transaction by Employee", 1, 1, -1

  LOCATE 2, 1
  INPUT "Enter the Search Date (MMDDYY): ", SearchDate$
  
  DIM Emp2 AS EmpData2Type
  ERecLen = LEN(Emp2)
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = ERecLen
  NumOfERec = LOF(1) \ ERecLen

  DIM Tr1 AS TransRecType
  TRecLen = LEN(Tr1)
  OPEN TransHistFileName FOR RANDOM AS #2 LEN = TRecLen
  NumOfTRec = LOF(2) \ TRecLen

  RptFileName$ = "PRRPTS\PRTRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS #3 LEN = 80
  LPTPort = 1
  
  ToPrint$ = SPACE$(80)
  SrchName$ = SPACE$(40)
  Srch$ = "Searching: "
  Date1 = Date2Num(SearchDate$)
  RptTitle$ = "PR Trans List"
  FF$ = CHR$(12)

  GOSUB PageHeader

  FOR cnt = 1 TO NumOfERec
    
    GET #1, cnt, Emp2

    LSET SrchName$ = QPTrim$(Emp2.EMPLNAME) + ", " + QPTrim$(Emp2.EMPFNAME)
    QPrintRC Srch$ + SrchName$, 2, 1, -1
    PRINT #3, SrchName$
    LineCnt = LineCnt + 1
    IF LineCnt >= 55 THEN
      PRINT #3, FF$
      GOSUB PageHeader
      PRINT #3, QPTrim$(SrchName$) + " cont'd"
    END IF
    
    TrNum = Emp2.LastTransRec
    Found = 0

    DO UNTIL TrNum = 0
      GET #2, TrNum, Tr1
      IF Tr1.CheckDate = Date1 THEN
        Found = Found + 1

        LineCnt = LineCnt + 1
        IF LineCnt >= 55 THEN
          PRINT #3, FF$
          GOSUB PageHeader
          PRINT #3, QPTrim$(SrchName$) + " cont'd"
        END IF
        
        TotNetPay# = Round(TotNetPay# + Tr1.NETPAY)
        SELECT CASE Tr1.TActive
        CASE 0
          MTot# = MTot# + Tr1.NETPAY
        CASE -1
          UTot# = UTot# + Tr1.NETPAY
        END SELECT

        LSET ToPrint$ = ""
        MID$(ToPrint$, 2) = STR$(TrNum)
        MID$(ToPrint$, 10) = Num2Date(Tr1.CheckDate)
        MID$(ToPrint$, 25) = STR$(Tr1.CHECKNUM)
        MID$(ToPrint$, 32) = FUsing(STR$(Tr1.NETPAY), "######,.##")
        MID$(ToPrint$, 45) = STR$(Tr1.TActive)
        PRINT #3, ToPrint$

        'IF TrNum = 622 THEN STOP
        'IF Tr1.CHECKNUM >= 23947 AND Tr1.CHECKNUM <= 23994 THEN
        'IF Tr1.CHECKNUM = 24053 THEN
          'Tr1.RetGrossPay = 1568.92
          'M# = 1568.92 * .0555
          'Tr1.MatchRetAmt = round(-95.32)
          'PUT #2, TrNum, Tr1
        'END IF
        
      END IF
      TrNum = Tr1.PrevTransRec

    LOOP

    IF Found = 0 THEN
       PRINT #3, "  No matching transactions for this employee."
       LineCnt = LineCnt + 1
       IF LineCnt >= 55 THEN
         PRINT #3, FF$
         GOSUB PageHeader
         PRINT #3, QPTrim$(SrchName$) + " cont'd"
       END IF
    END IF
    LSET SrchName$ = ""

  NEXT

  PRINT #3,
  PRINT #3, "Summary:"
  PRINT #3, USING "Grand Total Net Pay      : ######,.##"; TotNetPay#
  PRINT #3, USING "Total Marked to Delete   : ######,.##"; MTot#
  PRINT #3, USING "Total Good Transactions  : ######,.##"; UTot#

  PRINT #3, FF$

  CLOSE

  PrintRptFile RptTitle$, RptFileName$, LPTPort, RetCode, EntryPoint
  CLS

EXIT SUB
PageHeader:
  PRINT #3,
  PRINT #3,
  PRINT #3, "Payroll Transactions Dated: " + Num2Date(Date1)
  PRINT #3,
  PRINT #3, "Employee"
  PRINT #3, "  Tr#    Date          Check#         Net  Flag"
  PRINT #3, "-----------------------------------------------"
  LineCnt = 7
RETURN

END SUB

SUB PurgeDelTrans

  CLS
  LOCATE 1, 1, 0
  QPrintRC "Purging Deleted Records", 1, 1, -1

  DIM Tr1 AS TransRecType
  TRecLen = LEN(Tr1)
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = TRecLen
  NumTrans = LOF(1) \ TRecLen
  
  DIM Tr2 AS TransRecType
  OPEN "PRDATA\TRANS.NEW" FOR RANDOM AS #2 LEN = TRecLen
  
  FOR T = 1 TO NumTrans
    GET #1, T, Tr1
    P$ = STR$(INT((T / NumTrans) * 100))
    QPrintRC P$ + "% Complete.", 2, 1, -1
    IF Tr1.TActive = 0 THEN
      NT = NT + 1
      PUT #2, NT, Tr1
    END IF
  NEXT

  CLOSE

  NAME TransHistFileName AS "PRDATA\PRTRANSH.OLD"
  NAME "PRDATA\TRANS.NEW" AS TransHistFileName
  
END SUB

SUB ResetLeaveHrs

  CLS
  LOCATE 1, 1, 0
  QPrintRC "Reset Leave Balances", 1, 1, -1

  LOCATE 2, 1
  INPUT "Enter the Search Date (MMDDYY): ", SearchDate$

  DIM Emp2 AS EmpData2Type
  ERecLen = LEN(Emp2)
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = ERecLen
  NumOfERec = LOF(1) \ ERecLen

  DIM Tr AS TransRecType
  TRecLen = LEN(Tr)
  OPEN TransHistFileName FOR RANDOM AS #2 LEN = TRecLen
  NumOfTRec = LOF(2) \ TRecLen

  RptFileName$ = "PRRPTS\PRTRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS #3 LEN = 80
  LPTPort = 1

  ToPrint$ = SPACE$(80)
  SrchName$ = SPACE$(40)
  Srch$ = "Searching: "
  Date1 = Date2Num(SearchDate$)
  RptTitle$ = "PR Trans List"

  FOR cnt = 1 TO NumOfERec
    GET #1, cnt, Emp2

    LSET SrchName$ = QPTrim$(Emp2.EMPLNAME) + ", " + QPTrim$(Emp2.EMPFNAME)
    QPrintRC Srch$ + SrchName$, 2, 1, -1
    PRINT #3, SrchName$

    TrNum = Emp2.LastTransRec
    Found = 0

    DO UNTIL TrNum = 0
      GET #2, TrNum, Tr
      IF Tr.CheckDate = Date1 THEN
        Found = Found + 1

        LSET ToPrint$ = ""
        MID$(ToPrint$, 2) = "T V Used: " + STR$(Tr.VacUsed)
        MID$(ToPrint$, 20) = "T S Used: " + STR$(Tr.SickUsed)
        PRINT #3, ToPrint$

        LSET ToPrint$ = ""
        MID$(ToPrint$, 2) = "E V Used: " + STR$(Emp2.EMPVUSED)
        MID$(ToPrint$, 20) = "E S Used: " + STR$(Emp2.EMPSLUSE)
        PRINT #3, ToPrint$

        Emp2.EMPVUSED = Emp2.EMPVUSED - Tr.VacUsed
        Emp2.EMPSLUSE = Emp2.EMPSLUSE - Tr.SickUsed

        Emp2.EMPVBAL = Emp2.EMPVACE - Emp2.EMPVUSED
        Emp2.EMPSLBAL = Emp2.EMPSLE - Emp2.EMPSLUSE
        
        PUT #1, cnt, Emp2



        EXIT DO

      END IF
      TrNum = Tr.PrevTransRec
    LOOP

    IF Found = 0 THEN
       PRINT #3, "  No matching transactions for this employee."
    END IF
    LSET SrchName$ = ""

  NEXT

  PRINT #3, USING "Total Net Pay: ######,.##"; TotNetPay#

  CLOSE

  PrintRptFile RptTitle$, RptFileName$, LPTPort, RetCode, EntryPoint
  CLS
  

END SUB

SUB ResetPR

  CLS
  INPUT "Payroll Date to re-set: ", PRDate$
  TDate = Date2Num(PRDate$)

  CLS
  PRINT "Ready to reset payroll dated " + Num2Date(TDate)
  PRINT "Y=Go, Any other key to cancel"
  K$ = UCASE$(INPUT$(1))
  IF K$ <> "Y" THEN EXIT SUB
  
  REDIM PrdDefRec(1) AS PeriodDefaultRecType
  PeriodDefaultRecLen = LEN(PrdDefRec(1))
  
  DIM Tr AS TransRecType
  TRecLen = LEN(Tr)
  OPEN "PRDATA\PRTRANST.DAT" FOR RANDOM AS #1 LEN = TRecLen
  NumTrans = LOF(1) \ TRecLen

  'RptFileName$ = "PRRPTS\PRTRLIST.PRN"
  'OPEN RptFileName$ FOR OUTPUT AS #2 LEN = 80
  'LPTPort = 1
  'ToPrint$ = SPACE$(80)
  'RptTitle$ = "PR Trans List"

  FOR T = 1 TO NumTrans
    GET #1, T, Tr
    TrNum$ = STR$(T)
    QPrintRC "Processing:" + TrNum$, 1, 1, -1
    IF Tr.CheckDate = TDate THEN
      Tr.TActive = -1
      PUT #1, T, Tr
      Found = Found + 1
    END IF
    
  NEXT

  CLOSE
  PRINT "Found: "; Found
  
  FGetAH PPDefaultFileName, PrdDefRec(1), PeriodDefaultRecLen, 1
  PrdDefRec(1).PACTIVE = -1
  FPutAH PPDefaultFileName, PrdDefRec(1), PeriodDefaultRecLen, 1
  
  'PrintRptFile RptTitle$, RptFileName$, LPTPort, RetCode, EntryPoint
  'CLS
  
END SUB

SUB ResetTrPointers

  CLS
  LOCATE 1, 1, 0
  QPrintRC "Reset Transaction Pointers to Zero? (Y/N)", 1, 1, -1
  K$ = UCASE$(INPUT$(1))
  IF K$ <> "Y" THEN EXIT SUB
  
  DIM Emp2 AS EmpData2Type
  ERecLen = LEN(Emp2)
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = ERecLen
  NumOfERec = LOF(1) \ ERecLen

  FOR cnt = 1 TO NumOfERec
    GET #1, cnt, Emp2
    LSET SrchName$ = QPTrim$(Emp2.EMPLNAME) + ", " + QPTrim$(Emp2.EMPFNAME)
    QPrintRC "Re-Setting: " + SrchName$, 2, 1, -1
    Emp2.LastTransRec = 0
  NEXT

  CLS
  CLOSE


END SUB

