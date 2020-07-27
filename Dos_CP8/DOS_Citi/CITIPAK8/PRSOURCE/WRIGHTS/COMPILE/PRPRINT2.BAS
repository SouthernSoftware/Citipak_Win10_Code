DEFINT A-Z
DECLARE SUB RPTSetupPRN (RPTNum%, FHandle%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB MakeGLIFFile (TotEIC#, TotDeds() AS DOUBLE, Passed() AS DOUBLE, DistbSumAccts() AS ANY)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ReplaceString (Work$, Old$, New$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOFRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)

  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

  ''$INCLUDE: 'FieldInf.bi'    '        "          field information
  '$INCLUDE: 'FormEdit.bi'    '        "          form editing information

  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.biw'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'ErnCodes.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Biw'
  '$INCLUDE: 'prif.bi'
  '$INCLUDE: 'GLSETUP.bi'

  CONST False = 0, True = NOT False

'  REDIM SHARED DistbSumAccts(1 TO 1) AS DistWageRptType

SUB MakeGLIFFile (TotEIC#, TotDeds#(), Passed#(), DistbSumAccts() AS DistWageRptType)

  'SHARED DistbSumAccts() AS DistWageRptType

  REDIM SysRec(1)           AS RegDSysFileRecType
  REDIM PDR(1)              AS PeriodDefaultRecType
  REDIM DedCodes(1 TO 12)   AS DedCodeRecType

  REDIM GLSetupRec(1)       AS GLSetupRecType

  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  FGetAH PPDefaultFileName, PDR(1), LEN(PDR(1)), 1
  FGetAH SysFileName, SysRec(1), LEN(SysRec(1)), 1
  
  FACnt = SysRec(1).AcctCnt


'for new gl

  FundPad = 0
  FundLen = 2     'Default fund length

  GLSetUpName$ = QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT"
  IF Exist(GLSetUpName$) THEN
    FGetAH GLSetUpName$, GLSetupRec(1), LEN(GLSetupRec(1)), 1
    FundLen = GLSetupRec(1).FundLen
    FundPad = GLSetupRec(1).DetLen - GLSetupRec(1).FundLen
  END IF

  NumOfWageAccts = UBOUND(DistbSumAccts)
  
  'squeeze out all the "-" out of Acct numbers
  FOR Cnt = 1 TO NumOfWageAccts
    ReplaceString DistbSumAccts(Cnt).Acct, "-", ""
  NEXT
  
  'change Period ending date to nicks format
  
  GLIFTDate$ = Num2Date(PDR(1).PEREND)
  
  ReplaceString GLIFTDate$, "-", "/"
  ReplaceString GLIFTDate$, "1994", "94"
  ReplaceString GLIFTDate$, "1995", "95"
  ReplaceString GLIFTDate$, "1996", "96"
  ReplaceString GLIFTDate$, "1997", "97"
  ReplaceString GLIFTDate$, "1998", "98"
  ReplaceString GLIFTDate$, "1999", "99"
  ReplaceString GLIFTDate$, "2000", "00"
  '
  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$
  
  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
  
  REDIM GLIFRec(1 TO NumOfWageAccts + 17 + 1) AS GLIFDataType14
  
  FOR Cnt = 1 TO NumOfWageAccts
    GLIFRec(Cnt).TranAcct = DistbSumAccts(Cnt).Acct
    GLIFRec(Cnt).TranDate = GLIFTDate$
    GLIFRec(Cnt).TranDesc = "Wages"
    'GLIFRec(Cnt).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(Cnt).CrAmt = 0
    GLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).GROSSPAY
    GLIFRec(Cnt).Source = GLIFSource$
    GLIFRec(Cnt).FromFlag = "W"
  NEXT
  
  NextAcct = Cnt
  '
  FOR Cnt = 0 TO 4
    ReplaceString SysRec(1).Liab(Cnt + 1).Acct, "-", ""
    GLIFRec(NextAcct + Cnt).TranAcct = QPTrim$(SysRec(1).Liab(Cnt + 1).Acct)
    GLIFRec(NextAcct + Cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + Cnt).TranDesc = "Withholdings"
    GLIFRec(NextAcct + Cnt).DrAmt = 0
    GLIFRec(NextAcct + Cnt).Source = GLIFSource$
    GLIFRec(NextAcct + Cnt).FromFlag = "X"
    'get tax and ret account numbers
  NEXT
  
  GLIFRec(NextAcct).CrAmt = Passed#(1)          'federal
  GLIFRec(NextAcct + 1).CrAmt = Passed#(2)      'state
  GLIFRec(NextAcct + 2).CrAmt = Passed#(3)      'social sec
  GLIFRec(NextAcct + 3).CrAmt = Passed#(4)      'Medicare
  GLIFRec(NextAcct + 4).CrAmt = Passed#(5)      'Retirement total
  'good to here ;maybe
  
  ReplaceString SysRec(1).CashAcct, "-", ""
  
  SysCash$ = QPTrim$(SysRec(1).CashAcct)
  
  'EIC HERE
  '
  IF TotEIC# > 0 THEN
    REDIM EICGLIFRec(1 TO 2) AS GLIFDataType14
    EICGLIFRec(1).TranAcct = QPTrim$(SysRec(1).Liab(1).Acct)
    EICGLIFRec(1).TranDate = GLIFTDate$
    EICGLIFRec(1).TranDesc = "EIC Pmt"
    EICGLIFRec(1).CrAmt = 0
    EICGLIFRec(1).DrAmt = TotEIC#
    EICGLIFRec(1).Source = GLIFSource$
    EICGLIFRec(1).FromFlag = "E"
    '
    EICGLIFRec(2).TranAcct = LEFT$(QPTrim$(SysRec(1).Liab(1).Acct), FundLen) + SysCash$
    'EICGLIFRec(2).TranAcct = LEFT$(QPTrim$(SysRec(1).Liab(1).Acct), 2) + SysCash$
    EICGLIFRec(2).TranDate = GLIFTDate$
    EICGLIFRec(2).TranDesc = "EIC Pmt"
    EICGLIFRec(2).CrAmt = TotEIC#
    EICGLIFRec(2).DrAmt = 0
    EICGLIFRec(2).Source = GLIFSource$
    EICGLIFRec(2).FromFlag = "P"
  END IF

  NextAcct = NextAcct + Cnt
  '
  FOR Cnt = 0 TO 11
    ReplaceString DedCodes(Cnt + 1).DCACCT1, "-", ""
    GLIFRec(NextAcct + Cnt).TranAcct = QPTrim$(DedCodes(Cnt + 1).DCACCT1)
    GLIFRec(NextAcct + Cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + Cnt).TranDesc = "Deductions"
    GLIFRec(NextAcct + Cnt).CrAmt = TotDeds#(Cnt + 1)
    GLIFRec(NextAcct + Cnt).DrAmt = 0
    GLIFRec(NextAcct + Cnt).Source = GLIFSource$
    GLIFRec(NextAcct + Cnt).FromFlag = "D"
  NEXT
  
  REDIM DFunds$(1 TO NumOfWageAccts)
  NumDFunds = 1
  'fixed
  DFunds$(1) = LEFT$(DistbSumAccts(1).Acct, FundLen)
  FOR Cnt = 1 TO NumOfWageAccts - 1
    'fixed
    IF LEFT$(DistbSumAccts(Cnt).Acct, FundLen) <> LEFT$(DistbSumAccts(Cnt + 1).Acct, FundLen) THEN
      NumDFunds = NumDFunds + 1 'counting the total number of funds.
      'fixed
      DFunds$(NumDFunds) = LEFT$(DistbSumAccts(Cnt + 1).Acct, FundLen)
    END IF
  NEXT
  
  REDIM PRESERVE DFunds$(1 TO NumDFunds)
  REDIM DFund(1 TO NumDFunds) AS FundType
  FOR Cnt = 1 TO NumOfWageAccts
    'fixed
    CurrFund$ = LEFT$(DistbSumAccts(Cnt).Acct, FundLen)
    FOR ThisFund = 1 TO NumDFunds
      IF CurrFund$ = DFunds$(ThisFund) THEN
        DFund(ThisFund).FundCode = DFunds$(ThisFund)
        DFund(ThisFund).Debit = Round(DFund(ThisFund).Debit + DistbSumAccts(Cnt).GROSSPAY)
        EXIT FOR
      END IF
    NEXT
  NEXT
  
  'all gross pay by funds here!!
  'make funds and sumarize ded and taxs here
  '
  REDIM CFunds$(1 TO 17)
  '
  FirstC = NumOfWageAccts + 1
  LastC = NumOfWageAccts + 17 + 1
  
  NumCFunds = 1
  'fixed
  CFunds$(1) = LEFT$(GLIFRec(FirstC).TranAcct, FundLen)
  
  FOR Cnt = FirstC TO LastC - 1
    'fixed
    IF LEFT$(GLIFRec(Cnt).TranAcct, FundLen) <> LEFT$(GLIFRec(Cnt + 1).TranAcct, FundLen) THEN
      IF LEN(QPTrim$(GLIFRec(Cnt + 1).TranAcct)) THEN
        NumCFunds = NumCFunds + 1               'counting the total number of funds.
        'fixed
        CFunds$(NumCFunds) = LEFT$(GLIFRec(Cnt + 1).TranAcct, FundLen)
      END IF
    END IF
  NEXT

  REDIM PRESERVE CFunds$(1 TO NumCFunds)
  REDIM CFund(1 TO NumCFunds) AS FundType
  FOR Cnt = FirstC TO LastC - 1
    'fixed
    CurrFund$ = LEFT$(GLIFRec(Cnt).TranAcct, FundLen)
    FOR ThisFund = 1 TO NumCFunds
      IF CurrFund$ = CFunds$(ThisFund) THEN
        CFund(ThisFund).FundCode = CFunds$(ThisFund)
        CFund(ThisFund).Credit = Round(CFund(ThisFund).Credit + GLIFRec(Cnt).CrAmt)
        EXIT FOR
      END IF
    NEXT
  NEXT

  'combine all funds in one array here
  TotalFunds = NumDFunds + NumCFunds            '+ 1
  REDIM AllFunds(1 TO TotalFunds) AS FundType
  ThisFund = 1
  FOR Cnt = 1 TO NumDFunds
    AllFunds(ThisFund) = DFund(Cnt)
    ThisFund = ThisFund + 1
  NEXT
  '
  IF NoCFunds = False THEN
    FOR Cnt = 1 TO NumCFunds
      AllFunds(ThisFund) = CFund(Cnt)
      ThisFund = ThisFund + 1
    NEXT
  END IF

  ERASE DFund, CFund

'fixed: 05-27-96
  SortT AllFunds(1), TotalFunds, 0, LEN(AllFunds(1)), 0, LEN(AllFunds(1).FundCode)
                                                        
  'combine Debits and Credits for same fund numbers
  First = 1
  Start = 1
  Last = TotalFunds
  DO
    Start = Start + 1
    FOR Cnt2 = Start TO Last
      IF AllFunds(First).FundCode = AllFunds(Cnt2).FundCode THEN
        LSET AllFunds(Cnt2).FundCode = ""       'blank acct num as a flag
        AllFunds(First).Credit = Round#(AllFunds(First).Credit + AllFunds(Cnt2).Credit)
        AllFunds(Cnt2).Credit = 0
        AllFunds(First).Debit = Round#(AllFunds(First).Debit + AllFunds(Cnt2).Debit)
        AllFunds(Cnt2).Debit = 0
      END IF
    NEXT
    First = First + 1
  LOOP UNTIL First >= Last      '

  'calc net difference for funds
  FOR Cnt = 1 TO TotalFunds
    IF LEN(QPTrim$(AllFunds(Cnt).FundCode)) THEN
      AllFunds(Cnt).Net = Round#(AllFunds(Cnt).Debit - AllFunds(Cnt).Credit)
    END IF
  NEXT

  'add cash acct number to fund numbers
  FOR Cnt = 1 TO TotalFunds
    FundCash$ = QPTrim$(AllFunds(Cnt).FundCode)
    IF LEN(FundCash$) THEN
      LSET AllFunds(Cnt).FundCode = FundCash$ + SysCash$
    END IF
  NEXT

  TotalGLIFS = NumOfWageAccts + 17 + TotalFunds

  REDIM PRESERVE GLIFRec(1 TO TotalGLIFS) AS GLIFDataType14

  NextAcct = NumOfWageAccts + 17
  FOR Cnt = 1 TO TotalFunds
    AcctNum$ = QPTrim$(AllFunds(Cnt).FundCode)
    IF LEN(AcctNum$) THEN
      NextAcct = NextAcct + 1
      GLIFRec(NextAcct).TranAcct = AcctNum$
      GLIFRec(NextAcct).TranDate = GLIFTDate$
      GLIFRec(NextAcct).TranDesc = "PR Net "
      GLIFRec(NextAcct).Source = GLIFSource$
      GLIFRec(NextAcct).FromFlag = "N"
      
      IF AllFunds(Cnt).Net > 0 THEN
        GLIFRec(NextAcct).CrAmt = AllFunds(Cnt).Net
        GLIFRec(NextAcct).DrAmt = 0
      ELSEIF AllFunds(Cnt).Net < 0 THEN
        GLIFRec(NextAcct).DrAmt = ABS(AllFunds(Cnt).Net)
        GLIFRec(NextAcct).CrAmt = 0
      END IF
    END IF
  NEXT
  
  'if using the imprest account then
  SELECT CASE SysRec(1).USEIMP
  CASE "I"      'was Y            'I C P
    TotalGLIFS = TotalGLIFS + 2
    REDIM PRESERVE GLIFRec(1 TO TotalGLIFS) AS GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    ReplaceString SysRec(1).ICRACCT, "-", ""
    GLIFRec(TotalGLIFS - 1).TranAcct = QPTrim$(SysRec(1).ICRACCT)
    GLIFRec(TotalGLIFS - 1).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS - 1).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS - 1).Source = GLIFSource$
    GLIFRec(TotalGLIFS - 1).FromFlag = "i"
    GLIFRec(TotalGLIFS - 1).DrAmt = Passed#(6)
    GLIFRec(TotalGLIFS - 1).CrAmt = 0
    '
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "i"
    GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    GLIFRec(TotalGLIFS).DrAmt = 0
  CASE "C"      'NEW Central Depository
    TotalGLIFS = TotalGLIFS + 1
    REDIM PRESERVE GLIFRec(1 TO TotalGLIFS) AS GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "c"
    
    IF TotEIC# > 0 THEN
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6) - TotEIC#
    ELSE
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    END IF
    
    GLIFRec(TotalGLIFS).DrAmt = 0
    
    ReplaceString SysRec(1).ICRACCT, "-", ""
    REDIM CDGLIFRec(1 TO TotalFunds)  AS GLIFDataType14
    FOR Cnt = 1 TO TotalFunds
      IF AllFunds(Cnt).Net <> 0 THEN
        'fixed
        CDGLIFRec(Cnt).TranAcct = QPTrim$(SysRec(1).ICRACCT) + LEFT$(AllFunds(Cnt).FundCode, FundLen)
        IF FundPad > 0 THEN
          TempAcct$ = QPTrim$(CDGLIFRec(Cnt).TranAcct)
          TempAcct$ = TempAcct$ + STRING$(FundPad, "0")
          CDGLIFRec(Cnt).TranAcct = TempAcct$
        END IF
        CDGLIFRec(Cnt).TranDate = GLIFTDate$
        CDGLIFRec(Cnt).TranDesc = "PPE " + GLIFTDate$
        CDGLIFRec(Cnt).Source = GLIFSource$
        CDGLIFRec(Cnt).FromFlag = "C"
        CDGLIFRec(Cnt).DrAmt = AllFunds(Cnt).Net                'Passed(6)
        CDGLIFRec(Cnt).CrAmt = 0
      END IF
    NEXT
  CASE "P"
    
  END SELECT
  
  '*************END OF COMMON CODE SECTION
  
  IF QPTrim$(SysRec(1).EXPMETHD) = "" OR SysRec(1).EXPMETHD = "0" THEN
    GOTO WriteGLIFS
  END IF
  
  IF SysRec(1).EXPMETHD = "2" GOTO Type2Meth    'jump to type 2 here
  
  'calc and add Fringe GLIF recs
  REDIM FGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14

  ReplaceString SysRec(1).FRNGEXP, "-", ""
  FringeAcct$ = QPTrim$(SysRec(1).FRNGEXP)
  FringeRate# = SysRec(1).FRNGRATE

  FOR Cnt = 1 TO NumOfWageAccts
    AcctNum$ = LEFT$(GLIFRec(Cnt).TranAcct, FACnt)
    FGLIFRec(Cnt).TranAcct = AcctNum$ + FringeAcct$
    FGLIFRec(Cnt).DrAmt = Round(GLIFRec(Cnt).DrAmt * (FringeRate# * .01))
    FGLIFRec(Cnt).TranDate = GLIFTDate$
    FGLIFRec(Cnt).TranDesc = "FRINGE " + GLIFTDate$
    FGLIFRec(Cnt).Source = GLIFSource$
    FGLIFRec(Cnt).FromFlag = "F"
  NEXT

  'calc and add Indirect GLIF recs
  REDIM IGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14
  ReplaceString SysRec(1).INDDR, "-", ""
  'fixed
  INDFund$ = QPTrim$(LEFT$(SysRec(1).INDDR, FundLen))
  IndirectAcct$ = QPTrim$(SysRec(1).INDEXP)
  IndirectRate# = SysRec(1).INDRATE
  
  IF IndirectRate# < 0 THEN IndirectRate# = 0
  
  FOR Cnt = 1 TO NumOfWageAccts
    'look for acct that don't get indirect
    'fixed
    IF NOT QPTrim$(LEFT$(GLIFRec(Cnt).TranAcct, FundLen)) = INDFund$ THEN
      AcctNum$ = LEFT$(GLIFRec(Cnt).TranAcct, FACnt)
      IGLIFRec(Cnt).TranAcct = AcctNum$ + IndirectAcct$
      IGLIFRec(Cnt).DrAmt = Round((GLIFRec(Cnt).DrAmt + FGLIFRec(Cnt).DrAmt) * (IndirectRate#) * .01)
      IGLIFRec(Cnt).TranDate = GLIFTDate$
      IGLIFRec(Cnt).TranDesc = "INDIRECT " + GLIFTDate$
      IGLIFRec(Cnt).Source = GLIFSource$
      IGLIFRec(Cnt).FromFlag = "I"
    END IF
  NEXT

  REDIM IFGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14
  FOR Cnt = 1 TO NumOfWageAccts
    'fixed
    AcctNum$ = LEFT$(FGLIFRec(Cnt).TranAcct, FundLen)
    IFGLIFRec(Cnt).TranAcct = AcctNum$ + SysCash$
    IFGLIFRec(Cnt).CrAmt = Round(IGLIFRec(Cnt).DrAmt + FGLIFRec(Cnt).DrAmt)
    IFGLIFRec(Cnt).TranDate = GLIFTDate$
    IFGLIFRec(Cnt).TranDesc = "F&I COST " + GLIFTDate$
    IFGLIFRec(Cnt).Source = GLIFSource$
    IFGLIFRec(Cnt).FromFlag = "A"
  NEXT

  FOR Cnt = 1 TO NumOfWageAccts
    Indirect# = Round(Indirect# + IGLIFRec(Cnt).DrAmt)
    Fringe# = Round(Fringe# + FGLIFRec(Cnt).DrAmt)
  NEXT

  REDIM AGLIFRec(1 TO 4)  AS GLIFDataType14

  ReplaceString SysRec(1).FRNGDR, "-", ""
  ReplaceString SysRec(1).FRNGCR, "-", ""
  ReplaceString SysRec(1).INDDR, "-", ""
  ReplaceString SysRec(1).INDCR, "-", ""

  AGLIFRec(1).TranAcct = QPTrim$(SysRec(1).FRNGDR)
  AGLIFRec(1).DrAmt = Fringe#

  AGLIFRec(2).TranAcct = QPTrim$(SysRec(1).FRNGCR)
  AGLIFRec(2).CrAmt = Fringe#

  AGLIFRec(3).TranAcct = QPTrim$(SysRec(1).INDDR)
  AGLIFRec(3).DrAmt = Indirect#

  AGLIFRec(4).TranAcct = QPTrim$(SysRec(1).INDCR)
  AGLIFRec(4).CrAmt = Indirect#

  FOR Cnt = 1 TO 4
    AGLIFRec(Cnt).TranDate = GLIFTDate$
    AGLIFRec(Cnt).TranDesc = "PPE " + GLIFTDate$
    AGLIFRec(Cnt).Source = GLIFSource$
    AGLIFRec(Cnt).FromFlag = "T"
  NEXT
  GOTO WriteGLIFS
  '**********END TYPE 1 SECTION

Type2Meth:

  REDIM SocGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14
  REDIM MedGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14
  REDIM RetGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14

  SOCEXP$ = QPTrim$(SysRec(1).SOCEXP)
  MEDEXP$ = QPTrim$(SysRec(1).MEDEXP)
  RETEXP$ = QPTrim$(SysRec(1).RETEXP)

  SOCLIAB$ = QPTrim$(SysRec(1).SOCLIAB)
  MEDLIAB$ = QPTrim$(SysRec(1).MEDLIAB)
  RETLIAB$ = QPTrim$(SysRec(1).RETLIAB)

  ReplaceString SOCLIAB$, "-", ""
  ReplaceString MEDLIAB$, "-", ""
  ReplaceString RETLIAB$, "-", ""

  FOR Cnt = 1 TO NumOfWageAccts
    'social
    SocGLIFRec(Cnt).TranAcct = LEFT$(DistbSumAccts(Cnt).Acct, FACnt) + SOCEXP$
    SocGLIFRec(Cnt).TranDate = GLIFTDate$
    SocGLIFRec(Cnt).TranDesc = "Soc Match"
    SocGLIFRec(Cnt).Source = GLIFSource$
    SocGLIFRec(Cnt).FromFlag = "S"
    SocGLIFRec(Cnt).CrAmt = 0
    SocGLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).MATSocAmt
    'medicare
    MedGLIFRec(Cnt).TranAcct = LEFT$(DistbSumAccts(Cnt).Acct, FACnt) + MEDEXP$
    MedGLIFRec(Cnt).TranDate = GLIFTDate$
    MedGLIFRec(Cnt).TranDesc = "Med Match"
    MedGLIFRec(Cnt).Source = GLIFSource$
    MedGLIFRec(Cnt).FromFlag = "M"
    MedGLIFRec(Cnt).CrAmt = 0
    MedGLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).MATMedAmt
    'retirment
    RetGLIFRec(Cnt).TranAcct = LEFT$(DistbSumAccts(Cnt).Acct, FACnt) + RETEXP$
    RetGLIFRec(Cnt).TranDate = GLIFTDate$
    RetGLIFRec(Cnt).TranDesc = "Ret Match"
    RetGLIFRec(Cnt).Source = GLIFSource$
    RetGLIFRec(Cnt).FromFlag = "R"
    RetGLIFRec(Cnt).CrAmt = 0
    RetGLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).MATRetAmt
  NEXT

  REDIM SFGLIFRec(1 TO TotalFunds)  AS GLIFDataType14
  REDIM MFGLIFRec(1 TO TotalFunds)  AS GLIFDataType14
  REDIM RFGLIFRec(1 TO TotalFunds)  AS GLIFDataType14

  FOR Cnt = 1 TO TotalFunds
    'fixed
    SFGLIFRec(Cnt).TranAcct = LEFT$(AllFunds(Cnt).FundCode, FundLen) + SOCLIAB$
    MFGLIFRec(Cnt).TranAcct = LEFT$(AllFunds(Cnt).FundCode, FundLen) + MEDLIAB$
    RFGLIFRec(Cnt).TranAcct = LEFT$(AllFunds(Cnt).FundCode, FundLen) + RETLIAB$
  NEXT

  FOR Cnt = 1 TO NumOfWageAccts
    FOR Cnt2 = 1 TO TotalFunds
      'fixed
      IF LEFT$(SFGLIFRec(Cnt2).TranAcct, FundLen) = LEFT$(SocGLIFRec(Cnt).TranAcct, FundLen) THEN
        SFGLIFRec(Cnt2).CrAmt = Round(SFGLIFRec(Cnt2).CrAmt + SocGLIFRec(Cnt).DrAmt)
        SFGLIFRec(Cnt2).DrAmt = 0
        SFGLIFRec(Cnt2).TranDate = GLIFTDate$
        SFGLIFRec(Cnt2).TranDesc = "Soc Match Liab"
        SFGLIFRec(Cnt2).Source = GLIFSource$
        SFGLIFRec(Cnt2).FromFlag = "s"
      END IF
    NEXT
  NEXT

  FOR Cnt = 1 TO NumOfWageAccts
    FOR Cnt2 = 1 TO TotalFunds
      'fixed
      IF LEFT$(MFGLIFRec(Cnt2).TranAcct, FundLen) = LEFT$(MedGLIFRec(Cnt).TranAcct, FundLen) THEN
        MFGLIFRec(Cnt2).CrAmt = Round(MFGLIFRec(Cnt2).CrAmt + MedGLIFRec(Cnt).DrAmt)
        MFGLIFRec(Cnt2).DrAmt = 0
        MFGLIFRec(Cnt2).TranDate = GLIFTDate$
        MFGLIFRec(Cnt2).TranDesc = "Med Match Liab"
        MFGLIFRec(Cnt2).Source = GLIFSource$
        MFGLIFRec(Cnt2).FromFlag = "m"
      END IF
    NEXT
  NEXT

  FOR Cnt = 1 TO NumOfWageAccts
    FOR Cnt2 = 1 TO TotalFunds
      'fixed
      IF LEFT$(RFGLIFRec(Cnt2).TranAcct, FundLen) = LEFT$(RetGLIFRec(Cnt).TranAcct, FundLen) THEN
        RFGLIFRec(Cnt2).CrAmt = Round(RFGLIFRec(Cnt2).CrAmt + RetGLIFRec(Cnt).DrAmt)
        RFGLIFRec(Cnt2).DrAmt = 0
        RFGLIFRec(Cnt2).TranDate = GLIFTDate$
        RFGLIFRec(Cnt2).TranDesc = "Ret Match Liab"
        RFGLIFRec(Cnt2).Source = GLIFSource$
        RFGLIFRec(Cnt2).FromFlag = "r"
      END IF
    NEXT
  NEXT

WriteGLIFS:
  
  GLIFRecLen = LEN(GLIFRec(1))
  FCreate "TempIF.DAT"
  FOpenS "TempIF.DAT", GHandle
  RecNo& = 1
  
  FOR Cnt = 1 TO TotalGLIFS
    
    IF LEN(QPTrim$(GLIFRec(Cnt).TranAcct)) THEN
      IF GLIFRec(Cnt).DrAmt > 0 OR GLIFRec(Cnt).CrAmt > 0 THEN
        FPutRTA GHandle, GLIFRec(Cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      END IF
    END IF
  NEXT
  

  SELECT CASE SysRec(1).EXPMETHD
  CASE "1"
    FOR Cnt = 1 TO NumOfWageAccts
      IF LEN(QPTrim$(FGLIFRec(Cnt).TranAcct)) THEN
        IF FGLIFRec(Cnt).DrAmt > 0 OR FGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, FGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
    
    FOR Cnt = 1 TO NumOfWageAccts
      IF LEN(QPTrim$(IGLIFRec(Cnt).TranAcct)) THEN
        IF IGLIFRec(Cnt).DrAmt > 0 OR IGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, IGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
    
    FOR Cnt = 1 TO NumOfWageAccts
      IF LEN(QPTrim$(IFGLIFRec(Cnt).TranAcct)) THEN
        IF IFGLIFRec(Cnt).DrAmt > 0 OR IFGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, IFGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
    
    FOR Cnt = 1 TO 4
      IF LEN(QPTrim$(AGLIFRec(Cnt).TranAcct)) > 0 THEN
        FPutRTA GHandle, AGLIFRec(Cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      END IF
    NEXT
    
  CASE "2"
    FOR Cnt = 1 TO NumOfWageAccts
      IF LEN(QPTrim$(SocGLIFRec(Cnt).TranAcct)) THEN
        IF SocGLIFRec(Cnt).DrAmt > 0 OR SocGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, SocGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
    FOR Cnt = 1 TO NumOfWageAccts
      IF LEN(QPTrim$(MedGLIFRec(Cnt).TranAcct)) THEN
        IF MedGLIFRec(Cnt).DrAmt > 0 OR MedGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, MedGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
    FOR Cnt = 1 TO NumOfWageAccts
      IF LEN(QPTrim$(RetGLIFRec(Cnt).TranAcct)) THEN
        IF RetGLIFRec(Cnt).DrAmt > 0 OR RetGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, RetGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT

    FOR Cnt = 1 TO TotalFunds
      IF LEN(QPTrim$(SFGLIFRec(Cnt).TranAcct)) THEN
        IF SFGLIFRec(Cnt).DrAmt > 0 OR SFGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, SFGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
    FOR Cnt = 1 TO TotalFunds
      IF LEN(QPTrim$(MFGLIFRec(Cnt).TranAcct)) THEN
        IF MFGLIFRec(Cnt).DrAmt > 0 OR MFGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, MFGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
    FOR Cnt = 1 TO TotalFunds
      IF LEN(QPTrim$(RFGLIFRec(Cnt).TranAcct)) THEN
        IF RFGLIFRec(Cnt).DrAmt > 0 OR RFGLIFRec(Cnt).CrAmt > 0 THEN
          FPutRTA GHandle, RFGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
  END SELECT
  
  IF SysRec(1).USEIMP = "C" THEN
    FOR Cnt = 1 TO TotalFunds
      IF LEN(QPTrim$(CDGLIFRec(Cnt).TranAcct)) THEN
        IF CDGLIFRec(Cnt).DrAmt <> 0 OR CDGLIFRec(Cnt).CrAmt <> 0 THEN
          FPutRTA GHandle, CDGLIFRec(Cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        END IF
      END IF
    NEXT
  END IF

  'added EIC GLIF records if present 6/07/94
  IF TotEIC# > 0 THEN
    FOR Cnt = 1 TO 2
      FPutRTA GHandle, EICGLIFRec(Cnt), RecNo&, GLIFRecLen
      RecNo& = RecNo& + 1
    NEXT
  END IF

  FClose GHandle

  ERASE GLIFRec, EICGLIFRec, SysRec, PDR, AllFunds, DFunds$, CFunds$
  
  
END SUB

