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
DECLARE SUB MakeGLIFFile (TotEIC#, TotDeds() AS DOUBLE, Passed() AS DOUBLE)
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
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'ErnCodes.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'prif.bi'
  
  CONST False = 0, True = NOT False
  
  REDIM SHARED DistbSumAccts(1 TO 1) AS DistWageRptType

SUB MakeGLIFFile (TotEIC#, TotDeds#(), Passed#())
  
  REDIM SysRec(1)   AS RegDSysFileRecType
  REDIM PDR(1)      AS PeriodDefaultRecType
  REDIM DedCodes(1 TO 12)   AS DedCodeRecType
  
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  FGetAH PPDefaultFileName, PDR(1), LEN(PDR(1)), 1
  FGetAH SysFileName, SysRec(1), LEN(SysRec(1)), 1
  
  FACnt = SysRec(1).AcctCnt
  
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
  'ReplaceString GLIFTDate$, "1997", "97"
  '
  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$
  
  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
  
  REDIM GLIFRec(1 TO NumOfWageAccts + 17 + 1) AS GLIFDataType12
  
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
  
  ReplaceString SysRec(1).CASHACCT, "-", ""
  
  SysCash$ = QPTrim$(SysRec(1).CASHACCT)
  
  'EIC HERE
  '
  IF TotEIC# > 0 THEN
    REDIM EICGLIFRec(1 TO 2) AS GLIFDataType12
    EICGLIFRec(1).TranAcct = QPTrim$(SysRec(1).Liab(1).Acct)
    EICGLIFRec(1).TranDate = GLIFTDate$
    EICGLIFRec(1).TranDesc = "EIC Pmt"
    EICGLIFRec(1).CrAmt = 0
    EICGLIFRec(1).DrAmt = TotEIC#
    EICGLIFRec(1).Source = GLIFSource$
    EICGLIFRec(1).FromFlag = "E"
    '
    EICGLIFRec(2).TranAcct = LEFT$(QPTrim$(SysRec(1).Liab(1).Acct), 2) + SysCash$
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
  DFunds$(1) = LEFT$(DistbSumAccts(1).Acct, 2)
  FOR Cnt = 1 TO NumOfWageAccts - 1
    IF LEFT$(DistbSumAccts(Cnt).Acct, 2) <> LEFT$(DistbSumAccts(Cnt + 1).Acct, 2) THEN
      NumDFunds = NumDFunds + 1 'counting the total number of funds.
      DFunds$(NumDFunds) = LEFT$(DistbSumAccts(Cnt + 1).Acct, 2)
    END IF
  NEXT
  
  REDIM PRESERVE DFunds$(1 TO NumDFunds)
  REDIM DFund(1 TO NumDFunds) AS FundType
  FOR Cnt = 1 TO NumOfWageAccts
    CurrFund$ = LEFT$(DistbSumAccts(Cnt).Acct, 2)
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
  CFunds$(1) = LEFT$(GLIFRec(FirstC).TranAcct, 2)
  
  FOR Cnt = FirstC TO LastC - 1
    IF LEFT$(GLIFRec(Cnt).TranAcct, 2) <> LEFT$(GLIFRec(Cnt + 1).TranAcct, 2) THEN
      IF LEN(QPTrim$(GLIFRec(Cnt + 1).TranAcct)) THEN
        NumCFunds = NumCFunds + 1               'counting the total number of funds.
        CFunds$(NumCFunds) = LEFT$(GLIFRec(Cnt + 1).TranAcct, 2)
      END IF
    END IF
  NEXT

  REDIM PRESERVE CFunds$(1 TO NumCFunds)
  REDIM CFund(1 TO NumCFunds) AS FundType
  FOR Cnt = FirstC TO LastC - 1
    CurrFund$ = LEFT$(GLIFRec(Cnt).TranAcct, 2)
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
  SortT AllFunds(1), TotalFunds, 0, LEN(AllFunds(1)), 0, 9

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

  REDIM PRESERVE GLIFRec(1 TO TotalGLIFS) AS GLIFDataType12

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
    REDIM PRESERVE GLIFRec(1 TO TotalGLIFS) AS GLIFDataType12
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
    REDIM PRESERVE GLIFRec(1 TO TotalGLIFS) AS GLIFDataType12
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
    REDIM CDGLIFRec(1 TO TotalFunds)  AS GLIFDataType12
    FOR Cnt = 1 TO TotalFunds
      IF AllFunds(Cnt).Net <> 0 THEN
        CDGLIFRec(Cnt).TranAcct = QPTrim$(SysRec(1).ICRACCT) + LEFT$(AllFunds(Cnt).FundCode, 2)
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
  REDIM FGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType12

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
  REDIM IGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType12
  ReplaceString SysRec(1).INDDR, "-", ""
  INDFund$ = QPTrim$(LEFT$(SysRec(1).INDDR, 2))
  IndirectAcct$ = QPTrim$(SysRec(1).INDEXP)
  IndirectRate# = SysRec(1).INDRATE
  
  IF IndirectRate# < 0 THEN IndirectRate# = 0
  
  FOR Cnt = 1 TO NumOfWageAccts
    'look for acct that don't get indirect
    IF NOT QPTrim$(LEFT$(GLIFRec(Cnt).TranAcct, 2)) = INDFund$ THEN
      AcctNum$ = LEFT$(GLIFRec(Cnt).TranAcct, FACnt)
      IGLIFRec(Cnt).TranAcct = AcctNum$ + IndirectAcct$
      IGLIFRec(Cnt).DrAmt = Round((GLIFRec(Cnt).DrAmt + FGLIFRec(Cnt).DrAmt) * (IndirectRate#) * .01)
      IGLIFRec(Cnt).TranDate = GLIFTDate$
      IGLIFRec(Cnt).TranDesc = "INDIRECT " + GLIFTDate$
      IGLIFRec(Cnt).Source = GLIFSource$
      IGLIFRec(Cnt).FromFlag = "I"
    END IF
  NEXT

  REDIM IFGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType12
  FOR Cnt = 1 TO NumOfWageAccts
    AcctNum$ = LEFT$(FGLIFRec(Cnt).TranAcct, 2)
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

  REDIM AGLIFRec(1 TO 4)  AS GLIFDataType12

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

  REDIM SocGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType12
  REDIM MedGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType12
  REDIM RetGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType12

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

  REDIM SFGLIFRec(1 TO TotalFunds)  AS GLIFDataType12
  REDIM MFGLIFRec(1 TO TotalFunds)  AS GLIFDataType12
  REDIM RFGLIFRec(1 TO TotalFunds)  AS GLIFDataType12

  FOR Cnt = 1 TO TotalFunds
    SFGLIFRec(Cnt).TranAcct = LEFT$(AllFunds(Cnt).FundCode, 2) + SOCLIAB$
    MFGLIFRec(Cnt).TranAcct = LEFT$(AllFunds(Cnt).FundCode, 2) + MEDLIAB$
    RFGLIFRec(Cnt).TranAcct = LEFT$(AllFunds(Cnt).FundCode, 2) + RETLIAB$
  NEXT

  FOR Cnt = 1 TO NumOfWageAccts
    FOR Cnt2 = 1 TO TotalFunds
      IF LEFT$(SFGLIFRec(Cnt2).TranAcct, 2) = LEFT$(SocGLIFRec(Cnt).TranAcct, 2) THEN
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
      IF LEFT$(MFGLIFRec(Cnt2).TranAcct, 2) = LEFT$(MedGLIFRec(Cnt).TranAcct, 2) THEN
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
      IF LEFT$(RFGLIFRec(Cnt2).TranAcct, 2) = LEFT$(RetGLIFRec(Cnt).TranAcct, 2) THEN
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
      FPutRTA GHandle, AGLIFRec(Cnt), RecNo&, GLIFRecLen
      RecNo& = RecNo& + 1
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

SUB PCPrintPayRegister
  
  SHARED PrnDef$()
  
  RptTitle$ = "Register & G/L Interface Reports"
  ShowProcessingScrn RptTitle$
  
  REDIM TransRec(1) AS TransRecType
  REDIM EmpRec1(1)  AS EmpData1Type
  REDIM PDR(1)      AS PeriodDefaultRecType
  REDIM Unit(1)     AS UnitFileRecType
  
  REDIM DistbSumAccts(1 TO 1)  AS DistWageRptType
  
  REDIM SysRec(1) AS RegDSysFileRecType
  REDIM GLIFRec(1 TO 1)        AS GLIFDataType12
  
  REDIM EmpAct(1)             AS DistWageRptType
  
  REDIM DedCodes(1 TO 12) AS DedCodeRecType
  REDIM ErnCodes(1 TO 3) AS ErnCodeRecType
  
  REDIM TotDeds#(1 TO 12)
  REDIM TotErns(1 TO 3) AS DOUBLE
  
  REDIM EDesc(1)    AS STRING * 21
  REDIM EDAct(1)    AS STRING * 14
  REDIM EDPct(1)    AS STRING * 11
  REDIM EDRHrs(1)   AS STRING * 11
  REDIM EDOHrs(1)   AS STRING * 11
  REDIM EDRPay(1)   AS STRING * 11
  REDIM EDOPay(1)   AS STRING * 11
  REDIM EDEarn(1)   AS STRING * 11
  REDIM EDGroP(1)   AS STRING * 11
  
  REDIM EDSAmt(1)   AS STRING * 11
  REDIM EDMAmt(1)   AS STRING * 11
  REDIM EDRAmt(1)   AS STRING * 11
  
  REDIM ENum(1)     AS STRING * 14
  REDIM EName(1)    AS STRING * 33
  
  REDIM BRat(1)     AS STRING * 11
  REDIM ORat(1)     AS STRING * 11
  
  REDIM TaxFrn(1)   AS STRING * 11
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
  
  'added for EIC   6/07/94
  REDIM EEicP(1)   AS STRING * 11
  
  REDIM Ern(1)      AS STRING * 11
  
  REDIM Pg(1)       AS STRING * 5
  
  REDIM EMPLine(1) AS STRING * 132
  
  REDIM Dash0(1) AS STRING * 69
  REDIM Dash(1) AS STRING * 132
  REDIM Dash2(1) AS STRING * 124
  REDIM Dash3(1) AS STRING * 91
  
  FGetAH PPDefaultFileName, PDR(1), LEN(PDR(1)), 1
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  FGetAH ErnCodeFileName, ErnCodes(1), LEN(ErnCodes(1)), 3
  
  FGetAH SysFileName, SysRec(1), LEN(SysRec(1)), 1

  Image0$ = "####"
  Image$ = "####.##"
  Image3$ = ",######.##"
  Image4$ = "###.####"
  Image5$ = ",#######.##"
  
  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GTaxFrn# = 0
  TotDebit# = 0
  TotCredit# = 0
  
  LSET Fill11$(1) = ""
  
  LSET Dash0(1) = STRING$(69, "-")
  LSET Dash(1) = STRING$(132, "-")
  LSET Dash2(1) = STRING$(124, "-")
  RSET Dash3(1) = STRING$(63, "-")
  
  EmpActLen = LEN(EmpAct(1))
  DistbSumAcctsLen = LEN(DistbSumAccts(1))
  
  TransRecLen = LEN(TransRec(1))
  Emp1RecLen = LEN(EmpRec1(1))
  
  IdxRecLen = 2 'LEN(TempIdxNRec)
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOFRecs = IdxFileSize& \ IdxRecLen
  
  SalCnt = 0
  HrlCnt = 0
  
  DLineCnt = 0
  LineCnt = 0
  
  NumOfWageAccts = 0
  
  MaxLines = 45
  DMaxLines = 40
  EPage = 1
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
  ETitle$ = ""
  FOR Cnt = 1 TO 3
    TErn$ = QPTrim$(ErnCodes(Cnt).ERNCODE1)
    IF LEN(TErn$) > 0 THEN
      LastErn = LastErn + 1
      RSET Ern(1) = TErn$
      ETitle$ = ETitle$ + Ern(1)
    ELSE
      EXIT FOR
    END IF
  NEXT
  
  IF LastErn < 3 THEN
    ETitle$ = SPACE$(11 * (3 - LastErn)) + ETitle$
  END IF
  SumHeader2$ = "  Reg Wages  O/T Wages" + ETitle$
  ETitle$ = "   Reg Earn   O/T Earn" + ETitle$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"
  
  '------------------------------------------------------------------
  
  FCreate PayRegisterRptName
  FCreate DistributionRptName
  
  FOpenS PayRegisterRptName, RHandle
  FOpenS DistributionRptName, DHandle
  
  FOpenS EmpData1Name, NHandle  'open it
  FOpenS TransWorkFileName, THandle             'open it
  
  RPTSetupPRN 13, RHandle
  RPTSetupPRN 13, DHandle
  
  GOSUB PrintPayRollHeader
  GOSUB PrintDistHeader
  
  FOR Cnt = 1 TO NumOFRecs
    FGetRTA THandle, TransRec(1), CLNG(IdxBuff(Cnt)), TransRecLen
    
    IF TransRec(1).TActive = True THEN
      FGetRTA NHandle, EmpRec1(1), CLNG(IdxBuff(Cnt)), Emp1RecLen
      
      GOSUB SumAndPrintTime
      GOSUB ParseDistributions
      LineCnt = LineCnt + 5
      IF LineCnt >= MaxLines AND Cnt < NumOFRecs THEN
        LineCnt = 0
        FPut RHandle, FF$
        GOSUB PrintPayRollHeader
      END IF
      IF DLineCnt >= DMaxLines THEN
        DLineCnt = 0
        FPut DHandle, FF$
        GOSUB PrintDistHeader
      END IF
      
    END IF
    ShowPctComp Cnt, NumOFRecs
  NEXT
  
  GOSUB PrintSumTotal
  GOSUB PrintDistTotal
  
  RPTSetupPRN 0, RHandle
  RPTSetupPRN 0, DHandle
  
  FClose THandle
  FClose NHandle
  FClose RHandle
  FClose DHandle
  
  
  'HERE:  register report files have been written to disk
  
  'GLIF START
  'if there is a GL transfer directory then make GLIF file
  
  IF LEN(QPTrim$(SysRec(1).CITIDIR)) > 0 THEN
    REDIM Passed#(1 TO 6)
    Passed#(1) = FTax#
    Passed#(2) = STax#
    Passed#(3) = SSTax#
    Passed#(4) = MTax#
    Passed#(5) = RETTOT#
    Passed#(6) = TNetPay#
    MakeGLIFFile TotEIC#, TotDeds#(), Passed#() 'unrem
  END IF
  
  IF Exist(QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile) THEN
    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile
    REDIM JGLIdxRec(1)                    AS JGLAcctIdxType
    JFlag = True
  ELSE
    REDIM GLIdxRec(1)                     AS GLAcctIdxType
    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
  END IF
  
  IF JFlag THEN
    GLIdxRecLen = LEN(JGLIdxRec(1))
  ELSE
    GLIdxRecLen = LEN(GLIdxRec(1))
  END IF
  
  GLIFRecLen = LEN(GLIFRec(1))
  
  TotalGLIFS = FileSize("TempIF.DAT") \ GLIFRecLen
  TotalAccts = FileSize(GLIdxName$) \ GLIdxRecLen
  
  IF TotalGLIFS = 0 THEN
    PrintGLRpt = False
    GOTO SkipGLRpt
  ELSE
    PrintGLRpt = True
  END IF
  
  REDIM GLIFRec(1 TO TotalGLIFS)        AS GLIFDataType12
  FGetAH "TempIF.DAT", GLIFRec(1), GLIFRecLen, TotalGLIFS
  
  'SortT GLIFRec(1), TotalGLIFS, 0, GLIFRecLen, 17, 20
  'SortT GLIFRec(1), TotalGLIFS, 0, GLIFRecLen, 0, 9

  DO
    OutOfOrder = False                     'assume it's sorted
    FOR X = 1 TO UBOUND(GLIFRec) - 1
      IF GLIFRec(X).TranAcct > GLIFRec(X + 1).TranAcct THEN
        SWAP GLIFRec(X), GLIFRec(X + 1)    'if we had to swap
        OutOfOrder = True                'we're not done yet
      END IF
    NEXT
  LOOP WHILE OutOfOrder

  FPutAH "TempIF.DAT", GLIFRec(1), GLIFRecLen, TotalGLIFS

  'SortT GLIFRec(1), TotalGLIFS, 0, GLIFRecLen, 0, 12
  'Check into this^^^

  IF TotalAccts = 0 THEN GOTO SkipGLAccts
  
  QPrintRC "Reading G/L Accounts.", 11, 31, 112
  
  FOpenS GLIdxName$, GHandle
  SELECT CASE JFlag
  CASE False
    REDIM GoodAccts(1 TO TotalAccts)      AS DOUBLE
    FOR Cnt = 1 TO TotalAccts
      FGetRTA GHandle, GLIdxRec(1), CLNG(Cnt), GLIdxRecLen
      IF GLIdxRec(1).AcctNum > 0 THEN
        ActualAccts = ActualAccts + 1
        GoodAccts(ActualAccts) = GLIdxRec(1).AcctNum            'QPValL(AcctNum$)
        IF GoodAccts(ActualAccts) < 9999999 THEN
          GoodAccts(ActualAccts) = Round(GoodAccts(ActualAccts) * 100)
        END IF
      END IF
      ShowPctComp Cnt, TotalAccts
    NEXT
  CASE True
    REDIM JGoodAccts(1 TO TotalAccts)      AS STRING
    FOR Cnt = 1 TO TotalAccts
      FGetRTA GHandle, JGLIdxRec(1), CLNG(Cnt), GLIdxRecLen
      ActualAccts = ActualAccts + 1
      ReplaceString JGLIdxRec(1).AcctNum, "-", ""
      JGoodAccts(ActualAccts) = JGLIdxRec(1).AcctNum
      ShowPctComp Cnt, TotalAccts
    NEXT
    
  END SELECT
  
  FClose GHandle
  TotalAccts = ActualAccts
  
  QPrintRC "Checking for invalid Accounts.", 11, 27, 112


SkipGLAccts:
  Max = 55
  Lines = 0
  
  RptName$ = "prrpts\PRGLIF.RPT"
  FCreate RptName$
  FOpenS RptName$, RHandle
  RPTSetupPRN 14, RHandle
  FClose RHandle
  
  FOpenS "TempIF.DAT", THandle
  
  OPEN RptName$ FOR APPEND AS #1
  
  GOSUB GLIFHeader
  
  FOR Cnt = 1 TO TotalGLIFS
    ShowPctComp Cnt, TotalGLIFS
    IF Lines >= Max THEN
      PRINT #1, FF$
      GOSUB GLIFHeader
    END IF

    TotDebit# = Round(TotDebit# + GLIFRec(Cnt).DrAmt)
    TotCredit# = Round(TotCredit# + GLIFRec(Cnt).CrAmt)
    
    '10-13-94 ** added check for valid GL Account numbers
    IF TotalAccts > 0 THEN
      
      SELECT CASE JFlag
      CASE True
        FOR Cnt2 = 1 TO TotalAccts
          IF INSTR(JGoodAccts(Cnt2), GLIFRec(Cnt).TranAcct) AND LEN(QPTrim$(JGoodAccts(Cnt2))) = LEN(QPTrim$(GLIFRec(Cnt).TranAcct)) THEN
            AcctOk = True
            EXIT FOR
          END IF
        NEXT
        IF NOT AcctOk THEN
          LSET GLIFRec(Cnt).Fill = "MISSING "
          FPutRTA THandle%, GLIFRec(Cnt), CLNG(Cnt), LEN(GLIFRec(Cnt))
        ELSE
          LSET GLIFRec(Cnt).Fill = ""
        END IF
        AcctOk = False
        
      CASE False
        GLAcct@ = Round(VAL(GLIFRec(Cnt).TranAcct))
        IF GLAcct@ < 9999999 THEN
          GLAcct@ = GLAcct@ * 100
        END IF
        FOR Cnt2 = 1 TO TotalAccts
          IF GoodAccts(Cnt2) = GLAcct@ THEN
            AcctOk = True
            EXIT FOR
          END IF
        NEXT
        IF NOT AcctOk THEN
          LSET GLIFRec(Cnt).Fill = "MISSING "
          FPutRTA THandle%, GLIFRec(Cnt), CLNG(Cnt), LEN(GLIFRec(Cnt))
        ELSE
          LSET GLIFRec(Cnt).Fill = ""
        END IF
        AcctOk = False
      END SELECT
    END IF
    
    'NoCheckAccts:
    RSET EDSAmt(1) = FUsing(STR$(GLIFRec(Cnt).DrAmt), Image3$)
    RSET EDMAmt(1) = FUsing(STR$(GLIFRec(Cnt).CrAmt), Image3$)
    LSET EDesc(1) = QPTrim$(GLIFRec(Cnt).TranDesc)
    NoAcctNum = LEN(QPTrim$(GLIFRec(Cnt).TranAcct))
    IF NoAcctNum > 0 THEN
      PRINT #1, TAB(3); GLIFRec(Cnt).TranAcct; GLIFRec(Cnt).Fill; EDesc(1);
      PRINT #1, EDSAmt(1); "    "; EDMAmt(1)
      Lines = Lines + 1
    END IF
  NEXT
  GOSUB GLIFTotals
  
  Fund$ = LEFT$(QPTrim$(GLIFRec(1).TranAcct), 2)
  FDebit# = GLIFRec(1).DrAmt
  FCredit# = GLIFRec(1).CrAmt
  FOR Cnt = 2 TO TotalGLIFS
    NFund$ = LEFT$(QPTrim$(GLIFRec(Cnt).TranAcct), 2)
    IF NFund$ <> Fund$ THEN
      GOSUB PrintFundTotal
      Fund$ = NFund$
      FDebit# = GLIFRec(Cnt).DrAmt
      FCredit# = GLIFRec(Cnt).CrAmt
    ELSE
      FDebit# = Round(FDebit# + GLIFRec(Cnt).DrAmt)
      FCredit# = Round(FCredit# + GLIFRec(Cnt).CrAmt)
    END IF
  NEXT
  GOSUB PrintFundTotal
  PRINT #1, FF$; PrnDef$(1)
  
  CLOSE
  FClose THandle
  ERASE GoodAccts
  
SkipGLRpt:
  ERASE GLIFRec
  
  '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
  '05-10-94  fixed to insure everything was cleaned up after report
  '06-28-94  move so all is cleaned up BEFORE the report prints.
  '07-15-94  move again to add gl interface report.
  ERASE TransRec, EmpRec1, PDR, Unit, DistbSumAccts, SysRec, GLIFRec, EmpAct
  ERASE DedCodes, ErnCodes, TotDeds#, TotErns, EDAct, EDPct, EDRHrs, EDOHrs
  ERASE EDRPay, EDOPay, EDEarn, EDGroP, EDSAmt, EDMAmt, EDRAmt, ENum, EName
  ERASE BRat, ORat, Fill11, SCnt, HCnt, RHrs, VHrs, SHrs, HHrs, CHrs, THrs
  ERASE OTHrs, OTPaid, OTComp, RErnP, OErnP, GPayP, SSTaxP, MTaxP, FTaxP, STaxP
  ERASE RetirP, NetPayP, Ded, Ern, Pg, Dash, Dash2, Dash3
  ERASE IdxBuff, TaxFrn
  '----------------------------------------------------------------------------
  
  RptTitle$ = "Payroll Register Report"
  PrintRptFile RptTitle$, PayRegisterRptName, 1, RetCode, 1     'unrem
  
  IF RetCode = -1 THEN GOTO AltExit
  
  RptTitle$ = "Payroll Distribution Report"
  PrintRptFile RptTitle$, DistributionRptName, 1, RetCode, 1    '0    'unrem
  
  IF RetCode = -1 THEN GOTO AltExit
  
  IF PrintGLRpt THEN
    RptTitle$ = "G/L Interface Report"
    PrintRptFile RptTitle$, "prrpts\PRGLIF.RPT", 1, RetCode, 1  '0      'unrem
  END IF
  
AltExit:
  
  EXIT SUB
  
PrintFundTotal:
  RSET EDSAmt(1) = FUsing(STR$(FDebit#), Image3$)
  RSET EDMAmt(1) = FUsing(STR$(FCredit#), Image3$)
  LSET EDesc(1) = ""
  LSET GLIFRec(1).Fill = ""
  RSET GLIFRec(1).TranAcct = "Fund  " + Fund$
  
  PRINT #1, TAB(3); GLIFRec(1).TranAcct; GLIFRec(1).Fill; EDesc(1);
  PRINT #1, EDSAmt(1); "    "; EDMAmt(1)
  RETURN
  
GLIFTotals:
  PRINT #1,
  PRINT #1, Dash0(1)
  PRINT #1, "   Totals:"; TAB(45); FUsing(STR$(TotDebit#), Image3$); TAB(60); FUsing(STR$(TotCredit#), Image3$)
  IF TotalAccts = 0 THEN
    PRINT #1,
    PRINT #1, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory."
  END IF
  PRINT #1,
  PRINT #1, "SubTotals:"
  RETURN
  
  
GLIFHeader:
  '  RSET Pg(1) = STR$(Page)
  PRINT #1, Unit(1).UFEMPR
  PRINT #1, "G/L Interface Report."
  PRINT #1, "Period Ending: " + Num2Date$(PDR(1).PEREND)
  PRINT #1,
  PRINT #1, "Account No.           Description                Debit         Credit"
  PRINT #1, Dash0(1)
  'PRINT #1,
  Lines = 6
  
  RETURN
  
PrintPayRollHeader:
  RSET Pg(1) = STR$(Page)
  FPut RHandle, Unit(1).UFEMPR + SPACE$(87) + "Page:" + Pg(1) + CrLf$
  FPut RHandle, "Payroll Register" + CrLf$
  FPut RHandle, "Period Ending: " + Num2Date$(PDR(1).PEREND) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "Employee No   Name                                                                                                               EIC" + CrLf$
  FPut RHandle, "  Base Rate   O/T Rate   Tax Frng    Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid   O/T Comp" + CrLf$
  FPut RHandle, ETitle$ + CrLf$
  FPut RHandle, DTitle$ + CrLf$
  FPut RHandle, Dash(1) + CrLf$
  LincCnt = LineCnt + 9
  Page = Page + 1
  RETURN
  
SumAndPrintTime:
  
  REGHRS# = Round#(REGHRS# + TransRec(1).RegHrsWork)
  VACHRS# = Round#(VACHRS# + TransRec(1).VacUsed)
  SICKHRS# = Round#(SICKHRS# + TransRec(1).SickUsed)
  HOLHRS# = Round#(HOLHRS# + TransRec(1).HOLHOURS)
  COMPHRS# = Round#(COMPHRS# + TransRec(1).CompUsed)
  
  TotalHrs# = Round(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  
  'added for EIC   6/07/94
  TotEIC# = Round#(TotEIC# + TransRec(1).EICAmt)
  
  '-=-=-=-=-=-=-=
  TOTHrs# = Round#(TOTHrs# + TransRec(1).OTHours)
  IF TransRec(1).OTHrsPaid > 0 THEN
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
  IF TransRec(1).RetireAmt > 0 THEN
    RETTOT# = Round(RETTOT# + TransRec(1).RetireAmt)
  END IF
  'TOTDED# = Round(TOTDED# + TransRec(1).TotDedAmt)
  
  TNetPay# = Round#(TNetPay# + TransRec(1).NETPAY)
  GFedGross# = Round#(GFedGross# + TransRec(1).FedGrossPay)
  GStaGross# = Round#(GStaGross# + TransRec(1).StaGrossPay)
  GSocGross# = Round#(GSocGross# + TransRec(1).SocGrossPay)
  GMedGross# = Round#(GMedGross# + TransRec(1).MedGrossPay)
  
  GTaxFrn# = Round#(GTaxFrn# + TransRec(1).TaxFring)
  
  LSET ENum(1) = LTRIM$(EmpRec1(1).EMPNO)
  LSET EName(1) = QPTrim$(EmpRec1(1).EMPLNAME) + ", " + QPTrim$(EmpRec1(1).EMPFNAME)
  RSET BRat(1) = FUsing(STR$(TransRec(1).BaseRate), Image3$)
  RSET ORat(1) = FUsing(STR$(TransRec(1).OTRate), Image3$)
  
  RSET RHrs(1) = FUsing(STR$(TransRec(1).RegHrsWork), Image$)
  
  RSET VHrs(1) = FUsing(STR$(TransRec(1).VacUsed), Image$)
  RSET SHrs(1) = FUsing(STR$(TransRec(1).SickUsed), Image$)
  RSET HHrs(1) = FUsing(STR$(TransRec(1).HOLHOURS), Image$)
  RSET CHrs(1) = FUsing(STR$(TransRec(1).CompUsed), Image$)
  RSET THrs(1) = FUsing(STR$(TransRec(1).RegHrsPaid), Image$)
  
  RSET TaxFrn(1) = FUsing(STR$(TransRec(1).TaxFring), Image$)
  RSET OTHrs(1) = FUsing(STR$(TransRec(1).OTHours), Image$)
  RSET OTPaid(1) = FUsing(STR$(TransRec(1).OTHrsPaid), Image$)
  RSET OTComp(1) = FUsing(STR$(TransRec(1).OT2Comp), Image$)
  
  'added for EIC     6/07/94
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
    IF TransRec(1).DAMT(Cnt2) > 0 THEN
      TotDeds#(Cnt2) = Round#(TotDeds#(Cnt2) + TransRec(1).DAMT(Cnt2))
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
  SumErn$ = ""
  FOR Cnt2 = 1 TO LastErn
    TotErns(Cnt2) = Round#(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    RSET Ern(1) = FUsing(STR$(TransRec(1).EAmt(Cnt2)), Image3$)
    SumErn$ = SumErn$ + Ern(1)
  NEXT
  IF LastErn < 3 THEN
    SumErn$ = SPACE$(11 * (3 - LastErn)) + SumErn$
  END IF
  '-------------------------------------------------------
  
  RSET EMPLine$(1) = EEicP(1)
  MID$(EMPLine$(1), 1) = ENum(1) + EName(1)
  FPut RHandle, EMPLine$(1) + CrLf$
  FPut RHandle, BRat(1) + ORat(1) + TaxFrn(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
  FPut RHandle, CHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + OTComp(1) + CrLf$
  FPut RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1)
  FPut RHandle, RetirP(1) + NetPayP(1) + CrLf$
  FPut RHandle, SumDed$ + CrLf$
  FPut RHandle, CrLf$
  
  RETURN
  
PrintSumTotal:
  RSET SCnt(1) = FUsing(STR$(SalCnt), Image0$)
  RSET HCnt(1) = FUsing(STR$(HrlCnt), Image0$)
  
  RSET Fill11(1) = FUsing(STR$(GTaxFrn#), Image3$)
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
    RSET Ded(1) = FUsing(STR$(TotDeds#(Cnt2)), Image3$)
    SumDed$ = SumDed$ + Ded(1)
  NEXT
  IF LastDed < 12 THEN
    SumDed$ = SPACE$(11 * (12 - LastDed)) + SumDed$
  END IF
  '---------------------------------------------------------
  SumErn$ = ""
  FOR Cnt2 = 1 TO LastErn
    RSET Ern(1) = FUsing(STR$(TotErns(Cnt2)), Image3$)
    SumErn$ = SumErn$ + Ern(1)
  NEXT
  IF LastErn < 3 THEN
    SumErn$ = SPACE$(11 * (3 - LastErn)) + SumErn$
  END IF
  
  '--------------NEW----------------------------
  ' Page = Page + 1
  RSET Pg(1) = STR$(Page)
  FPut RHandle, FF$
  FPut RHandle, Unit(1).UFEMPR + SPACE$(87) + "Page:" + Pg(1) + CrLf$
  FPut RHandle, "Payroll Register Summary" + CrLf$
  FPut RHandle, "Period Ending: " + Num2Date$(PDR(1).PEREND) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, Dash(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "   Salaried     Hourly   Tax Frng    Reg Hrs      Vacat       Sick        Hol       Comp      Total    O/T Hrs   O/T Paid   O/T Comp" + CrLf$
  FPut RHandle, SCnt(1) + HCnt(1) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
  FPut RHandle, CHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + OTComp(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay" + CrLf$
  FPut RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1)
  FPut RHandle, STaxP(1) + RetirP(1) + NetPayP(1) + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, DTitle$ + CrLf$
  FPut RHandle, SumDed$ + CrLf$
  FPut RHandle, CrLf$
  FPut RHandle, "  Fed Gross  Sta Gross  Med Gross  Soc Gross  EIC Total" + CrLf$
  RSET FTaxP(1) = FUsing(STR$(GFedGross#), Image5$)
  RSET STaxP(1) = FUsing(STR$(GStaGross#), Image5$)
  RSET MTaxP(1) = FUsing(STR$(GMedGross#), Image5$)
  RSET SSTaxP(1) = FUsing(STR$(GSocGross#), Image5$)
  RSET EEicP(1) = FUsing(STR$(TotEIC#), Image5$)
  
  FPut RHandle, FTaxP(1) + STaxP(1) + MTaxP(1) + SSTaxP(1) + EEicP(1) + CrLf$
  
  FPut RHandle, CrLf$
  FPut RHandle, Dash(1) + CrLf$
  FPut RHandle, FF$
  
  RETURN
  
  
  '-----------------------------------------------------------------------
ParseDistributions:
  
  REDIM EmpAct(1 TO 14)  AS DistWageRptType
  ENumOfAct = 0
  
  'process wage distributions
  FOR Cnt2 = 1 TO 8
    Acct$ = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
    IF LEN(Acct$) > 0 THEN
      ENumOfAct = ENumOfAct + 1
      LSET EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).Pct = TransRec(1).TDist(Cnt2).DPct
      EmpAct(ENumOfAct).RHrs = TransRec(1).TDist(Cnt2).DRHrs
      EmpAct(ENumOfAct).OHrs = TransRec(1).TDist(Cnt2).DOHrs
      EmpAct(ENumOfAct).RWage = TransRec(1).TDist(Cnt2).DRWage
      EmpAct(ENumOfAct).OWage = TransRec(1).TDist(Cnt2).DOWage
      EmpAct(ENumOfAct).GROSSPAY = Round#(EmpAct(ENumOfAct).RWage + EmpAct(ENumOfAct).OWage)
    END IF
  NEXT
  
  'process earnings distributions
  FOR Cnt2 = 1 TO 6
    Acct$ = QPTrim$(TransRec(1).EDist(Cnt2).EAcct)
    IF LEN(Acct$) > 0 THEN
      ENumOfAct = ENumOfAct + 1
      LSET EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).AddEarn = TransRec(1).EDist(Cnt2).EAmt
      EmpAct(ENumOfAct).GROSSPAY = TransRec(1).EDist(Cnt2).EAmt
    END IF
  NEXT
  
  'HERE: got all accts for this employee
  
  First = 1
  Start = 1
  
  Last = ENumOfAct
  
  'purge and sum employee's dupelicate account distributions
  
  DO
    Start = Start + 1
    FOR Cnt2 = Start TO Last
      IF EmpAct(First).Acct = EmpAct(Cnt2).Acct THEN
        LSET EmpAct(Cnt2).Acct = ""             'blank acct num as a flag
        EmpAct(First).Pct = Round#(EmpAct(First).Pct + EmpAct(Cnt2).Pct)
        EmpAct(First).RHrs = Round#(EmpAct(First).RHrs + EmpAct(Cnt2).RHrs)
        EmpAct(First).OHrs = Round#(EmpAct(First).OHrs + EmpAct(Cnt2).OHrs)
        EmpAct(First).RWage = Round#(EmpAct(First).RWage + EmpAct(Cnt2).RWage)
        EmpAct(First).OWage = Round#(EmpAct(First).OWage + EmpAct(Cnt2).OWage)
        EmpAct(First).AddEarn = Round#(EmpAct(First).AddEarn + EmpAct(Cnt2).AddEarn)
        EmpAct(First).GROSSPAY = Round#(EmpAct(First).GROSSPAY + EmpAct(Cnt2).GROSSPAY)
      END IF
    NEXT
    First = First + 1
  LOOP UNTIL First >= Last
  
  'calc percentages of matching amts to each account distribution
  
  FOR Cnt2 = 1 TO ENumOfAct
    EmpAct(Cnt2).MATSocAmt = Round#(TransRec(1).MatchSocAmt * (EmpAct(Cnt2).Pct * .01))
    EmpAct(Cnt2).MATMedAmt = Round#(TransRec(1).MatchMedAmt * (EmpAct(Cnt2).Pct * .01))
    EmpAct(Cnt2).MATRetAmt = Round#(TransRec(1).MatchRetAmt * (EmpAct(Cnt2).Pct * .01))
  NEXT
  
  '---------------------------------------------------------------------------
  'calc and adjust matching distribution amts
  'find last active account
  'adjust Social Amt
  
  DO
    TotalSocAmt# = 0
    FOR Cnt2 = 1 TO 14          '8
      TotalSocAmt# = Round#(TotalSocAmt# + EmpAct(Cnt2).MATSocAmt)
    NEXT
    IF TotalSocAmt# = 0 GOTO SkipSocDist
    IF TotalSocAmt# <> TransRec(1).MatchSocAmt THEN
      FOR Cnt3 = 14 TO 1 STEP -1                '8 TO 1 STEP -1
        IF EmpAct(Cnt3).MATSocAmt > 0 THEN
          LastActive = Cnt3
          EXIT FOR
        END IF
      NEXT
      IF TotalSocAmt# > TransRec(1).MatchSocAmt THEN
        DistDif# = Round#(TotalSocAmt# - TransRec(1).MatchSocAmt)
        EmpAct(LastActive).MATSocAmt = Round(EmpAct(LastActive).MATSocAmt - DistDif#)
      ELSEIF TotalSocAmt# < TransRec(1).MatchSocAmt THEN
        DistDif# = Round#(TransRec(1).MatchSocAmt - TotalSocAmt#)
        EmpAct(LastActive).MATSocAmt = Round(EmpAct(LastActive).MATSocAmt + DistDif#)
      END IF
    END IF
  LOOP UNTIL TotalSocAmt# = TransRec(1).MatchSocAmt
  '-=-=-=-=-=-
  'adjust Medicare Amt
SkipSocDist:
  
  DO
    TotalMedAmt# = 0
    FOR Cnt2 = 1 TO 8
      TotalMedAmt# = Round#(TotalMedAmt# + EmpAct(Cnt2).MATMedAmt)
    NEXT
    IF TotalMedAmt# = 0 THEN GOTO SkipMedDist
    IF TotalMedAmt# <> TransRec(1).MatchMedAmt THEN
      FOR Cnt3 = 8 TO 1 STEP -1
        IF EmpAct(Cnt3).MATMedAmt > 0 THEN
          LastActive = Cnt3
          EXIT FOR
        END IF
      NEXT
      IF TotalMedAmt# > TransRec(1).MatchMedAmt THEN
        DistDif# = Round#(TotalMedAmt# - TransRec(1).MatchMedAmt)
        EmpAct(LastActive).MATMedAmt = Round(EmpAct(LastActive).MATMedAmt - DistDif#)
      ELSEIF TotalMedAmt# < TransRec(1).MatchMedAmt THEN
        DistDif# = Round#(TransRec(1).MatchMedAmt - TotalMedAmt#)
        EmpAct(LastActive).MATMedAmt = Round(EmpAct(LastActive).MATMedAmt + DistDif#)
      END IF
    END IF
  LOOP UNTIL TotalMedAmt# = TransRec(1).MatchMedAmt
  '-=-=-=-=-=-
SkipMedDist:
  
  'adjust Retire Amt
  DO
    TotalRetAmt# = 0
    FOR Cnt2 = 1 TO 8
      TotalRetAmt# = Round#(TotalRetAmt# + EmpAct(Cnt2).MATRetAmt)
    NEXT
    IF TotalRetAmt# = 0 THEN GOTO SkipRetDist
    IF TotalRetAmt# <> TransRec(1).MatchRetAmt THEN
      FOR Cnt3 = 8 TO 1 STEP -1
        IF EmpAct(Cnt3).MATRetAmt > 0 THEN
          LastActive = Cnt3
          EXIT FOR
        END IF
      NEXT
      IF TotalRetAmt# > TransRec(1).MatchRetAmt THEN
        DistDif# = Round#(TotalRetAmt# - TransRec(1).MatchRetAmt)
        EmpAct(LastActive).MATRetAmt = Round(EmpAct(LastActive).MATRetAmt - DistDif#)
      ELSEIF TotalRetAmt# < TransRec(1).MatchRetAmt THEN
        DistDif# = Round#(TransRec(1).MatchRetAmt - TotalRetAmt#)
        EmpAct(LastActive).MATRetAmt = Round(EmpAct(LastActive).MATRetAmt + DistDif#)
      END IF
    END IF
  LOOP UNTIL TotalRetAmt# = TransRec(1).MatchRetAmt

SkipRetDist:
  
  'print this employee's distributions
  FPut DHandle, ENum(1) + EName(1) + BRat(1) + ORat(1) + CrLf$
  DLineCnt = DLineCnt + 1
  FOR Cnt2 = 1 TO Last
    IF LEN(QPTrim$(EmpAct(Cnt2).Acct)) > 0 THEN
      GOSUB PrintEmpDist
    END IF
  NEXT
  
  IF Last > 1 THEN
    GOSUB PrintEmpSubTotal
  END IF
  
  FPut DHandle, CrLf$
  DLincCnt = DLineCnt + 1
  
  ' sum to master distrubtion list
  
  FOR Cnt2 = 1 TO Last          'process wage distributions
    Acct$ = QPTrim$(EmpAct(Cnt2).Acct)
    IF LEN(Acct$) > 0 THEN
      IF NumOfWageAccts > 0 THEN
        FOR Cnt3 = 1 TO NumOfWageAccts
          IF Acct$ = QPTrim$(DistbSumAccts(Cnt3).Acct) THEN
            DistbSumAccts(Cnt3).RWage = Round#(DistbSumAccts(Cnt3).RWage + EmpAct(Cnt2).RWage)
            DistbSumAccts(Cnt3).OWage = Round#(DistbSumAccts(Cnt3).OWage + EmpAct(Cnt2).OWage)
            DistbSumAccts(Cnt3).RHrs = Round#(DistbSumAccts(Cnt3).RHrs + EmpAct(Cnt2).RHrs)
            DistbSumAccts(Cnt3).OHrs = Round#(DistbSumAccts(Cnt3).OHrs + EmpAct(Cnt2).OHrs)
            DistbSumAccts(Cnt3).AddEarn = Round#(DistbSumAccts(Cnt3).AddEarn + EmpAct(Cnt2).AddEarn)
            DistbSumAccts(Cnt3).GROSSPAY = Round#(DistbSumAccts(Cnt3).GROSSPAY + EmpAct(Cnt2).GROSSPAY)
            DistbSumAccts(Cnt3).MATSocAmt = Round#(DistbSumAccts(Cnt3).MATSocAmt + EmpAct(Cnt2).MATSocAmt)
            DistbSumAccts(Cnt3).MATMedAmt = Round#(DistbSumAccts(Cnt3).MATMedAmt + EmpAct(Cnt2).MATMedAmt)
            DistbSumAccts(Cnt3).MATRetAmt = Round#(DistbSumAccts(Cnt3).MATRetAmt + EmpAct(Cnt2).MATRetAmt)
            
            EXIT FOR
          END IF
        NEXT
        IF Cnt3 > NumOfWageAccts THEN
          GOSUB AddDistbSumAcct 'add new sum dist acct
        END IF
      ELSE      'no previous sum accts. add new one
        GOSUB AddDistbSumAcct   'add new sum dist acct
      END IF
    END IF
  NEXT
  
  RETURN
  
AddDistbSumAcct:                'add amts to grand total acts summary
  
  NumOfWageAccts = NumOfWageAccts + 1
  IF NumOfWageAccts > 1 THEN
    REDIM PRESERVE DistbSumAccts(1 TO NumOfWageAccts) AS DistWageRptType
  END IF
  DistbSumAccts(NumOfWageAccts) = EmpAct(Cnt2)
  
  RETURN
  
PrintDistHeader:
  RSET Pg(1) = STR$(EPage)
  FPut DHandle, Unit(1).UFEMPR + SPACE$(87) + "Page:" + Pg(1) + CrLf$
  FPut DHandle, "Earnings Distribution" + CrLf$
  FPut DHandle, "Period Ending: " + Num2Date$(PDR(1).PEREND) + CrLf$
  FPut DHandle, CrLf$
  FPut DHandle, "Employee No  Name                                Base Rate   O/T Rate                          --------- Matching ----------" + CrLf$
  FPut DHandle, "Account Number       Sal%    Reg Hrs    O/T Hrs    Reg Pay    O/T Pay  Tot Other  Gross Pay    Soc Sec   Medicare     Retire" + CrLf$
  FPut DHandle, Dash2(1) + CrLf$
  DLineCnt = DLineCnt + 7
  EPage = EPage + 1
  
  RETURN
  
PrintEmpDist:
  LSET EDAct(1) = EmpAct(Cnt2).Acct
  RSET EDPct(1) = FUsing(STR$(EmpAct(Cnt2).Pct), Image$)
  RSET EDRHrs(1) = FUsing(STR$(EmpAct(Cnt2).RHrs), Image$)
  RSET EDOHrs(1) = FUsing(STR$(EmpAct(Cnt2).OHrs), Image$)
  RSET EDRPay(1) = FUsing(STR$(EmpAct(Cnt2).RWage), Image$)
  RSET EDOPay(1) = FUsing(STR$(EmpAct(Cnt2).OWage), Image$)
  RSET EDEarn(1) = FUsing(STR$(EmpAct(Cnt2).AddEarn), Image$)
  RSET EDGroP(1) = FUsing(STR$(EmpAct(Cnt2).GROSSPAY), Image$)
  
  RSET EDSAmt(1) = FUsing(STR$(EmpAct(Cnt2).MATSocAmt), Image$)
  RSET EDMAmt(1) = FUsing(STR$(EmpAct(Cnt2).MATMedAmt), Image$)
  RSET EDRAmt(1) = FUsing(STR$(EmpAct(Cnt2).MATRetAmt), Image$)
  
  FPut DHandle, EDAct(1) + EDPct(1) + EDRHrs(1) + EDOHrs(1) + EDRPay(1)
  FPut DHandle, EDOPay(1) + EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1) + CrLf$
  DLineCnt = DLineCnt + 1
  RETURN
  
PrintEmpSubTotal:
  RSET RHrs(1) = FUsing(STR$(TransRec(1).RegHrsPaid), Image$)
  RSET OTHrs(1) = FUsing(STR$(TransRec(1).OTHrsPaid), Image$)
  RSET EDRPay(1) = FUsing(STR$(TransRec(1).TotRegWage), Image$)
  RSET EDOPay(1) = FUsing(STR$(TransRec(1).TotOTWage), Image$)
  RSET EDEarn(1) = FUsing(STR$(TransRec(1).TotAdditEarn), Image$)
  RSET EDGroP(1) = FUsing(STR$(TransRec(1).GROSSPAY), Image$)
  
  RSET EDSAmt(1) = FUsing(STR$(TransRec(1).MatchSocAmt), Image$)
  RSET EDMAmt(1) = FUsing(STR$(TransRec(1).MatchMedAmt), Image$)
  RSET EDRAmt(1) = FUsing(STR$(TransRec(1).MatchRetAmt), Image$)
  
  'add sub totals for each employees soc, med + ret amts.
  FPut DHandle, Dash(1) + CrLf$
  FPut DHandle, "Employee Total" + Fill11(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1)
  FPut DHandle, EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1) + CrLf$
  DLineCnt = DLineCnt + 3
  
  RETURN
  
PrintDistTotal:
  
  REGHRS# = 0
  TOTHrs# = 0
  REGWAGE# = 0
  OTWAGE# = 0
  AddEarn# = 0
  DGPay# = 0
  
  DLineCnt = 0
  
  FPut DHandle, FF$
  
  SortT DistbSumAccts(1), NumOfWageAccts, 0, DistbSumAcctsLen, 0, 14
  
  'print the Summary of ALL Distrubtion Accounts
  GOSUB PrintAcctSumHeader
  
  FOR Cnt = 1 TO NumOfWageAccts
    LSET EDAct(1) = DistbSumAccts(Cnt).Acct
    RSET EDPct(1) = FUsing(STR$((DistbSumAccts(Cnt).GROSSPAY / GPay#) * 100), Image4$)
    RSET RHrs(1) = FUsing(STR$(DistbSumAccts(Cnt).RHrs), Image3$)
    REGHRS# = Round#(REGHRS# + DistbSumAccts(Cnt).RHrs)
    RSET OTHrs(1) = FUsing(STR$(DistbSumAccts(Cnt).OHrs), Image3$)
    TOTHrs# = Round#(TOTHrs# + DistbSumAccts(Cnt).OHrs)
    RSET EDRPay(1) = FUsing(STR$(DistbSumAccts(Cnt).RWage), Image3$)
    REGWAGE# = Round#(REGWAGE# + DistbSumAccts(Cnt).RWage)
    RSET EDOPay(1) = FUsing(STR$(DistbSumAccts(Cnt).OWage), Image3$)
    OTWAGE# = Round#(OTWAGE# + DistbSumAccts(Cnt).OWage)
    RSET EDEarn(1) = FUsing(STR$(DistbSumAccts(Cnt).AddEarn), Image3$)
    AddEarn# = Round#(AddEarn# + DistbSumAccts(Cnt).AddEarn)
    RSET EDGroP(1) = FUsing(STR$(DistbSumAccts(Cnt).GROSSPAY), Image3$)
    DGPay# = Round#(DGPay# + DistbSumAccts(Cnt).GROSSPAY)
    
    RSET EDSAmt(1) = FUsing(STR$(DistbSumAccts(Cnt).MATSocAmt), Image3$)
    ASAmt# = Round#(ASAmt# + DistbSumAccts(Cnt).MATSocAmt)
    
    RSET EDMAmt(1) = FUsing(STR$(DistbSumAccts(Cnt).MATMedAmt), Image3$)
    AMAmt# = Round#(AMAmt# + DistbSumAccts(Cnt).MATMedAmt)
    
    RSET EDRAmt(1) = FUsing(STR$(DistbSumAccts(Cnt).MATRetAmt), Image3$)
    ARAmt# = Round#(ARAmt# + DistbSumAccts(Cnt).MATRetAmt)
    
    FPut DHandle, EDAct(1) + EDPct(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1)
    FPut DHandle, EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1) + CrLf$
    
    DLineCnt = DLineCnt + 1
    IF DLineCnt >= DMaxLines THEN
      DLineCnt = 0
      GOSUB PrintAcctSumHeader
    END IF
  NEXT
  
  RSET RHrs(1) = FUsing(STR$(REGHRS#), Image3$)
  RSET OTHrs(1) = FUsing(STR$(TOTHrs#), Image3$)
  RSET EDRPay(1) = FUsing(STR$(REGWAGE#), Image3$)
  RSET EDOPay(1) = FUsing(STR$(OTWAGE#), Image3$)
  RSET EDEarn(1) = FUsing(STR$(AddEarn#), Image3$)
  RSET EDGroP(1) = FUsing(STR$(DGPay#), Image3$)
  
  RSET EDSAmt(1) = FUsing(STR$(ASAmt#), Image3$)
  RSET EDMAmt(1) = FUsing(STR$(AMAmt#), Image3$)
  RSET EDRAmt(1) = FUsing(STR$(ARAmt#), Image3$)
  
  FPut DHandle, Dash2(1) + CrLf$
  FPut DHandle, "Totals        " + Fill11(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1)
  FPut DHandle, EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1) + CrLf$
  
  FPut DHandle, FF$
  
  RETURN
  
PrintAcctSumHeader:
  FPut DHandle, Unit(1).UFEMPR + SPACE$(87) + "Page:" + Pg(1) + CrLf$
  FPut DHandle, "Earnings Distribution Totals" + CrLf$
  FPut DHandle, "Period Ending: " + Num2Date$(PDR(1).PEREND) + CrLf$
  FPut DHandle, "                                                                                               --------- Matching ----------" + CrLf$
  FPut DHandle, "Account Number        Pct    Reg Hrs    O/T Hrs    Reg Pay    O/T Pay  Tot Other  Gross Pay    Soc Sec   Medicare     Retire" + CrLf$
  FPut DHandle, Dash2(1) + CrLf$
  DLineCnt = DLineCnt + 6
  RETURN
  
END SUB

