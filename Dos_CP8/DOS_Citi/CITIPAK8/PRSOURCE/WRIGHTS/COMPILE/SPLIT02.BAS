DEFINT A-Z
DECLARE SUB RPTSetupPRN (RPTNum%, FHandle%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (Number%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRT (Handle%, Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB HideCursor ()
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
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
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Biw'
  '$INCLUDE: 'prif.bi'
  '$INCLUDE: 'GLSETUP.bi'
  
  CONST False = 0, True = NOT False

SUB MakeGLIFFileS (SplitDeds() AS SplitDedsType, DistbSumAccts() AS DistWageRptType, TotEIC#, TotDeds#(), Passed#(), PrintGLRpt)


  FEDWith$ = "FED Withholdings"      '    3

  MEDWith$ = "MED Withholdings"      '    4
  MEDMatLib$ = "MED Match Liab"      '    5

  SOCWith$ = "SOC Withholdings"      '    1
  SOCMatLib$ = "SOC Match Liab"      '    2

  RETWith$ = "RET Withholdings"      '    6
  RETMatLib$ = "RET Match Liab"      '    7

  STAWith$ = "STA Withholdings"      '    8

  MEDMat$ = "MED Match"
  SOCMat$ = "SOC Match"
  RETMat$ = "RET Match"              '
  
  REDIM SysRec(1)   AS RegDSysFileRecType
  REDIM PDR(1)      AS PeriodDefaultRecType
  REDIM DedCodes(1 TO 12)   AS DedCodeRecType
  REDIM Unit(1)     AS UnitFileRecType
  REDIM EDesc(1)    AS STRING * 21
  REDIM Dash0(1) AS STRING * 69
  LSET Dash0(1) = STRING$(69, "-")
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  FGetAH DedCodeFileName, DedCodes(1), LEN(DedCodes(1)), 12
  FGetAH PPDefaultFileName, PDR(1), LEN(PDR(1)), 1
  FGetAH SysFileName, SysRec(1), LEN(SysRec(1)), 1

  REDIM GLSetupRec(1) AS GLSetupRecType
  FundPad = 0
  FundLen = 2
  GLSetUpName$ = QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT"
  IF Exist(GLSetUpName$) THEN
    FGetAH GLSetUpName$, GLSetupRec(1), LEN(GLSetupRec(1)), 1
    FundLen = GLSetupRec(1).FundLen
    FundPad = GLSetupRec(1).DetLen - FundLen
  END IF

  Image3$ = ",######.##"
  Image4$ = "######,.##"

  FACnt = SysRec(1).AcctCnt
  
  NumOfWageAccts = UBOUND(DistbSumAccts)
  
  SplDedCnt = UBOUND(SplitDeds)
  
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
  
  FedAcct$ = SysRec(1).Liab(1).Acct
  ReplaceString FedAcct$, "-", ""
  StaAcct$ = SysRec(1).Liab(2).Acct
  ReplaceString StaAcct$, "-", ""
  SocAcct$ = SysRec(1).Liab(3).Acct
  ReplaceString SocAcct$, "-", ""
  MedAcct$ = SysRec(1).Liab(4).Acct
  ReplaceString MedAcct$, "-", ""
  RetAcct$ = SysRec(1).Liab(5).Acct
  ReplaceString RetAcct$, "-", ""
  
  'ReplaceString GLIFTDate$, "1997", "97"
  '
  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$
  
  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
  
  GLIFArrayAdj = 18
  
  GLIFArrayAdj = SplDedCnt * GLIFArrayAdj
  
  ArraySize = NumOfWageAccts + GLIFArrayAdj + SplDedCnt
  
  REDIM GLIFRec(1 TO ArraySize) AS GLIFDataType14
  
  FOR Cnt = 1 TO NumOfWageAccts
    GLIFRec(Cnt).TranAcct = DistbSumAccts(Cnt).Acct
    GLIFRec(Cnt).TranDate = GLIFTDate$
    GLIFRec(Cnt).TranDesc = "Wages"
    GLIFRec(Cnt).CrAmt = 0
    GLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).GROSSPAY
    GLIFRec(Cnt).Source = GLIFSource$
    GLIFRec(Cnt).FromFlag = "W"
  NEXT
  
  NextAcct = Cnt
  
'06-02-95 split deductions
  FOR Cnt = 0 TO (SplDedCnt - 1)
    GLIFRec(NextAcct + Cnt).TranAcct = QPTrim$(SplitDeds(Cnt + 1).Acct) + FedAcct$
    GLIFRec(NextAcct + Cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + Cnt).TranDesc = FEDWith$
    GLIFRec(NextAcct + Cnt).CrAmt = SplitDeds(Cnt + 1).FedTax
    GLIFRec(NextAcct + Cnt).DrAmt = 0
    GLIFRec(NextAcct + Cnt).Source = GLIFSource$
    GLIFRec(NextAcct + Cnt).FromFlag = "X"
  NEXT
  
  NextAcct = NextAcct + Cnt
'look here
  FOR Cnt = 0 TO (SplDedCnt - 1)
    GLIFRec(NextAcct + Cnt).TranAcct = QPTrim$(SplitDeds(Cnt + 1).Acct) + StaAcct$
    GLIFRec(NextAcct + Cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + Cnt).TranDesc = STAWith$
    GLIFRec(NextAcct + Cnt).CrAmt = SplitDeds(Cnt + 1).StaTax
    GLIFRec(NextAcct + Cnt).DrAmt = 0
    GLIFRec(NextAcct + Cnt).Source = GLIFSource$
    GLIFRec(NextAcct + Cnt).FromFlag = "X"
  NEXT
  
  NextAcct = NextAcct + Cnt
  
  FOR Cnt = 0 TO (SplDedCnt - 1)
    GLIFRec(NextAcct + Cnt).TranAcct = QPTrim$(SplitDeds(Cnt + 1).Acct) + SocAcct$
    GLIFRec(NextAcct + Cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + Cnt).TranDesc = SOCWith$
    GLIFRec(NextAcct + Cnt).CrAmt = SplitDeds(Cnt + 1).SocTax
    GLIFRec(NextAcct + Cnt).DrAmt = 0
    GLIFRec(NextAcct + Cnt).Source = GLIFSource$
    GLIFRec(NextAcct + Cnt).FromFlag = "X"
  NEXT
  
  NextAcct = NextAcct + Cnt
  
  FOR Cnt = 0 TO (SplDedCnt - 1)
    GLIFRec(NextAcct + Cnt).TranAcct = QPTrim$(SplitDeds(Cnt + 1).Acct) + MedAcct$
    GLIFRec(NextAcct + Cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + Cnt).TranDesc = MEDWith$
    GLIFRec(NextAcct + Cnt).CrAmt = SplitDeds(Cnt + 1).MedTax
    GLIFRec(NextAcct + Cnt).DrAmt = 0
    GLIFRec(NextAcct + Cnt).Source = GLIFSource$
    GLIFRec(NextAcct + Cnt).FromFlag = "X"
  NEXT
  
  NextAcct = NextAcct + Cnt
  
  FOR Cnt = 0 TO (SplDedCnt - 1)
    GLIFRec(NextAcct + Cnt).TranAcct = QPTrim$(SplitDeds(Cnt + 1).Acct) + RetAcct$
    GLIFRec(NextAcct + Cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + Cnt).TranDesc = RETWith$
    GLIFRec(NextAcct + Cnt).CrAmt = SplitDeds(Cnt + 1).RetTax
    GLIFRec(NextAcct + Cnt).DrAmt = 0
    GLIFRec(NextAcct + Cnt).Source = GLIFSource$
    GLIFRec(NextAcct + Cnt).FromFlag = "X"
  NEXT
  
  'jumpthis:
  
  NextAcct = NextAcct + Cnt
  
  'good to here ;maybe
  
  ReplaceString SysRec(1).CashAcct, "-", ""
  
  SysCash$ = QPTrim$(SysRec(1).CashAcct)
  
  
  IF TotEIC# > 0 THEN
    REDIM EICGLIFRec(1 TO (SplDedCnt * 2)) AS GLIFDataType14
    FOR Cnt = 1 TO SplDedCnt
      EICGLIFRec(Cnt).TranAcct = QPTrim$(SplitDeds(Cnt).Acct) + FedAcct$
      EICGLIFRec(Cnt).TranDate = GLIFTDate$
      EICGLIFRec(Cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(Cnt).CrAmt = 0
      EICGLIFRec(Cnt).DrAmt = SplitDeds(Cnt).EICPay
      EICGLIFRec(Cnt).Source = GLIFSource$
      EICGLIFRec(Cnt).FromFlag = "X"
    NEXT
    
    FOR Cnt = SplDedCnt + 1 TO SplDedCnt * 2
      EICGLIFRec(Cnt).TranAcct = QPTrim$(SplitDeds(Cnt - SplDedCnt).Acct) + SysCash$
      EICGLIFRec(Cnt).TranDate = GLIFTDate$
      EICGLIFRec(Cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(Cnt).CrAmt = SplitDeds(Cnt - SplDedCnt).EICPay
      EICGLIFRec(Cnt).DrAmt = 0
      EICGLIFRec(Cnt).Source = GLIFSource$
      EICGLIFRec(Cnt).FromFlag = "X"
    NEXT
    
  END IF

  FOR ZZCnt = 0 TO (SplDedCnt - 1)
    FOR Cnt = 1 TO 12
      ReplaceString DedCodes(Cnt).DCACCT1, "-", ""
      TAcct$ = QPTrim$(SplitDeds(ZZCnt + 1).Acct)
      TAcct$ = TAcct$ + QPTrim$(DedCodes(Cnt).DCACCT1)
      GLIFRec(NextAcct + (Cnt - 1)).TranAcct = TAcct$
      GLIFRec(NextAcct + (Cnt - 1)).TranDate = GLIFTDate$
      GLIFRec(NextAcct + (Cnt - 1)).TranDesc = "Deductions"
      GLIFRec(NextAcct + (Cnt - 1)).CrAmt = SplitDeds(ZZCnt + 1).DedAmts(Cnt)
      GLIFRec(NextAcct + (Cnt - 1)).DrAmt = 0
      GLIFRec(NextAcct + (Cnt - 1)).Source = GLIFSource$
      GLIFRec(NextAcct + (Cnt - 1)).FromFlag = "D"
    NEXT
    NextAcct = NextAcct + Cnt
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
  
  ''all gross pay by funds here!!
  
  REDIM CFunds$(1 TO GLIFArrayAdj)              '17)
  
  FirstC = NumOfWageAccts + 1
  LastC = NumOfWageAccts + GLIFArrayAdj + 1     '17
  
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
  '
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
  
  ''combine all funds in one array here
  TotalFunds = NumDFunds + NumCFunds            '+ 1
  REDIM AllFunds(1 TO TotalFunds) AS FundType
  ThisFund = 1
  FOR Cnt = 1 TO NumDFunds
    AllFunds(ThisFund) = DFund(Cnt)
    ThisFund = ThisFund + 1
  NEXT
  
'  IF NoCFunds = False THEN
    FOR Cnt = 1 TO NumCFunds
      AllFunds(ThisFund) = CFund(Cnt)
      ThisFund = ThisFund + 1
    NEXT
'  END IF
  
  ERASE DFund, CFund
  '
'SortT AllFunds(1), TotalFunds, 0, LEN(AllFunds(1)), 0, 9
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
  
  TotalGLIFS = NumOfWageAccts + GLIFArrayAdj + TotalFunds       '17
  
  REDIM PRESERVE GLIFRec(1 TO TotalGLIFS) AS GLIFDataType14
  
  NextAcct = NumOfWageAccts + GLIFArrayAdj      '17
  
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
    NoMatchFlag = True
    GOTO WriteGLIFS
  ELSE
    NoMatchFlag = False
  END IF
  
  REDIM SFGLIFRec(1 TO TotalFunds)  AS GLIFDataType14
  REDIM MFGLIFRec(1 TO TotalFunds)  AS GLIFDataType14
  REDIM RFGLIFRec(1 TO TotalFunds)  AS GLIFDataType14
  REDIM SocGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14
  REDIM MedGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14
  REDIM RetGLIFRec(1 TO NumOfWageAccts)  AS GLIFDataType14
  
  SOCEXP$ = QPTrim$(SysRec(1).SOCEXP)
  MEDEXP$ = QPTrim$(SysRec(1).MEDEXP)
  RETEXP$ = QPTrim$(SysRec(1).RETEXP)
  SOCLIAB$ = QPTrim$(SysRec(1).SOCLIAB)
  MEDLIAB$ = QPTrim$(SysRec(1).MEDLIAB)
  RETLIAB$ = QPTrim$(SysRec(1).RETLIAB)
  '
  ReplaceString SOCLIAB$, "-", ""
  ReplaceString MEDLIAB$, "-", ""
  ReplaceString RETLIAB$, "-", ""
  
  FOR Cnt = 1 TO NumOfWageAccts
    'social
    SocGLIFRec(Cnt).TranAcct = LEFT$(DistbSumAccts(Cnt).Acct, FACnt) + SOCEXP$
    SocGLIFRec(Cnt).TranDate = GLIFTDate$
    SocGLIFRec(Cnt).TranDesc = SOCMat$
    SocGLIFRec(Cnt).Source = GLIFSource$
    SocGLIFRec(Cnt).FromFlag = "S"
    SocGLIFRec(Cnt).CrAmt = 0
    SocGLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).MATSocAmt
    'medicare
    MedGLIFRec(Cnt).TranAcct = LEFT$(DistbSumAccts(Cnt).Acct, FACnt) + MEDEXP$
    MedGLIFRec(Cnt).TranDate = GLIFTDate$
    MedGLIFRec(Cnt).TranDesc = MEDMat$
    MedGLIFRec(Cnt).Source = GLIFSource$
    MedGLIFRec(Cnt).FromFlag = "M"
    MedGLIFRec(Cnt).CrAmt = 0
    MedGLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).MATMedAmt
    'retirment
    RetGLIFRec(Cnt).TranAcct = LEFT$(DistbSumAccts(Cnt).Acct, FACnt) + RETEXP$
    RetGLIFRec(Cnt).TranDate = GLIFTDate$
    RetGLIFRec(Cnt).TranDesc = RETMat$
    RetGLIFRec(Cnt).Source = GLIFSource$
    RetGLIFRec(Cnt).FromFlag = "R"
    RetGLIFRec(Cnt).CrAmt = 0
    RetGLIFRec(Cnt).DrAmt = DistbSumAccts(Cnt).MATRetAmt
  NEXT
  
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
        SFGLIFRec(Cnt2).TranDesc = SOCMatLib$
        SFGLIFRec(Cnt2).Source = GLIFSource$
        SFGLIFRec(Cnt2).FromFlag = "s"
      END IF
    NEXT
  NEXT
  '
  FOR Cnt = 1 TO NumOfWageAccts
    FOR Cnt2 = 1 TO TotalFunds
      'fixed
      IF LEFT$(MFGLIFRec(Cnt2).TranAcct, FundLen) = LEFT$(MedGLIFRec(Cnt).TranAcct, FundLen) THEN
        MFGLIFRec(Cnt2).CrAmt = Round(MFGLIFRec(Cnt2).CrAmt + MedGLIFRec(Cnt).DrAmt)
        MFGLIFRec(Cnt2).DrAmt = 0
        MFGLIFRec(Cnt2).TranDate = GLIFTDate$
        MFGLIFRec(Cnt2).TranDesc = MEDMatLib$
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
        RFGLIFRec(Cnt2).TranDesc = RETMatLib$
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

  IF NoMatchFlag THEN GOTO NoMatchJump

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

'  END SELECT
  
NoMatchJump:

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
    FOR Cnt = 1 TO (SplDedCnt * 2)
      IF EICGLIFRec(Cnt).DrAmt <> 0 OR EICGLIFRec(Cnt).CrAmt <> 0 THEN
        FPutRTA GHandle, EICGLIFRec(Cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      END IF
    NEXT
  END IF
  '
  FClose GHandle

  REDIM GLSAmt(1)   AS STRING * 11
  REDIM GLMAmt(1)   AS STRING * 11

'**** NEW

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
  
  '  GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile

  TotalGLIFS = FileSize("TempIF.DAT") \ GLIFRecLen
  TotalAccts = FileSize(GLIdxName$) \ GLIdxRecLen
  
  IF TotalGLIFS = 0 THEN
    PrintGLRpt = False
    GOTO SkipGLRpt
  ELSE
    PrintGLRpt = True
  END IF
  
  REDIM GLIFRec(1 TO TotalGLIFS)        AS GLIFDataType14
  FGetAH "TempIF.DAT", GLIFRec(1), GLIFRecLen, TotalGLIFS
  
  SortT GLIFRec(1), TotalGLIFS, 0, GLIFRecLen, 0, 14

'  DO
'    OutOfOrder = False                     'assume it's sorted
'    FOR X = 1 TO UBOUND(GLIFRec) - 1
'      IF GLIFRec(X).TranAcct > GLIFRec(X + 1).TranAcct THEN
'        SWAP GLIFRec(X), GLIFRec(X + 1)    'if we had to swap
'        OutOfOrder = True                  'we're not done yet
'      END IF
'    NEXT
'  LOOP WHILE OutOfOrder

'04-05-96 Started Add summary to glif report

  FPutAH "TempIF.DAT", GLIFRec(1), GLIFRecLen, TotalGLIFS

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
    REDIM JGoodAccts(1 TO TotalAccts)      AS STRING * 16
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
  FOpenS RptName$, GHandle
  RPTSetupPRN 14, GHandle
  
  FOpenS "TempIF.DAT", THandle

  GOSUB GLIFHeader

'**** NEW

  FOR Cnt = 1 TO TotalGLIFS
    ShowPctComp Cnt, TotalGLIFS
    IF Lines >= Max THEN
      FPut GHandle, FF$
      GOSUB GLIFHeader
    END IF

    TotDebit# = Round(TotDebit# + GLIFRec(Cnt).DrAmt)
    TotCredit# = Round(TotCredit# + GLIFRec(Cnt).CrAmt)

    '10-13-94 ** added check for valid GL Account numbers
    IF TotalAccts > 0 THEN
      AcctOk = False
      
      SELECT CASE JFlag
      CASE True
        FOR Cnt2 = 1 TO TotalAccts
          IF INSTR(JGoodAccts(Cnt2), QPTrim$(GLIFRec(Cnt).TranAcct)) AND LEN(QPTrim$(JGoodAccts(Cnt2))) = LEN(QPTrim$(GLIFRec(Cnt).TranAcct)) THEN
            AcctOk = True
            EXIT FOR
          END IF
        NEXT
        IF NOT AcctOk THEN
          LSET GLIFRec(Cnt).Fill = "Error "
          FPutRTA THandle%, GLIFRec(Cnt), CLNG(Cnt), LEN(GLIFRec(Cnt))
        ELSE
          LSET GLIFRec(Cnt).Fill = ""
        END IF
        
      CASE False
        GLAcct@ = Round(VAL(GLIFRec(Cnt).TranAcct))
        IF GLAcct@ < 9999999 THEN
          'ON ERROR RESUME NEXT
          GLAcct@ = GLAcct@ * 100
        END IF
        FOR Cnt2 = 1 TO TotalAccts
          IF GoodAccts(Cnt2) = GLAcct@ THEN
            AcctOk = True
            EXIT FOR
          END IF
        NEXT
        IF NOT AcctOk THEN
          LSET GLIFRec(Cnt).Fill = "Error "
          FPutRTA THandle%, GLIFRec(Cnt), CLNG(Cnt), LEN(GLIFRec(Cnt))
        ELSE
          LSET GLIFRec(Cnt).Fill = ""
        END IF
      END SELECT
    END IF

    'NoCheckAccts:
    RSET GLSAmt(1) = FUsing(STR$(GLIFRec(Cnt).DrAmt), Image3$)
    RSET GLMAmt(1) = FUsing(STR$(GLIFRec(Cnt).CrAmt), Image3$)
    LSET EDesc(1) = QPTrim$(GLIFRec(Cnt).TranDesc)
    NoAcctNum = LEN(QPTrim$(GLIFRec(Cnt).TranAcct))
    IF NoAcctNum > 0 THEN
      IF AcctOk = True THEN
        LSET GLIFRec(Cnt).Fill = " "
      END IF
      Trans$ = QPTrim$(GLIFRec(Cnt).TranAcct)
      LSET GLIFRec(Cnt).TranAcct = " "
      LSET GLIFRec(Cnt).TranAcct = Trans$
      'FPut GHandle, "  " + GLIFRec(Cnt).TranAcct + GLIFRec(Cnt).Fill + EDesc(1) + GLSAmt(1) + "    " + GLMAmt(1) + CrLf$
      'fixed: 05-27-96
      FPut GHandle, " " + GLIFRec(Cnt).TranAcct + GLIFRec(Cnt).Fill + " " + EDesc(1) + GLSAmt(1) + "    " + GLMAmt(1) + CrLf$
      Lines = Lines + 1
    END IF
  NEXT
  GOSUB GLIFTotals
  'fixed
  Fund$ = LEFT$(QPTrim$(GLIFRec(1).TranAcct), FundLen)
  FDebit# = GLIFRec(1).DrAmt
  FCredit# = GLIFRec(1).CrAmt
  FOR Cnt = 2 TO TotalGLIFS
    'fixed
    NFund$ = LEFT$(QPTrim$(GLIFRec(Cnt).TranAcct), FundLen)
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

  FPut GHandle, FF$

  FClose THandle

  GOSUB PrintTaxTotals

  RPTSetupPRN 0, GHandle

  FClose GHandle
  
  'GOSUB PrintTaxTotals

  ERASE GoodAccts, GLSAmt, GLMAmt, GLIFRec
  ERASE SFGLIFRec, MFGLIFRec, RFGLIFRec
  ERASE SocGLIFRec, MedGLIFRec, RetGLIFRec
  ERASE GLIdxRec, CDGLIFRec     '
  
SkipGLRpt:
  ERASE GLIFRec, SysRec, PDR, AllFunds, DFunds$, CFunds$
  ERASE DedCodes, Unit, EDesc, Dash0
  
  EXIT SUB
  
PrintFundTotal:
  RSET GLSAmt(1) = FUsing(STR$(FDebit#), Image3$)
  RSET GLMAmt(1) = FUsing(STR$(FCredit#), Image3$)
  LSET EDesc(1) = ""
  LSET GLIFRec(1).Fill = ""
  RSET GLIFRec(1).TranAcct = "Fund  " + Fund$
  FPut GHandle, "   " + GLIFRec(1).TranAcct + GLIFRec(1).Fill + EDesc(1) + GLSAmt(1) + "    " + GLMAmt(1) + CrLf$
RETURN

GLIFTotals:
  FPut GHandle, CrLf$
  FPut GHandle, Dash0(1) + CrLf$
  FPut GHandle, "   Totals:                                  " + FUsing(STR$(TotDebit#), Image3$) + "     " + FUsing(STR$(TotCredit#), Image3$) + CrLf$
  IF TotalAccts = 0 THEN
    FPut GHandle, CrLf$
    FPut GHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory." + CrLf$
  END IF
  FPut GHandle, CrLf$
  FPut GHandle, "SubTotals:" + CrLf$
RETURN
  
GLIFHeader:
  FPut GHandle, QPTrim$(Unit(1).UFEMPR) + CrLf$
  FPut GHandle, "G/L Interface Report." + CrLf$
  FPut GHandle, "Period Ending: " + Num2Date$(PDR(1).PEREND) + CrLf$
  FPut GHandle, CrLf$
  FPut GHandle, "Account No.           Description                Debit         Credit" + CrLf$
  FPut GHandle, Dash0(1) + CrLf$
  Lines = 6
RETURN

PrintTaxTotals:
  
  REDIM AcctSum(1 TO TotalGLIFS) AS AcctSumType
  REDIM GLIFRec(1 TO TotalGLIFS)  AS GLIFDataType14

'Tax summary array
'  REDIM TaxSum(1 TO 7) AS AcctSumType

  AcctSumLen = LEN(AcctSum(1))

  FGetAH "TempIF.DAT", GLIFRec(1), GLIFRecLen, TotalGLIFS
  SortT GLIFRec(1), TotalGLIFS, 0, GLIFRecLen, 0, 14

  WhatAcct = 0

  FirstInAcct = 1
  FOR Cnt = 1 TO TotalGLIFS - 1
    IF VAL(GLIFRec(Cnt).TranAcct) <> VAL(GLIFRec(Cnt + 1).TranAcct) THEN
      LastInAcct = Cnt
      GOSUB SumThisAcct
      FirstInAcct = Cnt + 1
    ELSE
    END IF
  NEXT

  GOSUB SumThisAcct:

  GOSUB PrintAcctSumHeader:

  FOR Cnt = 1 TO WhatAcct
    TempLine$ = SPACE$(44)
    Acct$ = QPTrim$(AcctSum(Cnt).FundCode)
    LSET TempLine$ = Acct$
    FPut GHandle, TempLine$ + FUsing$(STR$(AcctSum(Cnt).Debit), Image4$) + "     " + FUsing$(STR$(AcctSum(Cnt).Credit), Image4$) + CrLf$
    Lines = Lines + 1
    IF Lines >= Max THEN
      FPut GHandle, FF$
      GOSUB PrintAcctSumHeader:
    END IF
  NEXT
  FPut GHandle, FF$


'     Messing with tax summary
'  SortT AcctSum(1), WhatAcct, 0, AcctSumLen, FundLen - 1, 14 - (FundLen + 1)
'
'  FirstInAcct = 1
'  FOR Cnt = 1 TO WhatAcct - 1
'    IF VAL(QPTrim$(MID$(AcctSum(Cnt).FundCode, FundLen + 1))) <> VAL(QPTrim$(MID$(AcctSum(Cnt + 1).FundCode, FundLen + 1))) THEN
'      LastInAcct = Cnt
'      GOSUB SumThisObject
'      FirstInAcct = Cnt + 1
'    ELSE
'    END IF
'  NEXT
'
'  FOR Cnt = 1 TO 5
'    DashPos = INSTR(SysRec(1).Liab(Cnt).Acct, "-")
'    DO WHILE DashPos
'      SysRec(1).Liab(Cnt).Acct = LEFT$(SysRec(1).Liab(Cnt).Acct, DashPos - 1) + MID$(SysRec(1).Liab(Cnt).Acct, DashPos + 1)
'      DashPos = INSTR(SysRec(1).Liab(Cnt).Acct, "-")
'    LOOP
'    TaxSum(Cnt).FundCode = SysRec(1).Liab(Cnt).Acct
'  NEXT
'
'  FedAcctNo$ = QPTrim$(SysRec(1).Liab(1).Acct)
'  StaAcctNo$ = QPTrim$(SysRec(1).Liab(2).Acct)
'  SocAcctNo$ = QPTrim$(SysRec(1).Liab(3).Acct)
'  MedAcctNo$ = QPTrim$(SysRec(1).Liab(4).Acct)
'  RetAcctNo$ = QPTrim$(SysRec(1).Liab(5).Acct)
 ' SocLiabNo$ = QPTrim$(SysRec(1).SOCLIAB)
'  MedLiabNo$ = QPTrim$(SysRec(1).MEDLIAB)
'
'  DashPos = INSTR(SocLiabNo$, "-")
'  DO WHILE DashPos
'    SocLiabNo$ = LEFT$(SocLiabNo$, DashPos - 1) + MID$(SocLiabNo$, DashPos + 1)
'    DashPos = INSTR(SocLiabNo$, "-")
'  LOOP
'  DashPos = INSTR(MedLiabNo$, "-")
'  DO WHILE DashPos
'    MedLiabNo$ = LEFT$(MedLiabNo$, DashPos - 1) + MID$(MedLiabNo$, DashPos + 1)
'    DashPos = INSTR(MedLiabNo$, "-")
'  LOOP
'
'  TaxSum(6).FundCode = SocLiabNo$
'  TaxSum(7).FundCode = MedLiabNo$
'
'
'  FOR Cnt = 1 TO TWhatAcct
'
'   TempAcct$ = QPTrim$(AcctSum(Cnt).FundCode)
'   SELECT CASE TempAcct$
'   CASE FedAcctNo$, SocAcctNo$, MedAcctNo$
'     TaxSum(1).Credit = Round#(TaxSum(1).Credit + AcctSum(Cnt).Credit)
'     TaxSum(1).Debit = Round#(TaxSum(1).Debit + AcctSum(Cnt).Debit)
'
'   CASE StaAcctNo$
'     TaxSum(2).Credit = Round#(TaxSum(2).Credit + AcctSum(Cnt).Credit)
'     TaxSum(2).Debit = Round#(TaxSum(2).Debit + AcctSum(Cnt).Debit)
'
''   CASE SocAcctNo$
''     TaxSum(3).Credit = Round#(TaxSum(3).Credit + AcctSum(Cnt).Credit)
''     TaxSum(3).Debit = Round#(TaxSum(3).Debit + AcctSum(Cnt).Debit)
'
''   CASE MedAcctNo$
''     TaxSum(4).Credit = Round#(TaxSum(4).Credit + AcctSum(Cnt).Credit)
''     TaxSum(4).Debit = Round#(TaxSum(4).Debit + AcctSum(Cnt).Debit)
'
'   CASE RetAcctNo$
'     TaxSum(5).Credit = Round#(TaxSum(5).Credit + AcctSum(Cnt).Credit)
'     TaxSum(5).Debit = Round#(TaxSum(5).Debit + AcctSum(Cnt).Debit)
'
'   CASE SocLiabNo$
'     TaxSum(6).Credit = Round#(TaxSum(6).Credit + AcctSum(Cnt).Credit)
'     TaxSum(6).Debit = Round#(TaxSum(6).Debit + AcctSum(Cnt).Debit)
'
'   CASE MedLiabNo$
'     TaxSum(7).Credit = Round#(TaxSum(7).Credit + AcctSum(Cnt).Credit)
'     TaxSum(7).Debit = Round#(TaxSum(7).Debit + AcctSum(Cnt).Debit)
'
'   END SELECT
'
'  NEXT

'  FOR Cnt = 1 TO 7
'    IF Cnt = 3 OR Cnt = 4 THEN
'    ELSE
'    PRINT TaxSum(Cnt).FundCode, TaxSum(Cnt).Credit, TaxSum(Cnt).Debit
'    END IF
'  NEXT
'  CALL WaitForAction
'
  'GOSUB TaxSumHeader:
RETURN

'****for tax summary
'SumThisObject:
'  TWhatAcct = TWhatAcct + 1
'  AcctSum(TWhatAcct).FundCode = QPTrim$(MID$(AcctSum(FirstInAcct).FundCode, FundLen + 1))
'  FOR AddWhatAcct = FirstInAcct TO LastInAcct
'    'AcctSum(TWhatAcct).FundCode = QPTrim$(MID$(AcctSum(FirstInAcct).FundCode, FundLen + 1))
'    AcctSum(TWhatAcct).Credit = Round#(AcctSum(TWhatAcct).Credit + GLIFRec(AddWhatAcct).CrAmt)
'    AcctSum(TWhatAcct).Debit = Round#(AcctSum(TWhatAcct).Debit + GLIFRec(AddWhatAcct).DrAmt)
'  NEXT
'
'RETURN
'*****

SumThisAcct:
  WhatAcct = WhatAcct + 1
  AcctSum(WhatAcct).FundCode = GLIFRec(FirstInAcct).TranAcct
  FOR AddWhatAcct = FirstInAcct TO LastInAcct
    AcctSum(WhatAcct).Credit = Round#(AcctSum(WhatAcct).Credit + GLIFRec(AddWhatAcct).CrAmt)
    AcctSum(WhatAcct).Debit = Round#(AcctSum(WhatAcct).Debit + GLIFRec(AddWhatAcct).DrAmt)
  NEXT
RETURN

PrintAcctSumHeader:
  FPut GHandle, QPTrim$(Unit(1).UFEMPR) + CrLf$
  FPut GHandle, "GL Posting Summary" + CrLf$
  FPut GHandle, "Period Ending: " + Num2Date$(PDR(1).PEREND) + CrLf$
  FPut GHandle, CrLf$
  FPut GHandle, "Account No.                                      Debit         Credit" + CrLf$
  FPut GHandle, Dash0(1) + CrLf$
  Lines = 6
RETURN



END SUB

