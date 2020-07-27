Attribute VB_Name = "modSplit02"
Option Explicit
  

Sub MakeGLIFFileS(SplitDeds() As SplitDedsType, DistbSumAccts() As DistWageRptType, TotEIC#, TotDeds#(), Passed#(), PrintGLRpt)

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


  ReDim SysRec(1) As RegDSysFileRecType
  ReDim PDR(1) As PeriodDefaultRecType
  ReDim DedCodes(1 To 12) As DedCodeRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim EDesc(1) As String * 21
  ReDim Dash0(1) As String * 69
  LSet Dash0(1) = String$(69, "-")
  
  FGetAH UnitFileName, Unit(1), Len(Unit(1)), 1
  FGetAH DedCodeFileName, DedCodes(1), Len(DedCodes(1)), 12
  FGetAH PPDefaultFileName, PDR(1), Len(PDR(1)), 1
  FGetAH SysFileName, SysRec(1), Len(SysRec(1)), 1
  
  Image3$ = ",######.##"
  
  FACnt = SysRec(1).AcctCnt
  
  NumOfWageAccts = UBound(DistbSumAccts)
  

  SplDedCnt = UBound(SplitDeds)
  
  'squeeze out all the "-" out of Acct numbers
  For cnt = 1 To NumOfWageAccts
    ReplaceString DistbSumAccts(cnt).Acct, "-", ""
  Next
  
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
  ReplaceString GLIFTDate$, "2001", "01"
  ReplaceString GLIFTDate$, "2002", "02"
  ReplaceString GLIFTDate$, "2003", "03"

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
  

  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$

  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
  
  GLIFArrayAdj = 18
  GLIFArrayAdj = SplDedCnt * GLIFArrayAdj
  
  ArraySize = NumOfWageAccts + GLIFArrayAdj + SplDedCnt
  
  ReDim GLIFRec(1 To ArraySize) As GLIFDataType14
  
  For cnt = 1 To NumOfWageAccts
    GLIFRec(cnt).TranAcct = DistbSumAccts(cnt).Acct
    GLIFRec(cnt).TranDate = GLIFTDate$
    GLIFRec(cnt).TranDesc = "Wages"
    GLIFRec(cnt).CrAmt = 0
    GLIFRec(cnt).DrAmt = DistbSumAccts(cnt).GrossPay
    GLIFRec(cnt).Source = GLIFSource$
    GLIFRec(cnt).FromFlag = "W"
  Next
  
  NextAcct = cnt
  
'06-02-95 split deductions
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + FedAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = FEDWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).FEDTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next
  
  NextAcct = NextAcct + cnt
'look here
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + StaAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = STAWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).StaTax
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next
  
  NextAcct = NextAcct + cnt
  
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + SocAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = SOCWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).SOCTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next
  
  NextAcct = NextAcct + cnt
  
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + MedAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = MEDWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).MEDTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next
  
  NextAcct = NextAcct + cnt
  
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + RetAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = RETWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).RetTax
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next
  
  'jumpthis:
  
  NextAcct = NextAcct + cnt
  
  'good to here ;maybe
  
  ReplaceString SysRec(1).CASHACCT, "-", ""
  
  SysCash$ = QPTrim$(SysRec(1).CASHACCT)
  
  If TotEIC# > 0 Then
    ReDim EICGLIFRec(1 To (SplDedCnt * 2)) As GLIFDataType14
    For cnt = 1 To SplDedCnt
      EICGLIFRec(cnt).TranAcct = QPTrim$(SplitDeds(cnt).Acct) + FedAcct$
      EICGLIFRec(cnt).TranDate = GLIFTDate$
      EICGLIFRec(cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(cnt).CrAmt = 0
      EICGLIFRec(cnt).DrAmt = SplitDeds(cnt).EICPay
      EICGLIFRec(cnt).Source = GLIFSource$
      EICGLIFRec(cnt).FromFlag = "X"
    Next
    
    For cnt = SplDedCnt + 1 To SplDedCnt * 2
      EICGLIFRec(cnt).TranAcct = QPTrim$(SplitDeds(cnt - SplDedCnt).Acct) + SysCash$
      EICGLIFRec(cnt).TranDate = GLIFTDate$
      EICGLIFRec(cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(cnt).CrAmt = SplitDeds(cnt - SplDedCnt).EICPay
      EICGLIFRec(cnt).DrAmt = 0
      EICGLIFRec(cnt).Source = GLIFSource$
      EICGLIFRec(cnt).FromFlag = "X"
    Next
    
  End If

  For ZZCnt = 0 To (SplDedCnt - 1)
    For cnt = 1 To 12
      ReplaceString DedCodes(cnt).DCACCT1, "-", ""
      TAcct$ = QPTrim$(SplitDeds(ZZCnt + 1).Acct)
      TAcct$ = TAcct$ + QPTrim$(DedCodes(cnt).DCACCT1)
      GLIFRec(NextAcct + (cnt - 1)).TranAcct = TAcct$
      GLIFRec(NextAcct + (cnt - 1)).TranDate = GLIFTDate$
      GLIFRec(NextAcct + (cnt - 1)).TranDesc = DedCodes(cnt).DCDESC1
      GLIFRec(NextAcct + (cnt - 1)).CrAmt = SplitDeds(ZZCnt + 1).DedAmts(cnt)
      GLIFRec(NextAcct + (cnt - 1)).DrAmt = 0
      GLIFRec(NextAcct + (cnt - 1)).Source = GLIFSource$
      GLIFRec(NextAcct + (cnt - 1)).FromFlag = "D"
    Next
    NextAcct = NextAcct + cnt
  Next
  
  ReDim DFunds$(1 To NumOfWageAccts)

  NumDFunds = 1
  DFunds$(1) = Left$(DistbSumAccts(1).Acct, 2)
  For cnt = 1 To NumOfWageAccts - 1
    If Left$(DistbSumAccts(cnt).Acct, 2) <> Left$(DistbSumAccts(cnt + 1).Acct, 2) Then
      NumDFunds = NumDFunds + 1 'counting the total number of funds.
      DFunds$(NumDFunds) = Left$(DistbSumAccts(cnt + 1).Acct, 2)
    End If
  Next
  
  ReDim Preserve DFunds$(1 To NumDFunds)
  
  ReDim DFund(1 To NumDFunds) As FundType
  
  For cnt = 1 To NumOfWageAccts
    CurrFund$ = Left$(DistbSumAccts(cnt).Acct, 2)
    For ThisFund = 1 To NumDFunds
      If CurrFund$ = DFunds$(ThisFund) Then
        DFund(ThisFund).FundCode = DFunds$(ThisFund)
        DFund(ThisFund).Debit = Round(DFund(ThisFund).Debit + DistbSumAccts(cnt).GrossPay)
        Exit For
      End If
    Next
  Next
  
  ''all gross pay by funds here!!
  
  ReDim CFunds$(1 To GLIFArrayAdj)              '17)
  
  FirstC = NumOfWageAccts + 1
  LastC = NumOfWageAccts + GLIFArrayAdj + 1     '17
  
  NumCFunds = 1
  CFunds$(1) = Left$(GLIFRec(FirstC).TranAcct, 2)
  
  For cnt = FirstC To LastC - 1
    If Left$(GLIFRec(cnt).TranAcct, 2) <> Left$(GLIFRec(cnt + 1).TranAcct, 2) Then
      If Len(QPTrim$(GLIFRec(cnt + 1).TranAcct)) Then
        NumCFunds = NumCFunds + 1               'counting the total number of funds.
        CFunds$(NumCFunds) = Left$(GLIFRec(cnt + 1).TranAcct, 2)
      End If
    End If
  Next
  '
  ReDim Preserve CFunds$(1 To NumCFunds)
  ReDim CFund(1 To NumCFunds) As FundType
  For cnt = FirstC To LastC - 1
    CurrFund$ = Left$(GLIFRec(cnt).TranAcct, 2)
    For ThisFund = 1 To NumCFunds
      If CurrFund$ = CFunds$(ThisFund) Then
        CFund(ThisFund).FundCode = CFunds$(ThisFund)
        CFund(ThisFund).Credit = Round(CFund(ThisFund).Credit + GLIFRec(cnt).CrAmt)
        Exit For
      End If
    Next
  Next
  
  ''combine all funds in one array here
  TotalFunds = NumDFunds + NumCFunds            '+ 1
  ReDim AllFunds(1 To TotalFunds) As FundType
  ThisFund = 1
  For cnt = 1 To NumDFunds
    AllFunds(ThisFund) = DFund(cnt)
    ThisFund = ThisFund + 1
  Next
  
'  IF NoCFunds = False THEN
    For cnt = 1 To NumCFunds
      AllFunds(ThisFund) = CFund(cnt)
      ThisFund = ThisFund + 1
    Next
'  END IF
  
  Erase DFund, CFund
  '
  SortT AllFunds(1), TotalFunds, 0, Len(AllFunds(1)), 0, 9
  'combine Debits and Credits for same fund numbers
  
  First = 1
  Start = 1
  Last = TotalFunds
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If AllFunds(First).FundCode = AllFunds(Cnt2).FundCode Then
        LSet AllFunds(Cnt2).FundCode = ""       'blank acct num as a flag
        AllFunds(First).Credit = Round#(AllFunds(First).Credit + AllFunds(Cnt2).Credit)
        AllFunds(Cnt2).Credit = 0
        AllFunds(First).Debit = Round#(AllFunds(First).Debit + AllFunds(Cnt2).Debit)
        AllFunds(Cnt2).Debit = 0
      End If
    Next
    First = First + 1
  Loop Until First >= Last      '

  'calc net difference for funds
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(AllFunds(cnt).FundCode)) Then
      AllFunds(cnt).Net = Round#(AllFunds(cnt).Debit - AllFunds(cnt).Credit)
    End If
  Next
  
  'add cash acct number to fund numbers
  For cnt = 1 To TotalFunds
    FundCash$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(FundCash$) Then
      LSet AllFunds(cnt).FundCode = FundCash$ + SysCash$
    End If
  Next
  
  TotalGLIFS = NumOfWageAccts + GLIFArrayAdj + TotalFunds       '17
  
  ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14
  
  NextAcct = NumOfWageAccts + GLIFArrayAdj      '17
  
  For cnt = 1 To TotalFunds
    AcctNum$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(AcctNum$) Then
      NextAcct = NextAcct + 1
      GLIFRec(NextAcct).TranAcct = AcctNum$
      GLIFRec(NextAcct).TranDate = GLIFTDate$
      GLIFRec(NextAcct).TranDesc = "PR Net "
      GLIFRec(NextAcct).Source = GLIFSource$
      GLIFRec(NextAcct).FromFlag = "N"
      If AllFunds(cnt).Net > 0 Then
        GLIFRec(NextAcct).CrAmt = AllFunds(cnt).Net
        GLIFRec(NextAcct).DrAmt = 0
      ElseIf AllFunds(cnt).Net < 0 Then
        GLIFRec(NextAcct).DrAmt = Abs(AllFunds(cnt).Net)
        GLIFRec(NextAcct).CrAmt = 0
      End If
    End If
  Next
  
  'if using the imprest account then
  Select Case SysRec(1).USEIMP
  Case "I"      'was Y            'I C P
    TotalGLIFS = TotalGLIFS + 2
    ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    ReplaceString SysRec(1).ICRACCT, "-", ""
    GLIFRec(TotalGLIFS - 1).TranAcct = QPTrim$(SysRec(1).ICRACCT)
    GLIFRec(TotalGLIFS - 1).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS - 1).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS - 1).Source = GLIFSource$
    GLIFRec(TotalGLIFS - 1).FromFlag = "i"
    GLIFRec(TotalGLIFS - 1).DrAmt = Passed#(6)
    GLIFRec(TotalGLIFS - 1).CrAmt = 0
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "i"
    GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    GLIFRec(TotalGLIFS).DrAmt = 0
  Case "C"      'NEW Central Depository
    TotalGLIFS = TotalGLIFS + 1
    ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "c"
    
    If TotEIC# > 0 Then
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6) - TotEIC#
    Else
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    End If
    
    GLIFRec(TotalGLIFS).DrAmt = 0
    
    ReplaceString SysRec(1).ICRACCT, "-", ""
    ReDim CDGLIFRec(1 To TotalFunds) As GLIFDataType14
    For cnt = 1 To TotalFunds
      If AllFunds(cnt).Net <> 0 Then
        CDGLIFRec(cnt).TranAcct = QPTrim$(SysRec(1).ICRACCT) + Left$(AllFunds(cnt).FundCode, 2)
        CDGLIFRec(cnt).TranDate = GLIFTDate$
        CDGLIFRec(cnt).TranDesc = "PPE " + GLIFTDate$
        CDGLIFRec(cnt).Source = GLIFSource$
        CDGLIFRec(cnt).FromFlag = "C"
        CDGLIFRec(cnt).DrAmt = AllFunds(cnt).Net                'Passed(6)
        CDGLIFRec(cnt).CrAmt = 0
      End If
    Next
  Case "P"
    
  End Select
  
  '*************END OF COMMON CODE SECTION
  
  If QPTrim$(SysRec(1).EXPMETHD) = "" Or SysRec(1).EXPMETHD = "0" Then
    GoTo WriteGLIFS
  End If
  
  ReDim SFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim MFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim RFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim SocGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim MedGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim RetGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  
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
  
  For cnt = 1 To NumOfWageAccts
    'social
    SocGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + SOCEXP$
    SocGLIFRec(cnt).TranDate = GLIFTDate$
    SocGLIFRec(cnt).TranDesc = SOCMat$
    SocGLIFRec(cnt).Source = GLIFSource$
    SocGLIFRec(cnt).FromFlag = "S"
    SocGLIFRec(cnt).CrAmt = 0
    SocGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATSocAmt
    'medicare
    MedGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + MEDEXP$
    MedGLIFRec(cnt).TranDate = GLIFTDate$
    MedGLIFRec(cnt).TranDesc = MEDMat$
    MedGLIFRec(cnt).Source = GLIFSource$
    MedGLIFRec(cnt).FromFlag = "M"
    MedGLIFRec(cnt).CrAmt = 0
    MedGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATMedAmt
    'retirment
    RetGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + RETEXP$
    RetGLIFRec(cnt).TranDate = GLIFTDate$
    RetGLIFRec(cnt).TranDesc = RETMat$
    RetGLIFRec(cnt).Source = GLIFSource$
    RetGLIFRec(cnt).FromFlag = "R"
    RetGLIFRec(cnt).CrAmt = 0
    RetGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATRetAmt
  Next
  
  For cnt = 1 To TotalFunds
    SFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + SOCLIAB$
    MFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + MEDLIAB$
    RFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + RETLIAB$
  Next
  
  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(SFGLIFRec(Cnt2).TranAcct, 2) = Left$(SocGLIFRec(cnt).TranAcct, 2) Then
        SFGLIFRec(Cnt2).CrAmt = Round(SFGLIFRec(Cnt2).CrAmt + SocGLIFRec(cnt).DrAmt)
        SFGLIFRec(Cnt2).DrAmt = 0
        SFGLIFRec(Cnt2).TranDate = GLIFTDate$
        SFGLIFRec(Cnt2).TranDesc = SOCMatLib$
        SFGLIFRec(Cnt2).Source = GLIFSource$
        SFGLIFRec(Cnt2).FromFlag = "s"
      End If
    Next
  Next
  '
  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(MFGLIFRec(Cnt2).TranAcct, 2) = Left$(MedGLIFRec(cnt).TranAcct, 2) Then
        MFGLIFRec(Cnt2).CrAmt = Round(MFGLIFRec(Cnt2).CrAmt + MedGLIFRec(cnt).DrAmt)
        MFGLIFRec(Cnt2).DrAmt = 0
        MFGLIFRec(Cnt2).TranDate = GLIFTDate$
        MFGLIFRec(Cnt2).TranDesc = MEDMatLib$
        MFGLIFRec(Cnt2).Source = GLIFSource$
        MFGLIFRec(Cnt2).FromFlag = "m"
      End If
    Next
  Next
  
  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(RFGLIFRec(Cnt2).TranAcct, 2) = Left$(RetGLIFRec(cnt).TranAcct, 2) Then
        RFGLIFRec(Cnt2).CrAmt = Round(RFGLIFRec(Cnt2).CrAmt + RetGLIFRec(cnt).DrAmt)
        RFGLIFRec(Cnt2).DrAmt = 0
        RFGLIFRec(Cnt2).TranDate = GLIFTDate$
        RFGLIFRec(Cnt2).TranDesc = RETMatLib$
        RFGLIFRec(Cnt2).Source = GLIFSource$
        RFGLIFRec(Cnt2).FromFlag = "r"
      End If
    Next
  Next
  
WriteGLIFS:
  
  GLIFRecLen = Len(GLIFRec(1))
  FCreate "TempIF.DAT"
  FOpenS "TempIF.DAT", GHandle
  RecNo& = 1
  
  For cnt = 1 To TotalGLIFS
    If Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
      If GLIFRec(cnt).DrAmt > 0 Or GLIFRec(cnt).CrAmt > 0 Then
        FPutRTA GHandle, GLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  
  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(SocGLIFRec(cnt).TranAcct)) Then
      If SocGLIFRec(cnt).DrAmt > 0 Or SocGLIFRec(cnt).CrAmt > 0 Then
        FPutRTA GHandle, SocGLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(MedGLIFRec(cnt).TranAcct)) Then
      If MedGLIFRec(cnt).DrAmt > 0 Or MedGLIFRec(cnt).CrAmt > 0 Then
        FPutRTA GHandle, MedGLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  
  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(RetGLIFRec(cnt).TranAcct)) Then
      If RetGLIFRec(cnt).DrAmt > 0 Or RetGLIFRec(cnt).CrAmt > 0 Then
        FPutRTA GHandle, RetGLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(SFGLIFRec(cnt).TranAcct)) Then
      If SFGLIFRec(cnt).DrAmt > 0 Or SFGLIFRec(cnt).CrAmt > 0 Then
        FPutRTA GHandle, SFGLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(MFGLIFRec(cnt).TranAcct)) Then
      If MFGLIFRec(cnt).DrAmt > 0 Or MFGLIFRec(cnt).CrAmt > 0 Then
        FPutRTA GHandle, MFGLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(RFGLIFRec(cnt).TranAcct)) Then
      If RFGLIFRec(cnt).DrAmt > 0 Or RFGLIFRec(cnt).CrAmt > 0 Then
        FPutRTA GHandle, RFGLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  
  If SysRec(1).USEIMP = "C" Then
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(CDGLIFRec(cnt).TranAcct)) Then
        If CDGLIFRec(cnt).DrAmt <> 0 Or CDGLIFRec(cnt).CrAmt <> 0 Then
          FPutRTA GHandle, CDGLIFRec(cnt), RecNo&, GLIFRecLen
          RecNo& = RecNo& + 1
        End If
      End If
    Next
  End If
  
  'added EIC GLIF records if present 6/07/94
  If TotEIC# > 0 Then
    For cnt = 1 To (SplDedCnt * 2)
      If EICGLIFRec(cnt).DrAmt <> 0 Or EICGLIFRec(cnt).CrAmt <> 0 Then
        FPutRTA GHandle, EICGLIFRec(cnt), RecNo&, GLIFRecLen
        RecNo& = RecNo& + 1
      End If
    Next
  End If
  '
  FClose GHandle
  
  ReDim GLSAmt(1) As String * 11
  ReDim GLMAmt(1) As String * 11

  Erase CFunds$, DFunds$, AllFunds
  Erase SFGLIFRec, MFGLIFRec, RFGLIFRec, CDGLIFRec
  Erase SocGLIFRec, MedGLIFRec, RetGLIFRec, EICGLIFRec

'**** NEW

  If Exist(QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile) Then
    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  Else
    ReDim GLIdxRec(1) As GLAcctIdxType
    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
  End If
  
  GLIdxRecLen = Len(JGLIdxRec(1))
  
  GLIFRecLen = Len(GLIFRec(1))
  
  '  GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile

  TotalGLIFS = FileSize("TempIF.DAT") \ GLIFRecLen
  TotalAccts = FileSize(GLIdxName$) \ GLIdxRecLen
  
  If TotalGLIFS = 0 Then
    PrintGLRpt = False
    GoTo SkipGLRpt
  Else
    PrintGLRpt = True
  End If
  
  ReDim GLIFRec(1 To TotalGLIFS) As GLIFDataType14
  FGetAH "TempIF.DAT", GLIFRec(1), GLIFRecLen, TotalGLIFS
  
'  SortT GLIFRec(1), TotalGLIFS, 0, GLIFRecLen, 17, 20

  Do
    OutOfOrder = False                     'assume it's sorted
    For x = 1 To UBound(GLIFRec) - 1
      If GLIFRec(x).TranAcct > GLIFRec(x + 1).TranAcct Then
        SWAP GLIFRec(x), GLIFRec(x + 1)    'if we had to swap
        OutOfOrder = True                  'we're not done yet
      End If
    Next
  Loop While OutOfOrder

'04-05-96 Started Add summary to glif report

  FPutAH "TempIF.DAT", GLIFRec(1), GLIFRecLen, TotalGLIFS
  
  If TotalAccts = 0 Then GoTo SkipGLAccts
  
  QPrintRC "Reading G/L Accounts.", 11, 31, 112
  
  FOpenS GLIdxName$, GHandle
'heredale
  ReDim JGoodAccts(1 To TotalAccts) As String * 16
  'REDIM JGoodAccts(1 TO TotalAccts)      AS STRING * 16

  For cnt = 1 To TotalAccts
    FGetRTA GHandle, JGLIdxRec(1), CLng(cnt), GLIdxRecLen
    ActualAccts = ActualAccts + 1
    ReplaceString JGLIdxRec(1).AcctNum, "-", ""
    JGoodAccts(ActualAccts) = JGLIdxRec(1).AcctNum
    ShowPctComp cnt, TotalAccts
  Next
  
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

  GoSub GLIFHeader

'**** NEW

  For cnt = 1 To TotalGLIFS
    ShowPctComp cnt, TotalGLIFS
    If Lines >= Max Then
      FPut GHandle, FF$
      GoSub GLIFHeader
    End If

    TotDebit# = Round(TotDebit# + GLIFRec(cnt).DrAmt)
    TotCredit# = Round(TotCredit# + GLIFRec(cnt).CrAmt)

    '10-13-94 ** added check for valid GL Account numbers
    If TotalAccts > 0 Then
      AcctOk = False
      For Cnt2 = 1 To TotalAccts
        If InStr(JGoodAccts(Cnt2), QPTrim$(GLIFRec(cnt).TranAcct)) And Len(QPTrim$(JGoodAccts(Cnt2))) = Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
          AcctOk = True
          Exit For
        End If
      Next
      If Not AcctOk Then
        LSet GLIFRec(cnt).Fill = "ERR     "
        FPutRTA THandle%, GLIFRec(cnt), CLng(cnt), Len(GLIFRec(cnt))
      Else
        LSet GLIFRec(cnt).Fill = ""
      End If
        
    End If

    'NoCheckAccts:
    RSet GLSAmt(1) = FUsing(Str$(GLIFRec(cnt).DrAmt), Image3$)
    RSet GLMAmt(1) = FUsing(Str$(GLIFRec(cnt).CrAmt), Image3$)
    LSet EDesc(1) = QPTrim$(GLIFRec(cnt).TranDesc)
    NoAcctNum = Len(QPTrim$(GLIFRec(cnt).TranAcct))
    If NoAcctNum > 0 Then
      If AcctOk = True Then
        LSet GLIFRec(cnt).Fill = " "
      End If
      Trans$ = QPTrim$(GLIFRec(cnt).TranAcct)
      LSet GLIFRec(cnt).TranAcct = " "
      LSet GLIFRec(cnt).TranAcct = Trans$
      FPut GHandle, "  " + GLIFRec(cnt).TranAcct + GLIFRec(cnt).Fill + EDesc(1) + GLSAmt(1) + "    " + GLMAmt(1) + CrLf$
      Lines = Lines + 1
    End If
  Next
  GoSub GLIFTotals
  
  Fund$ = Left$(QPTrim$(GLIFRec(1).TranAcct), 2)
  FDebit# = GLIFRec(1).DrAmt
  FCredit# = GLIFRec(1).CrAmt
  For cnt = 2 To TotalGLIFS
    NFund$ = Left$(QPTrim$(GLIFRec(cnt).TranAcct), 2)
    If NFund$ <> Fund$ Then
      GoSub PrintFundTotal
      Fund$ = NFund$
      FDebit# = GLIFRec(cnt).DrAmt
      FCredit# = GLIFRec(cnt).CrAmt
    Else
      FDebit# = Round(FDebit# + GLIFRec(cnt).DrAmt)
      FCredit# = Round(FCredit# + GLIFRec(cnt).CrAmt)
    End If
  Next

  GoSub PrintFundTotal

  'GOSUB PrintTaxTotals

  FPut GHandle, FF$
  
  RPTSetupPRN 0, GHandle

  FClose GHandle
  FClose THandle
  
  Erase GLSAmt, GLMAmt, GLIFRec
  Erase GLIdxRec, CDGLIFRec     '
  
SkipGLRpt:
  Erase GLIFRec, SysRec, PDR, AllFunds, DFunds$, CFunds$
  Erase DedCodes, Unit, EDesc, Dash0
  
  Exit Sub
  
PrintFundTotal:
  RSet GLSAmt(1) = FUsing(Str$(FDebit#), Image3$)
  RSet GLMAmt(1) = FUsing(Str$(FCredit#), Image3$)
  LSet EDesc(1) = ""
  LSet GLIFRec(1).Fill = ""
  RSet GLIFRec(1).TranAcct = "Fund  " + Fund$
  FPut GHandle, "   " + GLIFRec(1).TranAcct + GLIFRec(1).Fill + EDesc(1) + GLSAmt(1) + "    " + GLMAmt(1) + CrLf$
Return

GLIFTotals:
  FPut GHandle, CrLf$
  FPut GHandle, Dash0(1) + CrLf$
  FPut GHandle, "   Totals:                                  " + FUsing(Str$(TotDebit#), Image3$) + "     " + FUsing(Str$(TotCredit#), Image3$) + CrLf$
  If TotalAccts = 0 Then
    FPut GHandle, CrLf$
    FPut GHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory." + CrLf$
  End If
  FPut GHandle, CrLf$
  FPut GHandle, "SubTotals:" + CrLf$
Return
  
GLIFHeader:
  FPut GHandle, QPTrim$(Unit(1).UFEMPR) + CrLf$
  FPut GHandle, "G/L Interface Report." + CrLf$
  FPut GHandle, "Period Ending: " + Num2Date$(PDR(1).PEREND) + CrLf$
  FPut GHandle, CrLf$
  FPut GHandle, "Account No.           Description                Debit         Credit" + CrLf$
  FPut GHandle, Dash0(1) + CrLf$
  Lines = 6
Return


'PrintTaxTotals:
'RETURN

End Sub

