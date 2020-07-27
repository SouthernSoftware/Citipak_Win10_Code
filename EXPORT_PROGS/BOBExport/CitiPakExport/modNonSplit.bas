Attribute VB_Name = "modNonSplit"
Option Explicit
  Dim GLCreditTotal As Double
  Dim GLDebitTotal As Double
  Dim GLError As Integer

Sub MakeGLIFFileG(TotEIC#, TotDeds#(), Passed#(), DistbSumAccts() As DistWageRptType)

  ReDim SysRec(1) As RegDSysFileRecType
  ReDim PDR(1) As PeriodDefaultRecType
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  ReDim GLSetupRec(1) As GLSetupRecType
  Dim DedCodeFileName As Integer
  Dim PPDefaultFileName As Integer
  Dim SysFileName As Integer
  Dim x As Integer
  Dim FACnt As Integer
  Dim FundPad As Integer
  Dim FundLen As Integer
  Dim GLSetUpName$, GHandle As Integer
  Dim GLSetUpRecLen As Integer
  Dim NumOfWageAccts As Integer
  Dim GLIFTDate$, GLIFSource$
  Dim cnt As Integer, NextAcct As Integer
  Dim SysCash$, NumDFunds As Integer
  Dim CurrFund$, ThisFund As Integer
  Dim FirstC As Integer, TotalFunds As Integer
  Dim LastC As Integer, NoCFunds As Boolean
  Dim NumCFunds As Integer, First As Integer
  Dim Start As Integer, Last As Integer
  Dim Cnt2 As Integer, TotalGLIFS
  Dim AcctNum$, TempAcct$
  Dim FringeAcct$, FringeRate#, RecNo&
  Dim INDFund$, IndirectAcct$, IndirectRate#
  Dim Indirect#, Fringe#, SOCEXP$, RETLIAB$
  Dim MEDEXP$, RETEXP$, SOCLIAB$, MEDLIAB$
  Dim GLIFRecName$, FundCash$
  Dim GLIFRecLen As Integer, GLHandle As Integer
  Dim NumOfDeds As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenDedCodeFile DedCodeFileName
  
  For x = 1 To 50
    Get DedCodeFileName, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
      NumOfDeds = NumOfDeds + 1
    End If
  Next x
  Close DedCodeFileName
  
  OpenPPDefaultFile PPDefaultFileName
  Get PPDefaultFileName, 1, PDR(1)
  Close PPDefaultFileName
  
  OpenSysFile SysFileName
  Get SysFileName, 1, SysRec(1)
  Close SysFileName
  
  FACnt = SysRec(1).AcctCnt

'for new gl
  
  FundPad = 0
  FundLen = 2     'Default fund length
  
'  GLSetUpName$ = QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT"
  If Mid(CurrCitiPath, Len(CurrCitiPath), 1) <> "\" Then
    GLSetUpName = CurrCitiPath + "\GLSETUP.DAT"
  ElseIf Mid(CurrCitiPath, Len(CurrCitiPath), 1) = "\" Then
    GLSetUpName = CurrCitiPath + "GLSETUP.DAT"
  End If
  
  GLSetUpRecLen = Len(GLSetupRec(1))
  GLHandle = FreeFile
  
  If Exist(GLSetUpName$) Then
    Open GLSetUpName$ For Random Shared As GLHandle Len = GLSetUpRecLen
    Get GLHandle, 1, GLSetupRec(1)
    FundLen = GLSetupRec(1).FundLen
    FundPad = GLSetupRec(1).DetLen - GLSetupRec(1).FundLen
  End If
  Close GLHandle

  NumOfWageAccts = UBound(DistbSumAccts)
  If NumOfWageAccts = 0 Then
    MsgBox "No active transactions are pending at this time"
    Exit Sub
  End If
  ReDim Preserve DistbSumAccts(1 To NumOfWageAccts) As DistWageRptType
  'squeeze out all the "-" out of Acct numbers
  For cnt = 1 To NumOfWageAccts
    ReplaceString DistbSumAccts(cnt).Acct, "-", ""
  Next
  
  'change Period ending date to nicks format
  
  GLIFTDate$ = MakeRegDate(PDR(1).PEREND)
  
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
  ReplaceString GLIFTDate$, "2004", "04"
  ReplaceString GLIFTDate$, "2005", "05"
  ReplaceString GLIFTDate$, "2006", "06"
  ReplaceString GLIFTDate$, "2007", "07"
  ReplaceString GLIFTDate$, "2008", "08"
  ReplaceString GLIFTDate$, "2009", "09"
  '
  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$
  
  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
  ReDim GLIFRec(1 To (NumOfWageAccts + 5 + NumOfDeds + 1)) As GLIFDataType14
'  ReDim GLIFRec(1 To NumOfWageAccts + 38) As GLIFDataType14
  'changed Redim from (1 To NumOfWageAccts + 38) on 10/7/02
  'NumOfWageAccts = first for loop
  '5 = second for loop
  'NumOfDeds = third for loop
  For cnt = 1 To NumOfWageAccts  'first for loop
    GLIFRec(cnt).TranAcct = QPTrim$(DistbSumAccts(cnt).Acct)
    GLIFRec(cnt).TranDate = GLIFTDate$
    GLIFRec(cnt).TranDesc = "Wages"
    GLIFRec(cnt).CrAmt = 0
    GLIFRec(cnt).DrAmt = DistbSumAccts(cnt).GrossPay
    GLIFRec(cnt).Source = GLIFSource$
    GLIFRec(cnt).FromFlag = "W"
  Next
  
  NextAcct = cnt ' - 1
  '
  For cnt = 0 To 4 'second for loop
    ReplaceString SysRec(1).Liab(cnt + 1).Acct, "-", ""
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SysRec(1).Liab(cnt + 1).Acct)
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = "Withholdings"
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
    'get tax and ret account numbers
  Next
  GLIFRec(NextAcct).CrAmt = Passed#(1)          'federal
  GLIFRec(NextAcct).TranDesc = "Fed Withholdings"
  GLIFRec(NextAcct + 1).CrAmt = Passed#(2)      'state
  GLIFRec(NextAcct + 1).TranDesc = "State Withholdings"
  GLIFRec(NextAcct + 2).CrAmt = Passed#(3)      'social sec
  GLIFRec(NextAcct + 2).TranDesc = "Soc Sec Withholdings"
  GLIFRec(NextAcct + 3).CrAmt = Passed#(4)      'Medicare
  GLIFRec(NextAcct + 3).TranDesc = "Med Withholdings"
  GLIFRec(NextAcct + 4).CrAmt = Passed#(5)      'Retirement total
  GLIFRec(NextAcct + 4).TranDesc = "Ret Withholdings"
  'good to here ;maybe
  
  ReplaceString SysRec(1).CashAcct, "-", ""
  
  SysCash$ = QPTrim$(SysRec(1).CashAcct)
  
  If TotEIC# > 0 Then
    ReDim EICGLIFRec(1 To 2) As GLIFDataType14
    EICGLIFRec(1).TranAcct = QPTrim$(SysRec(1).Liab(1).Acct)
    EICGLIFRec(1).TranDate = GLIFTDate$
    EICGLIFRec(1).TranDesc = "EIC Pmt"
    EICGLIFRec(1).CrAmt = 0
    EICGLIFRec(1).DrAmt = TotEIC#
    EICGLIFRec(1).Source = GLIFSource$
    EICGLIFRec(1).FromFlag = "E"
    '
    EICGLIFRec(2).TranAcct = Left$(QPTrim$(SysRec(1).Liab(1).Acct), FundLen) + SysCash$
    EICGLIFRec(2).TranDate = GLIFTDate$
    EICGLIFRec(2).TranDesc = "EIC Pmt"
    EICGLIFRec(2).CrAmt = TotEIC#
    EICGLIFRec(2).DrAmt = 0
    EICGLIFRec(2).Source = GLIFSource$
    EICGLIFRec(2).FromFlag = "P"
  End If

  NextAcct = NextAcct + cnt '+ (NumOfDeds - 1) 'problem
  
  For cnt = 0 To NumOfDeds - 1 'third for loop
    ReplaceString DedCodes(cnt + 1).DCACCT1, "-", ""
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(DedCodes(cnt + 1).DCACCT1)
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
'    If QPTrim$(DedCodes(cnt + 1).DCDESC1) = "MISC23" Then Stop
    GLIFRec(NextAcct + cnt).TranDesc = QPTrim$(DedCodes(cnt + 1).DCDESC1) '"Deductions"
    GLIFRec(NextAcct + cnt).CrAmt = TotDeds#(cnt + 1)
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "D"
  Next
  
  
  ReDim DFunds$(1 To NumOfWageAccts)
  NumDFunds = 1
  'fixed
  DFunds$(1) = Left$(DistbSumAccts(1).Acct, FundLen)
  For cnt = 1 To NumOfWageAccts - 1
    'fixed
    If Left$(DistbSumAccts(cnt).Acct, FundLen) <> Left$(DistbSumAccts(cnt + 1).Acct, FundLen) Then
      NumDFunds = NumDFunds + 1 'counting the total number of funds.
      'fixed
      DFunds$(NumDFunds) = Left$(DistbSumAccts(cnt + 1).Acct, FundLen)
    End If
  Next
  
  ReDim Preserve DFunds$(1 To NumDFunds)
  ReDim DFund(1 To NumDFunds) As FundType
  For cnt = 1 To NumOfWageAccts
    'fixed
    CurrFund$ = Left$(DistbSumAccts(cnt).Acct, FundLen)
    For ThisFund = 1 To NumDFunds
      If CurrFund$ = DFunds$(ThisFund) Then
        DFund(ThisFund).FundCode = DFunds$(ThisFund)
        DFund(ThisFund).Debit = OldRound(DFund(ThisFund).Debit + DistbSumAccts(cnt).GrossPay)
        Exit For
      End If
    Next
  Next
  
  'all gross pay by funds here!!
  'make funds and sumarize ded and taxs here
  '
'  ReDim CFunds$(1 To 17)
  ReDim CFunds$(1 To NumOfWageAccts + 5 + NumOfDeds)
  
  '
  FirstC = NumOfWageAccts + 1
'  LastC = NumOfWageAccts + 17 + 1
  LastC = NumOfWageAccts + 5 + NumOfDeds + 1 'NumOfWageAccts +  17 + 1
  
  NumCFunds = 1
  'fixed
  CFunds$(1) = Left$(GLIFRec(FirstC).TranAcct, FundLen)
  
  For cnt = FirstC To LastC - 1
    'fixed
    If Left$(GLIFRec(cnt).TranAcct, FundLen) <> Left$(GLIFRec(cnt + 1).TranAcct, FundLen) Then
      If Len(QPTrim$(GLIFRec(cnt + 1).TranAcct)) Then
        NumCFunds = NumCFunds + 1               'counting the total number of funds.
        'fixed
        CFunds$(NumCFunds) = Left$(GLIFRec(cnt + 1).TranAcct, FundLen)
      End If
    End If
  Next

  ReDim Preserve CFunds$(1 To NumCFunds)
  ReDim CFund(1 To NumCFunds) As FundType
  For cnt = FirstC To LastC - 1
    'fixed
    CurrFund$ = Left$(GLIFRec(cnt).TranAcct, FundLen)
    For ThisFund = 1 To NumCFunds
'      If cnt = 65 Then Stop
      If CurrFund$ = CFunds$(ThisFund) Then
        CFund(ThisFund).FundCode = CFunds$(ThisFund)
        CFund(ThisFund).Credit = OldRound(CFund(ThisFund).Credit + GLIFRec(cnt).CrAmt)
        Exit For
      End If
    Next
  Next

  'combine all funds in one array here
  TotalFunds = NumDFunds + NumCFunds            '+ 1
  ReDim AllFunds(1 To TotalFunds) As FundType
  ThisFund = 1
  For cnt = 1 To NumDFunds
    AllFunds(ThisFund) = DFund(cnt)
    ThisFund = ThisFund + 1
  Next
  '
  If NoCFunds = False Then
    For cnt = 1 To NumCFunds
      AllFunds(ThisFund) = CFund(cnt)
      ThisFund = ThisFund + 1
    Next
  End If

'fixed: 05-27-96
  SortT AllFunds(), TotalFunds
                                                        
  'combine Debits and Credits for same fund numbers
  First = 1
  Start = 1
  Last = TotalFunds
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If AllFunds(First).FundCode = AllFunds(Cnt2).FundCode Then
        LSet AllFunds(Cnt2).FundCode = ""       'blank acct num as a flag
        AllFunds(First).Credit = OldRound(AllFunds(First).Credit + AllFunds(Cnt2).Credit)
        AllFunds(Cnt2).Credit = 0
        AllFunds(First).Debit = OldRound(AllFunds(First).Debit + AllFunds(Cnt2).Debit)
        AllFunds(Cnt2).Debit = 0
      End If
    Next
    First = First + 1
  Loop Until First >= Last      '

  'calc net difference for funds
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(AllFunds(cnt).FundCode)) Then
      AllFunds(cnt).Net = OldRound(AllFunds(cnt).Debit - AllFunds(cnt).Credit)
    End If
  Next

  'add cash acct number to fund numbers
  For cnt = 1 To TotalFunds
    FundCash$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(FundCash$) Then
      LSet AllFunds(cnt).FundCode = FundCash$ + SysCash$
    End If
  Next

'  TotalGLIFS = NumOfWageAccts + 39 + TotalFunds
  TotalGLIFS = NumOfWageAccts + TotalFunds + NumOfDeds + 6 'added 8/12/04 to resolve
  'problem of not all deductions being printed  because this array as too small

  ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14

'  NextAcct = NumOfWageAccts + 39
  NextAcct = NumOfWageAccts + NumOfDeds + 6 'added 8/12/04 to resolve problem of this
  'array being too small to hold all deductions up to 50

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
    '
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
        'fixed
        CDGLIFRec(cnt).TranAcct = QPTrim$(SysRec(1).ICRACCT) + Left$(AllFunds(cnt).FundCode, FundLen)
        If FundPad > 0 Then
          TempAcct$ = QPTrim$(CDGLIFRec(cnt).TranAcct)
          TempAcct$ = TempAcct$ + String$(FundPad, "0")
          CDGLIFRec(cnt).TranAcct = TempAcct$
        End If
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
  
  If SysRec(1).EXPMETHD = "2" Then GoTo Type2Meth 'jump to type 2 here
  
  'calc and add Fringe GLIF recs
  ReDim FGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  
  ReplaceString SysRec(1).FRNGEXP, "-", ""
  FringeAcct$ = QPTrim$(SysRec(1).FRNGEXP)
  FringeRate# = SysRec(1).FRNGRATE

  For cnt = 1 To NumOfWageAccts
    AcctNum$ = Left$(GLIFRec(cnt).TranAcct, FACnt)
    FGLIFRec(cnt).TranAcct = AcctNum$ + FringeAcct$
    FGLIFRec(cnt).DrAmt = OldRound(GLIFRec(cnt).DrAmt * (FringeRate# * 0.01))
    FGLIFRec(cnt).TranDate = GLIFTDate$
    FGLIFRec(cnt).TranDesc = "FRINGE " + GLIFTDate$
    FGLIFRec(cnt).Source = GLIFSource$
    FGLIFRec(cnt).FromFlag = "F"
  Next

  'calc and add Indirect GLIF recs
  ReDim IGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReplaceString SysRec(1).INDDR, "-", ""
  'fixed
  INDFund$ = QPTrim$(Left$(SysRec(1).INDDR, FundLen))
  IndirectAcct$ = QPTrim$(SysRec(1).INDEXP)
  IndirectRate# = SysRec(1).INDRATE
  
  If IndirectRate# < 0 Then IndirectRate# = 0
  
  For cnt = 1 To NumOfWageAccts
    'look for acct that don't get indirect
    'fixed
    If Not QPTrim$(Left$(GLIFRec(cnt).TranAcct, FundLen)) = INDFund$ Then
      AcctNum$ = Left$(GLIFRec(cnt).TranAcct, FACnt)
      IGLIFRec(cnt).TranAcct = AcctNum$ + IndirectAcct$
      IGLIFRec(cnt).DrAmt = OldRound((GLIFRec(cnt).DrAmt + FGLIFRec(cnt).DrAmt) * (IndirectRate#) * 0.01)
      IGLIFRec(cnt).TranDate = GLIFTDate$
      IGLIFRec(cnt).TranDesc = "INDIRECT " + GLIFTDate$
      IGLIFRec(cnt).Source = GLIFSource$
      IGLIFRec(cnt).FromFlag = "I"
    End If
  Next

  ReDim IFGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  For cnt = 1 To NumOfWageAccts
    'fixed
    AcctNum$ = Left$(FGLIFRec(cnt).TranAcct, FundLen)
    IFGLIFRec(cnt).TranAcct = AcctNum$ + SysCash$
    IFGLIFRec(cnt).CrAmt = OldRound(IGLIFRec(cnt).DrAmt + FGLIFRec(cnt).DrAmt)
    IFGLIFRec(cnt).TranDate = GLIFTDate$
    IFGLIFRec(cnt).TranDesc = "F&I COST " + GLIFTDate$
    IFGLIFRec(cnt).Source = GLIFSource$
    IFGLIFRec(cnt).FromFlag = "A"
  Next

  For cnt = 1 To NumOfWageAccts
    Indirect# = OldRound(Indirect# + IGLIFRec(cnt).DrAmt)
    Fringe# = OldRound(Fringe# + FGLIFRec(cnt).DrAmt)
  Next

  ReDim AGLIFRec(1 To 4) As GLIFDataType14

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

  For cnt = 1 To 4
    AGLIFRec(cnt).TranDate = GLIFTDate$
    AGLIFRec(cnt).TranDesc = "PPE " + GLIFTDate$
    AGLIFRec(cnt).Source = GLIFSource$
    AGLIFRec(cnt).FromFlag = "T"
  Next
  GoTo WriteGLIFS
  '**********END TYPE 1 SECTION

Type2Meth:

  ReDim SocGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim MedGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim RetGLIFRec(1 To NumOfWageAccts) As GLIFDataType14

  SOCEXP$ = QPTrim$(SysRec(1).SOCEXP)
  MEDEXP$ = QPTrim$(SysRec(1).MEDEXP)
  RETEXP$ = QPTrim$(SysRec(1).RETEXP)

  SOCLIAB$ = QPTrim$(SysRec(1).SOCLIAB)
  MEDLIAB$ = QPTrim$(SysRec(1).MEDLIAB)
  RETLIAB$ = QPTrim$(SysRec(1).RETLIAB)

  ReplaceString SOCLIAB$, "-", ""
  ReplaceString MEDLIAB$, "-", ""
  ReplaceString RETLIAB$, "-", ""

  For cnt = 1 To NumOfWageAccts
    'social
    SocGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + SOCEXP$
    SocGLIFRec(cnt).TranDate = GLIFTDate$
    SocGLIFRec(cnt).TranDesc = "Soc Match"
    SocGLIFRec(cnt).Source = GLIFSource$
    SocGLIFRec(cnt).FromFlag = "S"
    SocGLIFRec(cnt).CrAmt = 0
    SocGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATSocAmt
    'medicare
    MedGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + MEDEXP$
    MedGLIFRec(cnt).TranDate = GLIFTDate$
    MedGLIFRec(cnt).TranDesc = "Med Match"
    MedGLIFRec(cnt).Source = GLIFSource$
    MedGLIFRec(cnt).FromFlag = "M"
    MedGLIFRec(cnt).CrAmt = 0
    MedGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATMedAmt
    'retirment
    RetGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + RETEXP$
    RetGLIFRec(cnt).TranDate = GLIFTDate$
    RetGLIFRec(cnt).TranDesc = "Ret Match"
    RetGLIFRec(cnt).Source = GLIFSource$
    RetGLIFRec(cnt).FromFlag = "R"
    RetGLIFRec(cnt).CrAmt = 0
    RetGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATRetAmt
  Next

  ReDim SFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim MFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim RFGLIFRec(1 To TotalFunds) As GLIFDataType14

  For cnt = 1 To TotalFunds
    'fixed
    SFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, FundLen) + SOCLIAB$
    MFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, FundLen) + MEDLIAB$
    RFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, FundLen) + RETLIAB$
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      'fixed
      If Left$(SFGLIFRec(Cnt2).TranAcct, FundLen) = Left$(SocGLIFRec(cnt).TranAcct, FundLen) Then
        SFGLIFRec(Cnt2).CrAmt = OldRound(SFGLIFRec(Cnt2).CrAmt + SocGLIFRec(cnt).DrAmt)
        SFGLIFRec(Cnt2).DrAmt = 0
        SFGLIFRec(Cnt2).TranDate = GLIFTDate$
        SFGLIFRec(Cnt2).TranDesc = "Soc Match Liab"
        SFGLIFRec(Cnt2).Source = GLIFSource$
        SFGLIFRec(Cnt2).FromFlag = "s"
      End If
    Next
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      'fixed
      If Left$(MFGLIFRec(Cnt2).TranAcct, FundLen) = Left$(MedGLIFRec(cnt).TranAcct, FundLen) Then
        MFGLIFRec(Cnt2).CrAmt = OldRound(MFGLIFRec(Cnt2).CrAmt + MedGLIFRec(cnt).DrAmt)
        MFGLIFRec(Cnt2).DrAmt = 0
        MFGLIFRec(Cnt2).TranDate = GLIFTDate$
        MFGLIFRec(Cnt2).TranDesc = "Med Match Liab"
        MFGLIFRec(Cnt2).Source = GLIFSource$
        MFGLIFRec(Cnt2).FromFlag = "m"
      End If
    Next
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      'fixed
      If Left$(RFGLIFRec(Cnt2).TranAcct, FundLen) = Left$(RetGLIFRec(cnt).TranAcct, FundLen) Then
        RFGLIFRec(Cnt2).CrAmt = OldRound(RFGLIFRec(Cnt2).CrAmt + RetGLIFRec(cnt).DrAmt)
        RFGLIFRec(Cnt2).DrAmt = 0
        RFGLIFRec(Cnt2).TranDate = GLIFTDate$
        RFGLIFRec(Cnt2).TranDesc = "Ret Match Liab"
        RFGLIFRec(Cnt2).Source = GLIFSource$
        RFGLIFRec(Cnt2).FromFlag = "r"
      End If
    Next
  Next

WriteGLIFS:
  
  GLIFRecLen = Len(GLIFRec(1))
  GLIFRecName$ = "TempIF.DAT"
  KillFile "TempIF.DAT"
  GHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GHandle Len = GLIFRecLen
  RecNo& = 1
  
  For cnt = 1 To TotalGLIFS
    If Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
      If GLIFRec(cnt).DrAmt > 0 Or GLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, GLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  
  Select Case SysRec(1).EXPMETHD
  Case "1"
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(FGLIFRec(cnt).TranAcct)) Then
        If FGLIFRec(cnt).DrAmt > 0 Or FGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, FGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(IGLIFRec(cnt).TranAcct)) Then
        If IGLIFRec(cnt).DrAmt > 0 Or IGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, IGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(IFGLIFRec(cnt).TranAcct)) Then
        If IFGLIFRec(cnt).DrAmt > 0 Or IFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, IFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    
    For cnt = 1 To 4
      If Len(QPTrim$(AGLIFRec(cnt).TranAcct)) > 0 Then
        Put GHandle, RecNo&, AGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    Next
    
  Case "2"
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(SocGLIFRec(cnt).TranAcct)) Then
        If SocGLIFRec(cnt).DrAmt > 0 Or SocGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, SocGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(MedGLIFRec(cnt).TranAcct)) Then
        If MedGLIFRec(cnt).DrAmt > 0 Or MedGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, MedGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(RetGLIFRec(cnt).TranAcct)) Then
        If RetGLIFRec(cnt).DrAmt > 0 Or RetGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, RetGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next

    For cnt = 1 To TotalFunds
      If Len(QPTrim$(SFGLIFRec(cnt).TranAcct)) Then
        If SFGLIFRec(cnt).DrAmt > 0 Or SFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, SFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(MFGLIFRec(cnt).TranAcct)) Then
        If MFGLIFRec(cnt).DrAmt > 0 Or MFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, MFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(RFGLIFRec(cnt).TranAcct)) Then
        If RFGLIFRec(cnt).DrAmt > 0 Or RFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, RFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
  End Select
  
  If SysRec(1).USEIMP = "C" Then
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(CDGLIFRec(cnt).TranAcct)) Then
        If CDGLIFRec(cnt).DrAmt <> 0 Or CDGLIFRec(cnt).CrAmt <> 0 Then
          Put GHandle, RecNo&, CDGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
  End If

  'added EIC GLIF records if present 6/07/94
  If TotEIC# > 0 Then
    For cnt = 1 To 2
      Put GHandle, RecNo&, EICGLIFRec(cnt)
      RecNo& = RecNo& + 1
    Next
  End If

  Close GHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "modNonSplit", "MakeGLIFFileG", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
End Sub

Sub PCPrintPayRegisterG(PathCode As Integer)
  
  Dim RptTitle$, PPDefaultFileName As Integer
  Dim FileHandle As Integer, x As Integer
  Dim DedCodeFileName As Integer
  Dim ErnCodeFileName As Integer
  Dim SysFileName As Integer, ASAmt#
  Dim FundPad As Integer, TOTPaid#
  Dim FundLen As Integer, TOTComp#
  Dim GLSetUpName$, GHandle As Integer
  Dim GLSetUpRecLen As Integer
  Dim GFedGross#, GStaGross#, GMedGross#
  Dim GSocGross#, GRetGross#, GTaxFrn#
  Dim TotDebit#, TotCredit#, EmpActLen
  Dim DistbSumAcctsLen As Integer, ARAmt#
  Dim TransRecLen As Integer, IdxRecLen As Integer
  Dim Emp1RecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Integer, SalCnt As Integer
  Dim HrlCnt As Integer, DLineCnt As Integer
  Dim LineCnt As Integer, NumOfWageAccts As Integer
  Dim MaxLines As Integer, DMaxLines As Integer
  Dim EPage As Integer, Page As Integer
  Dim EmpIdxNNameHandle As Integer, RptName$
  Dim DTitle$(1 To 5), cnt As Integer, TDed$, LastDed As Integer
  Dim ETitle$, SumHeader2$, RHandle As Integer
  Dim DHandle As Integer, NHandle As Integer
  Dim THandle As Integer, DistributionRptName$
  Dim FF$, TErn$, PayRegisterRptName$
  Dim JFlag As Boolean, TotalGLIFS As Integer
  Dim TotalAccts As Integer, PrintGLRpt As Boolean
  Dim GLIFRecLen As Integer, GLIFRecName$
  Dim GRHandle As Integer, GLHandle As Integer
  Dim ActualAccts As Integer, Max As Integer
  Dim Lines As Integer, GLIdxName$, AMAmt#
  Dim Cnt2 As Integer, AcctOk As Boolean, GLAcct@
  Dim NoAcctNum As Integer, Fund$, FDebit#, FCredit#
  Dim NFund$, RetCode As Integer, LincCnt As Integer
  Dim RegHrs#, VACHRS#, SICKHRS#, HOLHRS#, COMPHRS#, PerHours#
  Dim TotalHrs#, TotEIC#, TRegWage#, TOTWage#, GPay#
  Dim SSTax#, MTax#, FTax#, STax#, RETTOT#
  Dim TNetPay#, Emp1Handle As Integer
  Dim SumDed$(1 To 5), GLIdxRecLen As Integer
  Dim SumErn$, LastErn As Integer
  Dim ENumOfAct As Integer, Acct$, First As Integer
  Dim Last As Integer, Start As Integer
  Dim TotalSocAmt#, DistDif#, TotalMedAmt#, LastActive As Integer
  Dim TotalRetAmt#, DLincCnt As Integer, Cnt3 As Integer
  Dim TotHrs#, RegWage#, OTWage#, PrnDef$
  Dim AddEarn#, DGPay#, OutOfOrder As Boolean
  Dim Image0$, Image$, Image3$, Image4$, Image5$
  Dim foundIt As Boolean
  Dim NumOfDeds As Integer, Nextx As Integer
  Dim tripCnt As Integer
  Dim RptHandle#, dlm$, TFHandle As Integer, FundTotals$
  Dim PHandle As Integer, PayRegTotal$
  Dim DedDesc$(1 To 50)
  Dim DTHandle As Integer
  Dim DistAcctTotalName$
  Dim AcctCnt As Integer, TActEmps As Integer
  Dim DedDescTrimmed(1 To 50) As String * 8
'  Dim TempMedTaxAmt#
  '-------------Temp Void variables------------
  Dim CSocWHAcct$
  Dim CMedWHAcct$
  Dim CSocMatchAcct$
  Dim CMedMatchAcct$
  Dim CRetMatchAcct$
  Dim CFedWHAcct$
  Dim CStaWHAcct$
  Dim CRetWHAcct$
  Dim CDedAcct$
  Dim CPRNetAcct$
  Dim DWagesAcct$
  Dim DSocMatchAcct$
  Dim DMedMatchAcct$
  Dim DRetMatchAcct$
  Dim FundNumOnly$
  Dim FundAndAcctOnly$
  Dim TempVoid As VoidCheckType
  Dim TVHandle As Integer
  Dim TVCnt As Double
  Dim AcctLen As Integer
  Dim DetLen As Integer
  Dim ThisPR As Double
  Dim ThisFTax#
  Dim ThisMTax#
  Dim ThisSSTax#
  Dim ThisStaTax#
  Dim ThisRTax#
  Dim FACnt As Integer
  Dim ThisFACnt As Integer
  Dim TotalDeds#
  Dim DbtCnt As Integer
  Dim ThisCRGL$
  '---------------^^^^-------------------------
  Dim ActiveCnt As Integer
  Dim ThisDesc$
  Dim Thisx As Integer
  Dim z As Integer
  Dim PoolFundNum$ '9/17/04
  Dim PRNetPoolFound As Boolean
  Dim TotalWHAndDeds As Double
  Dim TotalWages As Double
  Dim AcctNumCnt As Integer
  Dim ThisEmpCnt As Integer
  Dim PRNetSum As Double
  Dim GL4PPETotal$
  Dim PRPoolProcessed As Boolean
  
  PRPoolProcessed = False
  GLDebitTotal = 0
  GLCreditTotal = 0
  GLError = 0
  
  ActiveCnt = 0
  dlm$ = "~"
  RptTitle$ = "Register & G/L Interface Reports"
  FrmShowPctComp.Label1 = RptTitle$
  FrmShowPctComp.Show
  ReDim TransRec(1) As TransRecType
  ReDim EmpRec1(1) As EmpData1Type
  ReDim PDR(1) As PeriodDefaultRecType
  ReDim Unit(1) As UnitFileRecType
  
  ReDim DistbSumAccts(1 To 1) As DistWageRptType
  
  ReDim SysRec(1) As RegDSysFileRecType
  ReDim GLIFRec(1 To 1) As GLIFDataType14
  
  ReDim EmpAct(1) As DistWageRptType
  
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType
  
  ReDim TotDeds#(1 To 50)
  ReDim TotErns(1 To 3) As Double
  
  ReDim EDesc(1) As String * 21
  ReDim EDAct(1) As String * 14
  ReDim EDPct(1) As String * 11
  ReDim EDRHrs(1) As String * 11
  ReDim EDOHrs(1) As String * 11
  ReDim EDRPay(1) As String * 11
  ReDim EDOPay(1) As String * 11
  ReDim EDEarn(1) As String * 11
  ReDim EDGroP(1) As String * 11
  
  ReDim EDSAmt(1) As String * 11
  ReDim EDMAmt(1) As String * 11
  ReDim EDRAmt(1) As String * 11
  
  ReDim ENumb(1) As String * 14
  ReDim EName(1) As String * 33
  
  ReDim BRat(1) As String * 11
  ReDim ORat(1) As String * 11
  
  ReDim TaxFrn(1) As String * 11
  ReDim Fill11(1) As String * 11
  
  ReDim SCnt(1) As String * 11
  ReDim HCnt(1) As String * 11
  
  ReDim RHrs(1) As String * 11
  ReDim VHrs(1) As String * 11
  ReDim SHrs(1) As String * 11
  ReDim HHrs(1) As String * 11
  ReDim CHrs(1) As String * 11
  ReDim THrs(1) As String * 11

  ReDim PHrs(1) As String * 11

  ReDim OTHrs(1) As String * 11
  ReDim OTPaid(1) As String * 11
  ReDim OTComp(1) As String * 11
  
  ReDim RErnP(1) As String * 11
  ReDim OErnP(1) As String * 11
  
  ReDim GPayP(1) As String * 11
  ReDim SSTaxP(1) As String * 11
  ReDim MTaxP(1) As String * 11
  ReDim FTaxP(1) As String * 11
  ReDim STaxP(1) As String * 11
  ReDim RetirP(1) As String * 11
  ReDim NetPayP(1) As String * 11
  ReDim Ded(1) As String * 11
  
  'added for EIC   6/07/94
  ReDim EEicP(1) As String * 11
  
  ReDim Ern(1) As String * 11
  
  Dim ThisFund As Integer '8/10/04
  Dim FundCount As Integer '8/10/04
  Dim y As Integer '8/10/04
  
  Dim TOTFEDTAX As Double '8/13/04
  Dim TOTMEDTAX As Double '8/13/04
  Dim TOTSOCTAX As Double '8/13/04
  Dim TOTSTATAX As Double '8/13/04
  Dim TOTRetTax As Double '8/13/04
  Dim TOTMEDMat As Double '8/13/04
  Dim TOTSOCMat As Double '8/13/04
  Dim TOTRETMat As Double '8/13/04
  ReDim TotDedAmt(1 To 50) As Double '8/13/04
  
  OpenPPDefaultFile PPDefaultFileName
  Get PPDefaultFileName, 1, PDR(1)
  Close PPDefaultFileName
  
  OpenSysFile SysFileName
  Get SysFileName, 1, SysRec(1)
  Close SysFileName
  PoolFundNum = Mid(SysRec(1).Liab(1).Acct, 1, 2)
'  Call GetAcctStruct(SysRec(1).CITIDIR, FundLen, AcctLen, DetLen)
  Call GetAcctStruct(CurrCitiPath, FundLen, AcctLen, DetLen)
  FACnt = FundLen + AcctLen
  If DetLen > FundLen Then
    FundPad = DetLen - FundLen
  Else
    FundPad = 0
  End If
  
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle
  
  OpenDedCodeFile DedCodeFileName
  ReDim DedCodeNums(1 To 50) As String '6/22/04
  ReDim DedCodeDesc(1 To 50) As String
  For x = 1 To 50
    Get DedCodeFileName, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      NumOfDeds = NumOfDeds + 1
      DedCodes(x) = DedRec
      DedCodeNums(x) = QPTrim$(DedRec.DCACCT1)
      DedCodeDesc(x) = QPTrim$(DedRec.DCDESC1)
    End If
  Next x
  Close DedCodeFileName
  
  OpenErnCodeFile ErnCodeFileName
  For x = 1 To 3
    Get ErnCodeFileName, x, ErnCodes(x)
  Next x
  Close ErnCodeFileName
  
  ReDim GLSetupRec(1) As GLSetupRecType
'for new gl
'  FundPad = 0
'  FundLen = 2     'Default fund length
  
'  If Exist(QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT") Then
'    foundIt = True
  If Exist(CurrCitiPath + "\GLSETUP.DAT") Then
    foundIt = True
  ElseIf Exist(CurrCitiPath + "GLSETUP.DAT") Then
    foundIt = True
  Else '7/26
    Unload FrmShowPctComp '7/26
    MsgBox "Path to GLSETUP.DAT cannot be found." '7/26
    GoTo SkipGLReport '7/26
  End If '7/26
  
'  GLSetUpName$ = QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT"
  If Mid(CurrCitiPath, Len(CurrCitiPath), 1) <> "\" Then
    GLSetUpName$ = CurrCitiPath + "\GLSETUP.DAT"
  ElseIf Mid(CurrCitiPath, Len(CurrCitiPath), 1) = "\" Then
    GLSetUpName$ = CurrCitiPath + "GLSETUP.DAT"
  End If
  
  GLSetUpRecLen = Len(GLSetupRec(1))
  GLHandle = FreeFile
  Open GLSetUpName$ For Random Shared As GLHandle Len = GLSetUpRecLen
  
  If foundIt = True Then
    Get GLHandle, 1, GLSetupRec(1)
    FundLen = GLSetupRec(1).FundLen
    FundPad = GLSetupRec(1).DetLen - GLSetupRec(1).FundLen
  End If
  Close GLHandle
'  FundLen = FundLen
  
SkipGLReport:
  Image0$ = "####"
  Image$ = "###0.00"
  Image3$ = "###,##0.00"
  Image4$ = "##0.0000"
  Image5$ = "####,##0.00"
  
  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GRetGross# = 0

  GTaxFrn# = 0
  TotDebit# = 0
  TotCredit# = 0
  
  EmpActLen = Len(EmpAct(1))
  DistbSumAcctsLen = Len(DistbSumAccts(1))
  
  TransRecLen = Len(TransRec(1))
  Emp1RecLen = Len(EmpRec1(1))
  
  OpenEmpData1File Emp1Handle
  NumOfRecs = LOF(Emp1Handle) / Len(EmpRec1(1))
  Close Emp1Handle

  SalCnt = 0
  HrlCnt = 0
  
  NumOfWageAccts = 0
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  For x = 1 To NumOfRecs
    Get EmpIdxNNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxNNameHandle
  
  
  PayRegisterRptName$ = "PRRPTS\REGISTERNSG.RPT"
  RHandle = FreeFile
'  On Error GoTo ErrorHandler
  Open PayRegisterRptName$ For Output As RHandle
  
  PayRegTotal$ = "PRRPTS\REGISTERNSTOTAL.RPT"
  PHandle = FreeFile
  Open PayRegTotal$ For Output As PHandle
  
  DistributionRptName$ = "PRRPTS\DISTRIBUNSG.RPT"
  DHandle = FreeFile
  Open DistributionRptName$ For Output As DHandle
  
  DistAcctTotalName$ = "PRRPTS\DISTTOTALSNS.RPT"
  DTHandle = FreeFile
  Open DistAcctTotalName$ For Output As DTHandle

  OpenEmpData1File NHandle
  
  OpenTransWorkFile THandle
  
  KillFile TempVoidFileName
  OpenTempVoidFile TVHandle
  
  For cnt = 1 To NumOfRecs
    Get THandle, IdxBuff(cnt), TransRec(1)
    If TransRec(1).TActive = True Then
      TActEmps = TActEmps + 1
    End If
  Next cnt
  
  ReDim ThisDedAmt(1 To 50) As Double
  For cnt = 1 To NumOfRecs + 1
    If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
      If TVCnt <> 0 And PRNetPoolFound = False Then
        GoSub NoPRNetForPoolCOrI
      End If
    ElseIf TVCnt <> 0 And PRNetPoolFound = False Then
      GoSub NoPRNetForPool
    End If
    If cnt = NumOfRecs + 1 Then Exit For
    Get THandle, IdxBuff(cnt), TransRec(1)
    If TransRec(1).TActive = True Then
      ReDim ThisPRDbtFund(1 To 1) As String
      ReDim ThisPRDbtAmt(1 To 1) As Double
      DbtCnt = 0
      Get NHandle, IdxBuff(cnt), EmpRec1(1)
      PRNetPoolFound = False
      TotalWHAndDeds = 0
      AcctNumCnt = 0
      ThisEmpCnt = 0
      For x = 1 To 50
        ThisDedAmt(x) = 0 '6/22/2004
      Next x
      ThisFTax# = 0 '6/22/2004
      ThisMTax# = 0 '6/22/2004
      ThisSSTax# = 0 '6/22/2004
      ThisStaTax# = 0 '6/22/2004
      ThisRTax# = 0 '6/22/2004
      TotalDeds# = 0 '6/22/2004
      GoSub SumAndPrintTime
      GoSub ParseDistributions
      ActiveCnt = ActiveCnt + 1
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  If ActiveCnt = 0 Then
    MsgBox "No employees have been designated for payroll processing."
    Close
    GoTo AltExit
  End If
  
  GoSub PrintSumTotal
  GoSub PrintDistTotal
  
  Close THandle
  Close NHandle
  Close RHandle
  Close PHandle
  Close DHandle
  Close DTHandle
  Close RptHandle
  Close TVHandle
  
  'HERE:  register report files have been written to disk
  Close TFHandle
  'GLIF START
  'if there is a GL transfer directory then make GLIF file
  
'  If Len(QPTrim$(SysRec(1).CITIDIR)) > 0 Then
  If Len(CurrCitiPath) > 0 Then
    ReDim Passed#(1 To 6)
    Passed#(1) = FTax#
    Passed#(2) = STax#
    Passed#(3) = SSTax#
    Passed#(4) = MTax#
    Passed#(5) = RETTOT#
    Passed#(6) = TNetPay#
    MakeGLIFFileG TotEIC#, TotDeds#(), Passed#(), DistbSumAccts() 'unrem
  End If
  
'  If Exist(QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile) Then
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile
  If Exist(CurrCitiPath + "\" + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + "\" + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + "\" + GLAcctIdxFile) Then
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + "\" + GLAcctIdxFile
  Else
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + GLAcctIdxFile
  End If
  
  If JFlag Then
    GLIdxRecLen = Len(JGLIdxRec(1))
  Else
    GLIdxRecLen = Len(GLIdxRec(1))
  End If
  
  GLIFRecLen = Len(GLIFRec(1))
  TotalGLIFS = FileSize("TempIF.DAT") \ GLIFRecLen
  TotalAccts = FileSize(GLIdxName$) \ GLIdxRecLen
  
  If TotalGLIFS = 0 Then
    PrintGLRpt = False
    GoTo SkipGLRpt
    
  Else
    PrintGLRpt = True
  End If
  
  GLIFRecLen = Len(GLIFRec(1))
  GLIFRecName$ = "TempIF.DAT"
  GRHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GRHandle Len = GLIFRecLen
  ReDim GLIFRec(1 To TotalGLIFS) As GLIFDataType14
  For x = 1 To TotalGLIFS
    Get GRHandle, x, GLIFRec(x)
  Next x
  Do
    OutOfOrder = False                     'assume it's sorted
    For x = 1 To UBound(GLIFRec) - 1
      If GLIFRec(x).TranAcct > GLIFRec(x + 1).TranAcct Then
        SWAP GLIFRec(x), GLIFRec(x + 1)    'if we had to swap
        OutOfOrder = True                'we're not done yet
      End If
    Next
  Loop While OutOfOrder

  For x = 1 To TotalGLIFS
    Put GRHandle, x, GLIFRec(x)
  Next x

  If TotalAccts = 0 Then
    Close
    GoTo SkipGLAccts
  End If
  Close GRHandle
  FrmShowPctComp.Label1 = "Reading G/L Accounts."
  FrmShowPctComp.Show ' , Me
  
  GLHandle = FreeFile
  Open GLIdxName$ For Random As GLHandle Len = GLIdxRecLen
  Select Case JFlag
  Case False
    ReDim GoodAccts(1 To TotalAccts) As Double
    For cnt = 1 To TotalAccts
      Get GLHandle, cnt, GLIdxRec(1)
      If GLIdxRec(1).AcctNum > 0 Then
        ActualAccts = ActualAccts + 1
        GoodAccts(ActualAccts) = GLIdxRec(1).AcctNum            'QPValL(AcctNum$)
        If GoodAccts(ActualAccts) < 9999999 Then
          GoodAccts(ActualAccts) = OldRound(GoodAccts(ActualAccts) * 100)
        End If
      End If
       FrmShowPctComp.ShowPctComp cnt, TotalAccts
       If FrmShowPctComp.Out = True Then
         Close
         FrmShowPctComp.Out = False
         Unload FrmShowPctComp
         Exit Sub
       End If
    Next
  Case True
    ReDim JGoodAccts(1 To TotalAccts) As String * 16
    For cnt = 1 To TotalAccts
      Get GLHandle, cnt, JGLIdxRec(1)
      ActualAccts = ActualAccts + 1
      ReplaceString JGLIdxRec(1).AcctNum, "-", ""
      JGoodAccts(ActualAccts) = JGLIdxRec(1).AcctNum
      FrmShowPctComp.ShowPctComp cnt, TotalAccts
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        Exit Sub
      End If

    Next
    
  End Select
  
  Close GLHandle
  TotalAccts = ActualAccts
  
  FrmShowPctComp.Label1 = "Checking for Invalid Accounts."
  FrmShowPctComp.Show

SkipGLAccts:
  FundTotals$ = "prrpts\GLFundTotals.RPT"
  TFHandle = FreeFile
  Open FundTotals$ For Output As TFHandle
  
  RptName$ = "prrpts\PRGLIFNSG.RPT"
  KillFile "prrpts\PRGLIFNSG.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  Close RHandle
  RptHandle = FreeFile
  Open RptName$ For Append As RptHandle
  
'  KillFile "TempIF.DAT" '12/27/2002
  THandle = FreeFile
  Open "TempIF.DAT" For Random As THandle Len = GLIFRecLen
  
  For cnt = 1 To TotalGLIFS
    FrmShowPctComp.ShowPctComp cnt, TotalGLIFS
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
    TotDebit# = OldRound(TotDebit# + GLIFRec(cnt).DrAmt)
    TotCredit# = OldRound(TotCredit# + GLIFRec(cnt).CrAmt)
    
    '10-13-94 ** added check for valid GL Account numbers
    If TotalAccts > 0 Then
      
      Select Case JFlag
      Case True
        For Cnt2 = 1 To TotalAccts
          If InStr(JGoodAccts(Cnt2), GLIFRec(cnt).TranAcct) And Len(QPTrim$(JGoodAccts(Cnt2))) = Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
            AcctOk = True
            Exit For
          End If
        Next
        If Not AcctOk Then
          GLError = -1
          LSet GLIFRec(cnt).Fill = "Error "
           Put THandle, cnt, GLIFRec(cnt)
        Else
          LSet GLIFRec(cnt).Fill = ""
        End If
        AcctOk = False
        
      Case False
        GLAcct@ = OldRound(Val(GLIFRec(cnt).TranAcct))
        If GLAcct@ < 9999999 Then
          GLAcct@ = GLAcct@ * 100
        End If
        For Cnt2 = 1 To TotalAccts
          If GoodAccts(Cnt2) = GLAcct@ Then
            AcctOk = True
            Exit For
          End If
        Next
        If Not AcctOk Then
          LSet GLIFRec(cnt).Fill = "Error "
           Put THandle, cnt, GLIFRec(cnt)
        Else
          LSet GLIFRec(cnt).Fill = ""
        End If
        AcctOk = False
      End Select
    End If
    
    'NoCheckAccts:
    RSet EDSAmt(1) = Using(Image3$, GLIFRec(cnt).DrAmt)
    RSet EDMAmt(1) = Using(Image3$, GLIFRec(cnt).CrAmt)
    LSet EDesc(1) = QPTrim$(GLIFRec(cnt).TranDesc)
    'added 8/18/04------------------------
    ThisDesc$ = QPTrim$(GLIFRec(cnt).TranDesc)
    If FundCount = 0 Then
      FundCount = FundCount + 1
      ReDim FundArray(1 To FundCount) As String
      FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
      ReDim FedTaxByFund(1 To FundCount) As Double
      ReDim MedTaxByFund(1 To FundCount) As Double
      ReDim SocTaxByFund(1 To FundCount) As Double
      ReDim StaTaxByFund(1 To FundCount) As Double
      ReDim RetTaxByFund(1 To FundCount) As Double
      ReDim MedMatByFund(1 To FundCount) As Double
      ReDim SocMatByFund(1 To FundCount) As Double
      ReDim RetMatByFund(1 To FundCount) As Double
      ReDim DedAmtByFund(1 To 50, 1 To FundCount)
      Thisx = 1
   Else
     For x = 1 To FundCount
       If FundArray(x) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen) Then
         Thisx = x
         Exit For
       End If
     Next x
     If x > FundCount Then
       FundCount = FundCount + 1
       ReDim Preserve FundArray(1 To FundCount) As String
       FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
       ReDim Preserve FedTaxByFund(1 To FundCount) As Double
       ReDim Preserve MedTaxByFund(1 To FundCount) As Double
       ReDim Preserve SocTaxByFund(1 To FundCount) As Double
       ReDim Preserve StaTaxByFund(1 To FundCount) As Double
       ReDim Preserve RetTaxByFund(1 To FundCount) As Double
       ReDim Preserve MedMatByFund(1 To FundCount) As Double
       ReDim Preserve SocMatByFund(1 To FundCount) As Double
       ReDim Preserve RetMatByFund(1 To FundCount) As Double
       ReDim Preserve DedAmtByFund(1 To 50, 1 To FundCount)
       Thisx = FundCount
     End If
   End If

    If ThisDesc = "Fed Withholdings" Then
      FedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Soc Sec Withholdings" Then
      SocTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Med Withholdings" Then
      MedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "State Withholdings" Then
      StaTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Ret Withholdings" Then
      RetTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Soc Match Liab" Then
      SocMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Med Match Liab" Then
      MedMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Ret Match Liab" Then
      RetMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    Else
      For z = 1 To 50
        If ThisDesc = QPTrim$(DedCodes(z).DCDESC1) Then
          DedAmtByFund(z, Thisx) = GLIFRec(cnt).CrAmt
        End If
      Next z
    End If
    'added 8/18/04--^^^^^^^^^^^^^^^-------
    
    
    NoAcctNum = Len(QPTrim$(GLIFRec(cnt).TranAcct))
    If NoAcctNum > 0 Then
    '                              0                    1                         2           3               4
      Print #RptHandle, GLIFRec(cnt).TranAcct; dlm; GLIFRec(cnt).Fill; dlm; EDesc(1); dlm; EDSAmt(1); dlm; EDMAmt(1); dlm;
    '                    5        6        7                8                   9
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; Unit(1).UFEMPR; dlm; MakeRegDate(PDR(1).PEREND); dlm;
    '                       10
      Print #RptHandle, NoAcctNum
    End If
  Next
  
  GoSub GLIFTotals
  
  Fund$ = Left$(QPTrim$(GLIFRec(1).TranAcct), FundLen)
  FDebit# = GLIFRec(1).DrAmt
  FCredit# = GLIFRec(1).CrAmt
  For cnt = 2 To TotalGLIFS
    NFund$ = Left$(QPTrim$(GLIFRec(cnt).TranAcct), FundLen)
    If NFund$ <> Fund$ Then
      GoSub PrintFundTotal
      Fund$ = NFund$
      FDebit# = GLIFRec(cnt).DrAmt
      FCredit# = GLIFRec(cnt).CrAmt
    Else
      FDebit# = OldRound(FDebit# + GLIFRec(cnt).DrAmt)
      FCredit# = OldRound(FCredit# + GLIFRec(cnt).CrAmt)
    End If
  Next
  
  GoSub PrintFundTotal
  
  Close RptHandle
  Close THandle
  Close TVHandle
  Close
SkipGLRpt:
  
  '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
  '05-10-94  fixed to insure everything was cleaned up after report
  '06-28-94  move so all is cleaned up BEFORE the report prints.
  '07-15-94  move again to add gl interface report.
  '----------------------------------------------------------------------------
  
  arPayRollRegisterNS.Show
  frmLoadingRpt.Show
  If RetCode = -1 Then GoTo AltExit
  
  arEarnDistRegNS.Show
  frmLoadingRpt.Show
  If RetCode = -1 Then GoTo AltExit
  
  If GLError <> -1 Then
    GLError = TotalAccts
  End If
  
  If PrintGLRpt Then
    arGLRegisterNS.Show
    frmLoadingRpt.Show
  End If
  
  If GLCreditTotal <> GLDebitTotal Then
    frmBackGround.Show
    frmMessage.Label1.Caption = "The General Ledger Interface is OUT OF BALANCE."
    frmMessage.Label1.Top = 800
    frmMessage.Show vbModal
    Unload frmBackGround
    MainLog "User warned that the GL Interface is OUT OF BALANCE (Debit Total = " + QPTrim$(Using$("$#,###,##0.00", GLDebitTotal)) + " and the Credit Total is " + QPTrim$(Using$("$#,###,##0.00", GLCreditTotal)) + ")."
  End If
  
  If GLError < 1 Then
    frmBackGround.Show
    frmMessage.Label1.Caption = "General Ledger number errors have been found in the GL Interface Report."
    frmMessage.Label1.Top = 800
    frmMessage.Show vbModal
    MainLog "User warned that the GL Interface Report has GL number errors."
    Unload frmBackGround
  End If
  
'  Dim y As Integer
'  Dim NumAccts As Integer
'  OpenTempVoidFile TVHandle
'  NumAccts = LOF(TVHandle) / Len(TempVoid)
'  For x = 1 To NumAccts
'    Get TVHandle, x, TempVoid
'    If QPTrim$(TempVoid.EmpNum) = "43128" Then
'      Debug.Print TempVoid.NumOfAccts
'      Debug.Print TempVoid.EmpNum
'      Debug.Print TempVoid.PPEGL + " PPE              " + CStr(TempVoid.PPEAmt)
'      Debug.Print TempVoid.PPETotGL + " PPE Total   " + CStr(TempVoid.PPETotAmt)
'      Debug.Print TempVoid.PRNetGL + " PRNET            " + CStr(TempVoid.PRNet)
'      Debug.Print TempVoid.SOCWHGL + " SOC Withholdings " + CStr(TempVoid.SOCWHAmt)
'      Debug.Print TempVoid.MEDWHGL + " MED Withholdings " + CStr(TempVoid.MEDWHAmt)
'      Debug.Print TempVoid.SOCMATCRGL + " SOC Match Liab   " + CStr(TempVoid.SOCMATCRAmt)
'      Debug.Print TempVoid.MEDMATCRGL + " MED Match Liab   " + CStr(TempVoid.MEDMATCRAmt)
'      Debug.Print TempVoid.FEDWHGL + " FED Withholdings " + CStr(TempVoid.FEDWHAmt)
'      Debug.Print TempVoid.STAWHGL + " STA Withholdings " + CStr(TempVoid.STAWHAmt)
'      Debug.Print TempVoid.RETWHGL + " RET Withholdings " + CStr(TempVoid.RETWHAmt)
'      Debug.Print TempVoid.RETMATCRGL + " RET Match Liab   " + CStr(TempVoid.RETMATCRAmt)
'      For Cnt2 = 1 To 50
'        If TempVoid.DedData(Cnt2).DAmt > 0 Then
'          Debug.Print TempVoid.DedData(Cnt2).DedGLNum + " Deduction        " + CStr(TempVoid.DedData(Cnt2).DAmt)
'        End If
'      Next Cnt2
'      Debug.Print TempVoid.WagesGL + "  Wages           " + CStr(TempVoid.WagesAmt)
'      Debug.Print TempVoid.SOCMATDBGL + " SOC Match        " + CStr(TempVoid.SOCMATDBAmt)
'      Debug.Print TempVoid.MEDMATDBGL + " MED Match        " + CStr(TempVoid.MEDMATDBAmt)
'      Debug.Print TempVoid.RETMATDBGL + " RET Match        " + CStr(TempVoid.RETMATDBAmt)
'    End If
'  Next x
'  Close TVHandle
'

AltExit:
  
  Exit Sub
  
PrintFundTotal:
  RSet EDSAmt(1) = Using(Image3$, FDebit#)
  RSet EDMAmt(1) = Using(Image3$, FCredit#)
  LSet EDesc(1) = ""
  LSet GLIFRec(1).Fill = ""
  RSet GLIFRec(1).TranAcct = Fund$
  If Len(QPTrim$(Fund$)) > 0 Then
    For x = 1 To FundCount
      If FundArray(x) = QPTrim$(GLIFRec(1).TranAcct) Then
        ThisFund = x
        Exit For
      End If
    Next x
    
    If x <= FundCount Then
      TOTFEDTAX = OldRound(TOTFEDTAX + FedTaxByFund(ThisFund))
      TOTMEDTAX = OldRound(TOTMEDTAX + MedTaxByFund(ThisFund))
      TOTSOCTAX = OldRound(TOTSOCTAX + SocTaxByFund(ThisFund))
      TOTSTATAX = OldRound(TOTSTATAX + StaTaxByFund(ThisFund))
      TOTRetTax = OldRound(TOTRetTax + RetTaxByFund(ThisFund))
      TOTMEDMat = OldRound(TOTMEDMat + MedMatByFund(ThisFund))
      TOTSOCMat = OldRound(TOTSOCMat + SocMatByFund(ThisFund))
      TOTRETMat = OldRound(TOTRETMat + RetMatByFund(ThisFund))
    '                            0                   1              2
      Print #TFHandle, GLIFRec(1).TranAcct; dlm; EDSAmt(1); dlm; EDMAmt(1); dlm;
  '                                3                         4
      Print #TFHandle, FedTaxByFund(ThisFund); dlm; MedTaxByFund(ThisFund); dlm;
  '                                5                         6
      Print #TFHandle, SocTaxByFund(ThisFund); dlm; StaTaxByFund(ThisFund); dlm;
  '                                7                       8
      Print #TFHandle, RetTaxByFund(ThisFund); dlm; MedMatByFund(ThisFund); dlm;
  '                                9                       10
      Print #TFHandle, SocMatByFund(ThisFund); dlm; RetMatByFund(ThisFund); dlm;
      
      For y = 1 To 50
        If DedAmtByFund(y, ThisFund) > 0 Then
          TotDedAmt(y) = OldRound(TotDedAmt(y) + DedAmtByFund(y, ThisFund))
          Print #TFHandle, QPTrim$(DedCodes(y).DCDESC1); dlm; DedAmtByFund(y, ThisFund); dlm;
        Else
          Print #TFHandle, QPTrim$(DedCodes(y).DCDESC1); dlm; 0; dlm;
        End If
      Next y
      
      Print #TFHandle, NumOfDeds; dlm;
    Else
      '                 0       1       2
      Print #TFHandle, ""; dlm; 0; dlm; 0; dlm;
      '                3       4       5       6
      Print #TFHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm;
      '                7       8       9       10
      Print #TFHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm;
      
      For y = 1 To 50
        Print #TFHandle, ""; dlm; 0; dlm;
      Next y
      
      Print #TFHandle, 0; dlm;
      
    End If
    
    Print #TFHandle, TOTFEDTAX; dlm; TOTMEDTAX; dlm; TOTSOCTAX; dlm; TOTSTATAX; dlm;
    
    Print #TFHandle, TOTRetTax; dlm; TOTMEDMat; dlm; TOTSOCMat; dlm; TOTRETMat; dlm;
    
    For y = 1 To 50
      If TotDedAmt(y) > 0 Then
        Print #TFHandle, QPTrim$(DedCodes(y).DCDESC1); dlm; TotDedAmt(y); dlm;
      Else
        Print #TFHandle, QPTrim$(DedCodes(y).DCDESC1); dlm; 0; dlm;
      End If
    Next y
    
    Print #TFHandle, TotDebit; dlm; TotCredit
  End If
  
  GLDebitTotal = OldRound(GLDebitTotal + EDSAmt(1))
  GLCreditTotal = OldRound(GLCreditTotal + EDMAmt(1))
  
  Return
  
GLIFTotals:
  '                  0        1        2       3        4
  Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  
  '                            5                             6
  Print #RptHandle, Using(Image3$, TotDebit#); dlm; Using(Image3$, TotCredit#); dlm;
  If TotalAccts = 0 Then
 '                                 7
    Print #RptHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory."; dlm;
  Else
  '                    7
    Print #RptHandle, ""; dlm;
  End If
  '                          8                    9                         10
  Print #RptHandle, Unit(1).UFEMPR; dlm; MakeRegDate(PDR(1).PEREND); dlm; NoAcctNum
  Return
  
SumAndPrintTime:
  RegHrs# = OldRound(RegHrs# + TransRec(1).RegHrsWork)
  VACHRS# = OldRound(VACHRS# + TransRec(1).VacUsed)
  SICKHRS# = OldRound(SICKHRS# + TransRec(1).SickUsed)
  HOLHRS# = OldRound(HOLHRS# + TransRec(1).HOLHOURS)
  COMPHRS# = OldRound(COMPHRS# + TransRec(1).CompUsed)
  PerHours# = OldRound(PerHours# + TransRec(1).PerHours)
  
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).PerHours)

  'added for EIC   6/07/94
  TotEIC# = OldRound(TotEIC# + TransRec(1).EICAmt)
  
  '-=-=-=-=-=-=-=
  TotHrs# = OldRound(TotHrs# + TransRec(1).OTHours)
  If TransRec(1).OTHrsPaid > 0 Then
    TOTPaid# = OldRound(TOTPaid# + TransRec(1).OTHrsPaid)
  End If
  TOTComp# = OldRound(TOTComp# + TransRec(1).OT2Comp)
  
  TRegWage# = OldRound(TRegWage# + TransRec(1).TotRegWage)
  
  If TransRec(1).TotOTWage > 0 Then
    TOTWage# = OldRound(TOTWage# + TransRec(1).TotOTWage)
  End If
  GPay# = OldRound(GPay# + TransRec(1).GrossPay)
  SSTax# = OldRound(SSTax# + TransRec(1).SocTaxAmt)
  MTax# = OldRound(MTax# + TransRec(1).MedTaxAmt)
  FTax# = OldRound(FTax# + TransRec(1).FedTaxAmt)
  STax# = OldRound(STax# + TransRec(1).StaTaxAmt)
  If TransRec(1).RetireAmt > 0 Then
    RETTOT# = OldRound(RETTOT# + TransRec(1).RetireAmt)
  End If
  
  TNetPay# = OldRound(TNetPay# + TransRec(1).NetPay)
  GFedGross# = OldRound(GFedGross# + TransRec(1).FedGrossPay)
  GStaGross# = OldRound(GStaGross# + TransRec(1).StaGrossPay)
  GSocGross# = OldRound(GSocGross# + TransRec(1).SocGrossPay)
  GMedGross# = OldRound(GMedGross# + TransRec(1).MedGrossPay)
  GRetGross# = OldRound(GRetGross# + TransRec(1).RetGrossPay)
  
  GTaxFrn# = OldRound(GTaxFrn# + TransRec(1).TaxFring)
  LSet ENumb(1) = LTrim$(EmpRec1(1).EmpNo)
  LSet EName(1) = QPTrim$(EmpRec1(1).EmpLName) + ", " + QPTrim$(EmpRec1(1).EmpFName)
  RSet BRat(1) = Using(Image3$, TransRec(1).BaseRate)
  RSet ORat(1) = Using(Image3$, TransRec(1).OTRate)
  
  RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsWork)
  
  RSet VHrs(1) = Using(Image$, TransRec(1).VacUsed)
  RSet SHrs(1) = Using(Image$, TransRec(1).SickUsed)
  RSet HHrs(1) = Using(Image$, TransRec(1).HOLHOURS)
  RSet CHrs(1) = Using(Image$, TransRec(1).CompUsed)
  RSet THrs(1) = Using(Image$, TransRec(1).RegHrsPaid)
  
  RSet TaxFrn(1) = Using(Image$, TransRec(1).TaxFring)

  RSet PHrs(1) = Using(Image$, TransRec(1).PerHours)

  RSet OTPaid(1) = Using(Image$, TransRec(1).OTHrsPaid)
  RSet OTComp(1) = Using(Image$, TransRec(1).OT2Comp)
  
  'added for EIC     6/07/94
  RSet EEicP(1) = Using(Image3$, TransRec(1).EICAmt)
  
  Select Case TransRec(1).PayType
  Case "S"
    RSet RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  Case Else
    HrlCnt = HrlCnt + 1
  End Select
  
  '=======
  RSet RErnP(1) = Using(Image3$, TransRec(1).TotRegWage)
  RSet OErnP(1) = Using(Image3$, TransRec(1).TotOTWage)
  
  RSet GPayP(1) = Using(Image3$, TransRec(1).GrossPay)
  RSet SSTaxP(1) = Using(Image3$, TransRec(1).SocTaxAmt)
  RSet MTaxP(1) = Using(Image3$, TransRec(1).MedTaxAmt)
  RSet FTaxP(1) = Using(Image3$, TransRec(1).FedTaxAmt)
  RSet STaxP(1) = Using(Image3$, TransRec(1).StaTaxAmt)
  
  RSet RetirP(1) = Using(Image3$, TransRec(1).RetireAmt)
  
  RSet NetPayP(1) = Using(Image3$, TransRec(1).NetPay)
  
  For Cnt2 = 1 To 3 'LastErn 'changed from LastErn to 3 on 8/5/05
    TotErns(Cnt2) = OldRound(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
  Next
  '                        0                               1
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm;
  '                  2             3
  Print #RHandle, ENumb(1); dlm; EName(1); dlm;
  '                  4            5              6              7             8
  Print #RHandle, BRat(1); dlm; ORat(1); dlm; TaxFrn(1); dlm; RHrs(1); dlm; VHrs(1); dlm;
  '                  9            10              11           12            13
  Print #RHandle, SHrs(1); dlm; HHrs(1); dlm; CHrs(1); dlm; PHrs(1); dlm; THrs(1); dlm;
  '                  14              15              16             17             18
  Print #RHandle, OTPaid(1); dlm; OTComp(1); dlm; RErnP(1); dlm; OErnP(1); dlm; TransRec(1).EAmt(3); dlm;
  '                     19                         20
  Print #RHandle, TransRec(1).EAmt(2); dlm; TransRec(1).EAmt(1); dlm;
  '                           21                                 22                                  23
  Print #RHandle, QPTrim$(ErnCodes(3).ERNCODE1); dlm; QPTrim$(ErnCodes(2).ERNCODE1); dlm; QPTrim$(ErnCodes(1).ERNCODE1); dlm;
  '                  24              25            26             27              28
  Print #RHandle, GPayP(1); dlm; SSTaxP(1); dlm; MTaxP(1); dlm; FTaxP(1); dlm; STaxP(1); dlm;
  '                  29              30                  31
  Print #RHandle, RetirP(1); dlm; NetPayP(1); dlm; TransRec(1).EICAmt; dlm;
  
  For Cnt2 = 1 To 50
    '                   32 - 81
    Print #RHandle, QPTrim$(DedCodes(Cnt2).DCDESC1); dlm;
  Next Cnt2
  
  For Cnt2 = 1 To 50
    TotDeds#(Cnt2) = OldRound(TotDeds#(Cnt2) + TransRec(1).DAmt(Cnt2))
    ' 82 - 131
    Print #RHandle, Using(Image3$, TransRec(1).DAmt(Cnt2)); dlm;
  Next
  '                  132
  Print #RHandle, NumOfDeds
  
  Return
  
PrintSumTotal:
  RSet SCnt(1) = Using(Image0$, SalCnt)
  RSet HCnt(1) = Using(Image0$, HrlCnt)
  
  RSet Fill11(1) = Using(Image3$, GTaxFrn#)
  RSet THrs(1) = Using(Image3$, TotalHrs#)
  RSet RHrs(1) = Using(Image3$, RegHrs#)
  RSet VHrs(1) = Using(Image3$, VACHRS#)
  RSet SHrs(1) = Using(Image3$, SICKHRS#)
  RSet HHrs(1) = Using(Image3$, HOLHRS#)
  RSet CHrs(1) = Using(Image3$, COMPHRS#)

  RSet PHrs(1) = Using(Image3$, PerHours#)

  RSet OTPaid(1) = Using(Image3$, TOTPaid#)
  RSet OTComp(1) = Using(Image3$, TOTComp#)
  
  RSet RErnP(1) = Using(Image3$, TRegWage#)
  RSet OErnP(1) = Using(Image3$, TOTWage#)
  
  RSet GPayP(1) = Using(Image3$, GPay#)
  RSet SSTaxP(1) = Using(Image3$, SSTax#)
  RSet MTaxP(1) = Using(Image3$, MTax#)
  RSet FTaxP(1) = Using(Image3$, FTax#)
  RSet STaxP(1) = Using(Image3$, STax#)
  RSet RetirP(1) = Using(Image3$, RETTOT#)
  RSet NetPayP(1) = Using(Image3$, TNetPay#)
  
  '                           0                             1
  Print #PHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm;
  '                 2             3             4             5              6             7              8
  Print #PHandle, SCnt(1); dlm; HCnt(1); dlm; Fill11(1); dlm; RHrs(1); dlm; VHrs(1); dlm; SHrs(1); dlm; HHrs(1); dlm;
  '                 9             10           11             12                13
  Print #PHandle, CHrs(1); dlm; PHrs(1); dlm; THrs(1); dlm; OTPaid(1); dlm; OTComp(1); dlm;
  '                14              15                        16                        17                                   18
  Print #PHandle, RErnP(1); dlm; OErnP(1); dlm; Using(Image3$, TotErns(3)); dlm; Using(Image3$, TotErns(2)); dlm; Using(Image3$, TotErns(1)); dlm;
  '                 19             20              21
  Print #PHandle, GPayP(1); dlm; SSTaxP(1); dlm; MTaxP(1); dlm;
  '                 22             23             24                25
  Print #PHandle, FTaxP(1); dlm; STaxP(1); dlm; RetirP(1); dlm; NetPayP(1); dlm;
  
  For Cnt2 = 1 To 50
    '                            26 - 75
    Print #PHandle, Using(Image3$, TotDeds#(Cnt2)); dlm;
  Next Cnt2
  '                       76                             77
  Print #PHandle, Using(Image5$, GFedGross#); dlm; Using(Image5$, GStaGross#); dlm;
  '                         78                            79
  Print #PHandle, Using(Image5$, GMedGross#); dlm; Using(Image5$, GSocGross#); dlm;
  '                           80                              81
  
  Print #PHandle, Using(Image5$, GRetGross#); dlm; Using(Image5$, TotEIC#); dlm;
  
  For Cnt2 = 1 To 50
    '
    Print #PHandle, QPTrim$(DedCodes(Cnt2).DCDESC1); dlm;
  Next Cnt2
  Print #PHandle, NumOfDeds; dlm; '8/5/05
  
  Print #PHandle, QPTrim$(ErnCodes(3).ERNCODE1); dlm; QPTrim$(ErnCodes(2).ERNCODE1); dlm; QPTrim$(ErnCodes(1).ERNCODE1) '8/5/05
  
  Return
  
  '-----------------------------------------------------------------------
ParseDistributions:
  
  ReDim EmpAct(1 To 14) As DistWageRptType
  ENumOfAct = 0
  
  'process wage distributions
  For Cnt2 = 1 To 8
    Acct$ = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
    If Len(Acct$) > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).Pct = TransRec(1).TDist(Cnt2).DPct
      EmpAct(ENumOfAct).RHrs = TransRec(1).TDist(Cnt2).DRHrs
      EmpAct(ENumOfAct).OHrs = TransRec(1).TDist(Cnt2).DOHrs
      EmpAct(ENumOfAct).RWage = TransRec(1).TDist(Cnt2).DRWage
      EmpAct(ENumOfAct).OWage = TransRec(1).TDist(Cnt2).DOWage
      EmpAct(ENumOfAct).GrossPay = OldRound(EmpAct(ENumOfAct).RWage + EmpAct(ENumOfAct).OWage)
    End If
  Next
  
  'process earnings distributions
  For Cnt2 = 1 To 6
    Acct$ = QPTrim$(TransRec(1).EDist(Cnt2).EAcct)
    If Len(Acct$) > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).AddEarn = TransRec(1).EDist(Cnt2).EAmt
      EmpAct(ENumOfAct).GrossPay = TransRec(1).EDist(Cnt2).EAmt
    End If
  Next
  
  'HERE: got all accts for this employee
  
  First = 1
  Start = 1
  
  Last = ENumOfAct
  
  'purge and sum employee's dupelicate account distributions
  
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If EmpAct(First).Acct = EmpAct(Cnt2).Acct Then
        LSet EmpAct(Cnt2).Acct = ""             'blank acct num as a flag
        EmpAct(First).Pct = OldRound(EmpAct(First).Pct + EmpAct(Cnt2).Pct)
        EmpAct(First).RHrs = OldRound(EmpAct(First).RHrs + EmpAct(Cnt2).RHrs)
        EmpAct(First).OHrs = OldRound(EmpAct(First).OHrs + EmpAct(Cnt2).OHrs)
        EmpAct(First).RWage = OldRound(EmpAct(First).RWage + EmpAct(Cnt2).RWage)
        EmpAct(First).OWage = OldRound(EmpAct(First).OWage + EmpAct(Cnt2).OWage)
        EmpAct(First).AddEarn = OldRound(EmpAct(First).AddEarn + EmpAct(Cnt2).AddEarn)
        EmpAct(First).GrossPay = OldRound(EmpAct(First).GrossPay + EmpAct(Cnt2).GrossPay)
      End If
    Next
Again:
    First = First + 1
  Loop Until First >= Last
  
  'calc percentages of matching amts to each account distribution
  
  For Cnt2 = 1 To ENumOfAct
    EmpAct(Cnt2).MATSocAmt = OldRound(TransRec(1).MatchSocAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATMedAmt = OldRound(TransRec(1).MatchMedAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATRetAmt = OldRound(TransRec(1).MatchRetAmt * (EmpAct(Cnt2).Pct * 0.01))
  Next
  
  '---------------------------------------------------------------------------
  'calc and adjust matching distribution amts
  'find last active account
  'adjust Social Amt
  
  Do
    TotalSocAmt# = 0
    For Cnt2 = 1 To 14          '8
      TotalSocAmt# = OldRound(TotalSocAmt# + EmpAct(Cnt2).MATSocAmt)
    Next
    If TotalSocAmt# = 0 Then GoTo SkipSocDist
    If TotalSocAmt# <> TransRec(1).MatchSocAmt Then
      For Cnt3 = 14 To 1 Step -1                '8 TO 1 STEP -1
        If EmpAct(Cnt3).MATSocAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalSocAmt# > TransRec(1).MatchSocAmt Then
        DistDif# = OldRound(TotalSocAmt# - TransRec(1).MatchSocAmt)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt - DistDif#)
      ElseIf TotalSocAmt# < TransRec(1).MatchSocAmt Then
        DistDif# = OldRound(TransRec(1).MatchSocAmt - TotalSocAmt#)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt + DistDif#)
      End If
    End If
  Loop Until TotalSocAmt# = OldRound(TransRec(1).MatchSocAmt)
  '-=-=-=-=-=-
  'adjust Medicare Amt
SkipSocDist:
  Do
    TotalMedAmt# = 0
    For Cnt2 = 1 To 8
      TotalMedAmt# = OldRound(TotalMedAmt# + EmpAct(Cnt2).MATMedAmt)
    Next
    If TotalMedAmt# = 0 Then GoTo SkipMedDist
    If TotalMedAmt# <> TransRec(1).MatchMedAmt Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATMedAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalMedAmt# > TransRec(1).MatchMedAmt Then
        DistDif# = OldRound(TotalMedAmt# - TransRec(1).MatchMedAmt)
        EmpAct(LastActive).MATMedAmt = OldRound(EmpAct(LastActive).MATMedAmt - DistDif#)
      ElseIf TotalMedAmt# < TransRec(1).MatchMedAmt Then
        DistDif# = OldRound(TransRec(1).MatchMedAmt - TotalMedAmt#)
        EmpAct(LastActive).MATMedAmt = OldRound(EmpAct(LastActive).MATMedAmt + DistDif#)
      End If
    End If
  Loop Until TotalMedAmt# = OldRound(TransRec(1).MatchMedAmt)
  '-=-=-=-=-=-
SkipMedDist:
  'adjust Retire Amt
  Do
    TotalRetAmt# = 0
    For Cnt2 = 1 To 8
      TotalRetAmt# = OldRound(TotalRetAmt# + EmpAct(Cnt2).MATRetAmt)
    Next
    If TotalRetAmt# = 0 Then GoTo SkipRetDist
    If TotalRetAmt# <> TransRec(1).MatchRetAmt Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATRetAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalRetAmt# > TransRec(1).MatchRetAmt Then
        DistDif# = OldRound(TotalRetAmt# - TransRec(1).MatchRetAmt)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt - DistDif#)
      ElseIf TotalRetAmt# < TransRec(1).MatchRetAmt Then
        DistDif# = OldRound(TransRec(1).MatchRetAmt - TotalRetAmt#)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt + DistDif#)
      End If
    End If
  Loop Until TotalRetAmt# = OldRound(TransRec(1).MatchRetAmt)
SkipRetDist:
  AcctCnt = 0
  For Cnt2 = 1 To Last
    If Len(QPTrim$(EmpAct(Cnt2).Acct)) > 0 Then
      AcctCnt = AcctCnt + 1
    End If
  Next
  
  'print this employee's distributions
  For Cnt2 = 1 To Last
    If Len(QPTrim$(EmpAct(Cnt2).Acct)) > 0 Then
      '                   0            1               2            3            4
      Print #DHandle, ENumb(1); dlm; EName(1); dlm; BRat(1); dlm; ORat(1); dlm; AcctCnt; dlm;
      
      GoSub PrintEmpDist
      End If
    
  Next
 
  ' sum to master distrubution list
  
  For Cnt2 = 1 To Last          'process wage distributions
    Acct$ = QPTrim$(EmpAct(Cnt2).Acct)
    If Len(Acct$) > 0 Then
      If NumOfWageAccts > 0 Then
        For Cnt3 = 1 To NumOfWageAccts
          If Acct$ = QPTrim$(DistbSumAccts(Cnt3).Acct) Then
            DistbSumAccts(Cnt3).RWage = OldRound(DistbSumAccts(Cnt3).RWage + EmpAct(Cnt2).RWage)
            DistbSumAccts(Cnt3).OWage = OldRound(DistbSumAccts(Cnt3).OWage + EmpAct(Cnt2).OWage)
            DistbSumAccts(Cnt3).RHrs = OldRound(DistbSumAccts(Cnt3).RHrs + EmpAct(Cnt2).RHrs)
            DistbSumAccts(Cnt3).OHrs = OldRound(DistbSumAccts(Cnt3).OHrs + EmpAct(Cnt2).OHrs)
            DistbSumAccts(Cnt3).AddEarn = OldRound(DistbSumAccts(Cnt3).AddEarn + EmpAct(Cnt2).AddEarn)
            DistbSumAccts(Cnt3).GrossPay = OldRound(DistbSumAccts(Cnt3).GrossPay + EmpAct(Cnt2).GrossPay)
            DistbSumAccts(Cnt3).MATSocAmt = OldRound(DistbSumAccts(Cnt3).MATSocAmt + EmpAct(Cnt2).MATSocAmt)
            DistbSumAccts(Cnt3).MATMedAmt = OldRound(DistbSumAccts(Cnt3).MATMedAmt + EmpAct(Cnt2).MATMedAmt)
            DistbSumAccts(Cnt3).MATRetAmt = OldRound(DistbSumAccts(Cnt3).MATRetAmt + EmpAct(Cnt2).MATRetAmt)
            Exit For
          End If
        Next
        If Cnt3 > NumOfWageAccts Then
          GoSub AddDistbSumAcct 'add new sum dist acct
        End If
      Else      'no previous sum accts. add new one
        GoSub AddDistbSumAcct   'add new sum dist acct
      End If
    End If
  Next
  Return
  
AddDistbSumAcct:                'add amts to grand total acts summary
  
  NumOfWageAccts = NumOfWageAccts + 1
  If NumOfWageAccts > 1 Then
    ReDim Preserve DistbSumAccts(1 To NumOfWageAccts) As DistWageRptType
  End If
  DistbSumAccts(NumOfWageAccts) = EmpAct(Cnt2)
  
  Return

PrintEmpDist:
  LSet EDAct(1) = EmpAct(Cnt2).Acct
  RSet EDPct(1) = Using(Image$, EmpAct(Cnt2).Pct)
  RSet EDRHrs(1) = Using(Image$, EmpAct(Cnt2).RHrs)
  RSet EDOHrs(1) = Using(Image$, EmpAct(Cnt2).OHrs)
  RSet EDRPay(1) = Using(Image3$, EmpAct(Cnt2).RWage)
  RSet EDOPay(1) = Using(Image$, EmpAct(Cnt2).OWage)
  RSet EDEarn(1) = Using(Image$, EmpAct(Cnt2).AddEarn)
  RSet EDGroP(1) = Using(Image$, EmpAct(Cnt2).GrossPay)
  
  RSet EDSAmt(1) = Using(Image$, EmpAct(Cnt2).MATSocAmt)
  RSet EDMAmt(1) = Using(Image$, EmpAct(Cnt2).MATMedAmt)
  RSet EDRAmt(1) = Using(Image$, EmpAct(Cnt2).MATRetAmt)
  '                         5                              6
  Print #DHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm;
  '                  7             8               9               10
  Print #DHandle, EDAct(1); dlm; EDPct(1); dlm; EDRHrs(1); dlm; EDOHrs(1); dlm;
  '                  11               12               13               14
  Print #DHandle, EDRPay(1); dlm; EDOPay(1); dlm; EDEarn(1); dlm; EDGroP(1); dlm;
  '                  15              16              17
  Print #DHandle, EDSAmt(1); dlm; EDMAmt(1); dlm; EDRAmt(1)
  
  
'----------Void Check Code----------------------------
  '6/22/2004
  RETTOT# = RETTOT#
  TNetPay# = TNetPay#
  
  'Withholdings comes thru as the same amount no matter how many loops are run...
  'That amount should be saved once (pool fund) because when Void Check posts to the GL
  'for each check only one WH or deduction amount needs to be posted for only one GL (pool fund)
  
  'The matching amounts are saved differently...with each loop the debit
  'amounts are saved for their individual GL number but the credit amounts are
  'accumulated and saved to the pool fund number + GL suffix. All debit
  'amounts are accumulated for each fund number and saved as one credit
  'amount for that fund number
  
  FundNumOnly = Mid(EDAct(1), 1, FundLen)
  AcctNumCnt = AcctNumCnt + 1
  ThisFACnt = FACnt
  If InStr(Mid(EDAct(1), 1, ThisFACnt), "-") Then
    ThisFACnt = ThisFACnt + 1
  End If
  
  TempVoid.EmpNum = ENumb(1)
  ThisPR = EDGroP(1)
'  If QPTrim$(ENumb(1)) = "27" Then Stop
  TempVoid.WagesAmt = EDGroP(1)
  TotalWages = TotalWages + EDGroP(1)
  TempVoid.WagesGL = EDAct(1)
  
  
  'Non-split Pool...this type has a pool fund (usually fund 10) that serves
  'as the fund from which all WH and deductions are taken. This fund can also serve
  'as a fund from which an employee can be paid (just like fund 30, 11, etc.).
  'When the pool fund also serves as a typical 'paid from' fund then this code
  'handles it without having to run through the sub 'NoPRNetForPool'. However,
  'if the program does not detect the pool fund also serving as a 'Paid From'
  'fund (PRNetPoolFound = True) then the program must activate the sub 'NoPRNetForPool'
  'to insert the activity required for the pool fund.
  
  'Non-split Pool with Central Depository...This case is just like without Central
  'depository except that the Central Depository fund (usually 01) must be factored
  'in to the mix. The code below handles employee data when there is a pool fund also
  'serving as a 'Pay From' fund. However, if there is no pool fund detected (PRNetPoolFound = False)
  'then the sub 'NoPRNetForPoolCOrI' must be activated in order to record pay activity
  'for the overlooked pool fund.
  
  If PoolFundNum$ = Mid(EDAct(1), 1, 2) Then
    If PRPoolProcessed = True Then
        TempVoid.FEDWHAmt = 0
        TempVoid.MEDWHAmt = 0
        TempVoid.SOCWHAmt = 0
        TempVoid.RETWHAmt = 0
        TempVoid.STAWHAmt = 0
        GoTo PoolAlreadyDone
    Else
      PRNetPoolFound = True 'dictates if the PRNetPoolFound sub or PRNetPoolFoundCOrI sub activates
    End If
  End If
  '--------9/17/04--------
  If PoolFundNum$ = Mid(EDAct(1), 1, 2) Then 'we found the pool fund through normal activity (pool fund
  'is also being used as a 'PayFrom' fund) so go ahead and assign WH values to this fund
    TempVoid.FEDWHAmt = TransRec(1).FedTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).FedTaxAmt
    TempVoid.MEDWHAmt = TransRec(1).MedTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).MedTaxAmt
    TempVoid.SOCWHAmt = TransRec(1).SocTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).SocTaxAmt
    TempVoid.RETWHAmt = TransRec(1).RetireAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).RetireAmt
    TempVoid.STAWHAmt = TransRec(1).StaTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).StaTaxAmt
  Else
    TempVoid.FEDWHAmt = 0
    TempVoid.MEDWHAmt = 0
    TempVoid.SOCWHAmt = 0
    TempVoid.RETWHAmt = 0
    TempVoid.STAWHAmt = 0
  End If
PoolAlreadyDone:
  ThisPR = OldRound(ThisPR - TempVoid.FEDWHAmt)
  'next employee data comes thru
  TempVoid.FEDWHGL = QPTrim$(SysRec(1).Liab(1).Acct)
  TempVoid.FEDWHGL = ReplaceString(TempVoid.FEDWHGL, "-", "")
  TempVoid.FEDWHGL = AddDashesToGLNumber(TempVoid.FEDWHGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.MEDWHAmt)
  'next employee data comes thru
  TempVoid.MEDWHGL = QPTrim$(SysRec(1).Liab(4).Acct)
  TempVoid.MEDWHGL = ReplaceString(TempVoid.MEDWHGL, "-", "")
  TempVoid.MEDWHGL = AddDashesToGLNumber(TempVoid.MEDWHGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATCRAmt = EDMAmt(1)
  TempVoid.MEDMATCRGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).MEDLIAB))
  TempVoid.MEDMATCRGL = ReplaceString(TempVoid.MEDMATCRGL, "-", "")
  TempVoid.MEDMATCRGL = AddDashesToGLNumber(TempVoid.MEDMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATDBAmt = EDMAmt(1)
  TempVoid.MEDMATDBGL = Mid(EDAct(1), 1, ThisFACnt) + QPTrim$(SysRec(1).MEDEXP)
  TempVoid.MEDMATDBGL = ReplaceString(TempVoid.MEDMATDBGL, "-", "")
  TempVoid.MEDMATDBGL = AddDashesToGLNumber(TempVoid.MEDMATDBGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.SOCWHAmt)
  TempVoid.SOCWHGL = QPTrim$(SysRec(1).Liab(3).Acct)
  TempVoid.SOCWHGL = ReplaceString(TempVoid.SOCWHGL, "-", "")
  TempVoid.SOCWHGL = AddDashesToGLNumber(TempVoid.SOCWHGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATCRAmt = EDSAmt(1)
  TempVoid.SOCMATCRGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).SOCLIAB))
  TempVoid.SOCMATCRGL = ReplaceString(TempVoid.SOCMATCRGL, "-", "")
  TempVoid.SOCMATCRGL = AddDashesToGLNumber(TempVoid.SOCMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATDBAmt = EDSAmt(1)
  TempVoid.SOCMATDBGL = Mid(EDAct(1), 1, ThisFACnt) + QPTrim$(SysRec(1).SOCEXP)
  TempVoid.SOCMATDBGL = ReplaceString(TempVoid.SOCMATDBGL, "-", "")
  TempVoid.SOCMATDBGL = AddDashesToGLNumber(TempVoid.SOCMATDBGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.RETWHAmt)
  TempVoid.RETWHGL = QPTrim$(SysRec(1).Liab(5).Acct)
  TempVoid.RETWHGL = ReplaceString(TempVoid.RETWHGL, "-", "")
  TempVoid.RETWHGL = AddDashesToGLNumber(TempVoid.RETWHGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATCRAmt = EDRAmt(1)
  TempVoid.RETMATCRGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).RETLIAB))
  TempVoid.RETMATCRGL = ReplaceString(TempVoid.RETMATCRGL, "-", "")
  TempVoid.RETMATCRGL = AddDashesToGLNumber(TempVoid.RETMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATDBAmt = EDRAmt(1)
  TempVoid.RETMATDBGL = Mid(EDAct(1), 1, ThisFACnt) + QPTrim$(SysRec(1).RETEXP)
  TempVoid.RETMATDBGL = ReplaceString(TempVoid.RETMATDBGL, "-", "")
  TempVoid.RETMATDBGL = AddDashesToGLNumber(TempVoid.RETMATDBGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.STAWHAmt)
  TempVoid.STAWHGL = QPTrim$(SysRec(1).Liab(2).Acct)
  TempVoid.STAWHGL = ReplaceString(TempVoid.STAWHGL, "-", "")
  TempVoid.STAWHGL = AddDashesToGLNumber(TempVoid.STAWHGL, FundLen, AcctLen, DetLen)
  
  For x = 1 To 50
    If TransRec(1).DAmt(x) > 0 Then 'all deductions always go into the first fund
      If PoolFundNum$ = Mid(EDAct(1), 1, 2) And PRPoolProcessed = False Then
        TempVoid.DedData(x).DAmt = TransRec(1).DAmt(x)
        TotalDeds = OldRound(TotalDeds + TransRec(1).DAmt(x))
        TotalWHAndDeds = TotalWHAndDeds + TransRec(1).DAmt(x)
      Else
        TempVoid.DedData(x).DAmt = 0
      End If
      TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedCodeDesc(x))
      TempVoid.DedData(x).DedGLNum = QPTrim$(DedCodeNums(x))
      TempVoid.DedData(x).DedGLNum = ReplaceString(TempVoid.DedData(x).DedGLNum, "-", "")
      TempVoid.DedData(x).DedGLNum = AddDashesToGLNumber(TempVoid.DedData(x).DedGLNum, FundLen, AcctLen, DetLen)
      ThisPR = OldRound(ThisPR - TempVoid.DedData(x).DAmt)
    Else
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DedDesc = ""
      TempVoid.DedData(x).DedGLNum = ""
    End If
  Next x
  
  TempVoid.NumOfAccts = ENumOfAct
  If ThisPR < 0 Then
    If Mid(EDAct(1), 1, FundLen) <> Mid(SysRec(1).Liab(3).Acct, 1, FundLen) Then
      DbtCnt = DbtCnt + 1
      ReDim Preserve ThisPRDbtFund(1 To DbtCnt) As String
      ThisPRDbtFund(DbtCnt) = Mid(EDAct(1), 1, FundLen)
      ReDim Preserve ThisPRDbtAmt(1 To DbtCnt) As Double
      ThisPRDbtAmt(DbtCnt) = TempVoid.WagesAmt
      ThisPR = OldRound(TempVoid.FEDWHAmt + TempVoid.MEDWHAmt + TempVoid.SOCWHAmt + TempVoid.RETWHAmt + TempVoid.STAWHAmt + TotalDeds)
      ThisPR = -ThisPR
      TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
      TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
      TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
    Else
      TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
      TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
      TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
    End If
  Else
    TempVoid.PRNetGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
    TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
    TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
  End If
  
  If DbtCnt > 0 Then
    For x = 1 To DbtCnt
      If QPTrim$(ThisPRDbtFund(x)) = Mid(TempVoid.PRNetGL, 1, FundLen) Then
        TempVoid.PRNet = OldRound(ThisPR + ThisPRDbtAmt(x))
        ThisPRDbtAmt(x) = 0
        Exit For
      End If
    Next x
    If x > DbtCnt Then TempVoid.PRNet = ThisPR
  Else
    TempVoid.PRNet = ThisPR
  End If
  
  TempVoid.TransRec = 0
  TempVoid.VoidFlag = False
  TempVoid.CheckAmt = 0
  TempVoid.CheckDate = 0
  TempVoid.CheckNum = 0
  TempVoid.Type = QPTrim$(SysRec(1).USEIMP)
  TempVoid.Pad = ""
  If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
    TempVoid.PPEAmt = ThisPR
    ThisCRGL = ReplaceString(SysRec(1).ICRACCT, "-", "")
    If FundPad > 0 Then
      If Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) <> "" Then
       ThisCRGL = QPTrim$(ThisCRGL) + String$(FundPad, "0") + FundNumOnly
        TempVoid.PPEGL = ThisCRGL
      ElseIf Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) = "" Then
        ThisCRGL = QPTrim$(ThisCRGL + Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen)) + FundNumOnly + String$(FundPad, "0")
        TempVoid.PPEGL = ThisCRGL
      End If
    Else
      TempVoid.PPEGL = QPTrim$(SysRec(1).ICRACCT) + FundNumOnly
    End If
    
    TempVoid.PPEGL = ReplaceString(TempVoid.PPEGL, "-", "")
    TempVoid.PPEGL = AddDashesToGLNumber(TempVoid.PPEGL, FundLen, AcctLen, DetLen)
    
    TempVoid.PPETotAmt = ThisPR
    TempVoid.PPETotGL = QPTrim$(SysRec(1).IDRACCT)
    TempVoid.PPETotGL = ReplaceString(TempVoid.PPETotGL, "-", "")
    TempVoid.PPETotGL = AddDashesToGLNumber(TempVoid.PPETotGL, FundLen, AcctLen, DetLen)
  Else
    TempVoid.PPEAmt = 0
    TempVoid.PPEGL = ""
    TempVoid.PPETotAmt = 0
    TempVoid.PPETotGL = ""
  End If
  TVCnt = TVCnt + 1
  
  Put TVHandle, TVCnt, TempVoid
  If PRNetPoolFound = True Then PRPoolProcessed = True
  
  Return
'--------------------^^^^--Void Check and Fund Summary Code---------
'----------------added 9/20/04---------------
NoPRNetForPool:
  For x = 1 To AcctNumCnt - 1
    Get TVHandle, TVCnt - x, TempVoid
      TempVoid.NumOfAccts = AcctNumCnt + 1 'the program looks at the number of
      'iterations for this employee to know which records to use if this check
      'is ever voided. Each iteration represents data coming through for a 'Paid
      'From' fund. If the pool fund is not a 'Paid From' fund then the number
      'of iterations saved so far will be one short because the code in this
      'sub represents another iteration. So we have to go back to the records
      'already saved for this paycheck and increase the iterations by one.
    Put TVHandle, TVCnt - x, TempVoid
  Next x
  Get TVHandle, TVCnt, TempVoid
  TempVoid.NumOfAccts = AcctNumCnt + 1
  Put TVHandle, TVCnt, TempVoid
  TotalWHAndDeds = 0
  TempVoid.NumOfAccts = AcctNumCnt + 1
  TempVoid.EmpNum = ENumb(1)
  TempVoid.TransRec = 0
  TempVoid.VoidFlag = False
  TempVoid.CheckAmt = 0
  TempVoid.CheckDate = 0
  TempVoid.CheckNum = 0
  TempVoid.PPEAmt = 0
  TempVoid.PPEGL = ""
  TempVoid.PPETotAmt = 0
  TempVoid.PPETotGL = ""
  For x = 1 To 50
    If TransRec(1).DAmt(x) > 0 Then 'all deductions always go into the pool fund
    'even if the pool fund is not a fund from which PRNet is naturally created
      TempVoid.DedData(x).DAmt = TransRec(1).DAmt(x)
      TotalDeds = OldRound(TotalDeds + TransRec(1).DAmt(x))
      TotalWHAndDeds = TotalWHAndDeds + TransRec(1).DAmt(x)
      TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedCodeDesc(x))
      TempVoid.DedData(x).DedGLNum = QPTrim$(DedCodeNums(x))
      TempVoid.DedData(x).DedGLNum = ReplaceString(TempVoid.DedData(x).DedGLNum, "-", "")
      TempVoid.DedData(x).DedGLNum = AddDashesToGLNumber(TempVoid.DedData(x).DedGLNum, FundLen, AcctLen, DetLen)
      ThisPR = OldRound(ThisPR - TempVoid.DedData(x).DAmt)
    Else
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DedDesc = ""
      TempVoid.DedData(x).DedGLNum = ""
    End If
  Next x
  TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
  TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
  TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
  TempVoid.FEDWHAmt = TransRec(1).FedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).FedTaxAmt
  TempVoid.MEDWHAmt = TransRec(1).MedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).MedTaxAmt
  TempVoid.SOCWHAmt = TransRec(1).SocTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).SocTaxAmt
  TempVoid.RETWHAmt = TransRec(1).RetireAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).RetireAmt
  TempVoid.STAWHAmt = TransRec(1).StaTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).StaTaxAmt
  TempVoid.PRNet = -TotalWHAndDeds
  TempVoid.MEDMATCRAmt = 0
  TempVoid.MEDMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).MEDLIAB))
  TempVoid.MEDMATCRGL = ReplaceString(TempVoid.MEDMATCRGL, "-", "")
  TempVoid.MEDMATCRGL = AddDashesToGLNumber(TempVoid.MEDMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATDBAmt = 0
  TempVoid.MEDMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).MEDEXP)
  TempVoid.MEDMATDBGL = ReplaceString(TempVoid.MEDMATDBGL, "-", "")
  TempVoid.MEDMATDBGL = AddDashesToGLNumber(TempVoid.MEDMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATCRAmt = 0
  TempVoid.RETMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).RETLIAB))
  TempVoid.RETMATCRGL = ReplaceString(TempVoid.RETMATCRGL, "-", "")
  TempVoid.RETMATCRGL = AddDashesToGLNumber(TempVoid.RETMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATDBAmt = 0
  TempVoid.RETMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).RETEXP)
  TempVoid.RETMATDBGL = ReplaceString(TempVoid.RETMATDBGL, "-", "")
  TempVoid.RETMATDBGL = AddDashesToGLNumber(TempVoid.RETMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATCRAmt = 0
  TempVoid.SOCMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).SOCLIAB))
  TempVoid.SOCMATCRGL = ReplaceString(TempVoid.SOCMATCRGL, "-", "")
  TempVoid.SOCMATCRGL = AddDashesToGLNumber(TempVoid.SOCMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATDBAmt = 0
  TempVoid.SOCMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).SOCEXP)
  TempVoid.SOCMATDBGL = ReplaceString(TempVoid.SOCMATDBGL, "-", "")
  TempVoid.SOCMATDBGL = AddDashesToGLNumber(TempVoid.SOCMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.WagesAmt = 0
  TempVoid.WagesGL = ""
  TVCnt = TVCnt + 1
  Put TVHandle, TVCnt, TempVoid
  PRNetPoolFound = True
  Return

NoPRNetForPoolCOrI:
  GL4PPETotal = ""
  PRNetSum = 0
  
  For x = 1 To AcctNumCnt - 1 'the program depends on the number of iterations
  'occuring to collect all the paycheck data (1 iteration per 'Pay From' fund.
  'Because there is no pool fund iteration we must go back and adjust the number
  'of iterations up 1 to include the rest of the code in this sub.
    Get TVHandle, TVCnt - x, TempVoid
      PRNetSum = PRNetSum + TempVoid.PRNet 'we'll be using this number below
      'to determine the PRNet for the pool fund
      TempVoid.NumOfAccts = AcctNumCnt + 1
    Put TVHandle, TVCnt - x, TempVoid
  Next x
  Get TVHandle, TVCnt, TempVoid
  PRNetSum = PRNetSum + TempVoid.PRNet
  GL4PPETotal = TempVoid.PPETotGL 'this GL number is constant and therefore
  'works fine for this iteration
  TempVoid.NumOfAccts = AcctNumCnt + 1
  Put TVHandle, TVCnt, TempVoid
  ThisCRGL = ReplaceString(SysRec(1).ICRACCT, "-", "")
  If FundPad > 0 Then
    If Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) <> "" Then
     ThisCRGL = QPTrim$(ThisCRGL) + String$(FundPad, "0") + PoolFundNum$
      TempVoid.PPEGL = ThisCRGL
    ElseIf Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) = "" Then
      ThisCRGL = QPTrim$(ThisCRGL + Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen)) + PoolFundNum$ + String$(FundPad, "0")
      TempVoid.PPEGL = ThisCRGL
    End If
  Else
    TempVoid.PPEGL = QPTrim$(SysRec(1).ICRACCT) + PoolFundNum$
  End If
  TempVoid.PPEGL = ReplaceString(TempVoid.PPEGL, "-", "")
  TempVoid.PPEGL = AddDashesToGLNumber(TempVoid.PPEGL, FundLen, AcctLen, DetLen)
  TotalWHAndDeds = 0
  TempVoid.NumOfAccts = AcctNumCnt + 1
  TempVoid.EmpNum = ENumb(1)
  TempVoid.TransRec = 0
  TempVoid.VoidFlag = False
  TempVoid.CheckAmt = 0
  TempVoid.CheckDate = 0
  TempVoid.CheckNum = 0
  For x = 1 To 50
    If TransRec(1).DAmt(x) > 0 Then 'all deductions always go into the pool fund
    'even if the pool fund is not a fund from which PRNet is naturally created
      TempVoid.DedData(x).DAmt = TransRec(1).DAmt(x)
      TotalDeds = OldRound(TotalDeds + TransRec(1).DAmt(x))
      TotalWHAndDeds = TotalWHAndDeds + TransRec(1).DAmt(x)
      TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedCodeDesc(x))
      TempVoid.DedData(x).DedGLNum = QPTrim$(DedCodeNums(x))
      TempVoid.DedData(x).DedGLNum = ReplaceString(TempVoid.DedData(x).DedGLNum, "-", "")
      TempVoid.DedData(x).DedGLNum = AddDashesToGLNumber(TempVoid.DedData(x).DedGLNum, FundLen, AcctLen, DetLen)
      ThisPR = OldRound(ThisPR - TempVoid.DedData(x).DAmt)
    Else
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DedDesc = ""
      TempVoid.DedData(x).DedGLNum = ""
    End If
  Next x
  TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
  TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
  TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
  TempVoid.FEDWHAmt = TransRec(1).FedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).FedTaxAmt
  TempVoid.MEDWHAmt = TransRec(1).MedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).MedTaxAmt
  TempVoid.SOCWHAmt = TransRec(1).SocTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).SocTaxAmt
  TempVoid.RETWHAmt = TransRec(1).RetireAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).RetireAmt
  TempVoid.STAWHAmt = TransRec(1).StaTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).StaTaxAmt
  TempVoid.PRNet = -TotalWHAndDeds
  TempVoid.PPEAmt = -TotalWHAndDeds
  TempVoid.PPETotAmt = -TotalWHAndDeds
  TempVoid.PPETotGL = GL4PPETotal
  TempVoid.MEDMATCRAmt = 0
  TempVoid.MEDMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).MEDLIAB))
  TempVoid.MEDMATCRGL = ReplaceString(TempVoid.MEDMATCRGL, "-", "")
  TempVoid.MEDMATCRGL = AddDashesToGLNumber(TempVoid.MEDMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATDBAmt = 0
  TempVoid.MEDMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).MEDEXP)
  TempVoid.MEDMATDBGL = ReplaceString(TempVoid.MEDMATDBGL, "-", "")
  TempVoid.MEDMATDBGL = AddDashesToGLNumber(TempVoid.MEDMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATCRAmt = 0
  TempVoid.RETMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).RETLIAB))
  TempVoid.RETMATCRGL = ReplaceString(TempVoid.RETMATCRGL, "-", "")
  TempVoid.RETMATCRGL = AddDashesToGLNumber(TempVoid.RETMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATDBAmt = 0
  TempVoid.RETMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).RETEXP)
  TempVoid.RETMATDBGL = ReplaceString(TempVoid.RETMATDBGL, "-", "")
  TempVoid.RETMATDBGL = AddDashesToGLNumber(TempVoid.RETMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATCRAmt = 0
  TempVoid.SOCMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).SOCLIAB))
  TempVoid.SOCMATCRGL = ReplaceString(TempVoid.SOCMATCRGL, "-", "")
  TempVoid.SOCMATCRGL = AddDashesToGLNumber(TempVoid.SOCMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATDBAmt = 0
  TempVoid.SOCMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).SOCEXP)
  TempVoid.SOCMATDBGL = ReplaceString(TempVoid.SOCMATDBGL, "-", "")
  TempVoid.SOCMATDBGL = AddDashesToGLNumber(TempVoid.SOCMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.WagesAmt = 0
  TempVoid.WagesGL = ""
  TVCnt = TVCnt + 1
  Put TVHandle, TVCnt, TempVoid
  PRNetPoolFound = True 'this tells the program not to come back to this sub until the next
  'time the pool fund number is not a 'Paid From' fund

  Return
'----^^^^^^^^^^^^^^added 9/20/2004 to allow for employees not paid out of the pool fund


PrintDistTotal:
  RegHrs# = 0
  TotHrs# = 0
  RegWage# = 0
  OTWage# = 0
  AddEarn# = 0
  DGPay# = 0
  
  SortD DistbSumAccts(), NumOfWageAccts
  
  'print the Summary of ALL Distrubution Accounts
  
  For cnt = 1 To NumOfWageAccts
    If GPay# <= 0 Then Return
    LSet EDAct(1) = DistbSumAccts(cnt).Acct
    RSet EDPct(1) = Using(Image4$, (DistbSumAccts(cnt).GrossPay / GPay#) * 100)
    RSet RHrs(1) = Using(Image3$, DistbSumAccts(cnt).RHrs)
    RegHrs# = OldRound(RegHrs# + DistbSumAccts(cnt).RHrs)
    RSet OTHrs(1) = Using(Image3$, DistbSumAccts(cnt).OHrs)
    TotHrs# = OldRound(TotHrs# + DistbSumAccts(cnt).OHrs)
    RSet EDRPay(1) = Using(Image3$, DistbSumAccts(cnt).RWage)
    RegWage# = OldRound(RegWage# + DistbSumAccts(cnt).RWage)
    RSet EDOPay(1) = Using(Image3$, DistbSumAccts(cnt).OWage)
    OTWage# = OldRound(OTWage# + DistbSumAccts(cnt).OWage)
    RSet EDEarn(1) = Using(Image3$, DistbSumAccts(cnt).AddEarn)
    AddEarn# = OldRound(AddEarn# + DistbSumAccts(cnt).AddEarn)
    RSet EDGroP(1) = Using(Image3$, DistbSumAccts(cnt).GrossPay)
    DGPay# = OldRound(DGPay# + DistbSumAccts(cnt).GrossPay)
    
    RSet EDSAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATSocAmt)
    ASAmt# = OldRound(ASAmt# + DistbSumAccts(cnt).MATSocAmt)
    
    RSet EDMAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATMedAmt)
    AMAmt# = OldRound(AMAmt# + DistbSumAccts(cnt).MATMedAmt)
    
    RSet EDRAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATRetAmt)
    ARAmt# = OldRound(ARAmt# + DistbSumAccts(cnt).MATRetAmt)
    '                   0              1              2             3
    Print #DTHandle, EDAct(1); dlm; EDPct(1); dlm; RHrs(1); dlm; OTHrs(1); dlm;
    '                   4              5              6                  7
    Print #DTHandle, EDRPay(1); dlm; EDOPay(1); dlm; EDEarn(1); dlm; EDGroP(1); dlm;
    '                   8              9               10
    Print #DTHandle, EDSAmt(1); dlm; EDMAmt(1); dlm; EDRAmt(1); dlm; QPTrim$(Unit(1).UFEMPR)
    
  Next
  
  Return
  
ErrorHandler:
  Close
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub
Private Sub SortT(AllFunds() As FundType, NumOfWageAccts As Integer)

  Dim TempDAccts As FundType
  Dim Temp As String
  Dim Largest As Double
  Dim Smallest As Double
  Dim thisNum As String
  Dim x As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  ReDim OrderedList(1 To NumOfWageAccts) As FundType
  ReDim IdxDAccts(1 To NumOfWageAccts) As Integer
  
  Largest = 0
  For x = 1 To NumOfWageAccts
    thisNum = QPTrim$(ReplaceString(AllFunds(x).FundCode, "-", ""))
    If Val(thisNum) > Largest Then
      Largest = Val(thisNum)
    End If
  Next x
  
  Smallest = Largest + 1
  Nextx = 1
  Thisx = Nextx
  Do
    For x = Nextx To NumOfWageAccts
    thisNum = QPTrim$(ReplaceString(AllFunds(x).FundCode, "-", ""))
    If Val(thisNum) < Smallest Then
      Smallest = Val(thisNum)
      Thisx = x
    End If
    Next x
    OrderedList(Nextx) = AllFunds(Thisx)
    TempDAccts = AllFunds(Nextx)
    AllFunds(Nextx) = AllFunds(Thisx)
    AllFunds(Thisx) = TempDAccts
    Nextx = Nextx + 1
    Smallest = Largest + 1
  Loop Until Nextx = NumOfWageAccts + 1

End Sub
Private Sub SWAP(This As GLIFDataType14, ForThis As GLIFDataType14)
  Dim TempThis As GLIFDataType14
  TempThis = This
  This = ForThis
  ForThis = TempThis
End Sub

Private Sub SortD(DistbSumAccts() As DistWageRptType, NumOfWageAccts As Integer)

  Dim TempDAccts As DistWageRptType
  Dim Temp As String
  Dim Largest As Double
  Dim Smallest As Double
  Dim thisNum As String
  Dim x As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  If NumOfWageAccts = 0 Then
    Exit Sub
  End If

  ReDim OrderedList(1 To NumOfWageAccts) As DistWageRptType
  ReDim IdxDAccts(1 To NumOfWageAccts) As Integer
  
  Largest = 0
  For x = 1 To NumOfWageAccts
    thisNum = ReplaceString(DistbSumAccts(x).Acct, "-", "")
    If Val(thisNum) > Largest Then
      Largest = Val(thisNum)
    End If
  Next x
  
  Smallest = Largest + 1
  Nextx = 1
  Thisx = Nextx
  Do
    For x = Nextx To NumOfWageAccts
    thisNum = ReplaceString(DistbSumAccts(x).Acct, "-", "")
    If Val(thisNum) < Smallest Then
      Smallest = Val(thisNum)
      Thisx = x
    End If
    Next x
    OrderedList(Nextx) = DistbSumAccts(Thisx)
    TempDAccts = DistbSumAccts(Nextx)
    DistbSumAccts(Nextx) = DistbSumAccts(Thisx)
    DistbSumAccts(Thisx) = TempDAccts
    Nextx = Nextx + 1
    Smallest = Largest + 1
  Loop Until Nextx = NumOfWageAccts + 1
  
End Sub

Sub MakeGLIFFileT(TotEIC#, TotDeds#(), Passed#(), DistbSumAccts() As DistWageRptType)

  ReDim SysRec(1) As RegDSysFileRecType
  ReDim PDR(1) As PeriodDefaultRecType
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  ReDim GLSetupRec(1) As GLSetupRecType
  Dim DedCodeFileName As Integer
  Dim PPDefaultFileName As Integer
  Dim SysFileName As Integer
  Dim x As Integer
  Dim FACnt As Integer
  Dim FundPad As Integer
  Dim FundLen As Integer
  Dim GLSetUpName$, GHandle As Integer
  Dim GLSetUpRecLen As Integer
  Dim NumOfWageAccts As Integer
  Dim GLIFTDate$, GLIFSource$
  Dim cnt As Integer, NextAcct As Integer
  Dim SysCash$, NumDFunds As Integer
  Dim CurrFund$, ThisFund As Integer
  Dim FirstC As Integer, TotalFunds As Integer
  Dim LastC As Integer, NoCFunds As Boolean
  Dim NumCFunds As Integer, First As Integer
  Dim Start As Integer, Last As Integer
  Dim Cnt2 As Integer, TotalGLIFS
  Dim AcctNum$, TempAcct$
  Dim FringeAcct$, FringeRate#, RecNo&
  Dim INDFund$, IndirectAcct$, IndirectRate#
  Dim Indirect#, Fringe#, SOCEXP$, RETLIAB$
  Dim MEDEXP$, RETEXP$, SOCLIAB$, MEDLIAB$
  Dim GLIFRecName$, FundCash$
  Dim GLIFRecLen As Integer, GLHandle As Integer
  Dim NumOfDeds As Integer
  
'  On Error GoTo ERRORSTUFF

  GLDebitTotal = 0
  GLCreditTotal = 0
  GLError = 0
  DistbSumAccts(1).Acct = DistbSumAccts(1).Acct
  OpenDedCodeFile DedCodeFileName
  
  For x = 1 To 50
    Get DedCodeFileName, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
      NumOfDeds = NumOfDeds + 1
    End If
  Next x
  Close DedCodeFileName
  
  OpenPPDefaultFile PPDefaultFileName
  Get PPDefaultFileName, 1, PDR(1)
  Close PPDefaultFileName
  
  OpenSysFile SysFileName
  Get SysFileName, 1, SysRec(1)
  Close SysFileName
  
  FACnt = SysRec(1).AcctCnt
  DistbSumAccts(1).Acct = DistbSumAccts(1).Acct

'for new gl
  
  FundPad = 0
  FundLen = 2     'Default fund length
  
'  GLSetUpName$ = QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT"
  If Mid(CurrCitiPath, Len(CurrCitiPath), 1) <> "\" Then
    GLSetUpName$ = CurrCitiPath + "\GLSETUP.DAT"
  ElseIf Mid(CurrCitiPath, Len(CurrCitiPath), 1) = "\" Then
    GLSetUpName$ = CurrCitiPath + "GLSETUP.DAT"
  End If
    
  GLSetUpRecLen = Len(GLSetupRec(1))
  GLHandle = FreeFile
  
  If Exist(GLSetUpName$) Then
    Open GLSetUpName$ For Random Shared As GLHandle Len = GLSetUpRecLen
    Get GLHandle, 1, GLSetupRec(1)
    FundLen = GLSetupRec(1).FundLen
    FundPad = GLSetupRec(1).DetLen - GLSetupRec(1).FundLen
  End If
  Close GLHandle

  NumOfWageAccts = UBound(DistbSumAccts)
  If NumOfWageAccts = 0 Then
    MsgBox "No active transactions are pending at this time"
    Exit Sub
  End If
  ReDim Preserve DistbSumAccts(1 To NumOfWageAccts) As DistWageRptType
  'squeeze out all the "-" out of Acct numbers
  For cnt = 1 To NumOfWageAccts
    ReplaceString DistbSumAccts(cnt).Acct, "-", ""
  Next
  
  'change Period ending date to nicks format
  
  GLIFTDate$ = MakeRegDate(PDR(1).PEREND)
  
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
  ReplaceString GLIFTDate$, "2004", "04"
  ReplaceString GLIFTDate$, "2005", "05"
  ReplaceString GLIFTDate$, "2006", "06"
  ReplaceString GLIFTDate$, "2007", "07"
  ReplaceString GLIFTDate$, "2008", "08"
  ReplaceString GLIFTDate$, "2009", "09"
  '
  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$
  
  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
'  ReDim GLIFRec(1 To NumOfWageAccts + 38) As GLIFDataType14
  ReDim GLIFRec(1 To (NumOfWageAccts + 17 + NumOfDeds + 1)) As GLIFDataType14
  '1st for loop as marked below = NumOfWageAccts + 18 is LastC
  '2nd for loop as marked below = 5
  '3rd for loop as marked below = NumOfDeds
  
  For cnt = 1 To NumOfWageAccts '1st for loop
    GLIFRec(cnt).TranAcct = DistbSumAccts(cnt).Acct
    GLIFRec(cnt).TranDate = GLIFTDate$
    GLIFRec(cnt).TranDesc = "Wages"
    GLIFRec(cnt).CrAmt = 0
    GLIFRec(cnt).DrAmt = DistbSumAccts(cnt).GrossPay
    GLIFRec(cnt).Source = GLIFSource$
    GLIFRec(cnt).FromFlag = "W"
  Next
  
  NextAcct = cnt
  '
  For cnt = 0 To 4 '2nd for loop
    ReplaceString SysRec(1).Liab(cnt + 1).Acct, "-", ""
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SysRec(1).Liab(cnt + 1).Acct)
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = "Withholdings"
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
    'get tax and ret account numbers
  Next
  GLIFRec(NextAcct).CrAmt = Passed#(1)          'federal
  GLIFRec(NextAcct).TranDesc = "Fed Withholdings"
  GLIFRec(NextAcct + 1).CrAmt = Passed#(2)      'state
  GLIFRec(NextAcct + 1).TranDesc = "State Withholdings"
  GLIFRec(NextAcct + 2).CrAmt = Passed#(3)      'social sec
  GLIFRec(NextAcct + 2).TranDesc = "Soc Sec Withholdings"
  GLIFRec(NextAcct + 3).CrAmt = Passed#(4)      'Medicare
  GLIFRec(NextAcct + 3).TranDesc = "Med Withholdings"
  GLIFRec(NextAcct + 4).CrAmt = Passed#(5)      'Retirement total
  GLIFRec(NextAcct + 4).TranDesc = "Ret Withholdings"
  'good to here ;maybe
  
  ReplaceString SysRec(1).CashAcct, "-", ""
  
  SysCash$ = QPTrim$(SysRec(1).CashAcct)
  
  If TotEIC# > 0 Then
    ReDim EICGLIFRec(1 To 2) As GLIFDataType14
    EICGLIFRec(1).TranAcct = QPTrim$(SysRec(1).Liab(1).Acct)
    EICGLIFRec(1).TranDate = GLIFTDate$
    EICGLIFRec(1).TranDesc = "EIC Pmt"
    EICGLIFRec(1).CrAmt = 0
    EICGLIFRec(1).DrAmt = TotEIC#
    EICGLIFRec(1).Source = GLIFSource$
    EICGLIFRec(1).FromFlag = "E"
    '
    EICGLIFRec(2).TranAcct = Left$(QPTrim$(SysRec(1).Liab(1).Acct), FundLen) + SysCash$
    EICGLIFRec(2).TranDate = GLIFTDate$
    EICGLIFRec(2).TranDesc = "EIC Pmt"
    EICGLIFRec(2).CrAmt = TotEIC#
    EICGLIFRec(2).DrAmt = 0
    EICGLIFRec(2).Source = GLIFSource$
    EICGLIFRec(2).FromFlag = "P"
  End If

  NextAcct = NextAcct + cnt
  '
  For cnt = 0 To NumOfDeds - 1 '3rd for loop
    ReplaceString DedCodes(cnt + 1).DCACCT1, "-", ""
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(DedCodes(cnt + 1).DCACCT1)
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = QPTrim$(DedCodes(cnt + 1).DCDESC1)  '"Deductions"
    GLIFRec(NextAcct + cnt).CrAmt = TotDeds#(cnt + 1)
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "D"
  Next
  
  
  ReDim DFunds$(1 To NumOfWageAccts)
  NumDFunds = 1
  'fixed
  DFunds$(1) = Left$(DistbSumAccts(1).Acct, FundLen)
  For cnt = 1 To NumOfWageAccts - 1
    'fixed
    If Left$(DistbSumAccts(cnt).Acct, FundLen) <> Left$(DistbSumAccts(cnt + 1).Acct, FundLen) Then
      NumDFunds = NumDFunds + 1 'counting the total number of funds.
      'fixed
      DFunds$(NumDFunds) = Left$(DistbSumAccts(cnt + 1).Acct, FundLen)
    End If
  Next
  
  ReDim Preserve DFunds$(1 To NumDFunds)
  ReDim DFund(1 To NumDFunds) As FundType
  For cnt = 1 To NumOfWageAccts
    'fixed
    CurrFund$ = Left$(DistbSumAccts(cnt).Acct, FundLen)
    For ThisFund = 1 To NumDFunds
      If CurrFund$ = DFunds$(ThisFund) Then
        DFund(ThisFund).FundCode = DFunds$(ThisFund)
        DFund(ThisFund).Debit = OldRound(DFund(ThisFund).Debit + DistbSumAccts(cnt).GrossPay)
        Exit For
      End If
    Next
  Next
  
  'all gross pay by funds here!!
  'make funds and sumarize ded and taxs here
  '
'  ReDim CFunds$(1 To 17)
  ReDim CFunds$(1 To NumOfWageAccts + 5 + NumOfDeds)
  '
  FirstC = NumOfWageAccts + 1
  'LastC = the number GLIFRec used to redim
  LastC = NumOfWageAccts + 17 + 1
  LastC = NumOfWageAccts + 5 + NumOfDeds + 1 'changed 04/15/2004 to match MakeGLIFFileG
  'NumOfWageAccts + 5 + NumOfDeds
  
  NumCFunds = 1
  'fixed
  CFunds$(1) = Left$(GLIFRec(FirstC).TranAcct, FundLen)
  
  For cnt = FirstC To LastC - 1
    'fixed
    If Left$(GLIFRec(cnt).TranAcct, FundLen) <> Left$(GLIFRec(cnt + 1).TranAcct, FundLen) Then
      If Len(QPTrim$(GLIFRec(cnt + 1).TranAcct)) Then
        NumCFunds = NumCFunds + 1               'counting the total number of funds.
        'fixed
        CFunds$(NumCFunds) = Left$(GLIFRec(cnt + 1).TranAcct, FundLen)
      End If
    End If
  Next

  ReDim Preserve CFunds$(1 To NumCFunds)
  ReDim CFund(1 To NumCFunds) As FundType
  For cnt = FirstC To LastC - 1
    'fixed
    CurrFund$ = Left$(GLIFRec(cnt).TranAcct, FundLen)
    For ThisFund = 1 To NumCFunds
      If CurrFund$ = CFunds$(ThisFund) Then
        CFund(ThisFund).FundCode = CFunds$(ThisFund)
        CFund(ThisFund).Credit = OldRound(CFund(ThisFund).Credit + GLIFRec(cnt).CrAmt)
        Exit For
      End If
    Next
  Next

  'combine all funds in one array here
  TotalFunds = NumDFunds + NumCFunds            '+ 1
  ReDim AllFunds(1 To TotalFunds) As FundType
  ThisFund = 1
  For cnt = 1 To NumDFunds
    AllFunds(ThisFund) = DFund(cnt)
    ThisFund = ThisFund + 1
  Next
  '
  If NoCFunds = False Then
    For cnt = 1 To NumCFunds
      AllFunds(ThisFund) = CFund(cnt)
      ThisFund = ThisFund + 1
    Next
  End If

'fixed: 05-27-96
  SortT AllFunds(), TotalFunds
                                                        
  'combine Debits and Credits for same fund numbers
  First = 1
  Start = 1
  Last = TotalFunds
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If AllFunds(First).FundCode = AllFunds(Cnt2).FundCode Then
        LSet AllFunds(Cnt2).FundCode = ""       'blank acct num as a flag
        AllFunds(First).Credit = OldRound(AllFunds(First).Credit + AllFunds(Cnt2).Credit)
        AllFunds(Cnt2).Credit = 0
        AllFunds(First).Debit = OldRound(AllFunds(First).Debit + AllFunds(Cnt2).Debit)
        AllFunds(Cnt2).Debit = 0
      End If
    Next
    First = First + 1
  Loop Until First >= Last      '

  'calc net difference for funds
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(AllFunds(cnt).FundCode)) Then
      AllFunds(cnt).Net = OldRound(AllFunds(cnt).Debit - AllFunds(cnt).Credit)
    End If
  Next

  'add cash acct number to fund numbers
  For cnt = 1 To TotalFunds
    FundCash$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(FundCash$) Then
      LSet AllFunds(cnt).FundCode = FundCash$ + SysCash$
    End If
  Next

'  TotalGLIFS = NumOfWageAccts + 39 + TotalFunds
  TotalGLIFS = NumOfWageAccts + TotalFunds + NumOfDeds + 6 'added 8/12/04 to resolve
  'problem of not all deductions being printed  because this array as too small

  ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14

'  NextAcct = NumOfWageAccts + 39
  NextAcct = NumOfWageAccts + NumOfDeds + 6 'added 8/12/04 to resolve problem of this
  'array being too small to hold all deductions up to 50

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
    '
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
        'fixed
        CDGLIFRec(cnt).TranAcct = QPTrim$(SysRec(1).ICRACCT) + Left$(AllFunds(cnt).FundCode, FundLen)
        If FundPad > 0 Then
          TempAcct$ = QPTrim$(CDGLIFRec(cnt).TranAcct)
          TempAcct$ = TempAcct$ + String$(FundPad, "0")
          CDGLIFRec(cnt).TranAcct = TempAcct$
        End If
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
  
  If SysRec(1).EXPMETHD = "2" Then GoTo Type2Meth 'jump to type 2 here
  
  'calc and add Fringe GLIF recs
  ReDim FGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  
  ReplaceString SysRec(1).FRNGEXP, "-", ""
  FringeAcct$ = QPTrim$(SysRec(1).FRNGEXP)
  FringeRate# = SysRec(1).FRNGRATE

  For cnt = 1 To NumOfWageAccts
    AcctNum$ = Left$(GLIFRec(cnt).TranAcct, FACnt)
    FGLIFRec(cnt).TranAcct = AcctNum$ + FringeAcct$
    FGLIFRec(cnt).DrAmt = OldRound(GLIFRec(cnt).DrAmt * (FringeRate# * 0.01))
    FGLIFRec(cnt).TranDate = GLIFTDate$
    FGLIFRec(cnt).TranDesc = "FRINGE " + GLIFTDate$
    FGLIFRec(cnt).Source = GLIFSource$
    FGLIFRec(cnt).FromFlag = "F"
  Next

  'calc and add Indirect GLIF recs
  ReDim IGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReplaceString SysRec(1).INDDR, "-", ""
  'fixed
  INDFund$ = QPTrim$(Left$(SysRec(1).INDDR, FundLen))
  IndirectAcct$ = QPTrim$(SysRec(1).INDEXP)
  IndirectRate# = SysRec(1).INDRATE
  
  If IndirectRate# < 0 Then IndirectRate# = 0
  
  For cnt = 1 To NumOfWageAccts
    'look for acct that don't get indirect
    'fixed
    If Not QPTrim$(Left$(GLIFRec(cnt).TranAcct, FundLen)) = INDFund$ Then
      AcctNum$ = Left$(GLIFRec(cnt).TranAcct, FACnt)
      IGLIFRec(cnt).TranAcct = AcctNum$ + IndirectAcct$
      IGLIFRec(cnt).DrAmt = OldRound((GLIFRec(cnt).DrAmt + FGLIFRec(cnt).DrAmt) * (IndirectRate#) * 0.01)
      IGLIFRec(cnt).TranDate = GLIFTDate$
      IGLIFRec(cnt).TranDesc = "INDIRECT " + GLIFTDate$
      IGLIFRec(cnt).Source = GLIFSource$
      IGLIFRec(cnt).FromFlag = "I"
    End If
  Next

  ReDim IFGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  For cnt = 1 To NumOfWageAccts
    'fixed
    AcctNum$ = Left$(FGLIFRec(cnt).TranAcct, FundLen)
    IFGLIFRec(cnt).TranAcct = AcctNum$ + SysCash$
    IFGLIFRec(cnt).CrAmt = OldRound(IGLIFRec(cnt).DrAmt + FGLIFRec(cnt).DrAmt)
    IFGLIFRec(cnt).TranDate = GLIFTDate$
    IFGLIFRec(cnt).TranDesc = "F&I COST " + GLIFTDate$
    IFGLIFRec(cnt).Source = GLIFSource$
    IFGLIFRec(cnt).FromFlag = "A"
  Next

  For cnt = 1 To NumOfWageAccts
    Indirect# = OldRound(Indirect# + IGLIFRec(cnt).DrAmt)
    Fringe# = OldRound(Fringe# + FGLIFRec(cnt).DrAmt)
  Next

  ReDim AGLIFRec(1 To 4) As GLIFDataType14

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

  For cnt = 1 To 4
    AGLIFRec(cnt).TranDate = GLIFTDate$
    AGLIFRec(cnt).TranDesc = "PPE " + GLIFTDate$
    AGLIFRec(cnt).Source = GLIFSource$
    AGLIFRec(cnt).FromFlag = "T"
  Next
  GoTo WriteGLIFS
  '**********END TYPE 1 SECTION

Type2Meth:

  ReDim SocGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim MedGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim RetGLIFRec(1 To NumOfWageAccts) As GLIFDataType14

  SOCEXP$ = QPTrim$(SysRec(1).SOCEXP)
  MEDEXP$ = QPTrim$(SysRec(1).MEDEXP)
  RETEXP$ = QPTrim$(SysRec(1).RETEXP)

  SOCLIAB$ = QPTrim$(SysRec(1).SOCLIAB)
  MEDLIAB$ = QPTrim$(SysRec(1).MEDLIAB)
  RETLIAB$ = QPTrim$(SysRec(1).RETLIAB)

  ReplaceString SOCLIAB$, "-", ""
  ReplaceString MEDLIAB$, "-", ""
  ReplaceString RETLIAB$, "-", ""

  For cnt = 1 To NumOfWageAccts
    'social
    SocGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + SOCEXP$
    SocGLIFRec(cnt).TranDate = GLIFTDate$
    SocGLIFRec(cnt).TranDesc = "Soc Match"
    SocGLIFRec(cnt).Source = GLIFSource$
    SocGLIFRec(cnt).FromFlag = "S"
    SocGLIFRec(cnt).CrAmt = 0
    SocGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATSocAmt
    'medicare
    MedGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + MEDEXP$
    MedGLIFRec(cnt).TranDate = GLIFTDate$
    MedGLIFRec(cnt).TranDesc = "Med Match"
    MedGLIFRec(cnt).Source = GLIFSource$
    MedGLIFRec(cnt).FromFlag = "M"
    MedGLIFRec(cnt).CrAmt = 0
    MedGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATMedAmt
    'retirment
    RetGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + RETEXP$
    RetGLIFRec(cnt).TranDate = GLIFTDate$
    RetGLIFRec(cnt).TranDesc = "Ret Match"
    RetGLIFRec(cnt).Source = GLIFSource$
    RetGLIFRec(cnt).FromFlag = "R"
    RetGLIFRec(cnt).CrAmt = 0
    RetGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATRetAmt
  Next

  ReDim SFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim MFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim RFGLIFRec(1 To TotalFunds) As GLIFDataType14

  For cnt = 1 To TotalFunds
    'fixed
    SFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, FundLen) + SOCLIAB$
    MFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, FundLen) + MEDLIAB$
    RFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, FundLen) + RETLIAB$
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      'fixed
      If Left$(SFGLIFRec(Cnt2).TranAcct, FundLen) = Left$(SocGLIFRec(cnt).TranAcct, FundLen) Then
        SFGLIFRec(Cnt2).CrAmt = OldRound(SFGLIFRec(Cnt2).CrAmt + SocGLIFRec(cnt).DrAmt)
        SFGLIFRec(Cnt2).DrAmt = 0
        SFGLIFRec(Cnt2).TranDate = GLIFTDate$
        SFGLIFRec(Cnt2).TranDesc = "Soc Match Liab"
        SFGLIFRec(Cnt2).Source = GLIFSource$
        SFGLIFRec(Cnt2).FromFlag = "s"
      End If
    Next
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      'fixed
      If Left$(MFGLIFRec(Cnt2).TranAcct, FundLen) = Left$(MedGLIFRec(cnt).TranAcct, FundLen) Then
        MFGLIFRec(Cnt2).CrAmt = OldRound(MFGLIFRec(Cnt2).CrAmt + MedGLIFRec(cnt).DrAmt)
        MFGLIFRec(Cnt2).DrAmt = 0
        MFGLIFRec(Cnt2).TranDate = GLIFTDate$
        MFGLIFRec(Cnt2).TranDesc = "Med Match Liab"
        MFGLIFRec(Cnt2).Source = GLIFSource$
        MFGLIFRec(Cnt2).FromFlag = "m"
      End If
    Next
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      'fixed
      If Left$(RFGLIFRec(Cnt2).TranAcct, FundLen) = Left$(RetGLIFRec(cnt).TranAcct, FundLen) Then
        RFGLIFRec(Cnt2).CrAmt = OldRound(RFGLIFRec(Cnt2).CrAmt + RetGLIFRec(cnt).DrAmt)
        RFGLIFRec(Cnt2).DrAmt = 0
        RFGLIFRec(Cnt2).TranDate = GLIFTDate$
        RFGLIFRec(Cnt2).TranDesc = "Ret Match Liab"
        RFGLIFRec(Cnt2).Source = GLIFSource$
        RFGLIFRec(Cnt2).FromFlag = "r"
      End If
    Next
  Next

WriteGLIFS:
  
  GLIFRecLen = Len(GLIFRec(1))
  GLIFRecName$ = "TempIF.DAT"
  KillFile "TempIF.DAT"
  GHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GHandle Len = GLIFRecLen
  RecNo& = 1
  
  For cnt = 1 To TotalGLIFS
    If Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
      If GLIFRec(cnt).DrAmt > 0 Or GLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, GLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next
  

  Select Case SysRec(1).EXPMETHD
  Case "1"
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(FGLIFRec(cnt).TranAcct)) Then
        If FGLIFRec(cnt).DrAmt > 0 Or FGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, FGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(IGLIFRec(cnt).TranAcct)) Then
        If IGLIFRec(cnt).DrAmt > 0 Or IGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, IGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(IFGLIFRec(cnt).TranAcct)) Then
        If IFGLIFRec(cnt).DrAmt > 0 Or IFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, IFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    
    For cnt = 1 To 4
      If Len(QPTrim$(AGLIFRec(cnt).TranAcct)) > 0 Then
        Put GHandle, RecNo&, AGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    Next
    
  Case "2"
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(SocGLIFRec(cnt).TranAcct)) Then
        If SocGLIFRec(cnt).DrAmt > 0 Or SocGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, SocGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(MedGLIFRec(cnt).TranAcct)) Then
        If MedGLIFRec(cnt).DrAmt > 0 Or MedGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, MedGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To NumOfWageAccts
      If Len(QPTrim$(RetGLIFRec(cnt).TranAcct)) Then
        If RetGLIFRec(cnt).DrAmt > 0 Or RetGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, RetGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next

    For cnt = 1 To TotalFunds
      If Len(QPTrim$(SFGLIFRec(cnt).TranAcct)) Then
        If SFGLIFRec(cnt).DrAmt > 0 Or SFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, SFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(MFGLIFRec(cnt).TranAcct)) Then
        If MFGLIFRec(cnt).DrAmt > 0 Or MFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, MFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(RFGLIFRec(cnt).TranAcct)) Then
        If RFGLIFRec(cnt).DrAmt > 0 Or RFGLIFRec(cnt).CrAmt > 0 Then
          Put GHandle, RecNo&, RFGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
  End Select
  
  If SysRec(1).USEIMP = "C" Then
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(CDGLIFRec(cnt).TranAcct)) Then
        If CDGLIFRec(cnt).DrAmt <> 0 Or CDGLIFRec(cnt).CrAmt <> 0 Then
          Put GHandle, RecNo&, CDGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
  End If

  'added EIC GLIF records if present 6/07/94
  If TotEIC# > 0 Then
    For cnt = 1 To 2
      Put GHandle, RecNo&, EICGLIFRec(cnt)
      RecNo& = RecNo& + 1
    Next
  End If

  Close GHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "modNonSplit", "MakeGLIFFileT", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
End Sub

Sub PCPrintPayRegisterT(PathCode As Integer)
  
  Dim RptTitle$, PPDefaultFileName As Integer
  Dim FileHandle As Integer, x As Integer
  Dim DedCodeFileName As Integer
  Dim ErnCodeFileName As Integer
  Dim SysFileName As Integer, ASAmt#
  Dim FundPad As Integer, TOTPaid#
  Dim FundLen As Integer, TOTComp#
  Dim GLSetUpName$, GHandle As Integer
  Dim GLSetUpRecLen As Integer
  Dim GFedGross#, GStaGross#, GMedGross#
  Dim GSocGross#, GRetGross#, GTaxFrn#
  Dim TotDebit#, TotCredit#, EmpActLen
  Dim DistbSumAcctsLen As Integer, ARAmt#
  Dim TransRecLen As Integer, IdxRecLen As Integer
  Dim Emp1RecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Integer, SalCnt As Integer
  Dim HrlCnt As Integer, DLineCnt As Integer
  Dim LineCnt As Integer, NumOfWageAccts As Integer
  Dim MaxLines As Integer, DMaxLines As Integer
  Dim EPage As Integer, Page As Integer
  Dim EmpIdxNNameHandle As Integer, RptName$
  Dim DTitle$(1 To 5), cnt As Integer, TDed$, LastDed As Integer
  Dim ETitle$, SumHeader2$, RHandle As Integer
  Dim DHandle As Integer, NHandle As Integer
  Dim THandle As Integer, DistributionRptName$
  Dim FF$, TErn$, PayRegisterRptName$
  Dim JFlag As Boolean, TotalGLIFS As Integer
  Dim TotalAccts As Integer, PrintGLRpt As Boolean
  Dim GLIFRecLen As Integer, GLIFRecName$
  Dim GRHandle As Integer, GLHandle As Integer
  Dim ActualAccts As Integer, Max As Integer
  Dim Lines As Integer, GLIdxName$, AMAmt#
  Dim Cnt2 As Integer, AcctOk As Boolean, GLAcct@
  Dim NoAcctNum As Integer, Fund$, FDebit#, FCredit#
  Dim NFund$, RetCode As Integer, LincCnt As Integer
  Dim RegHrs#, VACHRS#, SICKHRS#, HOLHRS#, COMPHRS#, PerHours#
  Dim TotalHrs#, TotEIC#, TRegWage#, TOTWage#, GPay#
  Dim SSTax#, MTax#, FTax#, STax#, RETTOT#
  Dim TNetPay#, Emp1Handle As Integer
  Dim SumDed$(1 To 5), GLIdxRecLen As Integer
  Dim SumErn$, LastErn As Integer
  Dim ENumOfAct As Integer, Acct$, First As Integer
  Dim Last As Integer, Start As Integer
  Dim TotalSocAmt#, DistDif#, TotalMedAmt#, LastActive As Integer
  Dim TotalRetAmt#, DLincCnt As Integer, Cnt3 As Integer
  Dim TotHrs#, RegWage#, OTWage#, PrnDef$
  Dim AddEarn#, DGPay#, OutOfOrder As Boolean
  Dim Image0$, Image$, Image3$, Image4$, Image5$
  Dim foundIt As Boolean
  Dim NumOfDeds As Integer, Nextx As Integer
  Dim tripCnt As Integer
  Dim RptHandle#
  
  '-------------Temp Void variables------------
  Dim CSocWHAcct$
  Dim CMedWHAcct$
  Dim CSocMatchAcct$
  Dim CMedMatchAcct$
  Dim CRetMatchAcct$
  Dim CFedWHAcct$
  Dim CStaWHAcct$
  Dim CRetWHAcct$
  Dim CDedAcct$
  Dim CPRNetAcct$
  Dim DWagesAcct$
  Dim DSocMatchAcct$
  Dim DMedMatchAcct$
  Dim DRetMatchAcct$
  Dim FundNumOnly$
  Dim FundAndAcctOnly$
  Dim TempVoid As VoidCheckType
  Dim TVHandle As Integer
  Dim TVCnt As Double
  Dim AcctLen As Integer
  Dim DetLen As Integer
  Dim ThisPR As Double
  Dim ThisFTax#
  Dim ThisMTax#
  Dim ThisSSTax#
  Dim ThisStaTax#
  Dim ThisRTax#
  Dim FACnt As Integer
  Dim ThisFACnt As Integer
  Dim TotalDeds#
  Dim DbtCnt As Integer
  Dim ThisCRGL$
  '---------------^^^^-------------------------
  Dim ActiveCnt As Integer
  Dim ThisDesc$
  Dim Thisx As Integer
  Dim z As Integer
  ReDim TotDedAmt(1 To 50) As Double '8/19/04
  Dim CIDebit# '8/19/04
  Dim CICredit# '8/19/04
  ReDim FundArray(1 To 1) As String '8/19/04
  Dim FirstFundSum As Boolean '8/19/04
  Dim FedTot As Double '8/19/04
  Dim MedTot As Double '8/19/04
  Dim SocTot As Double '8/19/04
  Dim StaTot As Double '8/19/04
  Dim FSMTot As Double '8/19/04
  Dim FundCount As Integer
  Dim PoolFundNum$ '9/17/04
  Dim PRNetPoolFound As Boolean
  Dim TotalWHAndDeds As Double
  Dim TotalWages As Double
  Dim AcctNumCnt As Integer
  Dim ThisEmpCnt As Integer
  Dim PRNetSum As Double
  Dim GL4PPETotal$
  Dim PRPoolProcessed As Boolean
  
  PRPoolProcessed = False
  
  FirstFundSum = True
  FundCount = 0
  
  GLDebitTotal = 0
  GLCreditTotal = 0
  GLError = 0
  
  ActiveCnt = 0
  FF$ = Chr$(12)
  
  RptTitle$ = "Register & G/L Interface Reports"
  FrmShowPctComp.Label1 = RptTitle$
  FrmShowPctComp.Show
  ReDim TransRec(1) As TransRecType
  ReDim EmpRec1(1) As EmpData1Type
  ReDim PDR(1) As PeriodDefaultRecType
  ReDim Unit(1) As UnitFileRecType
  
  ReDim DistbSumAccts(1 To 1) As DistWageRptType
  
  ReDim SysRec(1) As RegDSysFileRecType
  ReDim GLIFRec(1 To 1) As GLIFDataType14
  
  ReDim EmpAct(1) As DistWageRptType
  
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType
  
  ReDim TotDeds#(1 To 50)
  ReDim TotErns(1 To 3) As Double
  
  ReDim EDesc(1) As String * 21
  ReDim EDAct(1) As String * 14
  ReDim EDPct(1) As String * 11
  ReDim EDRHrs(1) As String * 11
  ReDim EDOHrs(1) As String * 11
  ReDim EDRPay(1) As String * 11
  ReDim EDOPay(1) As String * 11
  ReDim EDEarn(1) As String * 11
  ReDim EDGroP(1) As String * 11
  
  ReDim EDSAmt(1) As String * 11
  ReDim EDMAmt(1) As String * 11
  ReDim EDRAmt(1) As String * 11
  
  ReDim ENumb(1) As String * 14
  ReDim EName(1) As String * 33
  
  ReDim BRat(1) As String * 11
  ReDim ORat(1) As String * 11
  
  ReDim TaxFrn(1) As String * 11
  ReDim Fill11(1) As String * 11
  
  ReDim SCnt(1) As String * 11
  ReDim HCnt(1) As String * 11
  
  ReDim RHrs(1) As String * 11
  ReDim VHrs(1) As String * 11
  ReDim SHrs(1) As String * 11
  ReDim HHrs(1) As String * 11
  ReDim CHrs(1) As String * 11
  ReDim THrs(1) As String * 11

  ReDim PHrs(1) As String * 11

  ReDim OTHrs(1) As String * 11
  ReDim OTPaid(1) As String * 11
  ReDim OTComp(1) As String * 11
  
  ReDim RErnP(1) As String * 11
  ReDim OErnP(1) As String * 11
  
  ReDim GPayP(1) As String * 11
  ReDim SSTaxP(1) As String * 11
  ReDim MTaxP(1) As String * 11
  ReDim FTaxP(1) As String * 11
  ReDim STaxP(1) As String * 11
  ReDim RetirP(1) As String * 11
  ReDim NetPayP(1) As String * 11
  ReDim Ded(1) As String * 11
  
  'added for EIC   6/07/94
  ReDim EEicP(1) As String * 11
  
  ReDim Ern(1) As String * 11
  
  ReDim Pg(1) As String * 5
  
  ReDim EMPLine(1) As String * 132
  
  ReDim Dash0(1) As String * 69
  ReDim Dash1(1) As String * 132
  ReDim Dash2(1) As String * 124
  ReDim Dash3(1) As String * 91
  Dim ThisFund As Integer '8/10/04
  Dim y As Integer '8/10/04
  
  Dim TOTFEDTAX As Double '8/13/04
  Dim TOTMEDTAX As Double '8/13/04
  Dim TOTSOCTAX As Double '8/13/04
  Dim TOTSTATAX As Double '8/13/04
  Dim TOTRetTax As Double '8/13/04
  Dim TOTMEDMat As Double '8/13/04
  Dim TOTSOCMat As Double '8/13/04
  Dim TOTRETMat As Double '8/13/04
  ReDim TotDedAmt(1 To 50) As Double '8/13/04
  
  OpenPPDefaultFile PPDefaultFileName
  Get PPDefaultFileName, 1, PDR(1)
  Close PPDefaultFileName
  
  OpenSysFile SysFileName
  Get SysFileName, 1, SysRec(1)
  Close SysFileName
  PoolFundNum = Mid(SysRec(1).Liab(1).Acct, 1, 2)
  
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle
  
'  Call GetAcctStruct(SysRec(1).CITIDIR, FundLen, AcctLen, DetLen)
  Call GetAcctStruct(CurrCitiPath, FundLen, AcctLen, DetLen)
  FACnt = FundLen + AcctLen
  If DetLen > FundLen Then
    FundPad = DetLen - FundLen
  Else
    FundPad = 0
  End If
  
  OpenDedCodeFile DedCodeFileName
  ReDim DedCodeNums(1 To 50) As String '6/22/04
  ReDim DedCodeDesc(1 To 50) As String
  For x = 1 To 50
    Get DedCodeFileName, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      NumOfDeds = NumOfDeds + 1
      DedCodes(x) = DedRec
      DedCodeNums(x) = QPTrim$(DedRec.DCACCT1)
      DedCodeDesc(x) = QPTrim$(DedRec.DCDESC1)
    End If
  Next x
  Close DedCodeFileName

  OpenErnCodeFile ErnCodeFileName
  For x = 1 To 3
    Get ErnCodeFileName, x, ErnCodes(x)
  Next x
  Close ErnCodeFileName
  
  ReDim GLSetupRec(1) As GLSetupRecType
'for new gl
'  FundPad = 0
'  FundLen = 2     'Default fund length
  
'  If Exist(QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT") Then
  If Exist(CurrCitiPath + "\GLSETUP.DAT") Then
    foundIt = True
  ElseIf Exist(CurrCitiPath + "GLSETUP.DAT") Then
    foundIt = True
  Else '7/26
    Unload FrmShowPctComp '7/26
    MsgBox "Path to GLSETUP.DAT cannot be found." '7/26
    GoTo SkipGLReport '7/26
  End If '7/26
  
'  GLSetUpName$ = QPTrim$(SysRec(1).CITIDIR) + "\GLSETUP.DAT"
  If Mid(CurrCitiPath, Len(CurrCitiPath), 1) <> "\" Then
    GLSetUpName$ = CurrCitiPath + "\GLSETUP.DAT"
  ElseIf Mid(CurrCitiPath, Len(CurrCitiPath), 1) = "\" Then
    GLSetUpName$ = CurrCitiPath + "GLSETUP.DAT"
  End If
    
  GLSetUpRecLen = Len(GLSetupRec(1))
  GLHandle = FreeFile
  Open GLSetUpName$ For Random Shared As GLHandle Len = GLSetUpRecLen
  
  If foundIt = True Then
    Get GLHandle, 1, GLSetupRec(1)
    FundLen = GLSetupRec(1).FundLen
    FundPad = GLSetupRec(1).DetLen - GLSetupRec(1).FundLen
  End If
  Close GLHandle
  FundLen = FundLen
  
SkipGLReport:
  Image0$ = "####"
  Image$ = "###0.00"
  Image3$ = "###,##0.00"
  Image4$ = "##0.0000"
  Image5$ = "####,##0.00"
  
  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GRetGross# = 0

  GTaxFrn# = 0
  TotDebit# = 0
  TotCredit# = 0
  
  LSet Fill11$(1) = ""
  
  LSet Dash0(1) = String$(69, "-")
  LSet Dash1(1) = String$(132, "-")
  LSet Dash2(1) = String$(124, "-")
  RSet Dash3(1) = String$(63, "-")
  
  EmpActLen = Len(EmpAct(1))
  DistbSumAcctsLen = Len(DistbSumAccts(1))
  
  TransRecLen = Len(TransRec(1))
  Emp1RecLen = Len(EmpRec1(1))
  
  OpenEmpData1File Emp1Handle
  NumOfRecs = LOF(Emp1Handle) / Len(EmpRec1(1))
  Close Emp1Handle

  SalCnt = 0
  HrlCnt = 0
  
  DLineCnt = 0
  LineCnt = 0
  
  NumOfWageAccts = 0

  MaxLines = 45
  DMaxLines = 40
  EPage = 1
  Page = 1
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  For x = 1 To NumOfRecs
    Get EmpIdxNNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxNNameHandle
  
  For x = 1 To 5
    DTitle$(x) = ""
  Next x
  
  Nextx = 1
  tripCnt = 1
  For cnt = 1 To 50
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    TDed$ = QPTrim$(DedCodes(cnt).DCDESC1)
    If Len(TDed$) > 0 Then
      LastDed = LastDed + 1
      RSet Ded(1) = TDed$
      DTitle$(Nextx) = DTitle$(Nextx) + Ded(1)
    End If
    tripCnt = tripCnt + 1
  Next
  
  '---------------------------------------------
  ETitle$ = ""
  For cnt = 1 To 3
    TErn$ = QPTrim$(ErnCodes(cnt).ERNCODE1)
    If Len(TErn$) > 0 Then
      LastErn = LastErn + 1
      RSet Ern(1) = TErn$
      ETitle$ = ETitle$ + Ern(1)
    Else
      Exit For
    End If
  Next
  
  If LastErn < 3 Then
    ETitle$ = Space$(11 * (3 - LastErn)) + ETitle$
  End If
  SumHeader2$ = "  Reg Wages  O/T Wages" + ETitle$
  ETitle$ = "   Reg Earn   O/T Earn" + ETitle$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"
  
  '------------------------------------------------------------------
  
  PayRegisterRptName$ = "PRRPTS\REGISTER.RPT"
  RHandle = FreeFile
  Open PayRegisterRptName$ For Output As RHandle
  RPTSetupPRN 13, RHandle
  DistributionRptName$ = "PRRPTS\DISTRIBU.RPT"
  DHandle = FreeFile
  Open DistributionRptName$ For Output As DHandle
  RPTSetupPRN 13, DHandle
  OpenEmpData1File NHandle
  
  OpenTransWorkFile THandle
  
  
  GoSub PrintPayRollHeader
  GoSub PrintDistHeader
  
  KillFile TempVoidFileName
  OpenTempVoidFile TVHandle
  
  ReDim ThisDedAmt(1 To 50) As Double
  
  For cnt = 1 To NumOfRecs + 1
    If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
      If TVCnt <> 0 And PRNetPoolFound = False Then
        GoSub NoPRNetForPoolCOrI
      End If
    ElseIf TVCnt <> 0 And PRNetPoolFound = False Then
      GoSub NoPRNetForPool
    End If
    If cnt = NumOfRecs + 1 Then Exit For 'catches the last employee
    'if there is no 'Paid From'/pool fund
    Get THandle, IdxBuff(cnt), TransRec(1)
    If TransRec(1).TActive = True Then
      ReDim ThisPRDbtFund(1 To 1) As String
      ReDim ThisPRDbtAmt(1 To 1) As Double
      DbtCnt = 0
      Get NHandle, IdxBuff(cnt), EmpRec1(1)
      PRNetPoolFound = False
      TotalWHAndDeds = 0
      AcctNumCnt = 0
      ThisEmpCnt = 0
      For x = 1 To 50
        ThisDedAmt(x) = 0 '6/22/2004
      Next x
      ThisFTax# = 0 '6/22/2004
      ThisMTax# = 0 '6/22/2004
      ThisSSTax# = 0 '6/22/2004
      ThisStaTax# = 0 '6/22/2004
      ThisRTax# = 0 '6/22/2004
      TotalDeds# = 0 '6/22/2004
      GoSub SumAndPrintTime
      GoSub ParseDistributions
      ActiveCnt = ActiveCnt + 1
      LineCnt = LineCnt + 5
      If LineCnt >= MaxLines And cnt < NumOfRecs Then
        LineCnt = 0
        Print #RHandle, FF$
        GoSub PrintPayRollHeader
      End If
      If DLineCnt >= DMaxLines Then
        DLineCnt = 0
        Print #DHandle, FF$
        GoSub PrintDistHeader
      End If
      
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  If ActiveCnt = 0 Then
    MsgBox "No employees have been designated for payroll processing."
    Close
    GoTo AltExit
  End If
  
  GoSub PrintSumTotal
  GoSub PrintDistTotal
  
  Close THandle
  Close NHandle
  RPTSetupPRN 123, RHandle '7/24
  Close RHandle
  RPTSetupPRN 123, DHandle '7/24
  Close DHandle
  Close RptHandle
  Close TVHandle
  'HERE:  register report files have been written to disk
  
  'GLIF START
  'if there is a GL transfer directory then make GLIF file
  
'  If Len(QPTrim$(SysRec(1).CITIDIR)) > 0 Then
  If Len(CurrCitiPath) > 0 Then
    ReDim Passed#(1 To 6)
    Passed#(1) = FTax#
    Passed#(2) = STax#
    Passed#(3) = SSTax#
    Passed#(4) = MTax#
    Passed#(5) = RETTOT#
    Passed#(6) = TNetPay#
    DistbSumAccts(1).Acct = DistbSumAccts(1).Acct
    MakeGLIFFileT TotEIC#, TotDeds#(), Passed#(), DistbSumAccts() 'unrem
  End If
  
'  If Exist(QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile) Then
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile
  If Exist(CurrCitiPath + "\" + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + "\" + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + "\" + GLAcctIdxFile) Then
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + "\" + GLAcctIdxFile
  Else
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + GLAcctIdxFile
  End If
  
  If JFlag Then
    GLIdxRecLen = Len(JGLIdxRec(1))
  Else
    GLIdxRecLen = Len(GLIdxRec(1))
  End If
  
  GLIFRecLen = Len(GLIFRec(1))
  TotalGLIFS = FileSize("TempIF.DAT") \ GLIFRecLen
  TotalAccts = FileSize(GLIdxName$) \ GLIdxRecLen
  
  If TotalGLIFS = 0 Then
    PrintGLRpt = False
    GoTo SkipGLRpt
    
  Else
    PrintGLRpt = True
  End If
  
  GLIFRecLen = Len(GLIFRec(1))
  GLIFRecName$ = "TempIF.DAT"
  GRHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GRHandle Len = GLIFRecLen
  ReDim GLIFRec(1 To TotalGLIFS) As GLIFDataType14
  For x = 1 To TotalGLIFS
    Get GRHandle, x, GLIFRec(x)
  Next x
  Do
    OutOfOrder = False                     'assume it's sorted
    For x = 1 To UBound(GLIFRec) - 1
      If GLIFRec(x).TranAcct > GLIFRec(x + 1).TranAcct Then
        SWAP GLIFRec(x), GLIFRec(x + 1)    'if we had to swap
        OutOfOrder = True                'we're not done yet
      End If
    Next
  Loop While OutOfOrder

  For x = 1 To TotalGLIFS
    Put GRHandle, x, GLIFRec(x)
  Next x

  If TotalAccts = 0 Then
    Close
    GoTo SkipGLAccts
  End If
  Close GRHandle
  FrmShowPctComp.Label1 = "Reading G/L Accounts."
  FrmShowPctComp.Show ' , Me
  
  GLHandle = FreeFile
  Open GLIdxName$ For Random As GLHandle Len = GLIdxRecLen
  Select Case JFlag
  Case False
    ReDim GoodAccts(1 To TotalAccts) As Double
    For cnt = 1 To TotalAccts
      Get GLHandle, cnt, GLIdxRec(1)
      If GLIdxRec(1).AcctNum > 0 Then
        ActualAccts = ActualAccts + 1
        GoodAccts(ActualAccts) = GLIdxRec(1).AcctNum            'QPValL(AcctNum$)
        If GoodAccts(ActualAccts) < 9999999 Then
          GoodAccts(ActualAccts) = OldRound(GoodAccts(ActualAccts) * 100)
        End If
      End If
       FrmShowPctComp.ShowPctComp cnt, TotalAccts
       If FrmShowPctComp.Out = True Then
         Close
         FrmShowPctComp.Out = False
         Unload FrmShowPctComp
         Exit Sub
       End If
    Next
  Case True
    ReDim JGoodAccts(1 To TotalAccts) As String * 16
    For cnt = 1 To TotalAccts
      Get GLHandle, cnt, JGLIdxRec(1)
      ActualAccts = ActualAccts + 1
      ReplaceString JGLIdxRec(1).AcctNum, "-", ""
      JGoodAccts(ActualAccts) = JGLIdxRec(1).AcctNum
      FrmShowPctComp.ShowPctComp cnt, TotalAccts
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        Exit Sub
      End If

    Next
    
  End Select
  
  Close GLHandle
  TotalAccts = ActualAccts
  
  FrmShowPctComp.Label1 = "Checking for Invalid Accounts."
  FrmShowPctComp.Show

SkipGLAccts:
  Max = 55
  Lines = 0
  
  RptName$ = "prrpts\PRGLIF.RPT"
  KillFile "prrpts\PRGLIF.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  Close RHandle
  RptHandle = FreeFile
  Open RptName$ For Append As RptHandle
  RPTSetupPRN 14, RptHandle
'  KillFile "TempIF.DAT" '12/27/2002
  THandle = FreeFile
  Open "TempIF.DAT" For Random As THandle Len = GLIFRecLen
  
  GoSub GLIFHeader
  
  For cnt = 1 To TotalGLIFS
    FrmShowPctComp.ShowPctComp cnt, TotalGLIFS
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
    If Lines >= Max Then
      Print #RptHandle, FF$
      GoSub GLIFHeader
    End If
'    Debug.Print GLIFRec(cnt).DrAmt
   
    TotDebit# = OldRound(TotDebit# + GLIFRec(cnt).DrAmt)
    TotCredit# = OldRound(TotCredit# + GLIFRec(cnt).CrAmt)
    
    '10-13-94 ** added check for valid GL Account numbers
    If TotalAccts > 0 Then
      
      Select Case JFlag
      Case True
        For Cnt2 = 1 To TotalAccts
          If InStr(JGoodAccts(Cnt2), GLIFRec(cnt).TranAcct) And Len(QPTrim$(JGoodAccts(Cnt2))) = Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
            AcctOk = True
            Exit For
          End If
        Next
        If Not AcctOk Then
          GLError = -1
          LSet GLIFRec(cnt).Fill = "Error "
           Put THandle, cnt, GLIFRec(cnt)
        Else
          LSet GLIFRec(cnt).Fill = ""
        End If
        AcctOk = False
        
      Case False
        GLAcct@ = OldRound(Val(GLIFRec(cnt).TranAcct))
        If GLAcct@ < 9999999 Then
          GLAcct@ = GLAcct@ * 100
        End If
        For Cnt2 = 1 To TotalAccts
          If GoodAccts(Cnt2) = GLAcct@ Then
            AcctOk = True
            Exit For
          End If
        Next
        If Not AcctOk Then
          LSet GLIFRec(cnt).Fill = "Error "
           Put THandle, cnt, GLIFRec(cnt)
        Else
          LSet GLIFRec(cnt).Fill = ""
        End If
        AcctOk = False
      End Select
    End If

    'NoCheckAccts:
    RSet EDSAmt(1) = Using(Image3$, GLIFRec(cnt).DrAmt)
    RSet EDMAmt(1) = Using(Image3$, GLIFRec(cnt).CrAmt)
    LSet EDesc(1) = QPTrim$(GLIFRec(cnt).TranDesc)
    'added 8/19/04------------------------
    ThisDesc$ = QPTrim$(GLIFRec(cnt).TranDesc)
    If FundCount = 0 Then
      FundCount = FundCount + 1
      ReDim FundArray(1 To FundCount) As String
      FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
      ReDim FedTaxByFund(1 To FundCount) As Double
      ReDim MedTaxByFund(1 To FundCount) As Double
      ReDim SocTaxByFund(1 To FundCount) As Double
      ReDim StaTaxByFund(1 To FundCount) As Double
      ReDim RetTaxByFund(1 To FundCount) As Double
      ReDim MedMatByFund(1 To FundCount) As Double
      ReDim SocMatByFund(1 To FundCount) As Double
      ReDim RetMatByFund(1 To FundCount) As Double
      ReDim DedAmtByFund(1 To 50, 1 To FundCount)
      Thisx = 1
   Else
     For x = 1 To FundCount
       If FundArray(x) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen) Then
         Thisx = x
         Exit For
       End If
     Next x
     If x > FundCount Then
       FundCount = FundCount + 1
       ReDim Preserve FundArray(1 To FundCount) As String
       FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
       ReDim Preserve FedTaxByFund(1 To FundCount) As Double
       ReDim Preserve MedTaxByFund(1 To FundCount) As Double
       ReDim Preserve SocTaxByFund(1 To FundCount) As Double
       ReDim Preserve StaTaxByFund(1 To FundCount) As Double
       ReDim Preserve RetTaxByFund(1 To FundCount) As Double
       ReDim Preserve MedMatByFund(1 To FundCount) As Double
       ReDim Preserve SocMatByFund(1 To FundCount) As Double
       ReDim Preserve RetMatByFund(1 To FundCount) As Double
       ReDim Preserve DedAmtByFund(1 To 50, 1 To FundCount)
       Thisx = FundCount
     End If
   End If

    If ThisDesc = "Fed Withholdings" Then
      FedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Soc Sec Withholdings" Then
      SocTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Med Withholdings" Then
      MedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "State Withholdings" Then
      StaTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Ret Withholdings" Then
      RetTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Soc Match Liab" Then
      SocMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Med Match Liab" Then
      MedMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "Ret Match Liab" Then
      RetMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    Else
      For z = 1 To 50
        If ThisDesc = QPTrim$(DedCodes(z).DCDESC1) Then
          DedAmtByFund(z, Thisx) = GLIFRec(cnt).CrAmt
        End If
      Next z
    End If
    'added 8/19/04--^^^^^^^^^^^^^^^-------
    NoAcctNum = Len(QPTrim$(GLIFRec(cnt).TranAcct))
    If NoAcctNum > 0 Then
      Print #RptHandle, " "; GLIFRec(cnt).TranAcct; QPTrim(GLIFRec(cnt).Fill); " "; EDesc(1); EDSAmt(1); "    "; EDMAmt(1)
      Lines = Lines + 1
    End If
  Next
  GoSub GLIFTotals
  
  Fund$ = Left$(QPTrim$(GLIFRec(1).TranAcct), FundLen)
  FDebit# = GLIFRec(1).DrAmt
  FCredit# = GLIFRec(1).CrAmt
  For cnt = 2 To TotalGLIFS
    NFund$ = Left$(QPTrim$(GLIFRec(cnt).TranAcct), FundLen)
    If NFund$ <> Fund$ Then
      GoSub PrintFundTotal
      Fund$ = NFund$
      FDebit# = GLIFRec(cnt).DrAmt
      FCredit# = GLIFRec(cnt).CrAmt
    Else
      FDebit# = OldRound(FDebit# + GLIFRec(cnt).DrAmt)
      FCredit# = OldRound(FCredit# + GLIFRec(cnt).CrAmt)
    End If
  Next
  GoSub PrintFundTotal
  GoSub PrintFundSummary
  Print #RptHandle, FF$; 'PrnDef$
  RPTSetupPRN 123, RptHandle '7/24
  
  Close RptHandle
  Close THandle
  Close TVHandle
  Close
SkipGLRpt:
  
  '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
  '05-10-94  fixed to insure everything was cleaned up after report
  '06-28-94  move so all is cleaned up BEFORE the report prints.
  '07-15-94  move again to add gl interface report.
  '----------------------------------------------------------------------------
  
  RptTitle$ = "Payroll Register Report"
  ViewPrint PayRegisterRptName, RptTitle$, True
  
  If RetCode = -1 Then GoTo AltExit
  
  RptTitle$ = "Payroll Distribution Report"
  ViewPrint DistributionRptName, RptTitle$, True
  
  If RetCode = -1 Then GoTo AltExit
  
  If GLError <> -1 Then
    GLError = TotalAccts
  End If
  
  If PrintGLRpt Then
    If GLCreditTotal <> GLDebitTotal Then
      frmMessage.Label1.Caption = "The General Ledger Interface is OUT OF BALANCE."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      MainLog "User warned that the GL Interface is OUT OF BALANCE (Debit Total = " + QPTrim$(Using$("$#,###,##0.00", GLDebitTotal)) + " and the Credit Total is " + QPTrim$(Using$("$#,###,##0.00", GLCreditTotal)) + ")."
    End If
    If GLError < 1 Then
      frmMessage.Label1.Caption = "General Ledger number errors have been found in the GL Interface Report."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      MainLog "User warned that the GL Interface Report has GL number errors."
    End If
    RptTitle$ = "G/L Interface Report"
    ViewPrint "prrpts\PRGLIF.RPT", RptTitle$, True
  End If
  
'  Dim NumAccts As Integer

'  OpenTempVoidFile TVHandle
'  NumAccts = LOF(TVHandle) / Len(TempVoid)
'  For X = 1 To NumAccts
'    Get TVHandle, X, TempVoid
'      Debug.Print TempVoid.EmpNum
'      If QPTrim$(TempVoid.EmpNum) = "74" Then Stop
'      Debug.Print TempVoid.PPEGL + " PPE              " + CStr(TempVoid.PPEAmt)
'      Debug.Print TempVoid.PPETotGL + " PPE Total   " + CStr(TempVoid.PPETotAmt)
'      Debug.Print TempVoid.PRNetGL + " PRNET            " + CStr(TempVoid.PRNet)
'      Debug.Print TempVoid.SOCWHGL + " SOC Withholdings " + CStr(TempVoid.SOCWHAmt)
'      Debug.Print TempVoid.MEDWHGL + " MED Withholdings " + CStr(TempVoid.MEDWHAmt)
'      Debug.Print TempVoid.SOCMATCRGL + " SOC Match Liab   " + CStr(TempVoid.SOCMATCRAmt)
'      Debug.Print TempVoid.MEDMATCRGL + " MED Match Liab   " + CStr(TempVoid.MEDMATCRAmt)
'      Debug.Print TempVoid.FEDWHGL + " FED Withholdings " + CStr(TempVoid.FEDWHAmt)
'      Debug.Print TempVoid.STAWHGL + " STA Withholdings " + CStr(TempVoid.STAWHAmt)
'      Debug.Print TempVoid.RETWHGL + " RET Withholdings " + CStr(TempVoid.RETWHAmt)
'      Debug.Print TempVoid.RETMATCRGL + " RET Match Liab   " + CStr(TempVoid.RETMATCRAmt)
'      For Cnt2 = 1 To 50
'        If TempVoid.DedData(Cnt2).DAmt > 0 Then
'          Debug.Print TempVoid.DedData(Cnt2).DedGLNum + " Deduction        " + CStr(TempVoid.DedData(Cnt2).DAmt)
'        End If
'      Next Cnt2
'      Debug.Print TempVoid.WagesGL + "  Wages           " + CStr(TempVoid.WagesAmt)
'      Debug.Print TempVoid.SOCMATDBGL + " SOC Match        " + CStr(TempVoid.SOCMATDBAmt)
'      Debug.Print TempVoid.MEDMATDBGL + " MED Match        " + CStr(TempVoid.MEDMATDBAmt)
'      Debug.Print TempVoid.RETMATDBGL + " RET Match        " + CStr(TempVoid.RETMATDBAmt)
'  Next X
'  Close TVHandle
  
AltExit:
  
  Exit Sub
  
PrintFundTotal:
  RSet EDSAmt(1) = Using(Image3$, FDebit#)
  RSet EDMAmt(1) = Using(Image3$, FCredit#)
  LSet EDesc(1) = ""
  LSet GLIFRec(1).Fill = ""
  RSet GLIFRec(1).TranAcct = Fund$
  If FirstFundSum = True Then
    FirstFundSum = False
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #RptHandle, String$(80, "-")
    Lines = 4
  End If
  If Len(QPTrim$(Fund$)) > 0 Then
    For x = 1 To FundCount
      If FundArray(x) = QPTrim$(GLIFRec(1).TranAcct) Then
        ThisFund = x
        Exit For
      End If
    Next x
    
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
    Print #RptHandle, String$(80, "-")
    If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
      If QPTrim$(Fund) = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
        Lines = Lines + 3
        Return
      End If
    End If
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = Lines + 5
    
    If Lines >= Max Then
      Print #RptHandle, FF$
      Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
      Print #RptHandle, "G/L Interface Report Summary."
      Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
      Print #RptHandle, String$(80, "-")
      Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
      Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
      Print #RptHandle, String$(80, "-")
      Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
      Print #RptHandle,
      Lines = 9
    End If
    If x <= FundCount Then
      TOTFEDTAX = OldRound(TOTFEDTAX + FedTaxByFund(ThisFund))
      TOTMEDTAX = OldRound(TOTMEDTAX + MedTaxByFund(ThisFund))
      TOTSOCTAX = OldRound(TOTSOCTAX + SocTaxByFund(ThisFund))
      TOTSTATAX = OldRound(TOTSTATAX + StaTaxByFund(ThisFund))
      TOTRetTax = OldRound(TOTRetTax + RetTaxByFund(ThisFund))
      TOTMEDMat = OldRound(TOTMEDMat + MedMatByFund(ThisFund))
      TOTSOCMat = OldRound(TOTSOCMat + SocMatByFund(ThisFund))
      TOTRETMat = OldRound(TOTRETMat + RetMatByFund(ThisFund))
      FedTot = FedTaxByFund(ThisFund)
      MedTot = MedTaxByFund(ThisFund) + MedMatByFund(ThisFund)
      SocTot = SocTaxByFund(ThisFund) + SocMatByFund(ThisFund)
      StaTot = StaTaxByFund(ThisFund)
      RETTOT = RetTaxByFund(ThisFund) + RetMatByFund(ThisFund)
      FSMTot = FedTot + MedTot + SocTot
      
      Print #RptHandle, " Federal"; Tab(19); Using(Image3$, FedTaxByFund(ThisFund)); Tab(71); Using$(Image3$, FedTot)
      Lines = Lines + 1
      If Lines >= Max Then
        Print #RptHandle, FF$
        Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
        Print #RptHandle, "G/L Interface Report Summary."
        Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
        Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
        Print #RptHandle,
        Lines = 9
      End If
      Print #RptHandle, " Social Security"; Tab(19); Using(Image3$, SocTaxByFund(ThisFund)); Tab(46); Using(Image3$, SocMatByFund(ThisFund)); Tab(71); Using$(Image3$, SocTot)
      Lines = Lines + 1
      If Lines >= Max Then
        Print #RptHandle, FF$
        Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
        Print #RptHandle, "G/L Interface Report Summary."
        Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
        Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
        Print #RptHandle,
        Lines = 9
      End If
      
      Print #RptHandle, " Medicare"; Tab(19); Using(Image3$, MedTaxByFund(ThisFund)); Tab(46); Using(Image3$, MedMatByFund(ThisFund)); Tab(71); Using$(Image3$, MedTot)
      Lines = Lines + 1
      If Lines >= Max Then
        Print #RptHandle, FF$
        Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
        Print #RptHandle, "G/L Interface Report Summary."
        Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
        Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
        Print #RptHandle,
        Lines = 9
      End If
      Print #RptHandle, " Sub Total"; Tab(71); Using(Image3$, FSMTot)
      Lines = Lines + 1
      If Lines >= Max Then
        Print #RptHandle, FF$
        Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
        Print #RptHandle, "G/L Interface Report Summary."
        Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
        Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
        Print #RptHandle,
        Lines = 9
      End If
      Print #RptHandle,
      Lines = Lines + 1
      If Lines >= Max Then
        Print #RptHandle, FF$
        Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
        Print #RptHandle, "G/L Interface Report Summary."
        Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
        Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
        Print #RptHandle,
        Lines = 9
      End If
      Print #RptHandle, " State"; Tab(19); Using(Image3$, StaTaxByFund(ThisFund)); Tab(71); Using$(Image3$, StaTot)
      Lines = Lines + 1
      If Lines >= Max Then
        Print #RptHandle, FF$
        Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
        Print #RptHandle, "G/L Interface Report Summary."
        Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
        Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
        Print #RptHandle,
        Lines = 9
      End If
      Print #RptHandle, " Retirement"; Tab(19); Using(Image3$, RetTaxByFund(ThisFund)); Tab(46); Using(Image3$, RetMatByFund(ThisFund)); Tab(71); Using$(Image3$, RETTOT)
      Lines = Lines + 1
      If Lines >= Max Then
        Print #RptHandle, FF$
        Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
        Print #RptHandle, "G/L Interface Report Summary."
        Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
        Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
        Print #RptHandle, String$(80, "-")
        Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
        Print #RptHandle,
        Lines = 9
        End If
        Print #RptHandle, String$(80, "-")
        Lines = Lines + 1
        If Lines >= Max Then
          Print #RptHandle, FF$
          Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
          Print #RptHandle, "G/L Interface Report Summary."
          Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
          Print #RptHandle, String$(80, "-")
          Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
          Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
          Print #RptHandle, String$(80, "-")
          Lines = 7
        End If
        
        If Fund <> Mid(SysRec(1).Liab(1).Acct, 1, FundLen) Then
          GLDebitTotal = OldRound(GLDebitTotal + EDSAmt(1))
          GLCreditTotal = OldRound(GLCreditTotal + EDMAmt(1))
          Return
        End If
        
        Print #RptHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
        Lines = Lines + 1
        If Lines >= Max Then
          Print #RptHandle, FF$
          Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
          Print #RptHandle, "G/L Interface Report Summary."
          Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
          Print #RptHandle, String$(80, "-")
          Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
          Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
          Print #RptHandle, String$(80, "-")
          Print #RptHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
          Print #RptHandle,
          Lines = 9
        End If
        Print #RptHandle,
      
        Lines = Lines + 1
        If Lines >= Max Then
          Print #RptHandle, FF$
          Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
          Print #RptHandle, "G/L Interface Report Summary."
          Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
          Print #RptHandle, String$(80, "-")
          Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
          Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
          Print #RptHandle, String$(80, "-")
          Print #RptHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
          Print #RptHandle,
          Lines = 9
        End If
        Print #RptHandle,
        Lines = Lines + 1
        If Lines >= Max Then
          Print #RptHandle, FF$
          Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
          Print #RptHandle, "G/L Interface Report Summary."
          Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
          Print #RptHandle, String$(80, "-")
          Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
          Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
          Print #RptHandle, String$(80, "-")
          Print #RptHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
          Print #RptHandle,
          Lines = 9
        End If
       For z = 1 To 50
         TotDedAmt(z) = OldRound(TotDedAmt(z) + DedAmtByFund(z, ThisFund))
         If QPTrim$(DedCodes(z).DCDESC1) <> "" Then
           Print #RptHandle, Tab(2); QPTrim$(DedCodes(z).DCDESC1); Tab(19); Using(Image3$, DedAmtByFund(z, ThisFund));
           Lines = Lines + 1
         End If
         If z = 50 Then Exit For
         TotDedAmt(z + 1) = OldRound(TotDedAmt(z + 1) + DedAmtByFund(z + 1, ThisFund))
         If QPTrim$(DedCodes(z + 1).DCDESC1) <> "" Then
           Print #RptHandle, Tab(51); QPTrim$(DedCodes(z + 1).DCDESC1); Tab(71); Using(Image3$, DedAmtByFund(z + 1, ThisFund))
           z = z + 1
         End If
         If Lines >= Max Then
           Print #RptHandle, FF$
           Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
           Print #RptHandle, "G/L Interface Report Summary."
           Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
           Print #RptHandle, String$(80, "-")
           Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
           Print #RptHandle, "  Fund " + Fund; Tab(41); EDSAmt(1); Tab(70); EDMAmt(1)
           Print #RptHandle, String$(80, "-")
           Print #RptHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
           Print #RptHandle,
           Lines = 9
         End If
       Next z
       Print #RptHandle,
       Print #RptHandle,
       Print #RptHandle,
       Lines = Lines + 3
       If Lines >= Max - 4 Then
         Print #RptHandle, FF$
         Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
         Print #RptHandle, "G/L Interface Report Summary."
         Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
         Print #RptHandle, String$(80, "-")
         Print #RptHandle,
         Lines = 5
      End If
    End If
  End If
  GLDebitTotal = OldRound(GLDebitTotal + EDSAmt(1))
  GLCreditTotal = OldRound(GLCreditTotal + EDMAmt(1))
  
  Return
  
PrintFundSummary:
  If Lines >= Max - 10 Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #RptHandle, String$(80, "-")
    Lines = 4
  End If
  
  MedTot = TOTMEDTAX + TOTMEDMat
  SocTot = TOTSOCTAX + TOTSOCMat
  RETTOT = TOTRetTax + TOTRETMat
  FSMTot = TOTFEDTAX + MedTot + SocTot
  
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
  Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
  Print #RptHandle, String$(80, "-")
  Lines = Lines + 9
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Lines = 7
  End If
  Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
  Print #RptHandle,
  Lines = Lines + 2
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = 9
  End If
  Print #RptHandle, " Federal"; Tab(19); Using(Image3$, TOTFEDTAX); Tab(71); Using$(Image3$, TOTFEDTAX)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = 9
  End If
  Print #RptHandle, " Social Security"; Tab(19); Using(Image3$, TOTSOCTAX); Tab(46); Using(Image3$, TOTSOCMat); Tab(71); Using$(Image3$, SocTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = 9
  End If
  Print #RptHandle, " Medicare"; Tab(19); Using(Image3$, TOTMEDTAX); Tab(46); Using(Image3$, TOTMEDMat); Tab(71); Using$(Image3$, MedTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = 9
  End If
  Print #RptHandle, " Sub Total"; Tab(71); Using(Image3$, FSMTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = 9
  End If
  Print #RptHandle,
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = 9
  End If
  Print #RptHandle, " State"; Tab(19); Using(Image3$, TOTSTATAX); Tab(71); Using$(Image3$, TOTSTATAX)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #RptHandle,
    Lines = 9
  End If
  Print #RptHandle, " Retirement"; Tab(19); Using(Image3$, TOTRetTax); Tab(46); Using(Image3$, TOTRETMat); Tab(71); Using$(Image3$, RETTOT)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Lines = 6
  End If
  Print #RptHandle, String$(80, "-")
  Lines = Lines + 1
  If Lines >= Max Then
    Print #RptHandle, FF$
    Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
    Print #RptHandle, "G/L Interface Report Summary."
    Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #RptHandle, String$(80, "-")
    Lines = 7
  End If
  Print #RptHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
  Print #RptHandle,
  For z = 1 To 50
    If QPTrim$(DedCodes(z).DCDESC1) <> "" Then
      Print #RptHandle, Tab(2); QPTrim$(DedCodes(z).DCDESC1); Tab(19); Using(Image3$, TotDedAmt(z));
      Lines = Lines + 1
    End If
    If z = 50 Then Exit For
    If QPTrim$(DedCodes(z + 1).DCDESC1) <> "" Then
      Print #RptHandle, Tab(51); QPTrim$(DedCodes(z + 1).DCDESC1); Tab(71); Using(Image3$, TotDedAmt(z + 1))
      z = z + 1
    End If
    If Lines >= Max Then
      Print #RptHandle, FF$
      Print #RptHandle, QPTrim$(Unit(1).UFEMPR)
      Print #RptHandle, "G/L Interface Report Summary."
      Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
      Print #RptHandle, String$(80, "-")
      Print #RptHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
      Print #RptHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
      Print #RptHandle, String$(80, "-")
      Print #RptHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
      Print #RptHandle,
      Lines = 8
    End If
  Next z
 
 Return
  
  
GLIFTotals:
  Print #RptHandle,
  Print #RptHandle, Dash0(1)
  Print #RptHandle, "   Totals:"; Tab(39); Using(Image3$, TotDebit#); Tab(54); Using(Image3$, TotCredit#)
  If TotalAccts = 0 Then
    Print #RptHandle,
    Print #RptHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory."
  End If
  Print #RptHandle,
  Print #RptHandle, '"SubTotals:"
  Return
  
GLIFHeader:
  Print #RptHandle, Unit(1).UFEMPR
  Print #RptHandle, "G/L Interface Report."
  Print #RptHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  Print #RptHandle,
  Print #RptHandle, "Account No.      Description               Debit         Credit"
  Print #RptHandle, Dash0(1)
  Lines = 6
  
  Return
PrintPayRollHeader:
  RSet Pg(1) = Str$(Page)
  Print #RHandle, QPTrim$(Unit(1).UFEMPR) + Space$(87) + "Page:" + Pg(1)
  Print #RHandle, "Payroll Register"
  Print #RHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  Print #RHandle,
  Print #RHandle, "Employee No   Name                                                                                                               EIC"
  Print #RHandle, "  Base Rate   O/T Rate   Tax Frng    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid   O/T Comp"
  Print #RHandle, ETitle$
  For x = 1 To 5
    If Len(QPTrim$(DTitle(x))) > 0 Then
      Print #RHandle, DTitle$(x)
    End If
  Next x
  Print #RHandle, Dash1(1)
  LincCnt = LineCnt + 9
  Page = Page + 1
  Return
  
SumAndPrintTime:
  RegHrs# = OldRound(RegHrs# + TransRec(1).RegHrsWork)
  VACHRS# = OldRound(VACHRS# + TransRec(1).VacUsed)
  SICKHRS# = OldRound(SICKHRS# + TransRec(1).SickUsed)
  HOLHRS# = OldRound(HOLHRS# + TransRec(1).HOLHOURS)
  COMPHRS# = OldRound(COMPHRS# + TransRec(1).CompUsed)
  PerHours# = OldRound(PerHours# + TransRec(1).PerHours)
  
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).PerHours)

  'added for EIC   6/07/94
  TotEIC# = OldRound(TotEIC# + TransRec(1).EICAmt)
  
  '-=-=-=-=-=-=-=
  TotHrs# = OldRound(TotHrs# + TransRec(1).OTHours)
  If TransRec(1).OTHrsPaid > 0 Then
    TOTPaid# = OldRound(TOTPaid# + TransRec(1).OTHrsPaid)
  End If
  TOTComp# = OldRound(TOTComp# + TransRec(1).OT2Comp)
  
  TRegWage# = OldRound(TRegWage# + TransRec(1).TotRegWage)
  
  If TransRec(1).TotOTWage > 0 Then
    TOTWage# = OldRound(TOTWage# + TransRec(1).TotOTWage)
  End If
  GPay# = OldRound(GPay# + TransRec(1).GrossPay)
  SSTax# = OldRound(SSTax# + TransRec(1).SocTaxAmt)
  MTax# = OldRound(MTax# + TransRec(1).MedTaxAmt)
  FTax# = OldRound(FTax# + TransRec(1).FedTaxAmt)
  STax# = OldRound(STax# + TransRec(1).StaTaxAmt)
  If TransRec(1).RetireAmt > 0 Then
    RETTOT# = OldRound(RETTOT# + TransRec(1).RetireAmt)
  End If
  
  TNetPay# = OldRound(TNetPay# + TransRec(1).NetPay)
  GFedGross# = OldRound(GFedGross# + TransRec(1).FedGrossPay)
  GStaGross# = OldRound(GStaGross# + TransRec(1).StaGrossPay)
  GSocGross# = OldRound(GSocGross# + TransRec(1).SocGrossPay)
  GMedGross# = OldRound(GMedGross# + TransRec(1).MedGrossPay)
  GRetGross# = OldRound(GRetGross# + TransRec(1).RetGrossPay)
  
  GTaxFrn# = OldRound(GTaxFrn# + TransRec(1).TaxFring)
  LSet ENumb(1) = LTrim$(EmpRec1(1).EmpNo)
  LSet EName(1) = QPTrim$(EmpRec1(1).EmpLName) + ", " + QPTrim$(EmpRec1(1).EmpFName)
  RSet BRat(1) = Using(Image3$, TransRec(1).BaseRate)
  RSet ORat(1) = Using(Image3$, TransRec(1).OTRate)
  
  RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsWork)
  
  RSet VHrs(1) = Using(Image$, TransRec(1).VacUsed)
  RSet SHrs(1) = Using(Image$, TransRec(1).SickUsed)
  RSet HHrs(1) = Using(Image$, TransRec(1).HOLHOURS)
  RSet CHrs(1) = Using(Image$, TransRec(1).CompUsed)
  RSet THrs(1) = Using(Image$, TransRec(1).RegHrsPaid)
  
  RSet TaxFrn(1) = Using(Image$, TransRec(1).TaxFring)

  RSet PHrs(1) = Using(Image$, TransRec(1).PerHours)

  RSet OTPaid(1) = Using(Image$, TransRec(1).OTHrsPaid)
  RSet OTComp(1) = Using(Image$, TransRec(1).OT2Comp)
  
  'added for EIC     6/07/94
  RSet EEicP(1) = Using(Image3$, TransRec(1).EICAmt)
  
  Select Case TransRec(1).PayType
  Case "S"
    RSet RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  Case Else
    HrlCnt = HrlCnt + 1
  End Select
  
  '=======
  RSet RErnP(1) = Using(Image3$, TransRec(1).TotRegWage)
  RSet OErnP(1) = Using(Image3$, TransRec(1).TotOTWage)
  
  RSet GPayP(1) = Using(Image3$, TransRec(1).GrossPay)
  RSet SSTaxP(1) = Using(Image3$, TransRec(1).SocTaxAmt)
  RSet MTaxP(1) = Using(Image3$, TransRec(1).MedTaxAmt)
  RSet FTaxP(1) = Using(Image3$, TransRec(1).FedTaxAmt)
  RSet STaxP(1) = Using(Image3$, TransRec(1).StaTaxAmt)
  
  RSet RetirP(1) = Using(Image3$, TransRec(1).RetireAmt)
  
  RSet NetPayP(1) = Using(Image3$, TransRec(1).NetPay)
  
  For x = 1 To 5
    SumDed$(x) = ""
  Next x
  
  Nextx = 1
  tripCnt = 1
  For Cnt2 = 1 To NumOfDeds
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    TotDeds#(Cnt2) = OldRound(TotDeds#(Cnt2) + TransRec(1).DAmt(Cnt2))
    RSet Ded(1) = Using(Image3$, TransRec(1).DAmt(Cnt2))
    SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
    tripCnt = tripCnt + 1
  Next
  
  '----------------------------------------------
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    TotErns(Cnt2) = OldRound(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    RSet Ern(1) = Using(Image3$, TransRec(1).EAmt(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  '-------------------------------------------------------
  
  RSet EMPLine$(1) = EEicP(1)
  Mid$(EMPLine$(1), 1) = ENumb(1) + EName(1)
  Print #RHandle, EMPLine$(1)
  Print #RHandle, BRat(1) + ORat(1) + TaxFrn(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1) + CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + OTComp(1)
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1) + RetirP(1) + NetPayP(1)
  For x = 1 To 5
    If Len(QPTrim$(SumDed$(x))) > 0 Then
      Print #RHandle, SumDed$(x)
    End If
  Next x
  Print #RHandle,
  
  Return
  
PrintSumTotal:
  RSet SCnt(1) = Using(Image0$, SalCnt)
  RSet HCnt(1) = Using(Image0$, HrlCnt)
  
  RSet Fill11(1) = Using(Image3$, GTaxFrn#)
  RSet THrs(1) = Using(Image3$, TotalHrs#)
  RSet RHrs(1) = Using(Image3$, RegHrs#)
  RSet VHrs(1) = Using(Image3$, VACHRS#)
  RSet SHrs(1) = Using(Image3$, SICKHRS#)
  RSet HHrs(1) = Using(Image3$, HOLHRS#)
  RSet CHrs(1) = Using(Image3$, COMPHRS#)

  RSet PHrs(1) = Using(Image3$, PerHours#)

  RSet OTPaid(1) = Using(Image3$, TOTPaid#)
  RSet OTComp(1) = Using(Image3$, TOTComp#)
  
  RSet RErnP(1) = Using(Image3$, TRegWage#)
  RSet OErnP(1) = Using(Image3$, TOTWage#)
  
  RSet GPayP(1) = Using(Image3$, GPay#)
  RSet SSTaxP(1) = Using(Image3$, SSTax#)
  RSet MTaxP(1) = Using(Image3$, MTax#)
  RSet FTaxP(1) = Using(Image3$, FTax#)
  RSet STaxP(1) = Using(Image3$, STax#)
  RSet RetirP(1) = Using(Image3$, RETTOT#)
  RSet NetPayP(1) = Using(Image3$, TNetPay#)
  
  For x = 1 To 5
    SumDed$(x) = ""
  Next x
  
  Nextx = 1
  tripCnt = 1
  For Cnt2 = 1 To NumOfDeds
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    RSet Ded(1) = Using(Image3$, TotDeds#(Cnt2))
    SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
    tripCnt = tripCnt + 1
  Next
  '---------------------------------------------------------
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    RSet Ern(1) = Using(Image3$, TotErns(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  
  '--------------NEW----------------------------
  RSet Pg(1) = Str$(Page)
  Print #RHandle, FF$
  Print #RHandle, QPTrim$(Unit(1).UFEMPR) + Space$(87) + "Page:" + Pg(1)
  Print #RHandle, "Payroll Register Summary"
  Print #RHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  Print #RHandle,
  Print #RHandle, Dash1(1)
  Print #RHandle,
  Print #RHandle, "   Salaried     Hourly   Tax Frng    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid   O/T Comp"
  Print #RHandle, SCnt(1) + HCnt(1) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1) + CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + OTComp(1)
'  Print #RHandle, CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + OTComp(1)'commented out on 08/12/2003
  Print #RHandle,
  Print #RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay"
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1) + RetirP(1) + NetPayP(1)
  Print #RHandle,
  For x = 1 To 5
    If Len(QPTrim(DTitle(x))) > 0 Then
      Print #RHandle, DTitle$(x)
      Print #RHandle, SumDed$(x)
    End If
  Next x
  Print #RHandle,
  Print #RHandle, "  Fed Gross  Sta Gross  Med Gross  Soc Gross  Ret Gross  EIC Total"
  RSet FTaxP(1) = Using(Image5$, GFedGross#)
  RSet STaxP(1) = Using(Image5$, GStaGross#)
  RSet MTaxP(1) = Using(Image5$, GMedGross#)
  RSet SSTaxP(1) = Using(Image5$, GSocGross#)
  RSet RetirP(1) = Using(Image5$, GRetGross#)
  RSet EEicP(1) = Using(Image5$, TotEIC#)
  
  Print #RHandle, FTaxP(1) + STaxP(1) + MTaxP(1) + SSTaxP(1) + RetirP(1) + EEicP(1)
  
  Print #RHandle,
  Print #RHandle, Dash1(1)
  Print #RHandle, FF$
  
  Return
  
  '-----------------------------------------------------------------------
ParseDistributions:
  
  ReDim EmpAct(1 To 14) As DistWageRptType
  ENumOfAct = 0
  
  'process wage distributions
  For Cnt2 = 1 To 8
    Acct$ = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
    If Len(Acct$) > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).Pct = TransRec(1).TDist(Cnt2).DPct
      EmpAct(ENumOfAct).RHrs = TransRec(1).TDist(Cnt2).DRHrs
      EmpAct(ENumOfAct).OHrs = TransRec(1).TDist(Cnt2).DOHrs
      EmpAct(ENumOfAct).RWage = TransRec(1).TDist(Cnt2).DRWage
      EmpAct(ENumOfAct).OWage = TransRec(1).TDist(Cnt2).DOWage
      EmpAct(ENumOfAct).GrossPay = OldRound(EmpAct(ENumOfAct).RWage + EmpAct(ENumOfAct).OWage)
    End If
  Next
  
  'process earnings distributions
  For Cnt2 = 1 To 6
    Acct$ = QPTrim$(TransRec(1).EDist(Cnt2).EAcct)
    If Len(Acct$) > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).AddEarn = TransRec(1).EDist(Cnt2).EAmt
      EmpAct(ENumOfAct).GrossPay = TransRec(1).EDist(Cnt2).EAmt
    End If
  Next
  
  'HERE: got all accts for this employee
  
  First = 1
  Start = 1
  
  Last = ENumOfAct
  
  'purge and sum employee's dupelicate account distributions
  
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If EmpAct(First).Acct = EmpAct(Cnt2).Acct Then
        LSet EmpAct(Cnt2).Acct = ""             'blank acct num as a flag
        EmpAct(First).Pct = OldRound(EmpAct(First).Pct + EmpAct(Cnt2).Pct)
        EmpAct(First).RHrs = OldRound(EmpAct(First).RHrs + EmpAct(Cnt2).RHrs)
        EmpAct(First).OHrs = OldRound(EmpAct(First).OHrs + EmpAct(Cnt2).OHrs)
        EmpAct(First).RWage = OldRound(EmpAct(First).RWage + EmpAct(Cnt2).RWage)
        EmpAct(First).OWage = OldRound(EmpAct(First).OWage + EmpAct(Cnt2).OWage)
        EmpAct(First).AddEarn = OldRound(EmpAct(First).AddEarn + EmpAct(Cnt2).AddEarn)
        EmpAct(First).GrossPay = OldRound(EmpAct(First).GrossPay + EmpAct(Cnt2).GrossPay)
      End If
    Next
Again:
    First = First + 1
  Loop Until First >= Last
  
  'calc percentages of matching amts to each account distribution
  
  For Cnt2 = 1 To ENumOfAct
    EmpAct(Cnt2).MATSocAmt = OldRound(TransRec(1).MatchSocAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATMedAmt = OldRound(TransRec(1).MatchMedAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATRetAmt = OldRound(TransRec(1).MatchRetAmt * (EmpAct(Cnt2).Pct * 0.01))
  Next
  
  '---------------------------------------------------------------------------
  'calc and adjust matching distribution amts
  'find last active account
  'adjust Social Amt
  
  Do
    TotalSocAmt# = 0
    For Cnt2 = 1 To 14          '8
      TotalSocAmt# = OldRound(TotalSocAmt# + EmpAct(Cnt2).MATSocAmt)
    Next
    If TotalSocAmt# = 0 Then GoTo SkipSocDist
    If TotalSocAmt# <> TransRec(1).MatchSocAmt Then
      For Cnt3 = 14 To 1 Step -1                '8 TO 1 STEP -1
        If EmpAct(Cnt3).MATSocAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalSocAmt# > TransRec(1).MatchSocAmt Then
        DistDif# = OldRound(TotalSocAmt# - TransRec(1).MatchSocAmt)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt - DistDif#)
      ElseIf TotalSocAmt# < TransRec(1).MatchSocAmt Then
        DistDif# = OldRound(TransRec(1).MatchSocAmt - TotalSocAmt#)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt + DistDif#)
      End If
    End If
  Loop Until TotalSocAmt# = OldRound(TransRec(1).MatchSocAmt)
  '-=-=-=-=-=-
  'adjust Medicare Amt
SkipSocDist:
  Do
    TotalMedAmt# = 0
    For Cnt2 = 1 To 8
      TotalMedAmt# = OldRound(TotalMedAmt# + EmpAct(Cnt2).MATMedAmt)
    Next
    If TotalMedAmt# = 0 Then GoTo SkipMedDist
    If TotalMedAmt# <> TransRec(1).MatchMedAmt Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATMedAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalMedAmt# > TransRec(1).MatchMedAmt Then
        DistDif# = OldRound(TotalMedAmt# - TransRec(1).MatchMedAmt)
        EmpAct(LastActive).MATMedAmt = OldRound(EmpAct(LastActive).MATMedAmt - DistDif#)
      ElseIf TotalMedAmt# < TransRec(1).MatchMedAmt Then
        DistDif# = OldRound(TransRec(1).MatchMedAmt - TotalMedAmt#)
        EmpAct(LastActive).MATMedAmt = OldRound(EmpAct(LastActive).MATMedAmt + DistDif#)
      End If
    End If
  Loop Until TotalMedAmt# = OldRound(TransRec(1).MatchMedAmt)
  '-=-=-=-=-=-
SkipMedDist:
  'adjust Retire Amt
  Do
    TotalRetAmt# = 0
    For Cnt2 = 1 To 8
      TotalRetAmt# = OldRound(TotalRetAmt# + EmpAct(Cnt2).MATRetAmt)
    Next
    If TotalRetAmt# = 0 Then GoTo SkipRetDist
    If TotalRetAmt# <> TransRec(1).MatchRetAmt Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATRetAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalRetAmt# > TransRec(1).MatchRetAmt Then
        DistDif# = OldRound(TotalRetAmt# - TransRec(1).MatchRetAmt)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt - DistDif#)
      ElseIf TotalRetAmt# < TransRec(1).MatchRetAmt Then
        DistDif# = OldRound(TransRec(1).MatchRetAmt - TotalRetAmt#)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt + DistDif#)
      End If
    End If
  Loop Until TotalRetAmt# = OldRound(TransRec(1).MatchRetAmt)

SkipRetDist:
  
  'print this employee's distributions
  Print #DHandle, ENumb(1) + EName(1) + BRat(1) + ORat(1)
  DLineCnt = DLineCnt + 1
  For Cnt2 = 1 To Last
    If Len(QPTrim$(EmpAct(Cnt2).Acct)) > 0 Then
      GoSub PrintEmpDist
    End If
  Next
  
  If Last > 1 Then
    GoSub PrintEmpSubTotal
  End If
  
  Print #DHandle,
  DLincCnt = DLineCnt + 1
  
  ' sum to master distrubtion list
  
  For Cnt2 = 1 To Last          'process wage distributions
    Acct$ = QPTrim$(EmpAct(Cnt2).Acct)
    If Len(Acct$) > 0 Then
      If NumOfWageAccts > 0 Then
        For Cnt3 = 1 To NumOfWageAccts
          If Acct$ = QPTrim$(DistbSumAccts(Cnt3).Acct) Then
            DistbSumAccts(Cnt3).RWage = OldRound(DistbSumAccts(Cnt3).RWage + EmpAct(Cnt2).RWage)
            DistbSumAccts(Cnt3).OWage = OldRound(DistbSumAccts(Cnt3).OWage + EmpAct(Cnt2).OWage)
            DistbSumAccts(Cnt3).RHrs = OldRound(DistbSumAccts(Cnt3).RHrs + EmpAct(Cnt2).RHrs)
            DistbSumAccts(Cnt3).OHrs = OldRound(DistbSumAccts(Cnt3).OHrs + EmpAct(Cnt2).OHrs)
            DistbSumAccts(Cnt3).AddEarn = OldRound(DistbSumAccts(Cnt3).AddEarn + EmpAct(Cnt2).AddEarn)
            DistbSumAccts(Cnt3).GrossPay = OldRound(DistbSumAccts(Cnt3).GrossPay + EmpAct(Cnt2).GrossPay)
            DistbSumAccts(Cnt3).MATSocAmt = OldRound(DistbSumAccts(Cnt3).MATSocAmt + EmpAct(Cnt2).MATSocAmt)
            DistbSumAccts(Cnt3).MATMedAmt = OldRound(DistbSumAccts(Cnt3).MATMedAmt + EmpAct(Cnt2).MATMedAmt)
            DistbSumAccts(Cnt3).MATRetAmt = OldRound(DistbSumAccts(Cnt3).MATRetAmt + EmpAct(Cnt2).MATRetAmt)
            Exit For
          End If
        Next
        If Cnt3 > NumOfWageAccts Then
          GoSub AddDistbSumAcct 'add new sum dist acct
        End If
      Else      'no previous sum accts. add new one
        GoSub AddDistbSumAcct   'add new sum dist acct
      End If
    End If
  Next
  
  Return
  
AddDistbSumAcct:                'add amts to grand total acts summary
  
  NumOfWageAccts = NumOfWageAccts + 1
  If NumOfWageAccts > 1 Then
    ReDim Preserve DistbSumAccts(1 To NumOfWageAccts) As DistWageRptType
  End If
  DistbSumAccts(NumOfWageAccts) = EmpAct(Cnt2)
  
  Return
PrintDistHeader:
  RSet Pg(1) = Str$(EPage)
  Print #DHandle, QPTrim$(Unit(1).UFEMPR) + Space$(87) + "Page:" + Pg(1)
  Print #DHandle, "Earnings Distribution"
  Print #DHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  Print #DHandle,
  Print #DHandle, "Employee No  Name                                Base Rate   O/T Rate                          --------- Matching ----------"
  Print #DHandle, "Account Number       Sal%    Reg Hrs    O/T Hrs    Reg Pay    O/T Pay  Tot Other  Gross Pay    Soc Sec   Medicare     Retire"
  Print #DHandle, Dash2(1)
  DLineCnt = DLineCnt + 7
  EPage = EPage + 1
  
  Return
  
PrintEmpDist:
  LSet EDAct(1) = EmpAct(Cnt2).Acct
  RSet EDPct(1) = Using(Image$, EmpAct(Cnt2).Pct)
  RSet EDRHrs(1) = Using(Image$, EmpAct(Cnt2).RHrs)
  RSet EDOHrs(1) = Using(Image$, EmpAct(Cnt2).OHrs)
  RSet EDRPay(1) = Using(Image3$, EmpAct(Cnt2).RWage)
  RSet EDOPay(1) = Using(Image$, EmpAct(Cnt2).OWage)
  RSet EDEarn(1) = Using(Image$, EmpAct(Cnt2).AddEarn)
  RSet EDGroP(1) = Using(Image$, EmpAct(Cnt2).GrossPay)
  
  RSet EDSAmt(1) = Using(Image$, EmpAct(Cnt2).MATSocAmt)
  RSet EDMAmt(1) = Using(Image$, EmpAct(Cnt2).MATMedAmt)
  RSet EDRAmt(1) = Using(Image$, EmpAct(Cnt2).MATRetAmt)
  
  Print #DHandle, EDAct(1) + EDPct(1) + EDRHrs(1) + EDOHrs(1) + EDRPay(1) + EDOPay(1) + EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1)
  DLineCnt = DLineCnt + 1
  
'----------Void Check Code----------------------------
  '6/22/2004
  RETTOT# = RETTOT#
  TNetPay# = TNetPay#
  
  'Fed Tax (and state and SS) comes thru as the same amount no matter how many loops are run...
  'That amount should be saved once because when Void Check posts to the GL
  'for each check only one amount needs to be posted for only one GL
  
  'The matching amounts are saved differently...with each loop the debit
  'amounts are saved for their individual GL number but the credit amounts are
  'accumulated and saved to the first fund number + GL suffix. All debit
  'amounts are accumulated for each fund number and saved as one credit
  'amount for that fund number
  ThisFACnt = FACnt
  
  FundNumOnly = Mid(EDAct(1), 1, FundLen)
  AcctNumCnt = AcctNumCnt + 1
  If InStr(Mid(EDAct(1), 1, ThisFACnt), "-") Then
    ThisFACnt = ThisFACnt + 1
  End If
  
  TempVoid.EmpNum = ENumb(1)
  ThisPR = EDGroP(1)
  
  TempVoid.WagesAmt = EDGroP(1)
  TotalWages = TotalWages + EDGroP(1)
  TempVoid.WagesGL = EDAct(1)
  
  'Non-split Pool...this type has a pool fund (usually fund 10) that serves
  'as the fund from which all WH and deductions are taken. This fund can also serve
  'as a fund from which an employee can be paid (just like fund 30, 11, etc.).
  'When the pool fund also serves as a typical 'paid from' fund then this code
  'handles it without having to run through the sub 'NoPRNetForPool'. However,
  'if the program does not detect the pool fund also serving as a 'Paid From'
  'fund (PRNetPoolFound = True) then the program must activate the sub 'NoPRNetForPool'
  'to insert the activity required for the pool fund.
  
  'Non-split Pool with Central Depository...This case is just like without Central
  'depository except that the Central Depository fund (usually 01) must be factored
  'in to the mix. The code below handles employee data when there is a pool fund also
  'serving as a 'Pay From' fund. However, if there is no pool fund detected (PRNetPoolFound = False)
  'then the sub 'NoPRNetForPoolCOrI' must be activated in order to record pay activity
  'for the overlooked pool fund.
  
  If PoolFundNum$ = Mid(EDAct(1), 1, 2) Then
    If PRPoolProcessed = True Then
        TempVoid.FEDWHAmt = 0
        TempVoid.MEDWHAmt = 0
        TempVoid.SOCWHAmt = 0
        TempVoid.RETWHAmt = 0
        TempVoid.STAWHAmt = 0
        GoTo PoolAlreadyDone
    Else
      PRNetPoolFound = True 'dictates if the PRNetPoolFound sub or PRNetPoolFoundCOrI sub activates
    End If
  End If
  
  '--------9/17/04--------
  If PoolFundNum$ = Mid(EDAct(1), 1, 2) Then 'we found the pool fund through normal activity (pool fund
  'is also being used as a 'PayFrom' fund) so go ahead and assign WH values to this fund
    TempVoid.FEDWHAmt = TransRec(1).FedTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).FedTaxAmt
    TempVoid.MEDWHAmt = TransRec(1).MedTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).MedTaxAmt
    TempVoid.SOCWHAmt = TransRec(1).SocTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).SocTaxAmt
    TempVoid.RETWHAmt = TransRec(1).RetireAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).RetireAmt
    TempVoid.STAWHAmt = TransRec(1).StaTaxAmt
    TotalWHAndDeds = TotalWHAndDeds + TransRec(1).StaTaxAmt
  Else
    TempVoid.FEDWHAmt = 0
    TempVoid.MEDWHAmt = 0
    TempVoid.SOCWHAmt = 0
    TempVoid.RETWHAmt = 0
    TempVoid.STAWHAmt = 0
  End If
PoolAlreadyDone:
  ThisPR = OldRound(ThisPR - TempVoid.FEDWHAmt)
  'next employee data comes thru
  TempVoid.FEDWHGL = QPTrim$(SysRec(1).Liab(1).Acct)
  TempVoid.FEDWHGL = ReplaceString(TempVoid.FEDWHGL, "-", "")
  TempVoid.FEDWHGL = AddDashesToGLNumber(TempVoid.FEDWHGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.MEDWHAmt)
  'next employee data comes thru
  TempVoid.MEDWHGL = QPTrim$(SysRec(1).Liab(4).Acct)
  TempVoid.MEDWHGL = ReplaceString(TempVoid.MEDWHGL, "-", "")
  TempVoid.MEDWHGL = AddDashesToGLNumber(TempVoid.MEDWHGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATCRAmt = EDMAmt(1)
  TempVoid.MEDMATCRGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).MEDLIAB))
  TempVoid.MEDMATCRGL = ReplaceString(TempVoid.MEDMATCRGL, "-", "")
  TempVoid.MEDMATCRGL = AddDashesToGLNumber(TempVoid.MEDMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATDBAmt = EDMAmt(1)
  TempVoid.MEDMATDBGL = Mid(EDAct(1), 1, ThisFACnt) + QPTrim$(SysRec(1).MEDEXP)
  TempVoid.MEDMATDBGL = ReplaceString(TempVoid.MEDMATDBGL, "-", "")
  TempVoid.MEDMATDBGL = AddDashesToGLNumber(TempVoid.MEDMATDBGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.SOCWHAmt)
  TempVoid.SOCWHGL = QPTrim$(SysRec(1).Liab(3).Acct)
  TempVoid.SOCWHGL = ReplaceString(TempVoid.SOCWHGL, "-", "")
  TempVoid.SOCWHGL = AddDashesToGLNumber(TempVoid.SOCWHGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATCRAmt = EDSAmt(1)
  TempVoid.SOCMATCRGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).SOCLIAB))
  TempVoid.SOCMATCRGL = ReplaceString(TempVoid.SOCMATCRGL, "-", "")
  TempVoid.SOCMATCRGL = AddDashesToGLNumber(TempVoid.SOCMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATDBAmt = EDSAmt(1)
  TempVoid.SOCMATDBGL = Mid(EDAct(1), 1, ThisFACnt) + QPTrim$(SysRec(1).SOCEXP)
  TempVoid.SOCMATDBGL = ReplaceString(TempVoid.SOCMATDBGL, "-", "")
  TempVoid.SOCMATDBGL = AddDashesToGLNumber(TempVoid.SOCMATDBGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.RETWHAmt)
  TempVoid.RETWHGL = QPTrim$(SysRec(1).Liab(5).Acct)
  TempVoid.RETWHGL = ReplaceString(TempVoid.RETWHGL, "-", "")
  TempVoid.RETWHGL = AddDashesToGLNumber(TempVoid.RETWHGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATCRAmt = EDRAmt(1)
  TempVoid.RETMATCRGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).RETLIAB))
  TempVoid.RETMATCRGL = ReplaceString(TempVoid.RETMATCRGL, "-", "")
  TempVoid.RETMATCRGL = AddDashesToGLNumber(TempVoid.RETMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATDBAmt = EDRAmt(1)
  TempVoid.RETMATDBGL = Mid(EDAct(1), 1, ThisFACnt) + QPTrim$(SysRec(1).RETEXP)
  TempVoid.RETMATDBGL = ReplaceString(TempVoid.RETMATDBGL, "-", "")
  TempVoid.RETMATDBGL = AddDashesToGLNumber(TempVoid.RETMATDBGL, FundLen, AcctLen, DetLen)
  
  ThisPR = OldRound(ThisPR - TempVoid.STAWHAmt)
  TempVoid.STAWHGL = QPTrim$(SysRec(1).Liab(2).Acct)
  TempVoid.STAWHGL = ReplaceString(TempVoid.STAWHGL, "-", "")
  TempVoid.STAWHGL = AddDashesToGLNumber(TempVoid.STAWHGL, FundLen, AcctLen, DetLen)
  
  For x = 1 To 50
    If TransRec(1).DAmt(x) > 0 Then
      If PoolFundNum$ = Mid(EDAct(1), 1, 2) And PRPoolProcessed = False Then
        TempVoid.DedData(x).DAmt = TransRec(1).DAmt(x)
        TotalDeds = OldRound(TotalDeds + TransRec(1).DAmt(x))
        TotalWHAndDeds = TotalWHAndDeds + TransRec(1).DAmt(x)
      Else
        TempVoid.DedData(x).DAmt = 0
      End If
      TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedCodeDesc(x))
      TempVoid.DedData(x).DedGLNum = QPTrim$(DedCodeNums(x))
      TempVoid.DedData(x).DedGLNum = ReplaceString(TempVoid.DedData(x).DedGLNum, "-", "")
      TempVoid.DedData(x).DedGLNum = AddDashesToGLNumber(TempVoid.DedData(x).DedGLNum, FundLen, AcctLen, DetLen)
      ThisPR = OldRound(ThisPR - TempVoid.DedData(x).DAmt)
    Else
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DedDesc = ""
      TempVoid.DedData(x).DedGLNum = ""
    End If
  Next x
  
  TempVoid.NumOfAccts = ENumOfAct
  
  If ThisPR < 0 Then
    If Mid(EDAct(1), 1, FundLen) <> Mid(SysRec(1).Liab(3).Acct, 1, FundLen) Then
      DbtCnt = DbtCnt + 1
      ReDim Preserve ThisPRDbtFund(1 To DbtCnt) As String
      ThisPRDbtFund(DbtCnt) = Mid(EDAct(1), 1, FundLen)
      ReDim Preserve ThisPRDbtAmt(1 To DbtCnt) As Double
      ThisPRDbtAmt(DbtCnt) = TempVoid.WagesAmt
      ThisPR = OldRound(TempVoid.FEDWHAmt + TempVoid.MEDWHAmt + TempVoid.SOCWHAmt + TempVoid.RETWHAmt + TempVoid.STAWHAmt + TotalDeds)
      ThisPR = -ThisPR
      TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
      TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
      TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
    Else
      TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
      TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
      TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
    End If
  Else
    TempVoid.PRNetGL = Mid(EDAct(1), 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
    TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
    TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
  End If
  
  If DbtCnt > 0 Then
    For x = 1 To DbtCnt
      If QPTrim$(ThisPRDbtFund(x)) = Mid(TempVoid.PRNetGL, 1, FundLen) Then
        TempVoid.PRNet = OldRound(ThisPR + ThisPRDbtAmt(x))
        ThisPRDbtAmt(x) = 0
        Exit For
      End If
    Next x
    If x > DbtCnt Then TempVoid.PRNet = ThisPR
  Else
    TempVoid.PRNet = ThisPR
  End If
  TempVoid.TransRec = 0
  TempVoid.VoidFlag = False
  TempVoid.CheckAmt = 0
  TempVoid.CheckDate = 0
  TempVoid.CheckNum = 0
  TempVoid.Type = QPTrim$(SysRec(1).USEIMP)
  TempVoid.Pad = ""
  If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
    TempVoid.PPEAmt = ThisPR
    ThisCRGL = ReplaceString(SysRec(1).ICRACCT, "-", "")
    If FundPad > 0 Then
      If Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) <> "" Then
       ThisCRGL = QPTrim$(ThisCRGL) + String$(FundPad, "0") + FundNumOnly
        TempVoid.PPEGL = ThisCRGL
      ElseIf Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) = "" Then
        ThisCRGL = QPTrim$(ThisCRGL + Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen)) + FundNumOnly + String$(FundPad, "0")
        TempVoid.PPEGL = ThisCRGL
      End If
    Else
      TempVoid.PPEGL = QPTrim$(SysRec(1).ICRACCT) + FundNumOnly
    End If
    TempVoid.PPEGL = ReplaceString(TempVoid.PPEGL, "-", "")
    TempVoid.PPEGL = AddDashesToGLNumber(TempVoid.PPEGL, FundLen, AcctLen, DetLen)
    
    TempVoid.PPETotAmt = ThisPR
    TempVoid.PPETotGL = QPTrim$(SysRec(1).IDRACCT)
    TempVoid.PPETotGL = ReplaceString(TempVoid.PPETotGL, "-", "")
    TempVoid.PPETotGL = AddDashesToGLNumber(TempVoid.PPETotGL, FundLen, AcctLen, DetLen)
  Else
    TempVoid.PPEAmt = 0
    TempVoid.PPEGL = ""
    TempVoid.PPETotAmt = 0
    TempVoid.PPETotGL = ""
  End If
  TVCnt = TVCnt + 1
  If PRNetPoolFound = True Then PRPoolProcessed = True
  Put TVHandle, TVCnt, TempVoid
  
  Return
'--------------------^^^^--Void Check Code---------
'----------------added 9/20/04---------------
NoPRNetForPool:
  For x = 1 To AcctNumCnt - 1
    Get TVHandle, TVCnt - x, TempVoid
      TempVoid.NumOfAccts = AcctNumCnt + 1 'the program looks at the number of
      'iterations for this employee to know which records to use if this check
      'is ever voided. Each iteration represents data coming through for a 'Paid
      'From' fund. If the pool fund is not a 'Paid From' fund then the number
      'of iterations saved so far will be one short because the code in this
      'sub represents another iteration. So we have to go back to the records
      'already saved for this paycheck and increase the iterations by one.
    Put TVHandle, TVCnt - x, TempVoid
  Next x
  Get TVHandle, TVCnt, TempVoid
  TempVoid.NumOfAccts = AcctNumCnt + 1
  Put TVHandle, TVCnt, TempVoid
  TotalWHAndDeds = 0
  TempVoid.NumOfAccts = AcctNumCnt + 1
  TempVoid.EmpNum = ENumb(1)
  TempVoid.TransRec = 0
  TempVoid.VoidFlag = False
  TempVoid.CheckAmt = 0
  TempVoid.CheckDate = 0
  TempVoid.CheckNum = 0
  TempVoid.PPEAmt = 0
  TempVoid.PPEGL = ""
  TempVoid.PPETotAmt = 0
  TempVoid.PPETotGL = ""
  For x = 1 To 50
    If TransRec(1).DAmt(x) > 0 Then 'all deductions always go into the pool fund
    'even if the pool fund is not a fund from which PRNet is naturally created
      TempVoid.DedData(x).DAmt = TransRec(1).DAmt(x)
      TotalDeds = OldRound(TotalDeds + TransRec(1).DAmt(x))
      TotalWHAndDeds = TotalWHAndDeds + TransRec(1).DAmt(x)
      TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedCodeDesc(x))
      TempVoid.DedData(x).DedGLNum = QPTrim$(DedCodeNums(x))
      TempVoid.DedData(x).DedGLNum = ReplaceString(TempVoid.DedData(x).DedGLNum, "-", "")
      TempVoid.DedData(x).DedGLNum = AddDashesToGLNumber(TempVoid.DedData(x).DedGLNum, FundLen, AcctLen, DetLen)
      ThisPR = OldRound(ThisPR - TempVoid.DedData(x).DAmt)
    Else
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DedDesc = ""
      TempVoid.DedData(x).DedGLNum = ""
    End If
  Next x
  TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
  TempVoid.FEDWHAmt = TransRec(1).FedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).FedTaxAmt
  TempVoid.MEDWHAmt = TransRec(1).MedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).MedTaxAmt
  TempVoid.SOCWHAmt = TransRec(1).SocTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).SocTaxAmt
  TempVoid.RETWHAmt = TransRec(1).RetireAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).RetireAmt
  TempVoid.STAWHAmt = TransRec(1).StaTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).StaTaxAmt
  TempVoid.PRNet = -TotalWHAndDeds
  TempVoid.MEDMATCRAmt = 0
  TempVoid.MEDMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).MEDLIAB))
  TempVoid.MEDMATCRGL = ReplaceString(TempVoid.MEDMATCRGL, "-", "")
  TempVoid.MEDMATCRGL = AddDashesToGLNumber(TempVoid.MEDMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATDBAmt = 0
  TempVoid.MEDMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).MEDEXP)
  TempVoid.MEDMATDBGL = ReplaceString(TempVoid.MEDMATDBGL, "-", "")
  TempVoid.MEDMATDBGL = AddDashesToGLNumber(TempVoid.MEDMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATCRAmt = 0
  TempVoid.RETMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).RETLIAB))
  TempVoid.RETMATCRGL = ReplaceString(TempVoid.RETMATCRGL, "-", "")
  TempVoid.RETMATCRGL = AddDashesToGLNumber(TempVoid.RETMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATDBAmt = 0
  TempVoid.RETMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).RETEXP)
  TempVoid.RETMATDBGL = ReplaceString(TempVoid.RETMATDBGL, "-", "")
  TempVoid.RETMATDBGL = AddDashesToGLNumber(TempVoid.RETMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATCRAmt = 0
  TempVoid.SOCMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).SOCLIAB))
  TempVoid.SOCMATCRGL = ReplaceString(TempVoid.SOCMATCRGL, "-", "")
  TempVoid.SOCMATCRGL = AddDashesToGLNumber(TempVoid.SOCMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATDBAmt = 0
  TempVoid.SOCMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).SOCEXP)
  TempVoid.SOCMATDBGL = ReplaceString(TempVoid.SOCMATDBGL, "-", "")
  TempVoid.SOCMATDBGL = AddDashesToGLNumber(TempVoid.SOCMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.WagesAmt = 0
  TempVoid.WagesGL = ""
  TVCnt = TVCnt + 1
  Put TVHandle, TVCnt, TempVoid
  PRNetPoolFound = True
  Return

NoPRNetForPoolCOrI:
  GL4PPETotal = ""
  PRNetSum = 0
  
  For x = 1 To AcctNumCnt - 1 'the program depends on the number of iterations
  'occuring to collect all the paycheck data (1 iteration per 'Pay From' fund.
  'Because there is no pool fund iteration we must go back and adjust the number
  'of iterations up 1 to include the rest of the code in this sub.
    Get TVHandle, TVCnt - x, TempVoid
      PRNetSum = PRNetSum + TempVoid.PRNet 'we'll be using this number below
      'to determine the PRNet for the pool fund
      TempVoid.NumOfAccts = AcctNumCnt + 1
    Put TVHandle, TVCnt - x, TempVoid
  Next x
  Get TVHandle, TVCnt, TempVoid
  PRNetSum = PRNetSum + TempVoid.PRNet
  GL4PPETotal = TempVoid.PPETotGL 'this GL number is constant and therefore
  'works fine for this iteration
  TempVoid.NumOfAccts = AcctNumCnt + 1
  Put TVHandle, TVCnt, TempVoid
  ThisCRGL = ReplaceString(SysRec(1).ICRACCT, "-", "")
  If FundPad > 0 Then
    If Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) <> "" Then
     ThisCRGL = QPTrim$(ThisCRGL) + String$(FundPad, "0") + PoolFundNum$
      TempVoid.PPEGL = ThisCRGL
    ElseIf Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen) = "" Then
      ThisCRGL = QPTrim$(ThisCRGL + Mid(ThisCRGL, (1 + FundLen + AcctLen), DetLen)) + PoolFundNum$ + String$(FundPad, "0")
      TempVoid.PPEGL = ThisCRGL
    End If
  Else
    TempVoid.PPEGL = QPTrim$(SysRec(1).ICRACCT) + PoolFundNum$
  End If
  TempVoid.PPEGL = ReplaceString(TempVoid.PPEGL, "-", "")
  TempVoid.PPEGL = AddDashesToGLNumber(TempVoid.PPEGL, FundLen, AcctLen, DetLen)
  TotalWHAndDeds = 0
  TempVoid.NumOfAccts = AcctNumCnt + 1
  TempVoid.EmpNum = ENumb(1)
  TempVoid.TransRec = 0
  TempVoid.VoidFlag = False
  TempVoid.CheckAmt = 0
  TempVoid.CheckDate = 0
  TempVoid.CheckNum = 0
  For x = 1 To 50
    If TransRec(1).DAmt(x) > 0 Then 'all deductions always go into the pool fund
    'even if the pool fund is not a fund from which PRNet is naturally created
      TempVoid.DedData(x).DAmt = TransRec(1).DAmt(x)
      TotalDeds = OldRound(TotalDeds + TransRec(1).DAmt(x))
      TotalWHAndDeds = TotalWHAndDeds + TransRec(1).DAmt(x)
      TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedCodeDesc(x))
      TempVoid.DedData(x).DedGLNum = QPTrim$(DedCodeNums(x))
      TempVoid.DedData(x).DedGLNum = ReplaceString(TempVoid.DedData(x).DedGLNum, "-", "")
      TempVoid.DedData(x).DedGLNum = AddDashesToGLNumber(TempVoid.DedData(x).DedGLNum, FundLen, AcctLen, DetLen)
      ThisPR = OldRound(ThisPR - TempVoid.DedData(x).DAmt)
    Else
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DAmt = 0
      TempVoid.DedData(x).DedDesc = ""
      TempVoid.DedData(x).DedGLNum = ""
    End If
  Next x
  TempVoid.PRNetGL = Mid(SysRec(1).Liab(3).Acct, 1, FundLen) + QPTrim$(SysRec(1).CashAcct)
  TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
  TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
  TempVoid.FEDWHAmt = TransRec(1).FedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).FedTaxAmt
  TempVoid.MEDWHAmt = TransRec(1).MedTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).MedTaxAmt
  TempVoid.SOCWHAmt = TransRec(1).SocTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).SocTaxAmt
  TempVoid.RETWHAmt = TransRec(1).RetireAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).RetireAmt
  TempVoid.STAWHAmt = TransRec(1).StaTaxAmt
  TotalWHAndDeds = TotalWHAndDeds + TransRec(1).StaTaxAmt
  TempVoid.PRNet = -TotalWHAndDeds
  TempVoid.PPEAmt = -TotalWHAndDeds
  TempVoid.PPETotAmt = -TotalWHAndDeds
  TempVoid.PPETotGL = GL4PPETotal
  TempVoid.MEDMATCRAmt = 0
  TempVoid.MEDMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).MEDLIAB))
  TempVoid.MEDMATCRGL = ReplaceString(TempVoid.MEDMATCRGL, "-", "")
  TempVoid.MEDMATCRGL = AddDashesToGLNumber(TempVoid.MEDMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.MEDMATDBAmt = 0
  TempVoid.MEDMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).MEDEXP)
  TempVoid.MEDMATDBGL = ReplaceString(TempVoid.MEDMATDBGL, "-", "")
  TempVoid.MEDMATDBGL = AddDashesToGLNumber(TempVoid.MEDMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATCRAmt = 0
  TempVoid.RETMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).RETLIAB))
  TempVoid.RETMATCRGL = ReplaceString(TempVoid.RETMATCRGL, "-", "")
  TempVoid.RETMATCRGL = AddDashesToGLNumber(TempVoid.RETMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.RETMATDBAmt = 0
  TempVoid.RETMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).RETEXP)
  TempVoid.RETMATDBGL = ReplaceString(TempVoid.RETMATDBGL, "-", "")
  TempVoid.RETMATDBGL = AddDashesToGLNumber(TempVoid.RETMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATCRAmt = 0
  TempVoid.SOCMATCRGL = Mid(SysRec(1).Liab(1).Acct, 1, FundLen) + QPTrim$(QPTrim$(SysRec(1).SOCLIAB))
  TempVoid.SOCMATCRGL = ReplaceString(TempVoid.SOCMATCRGL, "-", "")
  TempVoid.SOCMATCRGL = AddDashesToGLNumber(TempVoid.SOCMATCRGL, FundLen, AcctLen, DetLen)
  TempVoid.SOCMATDBAmt = 0
  TempVoid.SOCMATDBGL = Mid(SysRec(1).Liab(1).Acct, 1, ThisFACnt) + QPTrim$(SysRec(1).SOCEXP)
  TempVoid.SOCMATDBGL = ReplaceString(TempVoid.SOCMATDBGL, "-", "")
  TempVoid.SOCMATDBGL = AddDashesToGLNumber(TempVoid.SOCMATDBGL, FundLen, AcctLen, DetLen)
  TempVoid.WagesAmt = 0
  TempVoid.WagesGL = ""
  TVCnt = TVCnt + 1
  Put TVHandle, TVCnt, TempVoid
  PRNetPoolFound = True 'this tells the program not to come back to this sub until the next
  'time the pool fund number is not a 'Paid From' fund

  Return
'----^^^^^^^^^^^^^^added 9/20/2004 to allow for employees not paid out of the pool fund

  
  
  
  Return
  
PrintEmpSubTotal:
  RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsPaid)
  RSet OTHrs(1) = Using(Image$, TransRec(1).OTHrsPaid)
  RSet EDRPay(1) = Using(Image3$, TransRec(1).TotRegWage)
  RSet EDOPay(1) = Using(Image$, TransRec(1).TotOTWage)
  RSet EDEarn(1) = Using(Image$, TransRec(1).TotAdditEarn)
  RSet EDGroP(1) = Using(Image$, TransRec(1).GrossPay)
  
  RSet EDSAmt(1) = Using(Image$, TransRec(1).MatchSocAmt)
  RSet EDMAmt(1) = Using(Image$, TransRec(1).MatchMedAmt)
  RSet EDRAmt(1) = Using(Image$, TransRec(1).MatchRetAmt)
  
  'add sub totals for each employees soc, med + ret amts.
  Print #DHandle, Dash2(1)
  Print #DHandle, "Employee Total" + Fill11(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1) + EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1)
  DLineCnt = DLineCnt + 3
  
  Return
  
PrintDistTotal:
  
  RegHrs# = 0
  TotHrs# = 0
  RegWage# = 0
  OTWage# = 0
  AddEarn# = 0
  DGPay# = 0
  
  DLineCnt = 0
  
  Print #DHandle, FF$
  
  SortD DistbSumAccts(), NumOfWageAccts
  
  'print the Summary of ALL Distrubtion Accounts
  GoSub PrintAcctSumHeader
  
  For cnt = 1 To NumOfWageAccts
    If GPay# <= 0 Then Return
    LSet EDAct(1) = DistbSumAccts(cnt).Acct
    RSet EDPct(1) = Using(Image4$, (DistbSumAccts(cnt).GrossPay / GPay#) * 100)
    RSet RHrs(1) = Using(Image3$, DistbSumAccts(cnt).RHrs)
    RegHrs# = OldRound(RegHrs# + DistbSumAccts(cnt).RHrs)
    RSet OTHrs(1) = Using(Image3$, DistbSumAccts(cnt).OHrs)
    TotHrs# = OldRound(TotHrs# + DistbSumAccts(cnt).OHrs)
    RSet EDRPay(1) = Using(Image3$, DistbSumAccts(cnt).RWage)
    RegWage# = OldRound(RegWage# + DistbSumAccts(cnt).RWage)
    RSet EDOPay(1) = Using(Image3$, DistbSumAccts(cnt).OWage)
    OTWage# = OldRound(OTWage# + DistbSumAccts(cnt).OWage)
    RSet EDEarn(1) = Using(Image3$, DistbSumAccts(cnt).AddEarn)
    AddEarn# = OldRound(AddEarn# + DistbSumAccts(cnt).AddEarn)
    RSet EDGroP(1) = Using(Image3$, DistbSumAccts(cnt).GrossPay)
    DGPay# = OldRound(DGPay# + DistbSumAccts(cnt).GrossPay)
    
    RSet EDSAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATSocAmt)
    ASAmt# = OldRound(ASAmt# + DistbSumAccts(cnt).MATSocAmt)
    
    RSet EDMAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATMedAmt)
    AMAmt# = OldRound(AMAmt# + DistbSumAccts(cnt).MATMedAmt)
    
    RSet EDRAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATRetAmt)
    ARAmt# = OldRound(ARAmt# + DistbSumAccts(cnt).MATRetAmt)
    
    Print #DHandle, EDAct(1) + EDPct(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1) + EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1)
    
    DLineCnt = DLineCnt + 1
    If DLineCnt >= DMaxLines Then
      DLineCnt = 0
      GoSub PrintAcctSumHeader
    End If
  Next
  
  RSet RHrs(1) = Using(Image3$, RegHrs#)
  RSet OTHrs(1) = Using(Image3$, TotHrs#)
  RSet EDRPay(1) = Using(Image3$, RegWage#)
  RSet EDOPay(1) = Using(Image3$, OTWage#)
  RSet EDEarn(1) = Using(Image3$, AddEarn#)
  RSet EDGroP(1) = Using(Image3$, DGPay#)
  
  RSet EDSAmt(1) = Using(Image3$, ASAmt#)
  RSet EDMAmt(1) = Using(Image3$, AMAmt#)
  RSet EDRAmt(1) = Using(Image3$, ARAmt#)
  
  Print #DHandle, Dash2(1)
  Print #DHandle, "Totals        " + Fill11(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1) + EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1)
  
  Print #DHandle, FF$
  
  Return
  
PrintAcctSumHeader:
  Print #DHandle, QPTrim$(Unit(1).UFEMPR) + Space$(87) + "Page:" + Pg(1)
  Print #DHandle, "Earnings Distribution Totals"
  Print #DHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  Print #DHandle, "                                                                                               --------- Matching ----------"
  Print #DHandle, "Account Number        Pct    Reg Hrs    O/T Hrs    Reg Pay    O/T Pay  Tot Other  Gross Pay    Soc Sec   Medicare     Retire"
  Print #DHandle, Dash2(1)
  DLineCnt = DLineCnt + 6
  Return
  
End Sub

