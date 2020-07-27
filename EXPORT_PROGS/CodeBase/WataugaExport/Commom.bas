Attribute VB_Name = "Module1"
Option Explicit

Public RecNum As Long
'Public TaxText(1 To 10) As String * 2

Public EHandle As Integer
Public TRHandle As Integer
Public SplitFlag As Boolean
Public Const Manual = 2
Public Const Normal = 1
Public EntryType As Integer
Public ScreenW As Long
Public coladj As Double
Public doAlign As Boolean
Public alnRpt$
Public OutFileNames(1 To 20) As String
Public InFileNames(1 To 20) As String
Public ComputerName As String
Public BadMaskFlag As Boolean
'Public EmpInfo(1 To 30) As String
'Public ToPrint1(1 To 10) As Integer
'Public ToPrint2(1 To 10) As Integer
'Public CurrCitiPath As String
'Public NewListFlag As Boolean
Public StartPath As String
'Public NumOfAligns As Integer
'Public GlblQtr$ 'used in ESC report
'Public FundCnt4Rpt As Integer 'used in YTD Wage Distribution report
'Public DeductionSelNum As Integer 'used in Deduction report
'Public ThisRpt$ 'used in reprint report
'Public RptOpt As Integer 'used to determine the type of reports; graphic or text
'Public AccrualDate As Integer '12/12/02
'Public AccrualDateString$ '12/12/02
'Public ErrAcct() As String
'Public ErrAmt() As Double
'Public ErrType() As String
'Public GlobalCheckNum$ 'used solely for voiding a check
'Public GlobalTransNum As Double 'used solely for voiding a check
'Public GlobalName As String 'used solely for voiding a check
'Public ErrEmpNum$
'Public ErrCnt As Integer
'Public PayType As String
'Public OTRate As Double
'Public RegRate As Double
'Public ThisFreq As String
'Public Twiddle As String
'Public FromPR As Boolean
'Public OverCnt As Integer
'Public bigName() As String

Type TaxMasterType      'Master Default Information in Setup
  Name As String * 35
  Add1 As String * 35
  Add2 As String * 35
  'ADD3 As String * 35
  'Change the add3 line to break out individual city,st,zip on 013103.
  City As String * 25
  'use taxst for state in address
  'State As String * 2
  Zip As String * 10
  TaxSt As String * 2
  'TaxForm As String * 20
  'Change taxform above to 2 byte integer
  TaxForm As Integer
  'add lateform 031303
  RTaxYear As Integer
  LateForm As Integer
'  pad As String * 16  'left from taxform string of 20
'change above pad to use for following changes as of 3-28-03
'  pad     As String * 3
  WarnInt As String * 1  'Flag to Warn if interest not applied
'  DisFlag As String * 1  'set discount flag if want interest calculated
  MinBill As Double      'amount to not print bills
  'CurRate As Single
  'PastRate As Single
  'PenRate As Single
  'use the 3 rates above (12) for other stuff
 'change rcptport to pad up above - will set printer ports when sign on
  'RcptPort As Integer
  AcctgMethod As String * 1
  'add interface option 031301
  MinTxOpt As Integer '1/26/05 '1) if the taxpayer is charged nothing if
  'their tax bill is equal to or less than this amt...2) the taxpayer is charged at least this
  'amt even if they owe nothing
  TownState As String * 2 '1/26/05
  CurrRYrInt(1 To 5) As Double  '12/14/05
  CurrRYrIntInUse As Double '12/14/05
  CurrPYrInt(1 To 5) As Double  '12/14/05
  CurrPYrIntInUse As Double '12/14/05
  PastYrInt As Double '1/26/05
  PenPct As Double '1/26/05
  PenIdx As Integer
  CntrlDepYN As String * 1
  PriorYrMltRevYN As String * 1
  OverPayGLNum As String * 14
  PenPrncTaxYN As String * 1
  PenIntYN As String * 1
  PenAdvYN As String * 1
  PenLateLstYN As String * 1
  PenOpt1YN As String * 1
  PenOpt2YN As String * 1
  PenOpt3YN As String * 1
  IntPrncTaxYN As String * 1
  IntIntYN As String * 1
  IntAdvYN As String * 1
  IntLateLstYN As String * 1
  IntOpt1YN As String * 1
  IntOpt2YN As String * 1
  IntOpt3YN As String * 1
  OptRev1 As String * 20
  OptRev2 As String * 20
  OptRev3 As String * 20
  DiscRXDate As Integer      'discount amount to calc on payment screen
  DisRPct As Double
  DiscPXDate As Integer      'discount amount to calc on payment screen
  DisPPct As Double
  OptSrchCust As String * 15
  OptSrchProp As String * 15
  CountyName(1 To 5) As String * 20
  CountyNum(1 To 5) As Integer
  UseCountyYN As String * 1
  RealPersSplit As String * 1
  CycleNum(1 To 5) As Long
  CycleName(1 To 5) As String * 20
  UseCyclesYN As String * 1
  CDCashGL  As String * 14
  CDSubGL  As String * 14
  ClassName(1 To 6) As String * 15
  MultiYear As Integer
  PPTRADisc As Double
  MaxVehTaxVal As Double
  LawChngDate As Integer 'on or about 9/6/2006 the VA law changes such that delinquents
  'no longer receive PPTRA discounts
  MinVehTaxVal As Double
  PPTRAYN As String * 1
  PenPenaltyYN As String * 1
  IntPenaltyYN As String * 1

  '---------------------------added for 2.05
  POptRev1 As String * 20
  POptRev2 As String * 20
  POptRev3 As String * 20
  PenPersYN As String * 1
  IntPersYN As String * 1
  PersPayOrder As Integer
  PenMTYN As String * 1
  IntMTYN As String * 1
  MTPayOrder As Integer
  PenMCYN As String * 1
  IntMCYN As String * 1
  MCPayOrder As Integer
  PenFEYN As String * 1
  IntFEYN As String * 1
  FEPayOrder As Integer
  PenMHYN As String * 1
  IntMHYN As String * 1
  MHPayOrder As Integer
  PenPIntYN As String * 1
  IntPIntYN As String * 1
  PIntPayOrder As Integer
  PenPPenYN As String * 1
  IntPPenYN As String * 1
  PPenPayOrder As Integer
  PenPOpt1YN As String * 1
  IntPOpt1YN As String * 1
  POpt1PayOrder As Integer
  PenPOpt2YN As String * 1
  IntPOpt2YN As String * 1
  POpt2PayOrder As Integer
  PenPOpt3YN As String * 1
  IntPOpt3YN As String * 1
  POpt3PayOrder As Integer
  '------------------------------------------------------
  Padding As String * 57
  PTaxYear As Integer
  OptSrchPers As String * 15 'added 8/16/06
End Type


Type TaxCustType
  Acct       As Long
  OPENDATE   As Integer
'  FName      As String * 15
'  LName      As String * 25
  CustName   As String * 50
  SName      As String * 10
  HPHONE     As String * 14
  WPHONE     As String * 14
  CSSN       As String * 11
  OSSN       As String * 11
  Addr1      As String * 35
  Addr2      As String * 35
  City       As String * 20
  State      As String * 2
  Zip        As String * 10
  Active     As String * 1    'Y if Active N if Inactive
  Interest   As String * 1    'Y/N to Charge Interest
  TaxExempt  As String * 1    'Y/N to Charge Taxes Period
  Penalty    As String * 1    'Y/N to Charge Penalty
  Employer   As String * 25
  Bankrupt   As String * 1    'Y/N to Charge Collect/Adv
  TownShip   As String * 25
'032400
  LateNotice As String * 1    'Y/N Allow late notice
'*  PAD1              As String * 202
'*Change Pad1 from 202 and added PrePayment Balance field
'*Also pointer to prepay transaction
'*added these 2 fields on 3/25/03 pks
  PrePayBal    As Double
  PrePayTrans  As Long
'032900 for New Market Va to Track Who Has Been Submitted to DMV
'  DMV1999           As String * 1'remmed out on 3/2/05
'  DMV2000           As String * 1'remmed out on 3/2/05
'  DMV2001           As String * 1'remmed out on 3/2/05
'  DMV2002           As String * 1'remmed out on 3/2/05
  CountyAcctString  As String * 18    'County Account in String Format when lo
  CountyAcct    As Long        'County Account Number to Link to County Record
  LastTrans     As Long        'Pointer to last transaction
  FirstPropRec  As Long        'Pointer to first property rec
  FirstPersRec  As Long        'Pointer to first personal rec
  PIN           As Long        'Cust internal id number.
  Deleted       As Integer     'deleted flag
  FileVer       As Integer     'this is the file struct version number
  OptSrchDesc   As String * 15 '3/1/05
  ServiceAdd    As String * 35
  DrvrsLic      As String * 10
  DeliveryPt      As String * 2
  PostalRt     As String * 4
  Cycle        As Long
  CycleName    As String * 20
  County4BillNum As Long 'used as option for billing
  County4BillName As String * 20
  Pad1         As String * 190  '*remainder after additional fields
End Type

Type MortCodeRecType
  MORTCODE As String * 8
  BName    As String * 32
  Add1     As String * 32
  Add2     As String * 32
  Add3     As String * 32
  Contact  As String * 32
  PHONE    As String * 14
'Add deleted field 021003
  Deleted  As Integer
  XFileNme As String * 8
  pad      As String * 252
End Type

Type TownshipType
  TownShip As String * 30
End Type
'
Type MessLineType
  Msg As String * 69
End Type
'

Type TaxMessRecType
  MessLine(1 To 15) As MessLineType
  TaxRec As Long
End Type

Type OptCustIdxType
  OptDesc As String * 20
  CustRec As Long
  CustPin As Long
End Type

Type TAXLateLetterType
  Head1    As String * 40
  Head2    As String * 40
  Head3    As String * 40
  Head4    As String * 40
  Head5    As String * 40
  Body(1 To 20) As String * 75
End Type

Type OptPersIdxType
  OptDesc As String * 20
  PersRec As Long
  PersPin As String * 20
End Type

Public Function ParseBillNum$(Text$)
  Dim BillNum$
  Dim BNumLen As Integer
  Dim thischar$
  Dim GoodPos As Integer
  Dim cnt As Integer
  
  BillNum$ = QPTrim$(Text$)
  BNumLen = Len(BillNum$)
  If BNumLen > 0 Then
    For cnt = BNumLen To 1 Step -1
      thischar$ = Mid$(BillNum$, cnt, 1)
      If InStr("0123456789", thischar$) <= 0 Then
        Exit For
      End If
    Next
    GoodPos = cnt + 1
    BillNum$ = Mid$(BillNum$, GoodPos)
  End If
  If Not IsNumeric(BillNum$) Then
    BillNum = "-911"
  End If
  ParseBillNum$ = BillNum$
End Function

Public Function OldRound#(N As Double)
  OldRound# = Int(N * 100 + 0.50000001) / 100
End Function

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

Public Sub KillFile(FileName As String)
  On Local Error Resume Next
  If Exist(FileName$) Then 'added 7/24
    Kill FileName$
  End If
  On Error GoTo 0
End Sub

Public Function Exist(FileName$) As Boolean
  Dim FileHandle As Integer
  Dim TempSize As Long
  On Local Error Resume Next
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  TempSize = LOF(FileHandle)
  Close FileHandle
  If TempSize <= 0 Then
    Kill FileName$
    Exist = False
  Else
    Exist = True
  End If
  On Error GoTo 0
End Function

Public Static Function Using$(ByVal fmt As String, ByVal Number As Double)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If BuckPos = 1 Then
    fmt = Right$(fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(FmtLen)
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
  RSet FmtNumber = TempNumber
  Using = FmtNumber
End Function

Public Function MakeRegDate(ByVal DateNumb)
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function

Public Function FileExists(ByVal strFileName As String) As Boolean
  On Error Resume Next
'
  If (Len(Dir$(strFileName)) > 0) Then
    FileExists = True
  Else
    FileExists = False
  End If
  On Error GoTo 0
End Function

Public Function DirExists(ByVal strDirName As String) As Boolean
  On Error Resume Next
  Dim strFileName As String
  strFileName = strDirName & "\Nul"
  If (FileExists(strFileName)) Then
    DirExists = True
  Else
    DirExists = False
  End If
  On Error GoTo 0
End Function



