VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} RptChartAccts 
   Caption         =   "Chart of Accounts"
   ClientHeight    =   5532
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   9372
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16531
   _ExtentY        =   9758
   _Version        =   393216
   _DesignerVersion=   100685828
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   10
   GridY           =   10
   LeftMargin      =   1440
   RightMargin     =   1440
   TopMargin       =   1440
   BottomMargin    =   1440
   NumSections     =   5
   SectionCode0    =   1
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section4"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode1    =   2
   BeginProperty Section1 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section2"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode2    =   4
   BeginProperty Section2 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section1"
      Object.Height          =   1440
      NumControls     =   0
   EndProperty
   SectionCode3    =   7
   BeginProperty Section3 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section3"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode4    =   8
   BeginProperty Section4 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section5"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
End
Attribute VB_Name = "RptChartAccts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub PrintAcctListReport(LookFor)
'  Dim MaxLines As Integer, AcctIdxFileNum As Integer, NumAIdxRecs As Integer
'  Dim AcctFileNum As Integer, NumAccts As Integer, Linecnt As Integer
'  Dim PRNFile As Integer, cnt As Integer, HowMany As Integer
'  Dim ReportFile As String, ToPrint As String, PageNum As Integer
'  Dim FF As String, Header As String ', LookFor As String
'  Dim AcctIdx As GLAcctIndexType
'  Dim GLAcct As GLAcctRecType
'  Dim FundCode As String
'  Dim ChkFund As Boolean
'  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
'  LookFor = Trim(LookFor)
'  If InStr(LookFor, "All") > 0 Then
'    ChkFund = False
'  Else
'    ChkFund = True
'  End If
'   'Define vars used for printing
'   MaxLines = 55
'   FF$ = Chr$(12)
'   Header$ = "Master Account Listing"
'   OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
'   OpenAcctFile AcctFileNum
'   PRNFile = FreeFile
'   ReportFile$ = "ACCTLIST.PRN"
'   Open ReportFile$ For Output As #PRNFile
'   GoSub PrintAcctPageHeader
'   FrmShowPctComp.Label1 = "Creating Account Report"
'   FrmShowPctComp.Show , Me
'   DoEvents
'   EnableCloseButton Me.hwnd, False
'   Me.cmdExit.Enabled = False
'   Me.cmdPrint.Enabled = False
'
'   For cnt = 1 To NumAIdxRecs
'      Get AcctIdxFileNum, cnt, AcctIdx
'      FrmShowPctComp.ShowPctComp cnt, NumAIdxRecs
'      If FrmShowPctComp.Out = True Then
'        Close
'        FrmShowPctComp.Out = False
'        Me.cmdExit.Enabled = True
'        Me.cmdPrint.Enabled = True
'        EnableCloseButton Me.hwnd, True
'        Unload FrmShowPctComp
'        GoTo CancelExit
'      End If
'
'      If ChkFund Then
'        FundCode = Left$(AcctIdx.AcctNum, GLFundLen)
'        If FundCode <> LookFor Then
'          GoTo NotThisFund
'        End If
'      End If
'      Get AcctFileNum, AcctIdx.RecNum, GLAcct
'      HowMany = HowMany + 1
'      ToPrint$ = Space$(80)
'      Mid$(ToPrint$, 2) = GLAcct.Num
'      Mid$(ToPrint$, 18) = GLAcct.Title
'      Mid$(ToPrint$, 66) = GLAcct.Typ
'      Print #PRNFile, ToPrint$
'      Linecnt = Linecnt + 1
'      If Linecnt > MaxLines Then
'        Print #PRNFile, FF$
'        GoSub PrintAcctPageHeader
'      End If
'NotThisFund:
'   Next
'
'   Print #PRNFile,
'   Print #PRNFile, HowMany; "Accounts listed."
'   Print #PRNFile, FF$
'
'  Close
'  ViewPrint ReportFile$, "Account Listing Report"
'  Me.cmdExit.Enabled = True
'  Me.cmdPrint.Enabled = True
'  EnableCloseButton Me.hwnd, True
'
'
'Exit Sub
'PrintAcctPageHeader:
'  PageNum = PageNum + 1
'  Print #PRNFile, Header$
'  Print #PRNFile, Tab(70); "Page: "; PageNum
'  Print #PRNFile, "Acct Number     Title                                          Type"
'  Print #PRNFile, String$(78, "-")
'  Linecnt = 4
'Return
'CancelExit:
'Exit Sub
'End Sub

