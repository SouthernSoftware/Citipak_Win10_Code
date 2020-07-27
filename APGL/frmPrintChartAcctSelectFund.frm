VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrintChartAcctSelectFund 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Chart of Accounts"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmPrintChartAcctSelectFund.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo txtFundCode 
      Height          =   405
      Left            =   8070
      TabIndex        =   0
      Top             =   3450
      Width           =   960
      _Version        =   196608
      _ExtentX        =   1693
      _ExtentY        =   714
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrintChartAcctSelectFund.frx":08CA
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8280
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2:52 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "3/2/2007"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7368
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10080
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7368
      Width           =   1332
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or Select All To Print All Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   612
      Index           =   1
      Left            =   4080
      TabIndex        =   6
      Top             =   3960
      Width           =   3612
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2172
      Left            =   2640
      Top             =   2640
      Width           =   6972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Chart Of Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3240
      Top             =   1200
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3240
      Top             =   1080
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   3240
      Picture         =   "frmPrintChartAcctSelectFund.frx":0BC1
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select The Fund You Wish To Print A Chart of Accounts For"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   612
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   3240
      Width           =   3732
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPrintChartAcctSelectFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLFundIdx As GLFundIndexType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdExit_Click()
  frmChartAcctMaintMenu.Show
  Unload frmPrintChartAcctSelectFund
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub
Private Sub cmdPrint_Click()
  Dim LookFor As String
  LookFor = Me.txtFundCode.Text
  frmReportOpt.Show 1
  If rptopt = 1 Then
    PrintAcctListReport LookFor
  ElseIf rptopt = 2 Then
    PrintAcctListReport2 LookFor
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtfundCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    txtFundCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    txtFundCode.ListIndex = -1
    txtFundCode.Action = ActionClearSearchBuffer
  End If
  If txtFundCode.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub FundList()
  Dim FundIdxFileNum As Integer
  Dim NumFIdxRecs As Integer
  Dim cnt As Integer
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  txtFundCode.AddItem "All", 0
  For cnt = 1 To NumFIdxRecs
    Get FundIdxFileNum, cnt, GLFundIdx
    txtFundCode.AddItem GLFundIdx.FundNum
  Next
  Close FundIdxFileNum
  txtFundCode.ListIndex = 0
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  FundList
  Me.HelpContextID = hlpChartOfAccounts
End Sub
Private Sub PrintAcctListReport(LookFor)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFileNum As Integer, NumAccts As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer, newrpt As String
  Dim ReportFile As String, ToPrint As String, PageNum As Integer
  Dim Header As String, Tempcode As String
  Dim AcctIdx As GLAcctIndexType
  Dim GLAcct As GLAcctRecType
  Dim FundCode As String
  Dim ChkFund As Boolean
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  LookFor = Trim(LookFor)
  If InStr(LookFor, "All") > 0 Then
    ChkFund = False
  Else
    ChkFund = True
  End If
   'Define vars used for printing
   Header$ = "Master Account Listing"
   OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
   OpenAcctFile AcctFileNum
   PRNFile = FreeFile
   newrpt = "Acct"
   'get report name based on timer with acct as prefix and .prn as extention
   GetRPTName newrpt
   ReportFile$ = newrpt
   Open ReportFile$ For Output As #PRNFile
   FrmShowPctComp.Label1 = "Creating Account Report"
   FrmShowPctComp.Show , Me
   DoEvents
   EnableCloseButton Me.hwnd, False
   Me.cmdExit.Enabled = False
   Me.cmdPrint.Enabled = False
   Me.mnuOptions.Enabled = False
   For cnt = 1 To NumAIdxRecs
      Get AcctIdxFileNum, cnt, AcctIdx
      FrmShowPctComp.ShowPctComp cnt, NumAIdxRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Me.cmdExit.Enabled = True
        Me.cmdPrint.Enabled = True
        Me.mnuOptions.Enabled = True
        EnableCloseButton Me.hwnd, True
        Unload FrmShowPctComp
        GoTo CancelExit
      End If

      If ChkFund Then
        FundCode = Left$(AcctIdx.AcctNum, GLFundLen)
        If FundCode <> LookFor Then
          GoTo NotThisFund
        End If
      End If
      Get AcctFileNum, AcctIdx.RecNum, GLAcct
      Tempcode$ = ""
      If GLAcct.FNCTRec > 0 Then
        Tempcode$ = GetFnctCode(GLAcct.FNCTRec)
      End If
      Howmany = Howmany + 1
      ToPrint$ = ""
      If Len(Tempcode$) > 0 Then
        ToPrint$ = GLAcct.Num + "/" + Tempcode$ + "~" + GLAcct.Title + "~" + GLAcct.Typ
      Else
        ToPrint$ = GLAcct.Num + "~" + GLAcct.Title + "~" + GLAcct.Typ
      End If
      Print #PRNFile, ToPrint$
NotThisFund:
   Next

  Close
  'ViewPrint ReportFile$, "Account Listing Report"
  Me.cmdExit.Enabled = True
  Me.cmdPrint.Enabled = True
  Me.mnuOptions.Enabled = True
  EnableCloseButton Me.hwnd, True
  'Kill ReportFile$
   Load frmLoadingRpt
  ARptListings.Label1.Caption = "Acct Number/Function"
  ARptListings.Label2.Caption = "Title"
  ARptListings.Label3.Caption = "Type"
  ARptListings.Label4.Caption = "Accounts Listed"
  ARptListings.Total = Howmany
  ARptListings.txtDate = Now
  ARptListings.txtTown = GLUserName$
  ARptListings.Title.Caption = Header$
  ARptListings.GetName ReportFile$
  ARptListings.startrpt

Exit Sub
CancelExit:
Exit Sub
End Sub
Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Sub PrintAcctListReport2(LookFor)
  Dim MaxLines As Integer, AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFileNum As Integer, NumAccts As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer, newrpt As String
  Dim ReportFile As String, ToPrint As String, PageNum As Integer
  Dim FF As String, Header As String, Tempcode As String
  Dim AcctIdx As GLAcctIndexType
  Dim GLAcct As GLAcctRecType
  Dim FundCode As String
  Dim ChkFund As Boolean
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  LookFor = Trim(LookFor)
  If InStr(LookFor, "All") > 0 Then
    ChkFund = False
  Else
    ChkFund = True
  End If
   'Define vars used for printing
   MaxLines = 55
   FF$ = Chr$(12)
   Header$ = "Master Account Listing"
   OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
   OpenAcctFile AcctFileNum
   PRNFile = FreeFile
   newrpt = "Acct"
   'get report name based on timer with acct as prefix and .prn as extention
   GetRPTName newrpt
   ReportFile$ = newrpt
   Open ReportFile$ For Output As #PRNFile
   GoSub PrintAcctPageHeader
   FrmShowPctComp.Label1 = "Creating Account Report"
   FrmShowPctComp.Show , Me
   DoEvents
   EnableCloseButton Me.hwnd, False
   Me.cmdExit.Enabled = False
   Me.cmdPrint.Enabled = False
   Me.mnuOptions.Enabled = False
   For cnt = 1 To NumAIdxRecs
      Get AcctIdxFileNum, cnt, AcctIdx
      FrmShowPctComp.ShowPctComp cnt, NumAIdxRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Me.cmdExit.Enabled = True
        Me.cmdPrint.Enabled = True
        Me.mnuOptions.Enabled = True
        EnableCloseButton Me.hwnd, True
        Unload FrmShowPctComp
        GoTo CancelExit
      End If

      If ChkFund Then
        FundCode = Left$(AcctIdx.AcctNum, GLFundLen)
        If FundCode <> LookFor Then
          GoTo NotThisFund
        End If
      End If
      Get AcctFileNum, AcctIdx.RecNum, GLAcct
      Tempcode$ = ""
      If GLAcct.FNCTRec > 0 Then
        Tempcode$ = GetFnctCode(GLAcct.FNCTRec)
      End If
      Howmany = Howmany + 1
      ToPrint$ = Space$(80)
      If Len(Tempcode$) > 0 Then
        Mid$(ToPrint$, 2) = GLAcct.Num + "  /" + Tempcode$
      Else
        Mid$(ToPrint$, 2) = GLAcct.Num
      End If
      Mid$(ToPrint$, 28) = GLAcct.Title
      Mid$(ToPrint$, 66) = GLAcct.Typ
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFile, FF$
        GoSub PrintAcctPageHeader
      End If
NotThisFund:
   Next

   Print #PRNFile,
   Print #PRNFile, Howmany; "Accounts listed."
   Print #PRNFile, FF$

  Close
  ViewPrint ReportFile$, "Account Listing Report"
  Me.cmdExit.Enabled = True
  Me.cmdPrint.Enabled = True
  Me.mnuOptions.Enabled = True
  EnableCloseButton Me.hwnd, True
  Kill ReportFile$
 
Exit Sub
PrintAcctPageHeader:
  PageNum = PageNum + 1
  Print #PRNFile, Header$
  Print #PRNFile, Tab(70); "Page: "; PageNum
  Print #PRNFile, "Acct Number/Function     Title                                          Type"
  Print #PRNFile, String$(78, "-")
  Linecnt = 4
Return
CancelExit:
Exit Sub
End Sub

