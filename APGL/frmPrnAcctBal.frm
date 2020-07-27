VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPrnAcctBal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Balance"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPrnAcctBal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAcct 
      Height          =   384
      Left            =   6480
      TabIndex        =   0
      Top             =   3504
      Width           =   2220
      _Version        =   196608
      _ExtentX        =   3916
      _ExtentY        =   677
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
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
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
      ColDesigner     =   "frmPrnAcctBal.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   6480
      TabIndex        =   2
      Top             =   4896
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   677
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
      Columns         =   1
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      ColDesigner     =   "frmPrnAcctBal.frx":0C51
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "1:27 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "10/15/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      Left            =   8256
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
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
      Left            =   10032
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   1332
   End
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   6480
      TabIndex        =   1
      Top             =   4200
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
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
      Height          =   372
      Left            =   4056
      TabIndex        =   9
      Top             =   4920
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account And Object Codes:"
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
      Height          =   420
      Left            =   3360
      TabIndex        =   8
      Top             =   3552
      Width           =   3012
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2652
      Left            =   2688
      Top             =   2976
      Width           =   6828
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Account Balance Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3984
      TabIndex        =   7
      Top             =   1608
      Width           =   4332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   3552
      Picture         =   "frmPrnAcctBal.frx":0FB7
      Top             =   4416
      Width           =   288
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
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
      Height          =   420
      Left            =   4800
      TabIndex        =   6
      Top             =   4248
      Width           =   1572
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   1248
      Width           =   5772
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
Attribute VB_Name = "frmPrnAcctBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct As GLAcctRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim GLFundIdx As GLFundIndexType
Dim GLTrans   As GLTransRecType
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Private Sub cmdPrint_Click()
  If fpcboAcct.ListIndex > -1 Then
    If CheckValDate(txtDate) = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrintAcctBal
      ElseIf rptopt = 2 Then
        PrintAcctBal2
      End If
    Else
      MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    End If
  Else
    MsgBox "You Must Select An Account, Retry", vbOKOnly, "Invalid Selection"
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fpcboAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcct.ListIndex = -1
    fpcboAcct.Action = ActionClearSearchBuffer
  End If
  If fpcboAcct.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpAccountBalance
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  FillAcctListp fpcboAcct
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Sub cmdExit_Click()
  frmGLReportsMenu.Show
  Unload frmPrnAcctBal
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

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Function FillAcctListp(txtField As fpCombo)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, newac As String
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  txtField.Row = -1
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct
      If GLAcct.Deleted = 0 Then
        newac = QPStrip(GLAcct.Num)
        txtField.InsertRow = Mid$(QPTrim(GLAcct.Num), 2 + GLFundLen, (GLAcctLen + 1 + GLDetLen)) & Chr(9) & newac
      End If
  Next
  Close AcctIdxFileNum
  Close AcctFile
  End Function
Private Sub PrintAcctBal()
  Dim PRNfileName As String, PRNFileNum As Integer, FundCnt As Integer, cnt As Integer
  Dim ReportFile As String, ToPrint As String, CommaFmt As String
  Dim TotalFmt As String, BankGLAcct As String, cntp As Long
  Dim TotalCashBal As Double, CalcBal As Double, Newrp As String
  Dim NumFunds As Integer, EndDate As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, AcctRec As Integer, AcctNum As String
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long
  Dim FundIdxFile As Integer, Acct2Bal As String, ToPrintA As String
 
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars

  '=====================================================
  EndDate = DateDiff("d", "12/31/1979", txtDate)

  '--Clear help row
  Newrp = "AcctBal"
  GetRPTName Newrp
  fpcboAcct.Col = 0
  Acct2Bal$ = "-" + QPTrim$(fpcboAcct.ColText)
  PRNfileName$ = Newrp
  PRNFileNum = FreeFile
  Open PRNfileName$ For Output As #PRNFileNum

  OpenFundIdx FundIdxFile, NumFunds
  OpenAcctFile AcctFileNum
  NumGLAccts = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&

'  Print #PRNFileNum, GLUserName; Tab(43); "Run Date: " + Date$
'  Print #PRNFileNum, "Account Summary"
'  Print #PRNFileNum, "Ending Date: " + txtDate
'  Print #PRNFileNum,
'  Print #PRNFileNum, "Account                 Balance"
'  Print #PRNFileNum, "---------------------------------"
  FrmShowPctComp.Label1 = "Printing Account Balance Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnAcctBal, True
  For cnt = 1 To NumFunds
    
    Get FundIdxFile, cnt, GLFundIdx

    AcctNum$ = QPTrim$(GLFundIdx.FundNum) + Acct2Bal$
    fpcboAcct.Col = 0
    AcctRec = AcctFind(AcctNum$)
    If AcctRec > 0 Then
      ToPrint$ = Space(80)

      Get AcctFileNum, AcctRec, GLAcct
      CalcBal# = Round#(GLAcct.BegBal)
      NextTr& = GLAcct.FrstTran
      Do Until NextTr& = 0
        Get TransFileNum, NextTr&, GLTrans
        If GLTrans.TRDATE <= EndDate Then
          CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
        End If
        NextTr& = GLTrans.NextTran
      Loop

      ToPrint$ = Acct2Bal$ + "~" + AcctNum$ + "~" + Using(CommaFmt$, Str$(CalcBal#))
      Print #PRNFileNum, ToPrint$
      TotalCashBal# = TotalCashBal# + CalcBal#

    End If
       FrmShowPctComp.ShowPctComp cnt, NumFunds
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAcctBal, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
 

'  Print #PRNFileNum, Tab(18); "----------------"
'  Print #PRNFileNum, "Total"; Tab(18); Using(TotalFmt$, Str$(TotalCashBal#))
'  Print #PRNFileNum, Chr$(12)
  Close

  'ViewPrint PRNfileName$, "Account Balance Report"

  'KillFile PRNfileName$
  ActivateControls frmPrnAcctBal, True
   Load frmLoadingRpt
   ARptAcctBal.txtRptDate = "Period Ending: " + txtDate
   ARptAcctBal.txtDate = Now
   ARptAcctBal.txtTown = GLUserName$
   ARptAcctBal.GetName PRNfileName$
   ARptAcctBal.startrpt

CancelExit:
Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
  End If
End Sub
Private Sub PrintAcctBal2()
  Dim PRNfileName As String, PRNFileNum As Integer, FundCnt As Integer, cnt As Integer
  Dim ReportFile As String, ToPrint As String, CommaFmt As String
  Dim TotalFmt As String, BankGLAcct As String, cntp As Long
  Dim TotalCashBal As Double, CalcBal As Double, Newrp As String
  Dim NumFunds As Integer, EndDate As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, AcctRec As Integer, AcctNum As String
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long
  Dim FundIdxFile As Integer, Acct2Bal As String
 
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars

  '=====================================================
  EndDate = DateDiff("d", "12/31/1979", txtDate)

  '--Clear help row
  Newrp = "AcctBal"
  GetRPTName Newrp
  fpcboAcct.Col = 0
  Acct2Bal$ = "-" + QPTrim$(fpcboAcct.ColText)
  PRNfileName$ = Newrp
  PRNFileNum = FreeFile
  Open PRNfileName$ For Output As #PRNFileNum

  OpenFundIdx FundIdxFile, NumFunds
  OpenAcctFile AcctFileNum
  NumGLAccts = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&

  Print #PRNFileNum, GLUserName; Tab(43); "Run Date: " + Date$
  Print #PRNFileNum, "Account Summary"
  Print #PRNFileNum, "Ending Date: " + txtDate
  Print #PRNFileNum,
  Print #PRNFileNum, "Account                 Balance"
  Print #PRNFileNum, "---------------------------------"
  FrmShowPctComp.Label1 = "Printing Account Balance Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnAcctBal, True
  For cnt = 1 To NumFunds
    
    Get FundIdxFile, cnt, GLFundIdx

    AcctNum$ = QPTrim$(GLFundIdx.FundNum) + Acct2Bal$
    fpcboAcct.Col = 0
    AcctRec = AcctFind(AcctNum$)
    If AcctRec > 0 Then


      Get AcctFileNum, AcctRec, GLAcct
      CalcBal# = Round#(GLAcct.BegBal)
      NextTr& = GLAcct.FrstTran
      Do Until NextTr& = 0
        Get TransFileNum, NextTr&, GLTrans
        If GLTrans.TRDATE <= EndDate Then
          CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
        End If
        NextTr& = GLTrans.NextTran
      Loop

      Print #PRNFileNum, AcctNum$; Tab(20); Using(CommaFmt$, Str$(CalcBal#))
      TotalCashBal# = TotalCashBal# + CalcBal#

    End If
       FrmShowPctComp.ShowPctComp cnt, NumFunds
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAcctBal, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
 

  Print #PRNFileNum, Tab(18); "----------------"
  Print #PRNFileNum, "Total"; Tab(18); Using(TotalFmt$, Str$(TotalCashBal#))
  Print #PRNFileNum, Chr$(12)
  Close

  ViewPrint PRNfileName$, "Account Balance Report"

  KillFile PRNfileName$
  ActivateControls frmPrnAcctBal, True

CancelExit:
Exit Sub
End Sub

