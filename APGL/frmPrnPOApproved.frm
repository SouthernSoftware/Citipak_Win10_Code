VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnPOApproved 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Approved Purchase Orders"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnPOApproved.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5910
      TabIndex        =   2
      Top             =   4605
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ColDesigner     =   "frmPrnPOApproved.frx":08CA
   End
   Begin LpLib.fpCombo fpcboDepartment 
      Height          =   405
      Left            =   5910
      TabIndex        =   0
      Top             =   3120
      Width           =   2160
      _Version        =   196608
      _ExtentX        =   3810
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnPOApproved.frx":0C30
   End
   Begin LpLib.fpCombo fpcboPOs 
      Height          =   405
      Left            =   5925
      TabIndex        =   1
      Top             =   3870
      Width           =   1485
      _Version        =   196608
      _ExtentX        =   2619
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
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnPOApproved.frx":0FE3
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   468
      Left            =   4464
      TabIndex        =   3
      Top             =   5328
      Width           =   1428
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   468
      Left            =   6300
      TabIndex        =   4
      Top             =   5328
      Width           =   1428
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8484
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "11:47 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "3/14/2007"
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
   Begin VB.Label Label4 
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
      Left            =   3312
      TabIndex        =   9
      Top             =   4656
      Width           =   2388
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number:"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   3936
      Width           =   1524
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
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
      Height          =   324
      Left            =   4260
      TabIndex        =   7
      Top             =   3192
      Width           =   1356
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Orders Approved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3684
      TabIndex        =   6
      Top             =   1200
      Width           =   4836
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   960
      Width           =   7020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3372
      Left            =   3072
      Top             =   2736
      Width           =   5748
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   840
      Width           =   7020
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
Attribute VB_Name = "frmPrnPOApproved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim GLTrans   As GLTransRecType
Dim PO As POFORMRecType2
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdExit_Click()
  frmPOProcessMenu.Show
  Unload frmPrnPOApproved
End Sub

Private Sub cmdOk_Click()
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt = 1 Then
    PrnAPEditList
  ElseIf rptopt = 2 Then
    PrnAPEditList2
  End If
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdOk.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPOs.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpApprPO
  DeptList fpcboDepartment
  fpcboDepartment.ListIndex = 0
  POList fpcboPOs
  fpcboPOs.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PrnAPEditList()
  Dim cnt As Integer, DebitCol As Integer, ReportFile2 As String
  Dim CreditCol As Integer, PrnFileNum2 As Integer
  Dim PRNFile As Integer, HowMany As Integer, CommaFmt As String
  Dim ReportFile As String, ToPrint As String, OverFmt As String
  Dim Header As String, DistSumLine As String, ToPrint2 As String
  Dim TransTotal As Double, TranCnt As Integer, FileName As String
  Dim RegTitle As String, TranCol As Integer, CashCol As Integer
  Dim DeptNumber As String, NYBeg As Integer, NYEnd As Integer
  Dim ThisDist As Double, Accttotal As Double, Over As String
  Dim ThisAcct As String, AcctCnt As Integer, WhatAcct As Integer
  Dim Found As Boolean, Fund As Integer, FundNum As String
  Dim GrdTot As Double, HCnt As Integer, TempPO As String, Newrp As String
  Dim POEditFile As Integer, NumEdTrans As Integer, PODept As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, AcctDist As Integer
  Dim IdxFile As Integer, NumIdxRecs As Integer, Transaction As Integer
  Dim AcctFile As Integer, NumAccts As Integer, DidCnt As Integer
  Dim TransFileNum As Integer, NumTrans As Long, PrnFileNum As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, TotTranDist As Double
  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "###,###,###.##"   'ten millions
  DistSumLine$ = "--------------"
  ReDim Title$(5)
  TransTotal# = 0
  TranCnt = 0
  ToPrint$ = Space$(80)
  Newrp = "POApr"
  GetRPTName Newrp
  FileName$ = Newrp
  RegTitle$ = "Purchase Orders"
  TranCol = CreditCol
  CashCol = DebitCol
  FrmShowPctComp.Label1 = "Creating Approved PO Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  mnuOptions.Enabled = False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  fpcboDepartment.col = 1
  DeptNumber$ = QPTrim$(fpcboDepartment.ColText)
  TempPO$ = QPTrim(fpcboPOs.Text)

  OpenPOEditFile POEditFile, NumEdTrans
  PrnFileNum = FreeFile
  Open FileName$ For Output As #PrnFileNum
  ReportFile2$ = "POFundSum.prn"
  PrnFileNum2 = FreeFile
  Open ReportFile2$ For Output As #PrnFileNum2
  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  ReDim FundGrdTot#(1 To NumFunds)

  For Transaction = 1 To NumEdTrans
    Get POEditFile, Transaction, PO
    FrmShowPctComp.ShowPctComp Transaction, NumEdTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    PODept$ = QPTrim$(PO.REQNUM)
    If PO.Deleted <> True Then
      If PODept$ = DeptNumber$ Or DeptNumber$ = "All" Then
        If Left$(PO.PONum, 3) <> "N/A" Then
          If TempPO$ = QPTrim(PO.PONum) Or TempPO$ = "All" Then
          TranCnt = TranCnt + 1
          TransTotal# = Round#(TransTotal# + PO.POAmt)
          '--Print 1st Line - Transaction details
          ToPrint$ = ""    'SPACE$(78)

          ToPrint$ = QPTrim(PO.VNDRCODE) + " " + Str$(PO.VNDRREC)
          ToPrint$ = ToPrint$ + "~" + Format(DateAdd("d", (PO.PODATE), "12-31-1979"), "mm/dd/yyyy")
          ToPrint$ = ToPrint$ + "~" + PO.PONum
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(PO.POAmt))
          'Print #PrnFileNum, ToPrint$

          '--Blank line between detail and acct
          ToPrint$ = ToPrint$ + "~" + "Dept # " + PO.REQNUM
          TotTranDist# = 0
          '--Loop Thru distributions to print and summarize
          For AcctDist = 1 To 36
            If Len(QPTrim$(PO.ITEMS(AcctDist).ACCTNO)) Then
              If AcctFind(PO.ITEMS(AcctDist).ACCTNO) > 0 Then
              '--Print this distribution
              ToPrint2$ = ""
              ToPrint2$ = PO.ITEMS(AcctDist).ACCTNO

              ToPrint2$ = ToPrint2$ + "~" + Using$(CommaFmt$, Str$(PO.ITEMS(AcctDist).EXT))
              ToPrint2$ = ToPrint2$ + "~ ~" + QPTrim$(PO.ITEMS(AcctDist).Desc) + "~"
              Print #PrnFileNum, ToPrint$ + "~" + ToPrint2$

              TotTranDist# = TotTranDist# + PO.ITEMS(AcctDist).EXT


              'Sum by fund
              Found = False
              For Fund = 1 To NumFunds
                FundNum$ = Left$(PO.ITEMS(AcctDist).ACCTNO, GLFundLen)
                If FundNum$ = FundList$(Fund) Then
                  Found = True
                  FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + PO.ITEMS(AcctDist).EXT)
                  Exit For
                End If
              Next

              If Not Found Then 'Oh.Shit = True
                 Unload FrmShowPctComp
                 MsgBox "Error, Invalid Fund Found.", vbOKOnly, "Error"
                 ActivateControls frmPrnPOApproved
                 Exit Sub
              End If
            End If
            End If              'Active transaction test

          Next  'Distribution


          End If
        End If  'Not deleted test
      End If    'Not Correct Dept
    End If      'Not Approved Flag
  Next          'Transaction
  If TranCnt < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If


  For cnt = 1 To NumFunds
    If FundGrdTot#(cnt) > 0 Then
      ToPrint$ = ""
      ToPrint$ = FundList$(cnt)
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(FundGrdTot#(cnt)))
      Print #PrnFileNum2, ToPrint$
      GrdTot# = Round#(GrdTot# + FundGrdTot#(cnt))
    End If
  Next
  If TranCnt < 1 Then
    ARptPOEdits.Label13.Visible = True
    ARptPOEdits.Label13.Caption = "** No Approved Purchase Orders **"
  End If

  Close
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  Load frmLoadingRpt
  ARptPOEdits.totTrans = Using$("####", Str$(TranCnt))
  ARptPOEdits.totGrand = Using$(CommaFmt$, Str$(TransTotal#))
  ARptPOEdits.totFunds = Using$(CommaFmt$, Str$(GrdTot#))
  
  ARptPOEdits.txtDate = Now
  ARptPOEdits.txtTown = GLUserName$
  ARptPOEdits.Label1.Caption = "Purchase Orders : APPROVED!"
  ARptPOEdits.Label2.Caption = "Dept #: " + DeptNumber$
  ARptPOEdits.Caption = "Purchase Orders : Approved"
  ARptPOEdits.GetName FileName$, ReportFile2$
  ARptPOEdits.startrpt

  Exit Sub

CancelExit:
  Exit Sub
End Sub
Private Sub PrnAPEditList2()
  Dim MaxLines As Integer, cnt As Integer, DebitCol As Integer
  Dim Linecnt As Integer, Page As Integer, CreditCol As Integer
  Dim PRNFile As Integer, HowMany As Integer, CommaFmt As String
  Dim ReportFile As String, ToPrint As String, OverFmt As String
  Dim FF As String, Header As String, DistSumLine As String
  Dim TransTotal As Double, TranCnt As Integer, FileName As String
  Dim RegTitle As String, TranCol As Integer, CashCol As Integer
  Dim DeptNumber As String, NYBeg As Integer, NYEnd As Integer
  Dim ThisDist As Double, Accttotal As Double, Over As String
  Dim ThisAcct As String, AcctCnt As Integer, WhatAcct As Integer
  Dim Found As Boolean, Fund As Integer, FundNum As String
  Dim GrdTot As Double, HCnt As Integer, TempPO As String, Newrp As String
  Dim POEditFile As Integer, NumEdTrans As Integer, PODept As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, AcctDist As Integer
  Dim IdxFile As Integer, NumIdxRecs As Integer, Transaction As Integer
  Dim AcctFile As Integer, NumAccts As Integer, DidCnt As Integer
  Dim TransFileNum As Integer, NumTrans As Long, PrnFileNum As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, TotTranDist As Double
  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "###,###,###.##"   'ten millions
  DistSumLine$ = "--------------"
  FF$ = Chr$(12)
  ReDim Title$(5)
  MaxLines = 56
  TransTotal# = 0
  TranCnt = 0
  ToPrint$ = Space$(80)
  Page = 0
  Newrp = "POApr"
  GetRPTName Newrp
  FileName$ = Newrp
  RegTitle$ = "Purchase Orders"
  TranCol = CreditCol
  CashCol = DebitCol
  FrmShowPctComp.Label1 = "Creating Approved PO Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  mnuOptions.Enabled = False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  fpcboDepartment.col = 1
  DeptNumber$ = QPTrim$(fpcboDepartment.ColText)
  TempPO$ = QPTrim(fpcboPOs.Text)

  OpenPOEditFile POEditFile, NumEdTrans
  PrnFileNum = FreeFile

  Open FileName$ For Output As #PrnFileNum

  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  ReDim FundGrdTot#(1 To NumFunds)

  GoSub PrintHeader
  For Transaction = 1 To NumEdTrans
    Get POEditFile, Transaction, PO
    FrmShowPctComp.ShowPctComp Transaction, NumEdTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    PODept$ = QPTrim$(PO.REQNUM)
    If PO.Deleted <> True Then
      If PODept$ = DeptNumber$ Or DeptNumber$ = "All" Then
        If Left$(PO.PONum, 3) <> "N/A" Then
          If TempPO$ = QPTrim(PO.PONum) Or TempPO$ = "All" Then
          TranCnt = TranCnt + 1
          TransTotal# = Round#(TransTotal# + PO.POAmt)
          '--Print 1st Line - Transaction details
          LSet ToPrint$ = ""    'SPACE$(78)

          LSet ToPrint$ = PO.VNDRCODE + " " + Str$(PO.VNDRREC)
          Mid$(ToPrint$, 30) = Format(DateAdd("d", (PO.PODATE), "12-31-1979"), "mm/dd/yyyy")
          Mid$(ToPrint$, 42) = PO.PONum
          Mid$(ToPrint$, 66) = Using$(CommaFmt$, Str$(PO.POAmt))
          Print #PrnFileNum, ToPrint$

          '--Blank line between detail and acct'g distributions
          Print #PrnFileNum, "Dept # "; PO.REQNUM
          '--Print Distribution Label
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 2) = "Accounting Distribution:"
          Print #PrnFileNum, ToPrint$
          '--Print Field Titles
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 4) = "Item Desc          Account Number           Distribution"
          Print #PrnFileNum, ToPrint$
          '--Print Accounting Distributions
          TotTranDist# = 0
          '--Loop Thru distributions to print and summarize
          For AcctDist = 1 To 36
            If Len(QPTrim$(PO.ITEMS(AcctDist).ACCTNO)) Then
              If AcctFind(PO.ITEMS(AcctDist).ACCTNO) > 0 Then
              '--Print this distribution
              LSet ToPrint$ = ""
              Mid$(ToPrint$, 4) = Mid$(QPTrim$(PO.ITEMS(AcctDist).Desc), 1, 24)
              Mid$(ToPrint$, 23) = PO.ITEMS(AcctDist).ACCTNO

              Mid$(ToPrint$, TranCol) = Using$(CommaFmt$, Str$(PO.ITEMS(AcctDist).EXT))
              Print #PrnFileNum, ToPrint$
              
              TotTranDist# = TotTranDist# + PO.ITEMS(AcctDist).EXT

              Linecnt = Linecnt + 1
              If Linecnt >= MaxLines Then
                Print #PrnFileNum, FF$
                GoSub PrintHeader
              End If

              'Sum by fund
              Found = False
              For Fund = 1 To NumFunds
                FundNum$ = Left$(PO.ITEMS(AcctDist).ACCTNO, GLFundLen)
                If FundNum$ = FundList$(Fund) Then
                  Found = True
                  FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + PO.ITEMS(AcctDist).EXT)
                  Exit For
                End If
              Next

              If Not Found Then 'Oh.Shit = True
                 Unload FrmShowPctComp
                 MsgBox "Error, Invalid Fund Found.", vbOKOnly, "Error"
                 ActivateControls frmPrnPOApproved
                 Exit Sub
              End If
            End If
            End If              'Active transaction test

          Next  'Distribution

          '--Summary line after last distribution
          LSet ToPrint$ = ""
          Mid$(ToPrint$, TranCol) = DistSumLine$
          Print #PrnFileNum, ToPrint$

          '--Transaction Distribution Totals
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 4) = "Total Distributed"
          Mid$(ToPrint$, TranCol) = Using$(CommaFmt$, Str$(TotTranDist#))
          Print #PrnFileNum, ToPrint$
          Linecnt = Linecnt + 2
          If Linecnt >= MaxLines Then
            Print #PrnFileNum, FF$
            GoSub PrintHeader
          End If

          '--2 blank lines before next distribution
          ToPrint$ = String$(80, "=")
          Print #PrnFileNum, ToPrint$
          Print #PrnFileNum,
          Linecnt = Linecnt + 2
          If Linecnt >= MaxLines Then
            Print #PrnFileNum, FF$
            GoSub PrintHeader
          End If
          End If
        End If  'Not deleted test
      End If    'Not Correct Dept
    End If      'Not Approved Flag
  Next          'Transaction
  If TranCnt < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If

  'PRINT #PrnFileNum, STRING$(80, "-")
  If Linecnt > 45 Then
    Print #PrnFileNum, FF$
  End If

  '--Summary
  LSet ToPrint$ = ""
  LSet ToPrint$ = "File Totals:"
  Print #PrnFileNum, ToPrint$

  LSet ToPrint$ = ""
  LSet ToPrint$ = "Number of Transactions"
  Mid$(ToPrint$, 31) = Using$("####", Str$(TranCnt))
  Print #PrnFileNum, ToPrint$

  LSet ToPrint$ = ""
  LSet ToPrint$ = "Grand Totals"
  Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(TransTotal#))
  Print #PrnFileNum, ToPrint$

  Print #PrnFileNum,
  LSet ToPrint$ = ""
  LSet ToPrint$ = "Summary by Fund:"
  Print #PrnFileNum, ToPrint$

  For cnt = 1 To NumFunds
    If FundGrdTot#(cnt) > 0 Then
      LSet ToPrint$ = ""
      LSet ToPrint$ = "Fund" + " " + FundList$(cnt)
      Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(FundGrdTot#(cnt)))
      Print #PrnFileNum, ToPrint$
      GrdTot# = Round#(GrdTot# + FundGrdTot#(cnt))
    End If
  Next

  LSet ToPrint$ = ""
  LSet ToPrint$ = "Total All Funds"
  Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(GrdTot#))
  Print #PrnFileNum, ToPrint$
  Print #PrnFileNum, FF$
  If TranCnt < 1 Then
    Print #PrnFileNum, "** No Approved Purchase Orders **"
  End If
  Close
  ViewPrint FileName$, RegTitle$
  KillFile FileName$
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  Exit Sub

PrintHeader:
  Page = Page + 1
  Title$(1) = "Purchase Orders : APPROVED PO's"
  Title$(2) = "Run Date: " + Date$ + "                                                  Page : " & Page
  Title$(3) = "Dept #: " + DeptNumber$
  Title$(4) = "Vendor                       Date        PO No"
  Title$(5) = String$(80, "-")

  Print #PrnFileNum, GLUserName$
  For HCnt = 1 To 5
    Print #PrnFileNum, Title$(HCnt)
  Next
  Linecnt = 6
  Return
CancelExit:
  Exit Sub
End Sub

Private Sub fpcboDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDepartment.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboDepartment.ListIndex = -1
    fpcboDepartment.Action = ActionClearSearchBuffer
  End If
  If fpcboDepartment.ListDown <> True Then
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

Private Sub fpcboPOs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPOs.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboPOs.ListIndex = -1
    fpcboPOs.Action = ActionClearSearchBuffer
  End If
  If fpcboPOs.ListDown <> True Then
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

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
