VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnPOForms 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Purchase Order Forms"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnPOForms.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboDepartment 
      Height          =   405
      Left            =   5910
      TabIndex        =   0
      Top             =   3450
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
      ColDesigner     =   "frmPrnPOForms.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPOs 
      Height          =   405
      Left            =   5925
      TabIndex        =   1
      Top             =   4200
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
      ColDesigner     =   "frmPrnPOForms.frx":0C7D
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5925
      TabIndex        =   2
      Top             =   4920
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
      ColDesigner     =   "frmPrnPOForms.frx":0FAC
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
      Left            =   6288
      TabIndex        =   4
      Top             =   5664
      Width           =   1452
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
      Left            =   4452
      TabIndex        =   3
      Top             =   5664
      Width           =   1452
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
            TextSave        =   "4:02 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "11/10/2006"
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
      Caption         =   "Form Type: "
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
      Top             =   4944
      Width           =   1620
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3516
      Left            =   3336
      Top             =   3000
      Width           =   5532
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   1296
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Forms"
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
      TabIndex        =   8
      Top             =   1536
      Width           =   4836
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
      Top             =   3528
      Width           =   1356
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
      TabIndex        =   6
      Top             =   4272
      Width           =   1524
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   1176
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
Attribute VB_Name = "frmPrnPOForms"
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
Dim PO As POFORMRecType2
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Sub cmdExit_Click()
  frmPOProcessMenu.Show
  Unload frmPrnPOForms
End Sub

Private Sub cmdOk_Click()
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt = 1 Then
    PrintPOForms
  ElseIf rptopt = 2 Then
    PrintPOForms2
  End If
  cmdExit_Click
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
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpPOForms
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


Private Sub PrintPOForms()
  Dim i As Integer, ItemsPrinted As Integer, TPTop As String, TPShip As String
  Dim LCnt As Single, ToPrintH As String, TPBot As String
  Dim PRNFile As Integer, TempPO As String
  Dim PRNfileName As String, ToPrint As String, DetailLines As Integer
  Dim Header As String, DistSumLine As String, TPBod As String
  Dim TransTotal As Double, TranCnt As Integer, FileName As String
  Dim RegTitle As String, TranCol As Integer, CashCol As Integer
  Dim Header1 As String, Header2 As String, Header3 As String, Header4 As String
  Dim DeptNumber As String, NYBeg As Integer, NYEnd As Integer
  Dim ThisDist As Double, Accttotal As Double, Title As String
  Dim ThisAcct As String, AcctCnt As Integer, WhatAcct As Integer
  Dim Found As Boolean, Fund As Integer, FundNum As String, TempPrice As Double
  Dim POFile As Integer, NumRecs As Integer, HamFlag As Integer
  Dim POEditFile As Integer, NumEdTrans As Integer, PODept As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, AcctDist As Integer
  Dim IdxFile As Integer, NumIdxRecs As Integer, Transaction As Integer
  Dim AcctFile As Integer, NumAccts As Integer, T As Integer
  Dim TransFileNum As Integer, NumTrans As Long, PrnFileNum As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, TotTranDist As Double
  Dim ActuallyPrn As Long
  Dim TPBody$(1 To 36)
  Title$ = "PO Forms"
  fpcboDepartment.col = 1
  DeptNumber$ = QPTrim$(fpcboDepartment.ColText)
  TempPO$ = QPTrim(fpcboPOs.Text)
  
  '--Start printing forms
  
  ReDim POCont(1) As POControlRecType
  OpenPOFile POFile, NumRecs
  
  If LOF(POFile) > 0 Then
    Get POFile, 1, POCont(1)
  End If

'  If InStr(UCase$(GLUserName$), "HAMLET") Then
'    HamFlag = 1
'  End If
  FrmShowPctComp.Label1 = "Processing PO Forms"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  mnuOptions.Enabled = False
  OpenPOEditFile POEditFile, NumEdTrans
  PRNfileName$ = "POFORMS.PRN"
  PRNFile = FreeFile
  Open PRNfileName$ For Output As #PRNFile
  ActuallyPrn = 0
  For T = 1 To NumEdTrans
    Get POEditFile, T, PO
    FrmShowPctComp.ShowPctComp T, NumEdTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    ToPrint$ = " "
    ToPrintH$ = " "
    TPBot$ = " "
    TPTop$ = " "
    TPShip$ = " "
    TPBod$ = " "
    PODept$ = QPTrim$(PO.REQNUM)
    If PO.Deleted <> True Then 'Deleted Skip
      If Left$(PO.PONum, 3) <> "N/A" Then
        If PODept$ = DeptNumber$ Or DeptNumber$ = "All" Then
          If TempPO$ = QPTrim(PO.PONum) Or TempPO$ = "All" Then

          PO.Deleted = 1        'Mark as Printed
          Put POEditFile, T, PO
          ItemsPrinted = 0
          GoSub PrintPOHeader
          '--Loop Thru Items
          DetailLines = 0
          For i = 1 To 36
            TPBody$(i) = " "
            
            If Len(QPTrim$(PO.ITEMS(i).ACCTNO)) Then
              TPBody$(i) = QPTrim(PO.ITEMS(i).STKNO)
              TPBody$(i) = TPBody$(i) + "~" + Left$(PO.ITEMS(i).Desc, 40)
              If PO.ITEMS(i).QUAN > -999999.99 Then
                TPBody$(i) = TPBody$(i) + "~" + Using("#######.####", PO.ITEMS(i).QUAN)
              Else
                TPBody$(i) = TPBody$(i) + "~ ~"
              End If
              If PO.ITEMS(i).PRICE > -999999.99 Then
                If PO.ITEMS(i).PRICE <> 0 Then
                  TPBody$(i) = TPBody$(i) + "~" + Using("###########.####", PO.ITEMS(i).PRICE)
                Else
                  TempPrice = PO.ITEMS(i).EXT / PO.ITEMS(i).QUAN
                  TPBody$(i) = TPBody$(i) + "~" + Using("###########.####", TempPrice)
                End If
              Else
                TPBody$(i) = TPBody$(i) + "~ ~"
              End If
              TPBody$(i) = TPBody$(i) + "~" + Using("###########.##", PO.ITEMS(i).EXT)
              TPBody$(i) = TPBody$(i) + "~" + QPStrip(PO.ITEMS(i).ACCTNO) + "~"
            Else
              TPBody$(i) = " ~ ~ ~ ~ ~ ~"
            End If    'Active transaction test
            TPBod$ = TPBod$ + TPBody$(i)
          Next
          TPBot$ = Using("$##,###,###,###.##", PO.POAmt)
          TPBot$ = TPBot$ + "~" + QPTrim(PO.Addinst1)
          TPBot$ = TPBot$ + "~" + QPTrim(PO.Addinst2)
          TPBot$ = TPBot$ + "~" + QPTrim(PO.Addinst3)
 'WRITE RECORD TO PRINT FILE
        ToPrint$ = ToPrintH$ + "~" + TPTop$ + "~" + TPShip$ + "~" + TPBod$ + TPBot$
        Print #PRNFile, ToPrint$
        ActuallyPrn = ActuallyPrn + 1
        End If
        End If
        End If
      End If
  Next
  Close
  If ActuallyPrn < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    MsgBox "No PO's To Print", vbOKOnly, "No PO's"
    Exit Sub
  End If
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  frmLoadingRpt.Show
  ARptPOForms.GetName PRNfileName$
  ARptPOForms.startrpt
  Exit Sub
'ViewPrint PRNfileName$, Title$
PrintPOHeader:
  Header1$ = " "
  Header2$ = " "
  Header3$ = " "
  Header4$ = " "
  ToPrintH$ = " "
  TPTop$ = " "
  TPShip$ = " "
  Header1$ = POCont(1).Header1
  Header1$ = QPTrim$(Header1$)
  Header2$ = POCont(1).Header2
  Header2$ = QPTrim$(Header2$)
  Header3$ = POCont(1).Header3
  Header3$ = QPTrim$(Header3$)
  Header4$ = POCont(1).Header4
  Header4$ = QPTrim$(Header4$)
  ToPrintH$ = Header1$ + "~" + Header2$ + "~" + Header3$ + "~" + Header4$
  
  TPTop$ = Format(DateAdd("d", (PO.PODATE), "12-31-1979"), "mm/dd/yyyy") + "~" + PO.PONum
  
  TPTop$ = TPTop$ + "~" + PO.VNDRINF1 + "~" + PO.SHPLINE1
  TPTop$ = TPTop$ + "~" + PO.VNDRINF2 + "~" + PO.SHPLINE2
  TPTop$ = TPTop$ + "~" + PO.VNDRINF3 + "~" + PO.SHPLINE3
  TPTop$ = TPTop$ + "~" + PO.VNDRINF4 + "~" + PO.SHPLINE4 + "~" + PO.SHPLINE5
  TPShip$ = PO.Shipvia + "~" + PO.FOB + "~" + PO.REQNUM
  TPShip$ = TPShip$ + "~" + PO.SHIPON + "~" + PO.Terms
  Return
CancelExit:
  Exit Sub
End Sub
Private Sub PrintPOForms2()
  Dim MaxLines As Integer, i As Integer, ItemsPrinted As Integer
  Dim Linecnt As Integer, Page As Integer, LCnt As Single
  Dim PRNFile As Integer, TempPO As String
  Dim PRNfileName As String, ToPrint As String, DetailLines As Integer
  Dim FF As String, Header As String, DistSumLine As String
  Dim TransTotal As Double, TranCnt As Integer, FileName As String
  Dim RegTitle As String, TranCol As Integer, CashCol As Integer
  Dim Header1 As String, Header2 As String, Header3 As String, Header4 As String
  Dim DeptNumber As String, NYBeg As Integer, NYEnd As Integer
  Dim ThisDist As Double, Accttotal As Double, Title As String
  Dim ThisAcct As String, AcctCnt As Integer, WhatAcct As Integer
  Dim Found As Boolean, Fund As Integer, FundNum As String
  Dim POFile As Integer, NumRecs As Integer, HamFlag As Integer
  Dim POEditFile As Integer, NumEdTrans As Integer, PODept As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, AcctDist As Integer
  Dim IdxFile As Integer, NumIdxRecs As Integer, Transaction As Integer
  Dim AcctFile As Integer, NumAccts As Integer, T As Integer
  Dim TransFileNum As Integer, NumTrans As Long, PrnFileNum As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, TotTranDist As Double
  Dim TempPrice As Double, ActuallyPrn As Long
  Title$ = "PO Forms"
  FF$ = Chr$(12)
  fpcboDepartment.col = 1
  DeptNumber$ = QPTrim$(fpcboDepartment.ColText)
  TempPO$ = QPTrim(fpcboPOs.Text)
  ToPrint$ = Space(80)
  '--Start printing forms

  ReDim POCont(1) As POControlRecType
  OpenPOFile POFile, NumRecs
  
  If LOF(POFile) > 0 Then
    Get POFile, 1, POCont(1)
  End If

  If InStr(UCase$(GLUserName$), "HAMLET") Then
    HamFlag = 1
  End If
  FrmShowPctComp.Label1 = "Processing PO Forms"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  mnuOptions.Enabled = False
  OpenPOEditFile POEditFile, NumEdTrans
  PRNfileName$ = "POFORMS.PRN"
  PRNFile = FreeFile
  Open PRNfileName$ For Output As #PRNFile
  ActuallyPrn = 0
  For T = 1 To NumEdTrans
    Get POEditFile, T, PO
    FrmShowPctComp.ShowPctComp T, NumEdTrans
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
    If PO.Deleted <> True Then 'Deleted Skip
      If Left$(PO.PONum, 3) <> "N/A" Then
        If PODept$ = DeptNumber$ Or DeptNumber$ = "All" Then
          If TempPO$ = QPTrim(PO.PONum) Or TempPO$ = "All" Then

          PO.Deleted = 1        'Mark as Printed
          Put POEditFile, T, PO
          ActuallyPrn = ActuallyPrn + 1
          If HamFlag = 1 Then
            ItemsPrinted = 0
            For LCnt! = 1 To 3
              Print #PRNFile, ""
            Next LCnt!
            ToPrint$ = Space(90)
            Mid$(ToPrint$, 67) = PO.PONum
            Print #PRNFile, ToPrint$
            For LCnt! = 5 To 7
              Print #PRNFile, ""
            Next LCnt!
            Print #PRNFile, Tab(12); PO.SHPLINE1
            Print #PRNFile, Tab(12); PO.SHPLINE2;
            Print #PRNFile, Tab(56); Format(DateAdd("d", (PO.PODATE), "12-31-1979"), "mm/dd/yyyy")
            Print #PRNFile, Tab(12); PO.SHPLINE3
            Print #PRNFile, Tab(12); PO.SHPLINE4
            Print #PRNFile, Tab(12); PO.SHPLINE5 '""  'Line 12 Here
            Print #PRNFile, Tab(63); PO.Terms
            Print #PRNFile, Tab(12); PO.VNDRINF1
            Print #PRNFile, Tab(12); PO.VNDRINF2
            Print #PRNFile, Tab(12); PO.VNDRINF3
            Print #PRNFile, Tab(12); PO.VNDRINF4
            Print #PRNFile, Tab(12); ""
            For LCnt! = 19 To 21
              Print #PRNFile, ""
            Next LCnt!

            '--Loop Thru Items
            DetailLines = 0
            For i = 1 To 36
              ToPrint$ = Space(90)
              If Len(QPTrim$(PO.ITEMS(i).ACCTNO)) Then
                Mid$(ToPrint$, 1) = i
                Mid$(ToPrint$, 5) = Left$(PO.ITEMS(i).Desc, 20)
                Mid$(ToPrint$, 38) = QPStrip(PO.ITEMS(i).ACCTNO)

                If PO.ITEMS(i).QUAN > -999999.99 Then
                  Mid$(ToPrint$, 52) = Using("####.##", PO.ITEMS(i).QUAN)
                End If
                If PO.ITEMS(i).PRICE > -999999.99 Then
                If PO.ITEMS(i).PRICE <> 0 Then
                  Mid$(ToPrint$, 63) = Using("#####.##", PO.ITEMS(i).PRICE)
                Else
                  TempPrice = PO.ITEMS(i).EXT / PO.ITEMS(i).QUAN
                  Mid$(ToPrint$, 63) = Using("###.####", TempPrice)
                End If
                End If
                Mid$(ToPrint$, 74) = Using("$###,###.##", PO.ITEMS(i).EXT)
                Print #PRNFile, ToPrint$
                ToPrint$ = Space(90)
                Mid$(ToPrint$, 5) = Right$(PO.ITEMS(i).Desc, 20)
                Print #PRNFile, ToPrint$
                DetailLines = DetailLines + 2
              End If            'Active transaction test

              If Len(QPTrim$(PO.ITEMS(i).ACCTNO)) = 0 And Len(QPTrim$(PO.ITEMS(i).Desc)) Then
                Print #PRNFile, Tab(5); Left$(PO.ITEMS(i).Desc, 20)
                Print #PRNFile, Tab(5); Right$(PO.ITEMS(i).Desc, 20)
                DetailLines = DetailLines + 2
              End If            'Active transaction test

            Next

           For LCnt! = 21 + DetailLines To 52
             Print #PRNFile, ""
           Next LCnt!
      
           Print #PRNFile, Tab(74); Using("$###,###.##", PO.POAmt)
           Print #PRNFile, FF$

        Else
          ItemsPrinted = 0

          GoSub PrintPOHeader

          '--Loop Thru Items
          DetailLines = 0
          For i = 1 To 36
            LSet ToPrint$ = ""
            If Len(QPTrim$(PO.ITEMS(i).ACCTNO)) Then
              LSet ToPrint$ = PO.ITEMS(i).STKNO
              Mid$(ToPrint$, 10) = Left$(PO.ITEMS(i).Desc, 20)
              If PO.ITEMS(i).QUAN > -999999.99 Then
                Mid$(ToPrint$, 34) = Using("####.##", PO.ITEMS(i).QUAN)
              End If
              If PO.ITEMS(i).PRICE > -999999.99 Then
              If PO.ITEMS(i).PRICE <> 0 Then
                  Mid$(ToPrint$, 44) = Using("######.##", PO.ITEMS(i).PRICE)
                Else
                  TempPrice = PO.ITEMS(i).EXT / PO.ITEMS(i).QUAN
                  Mid$(ToPrint$, 44) = Using("####.####", TempPrice)
                End If
              End If
              Mid$(ToPrint$, 54) = Using("######.##", PO.ITEMS(i).EXT)
              Mid$(ToPrint$, 65) = QPStrip(PO.ITEMS(i).ACCTNO)
              Print #PRNFile, ToPrint$
              Mid$(ToPrint$, 10) = Right$(PO.ITEMS(i).Desc, 20)
              ToPrint$ = Space(80)
              Print #PRNFile, ToPrint$
              DetailLines = DetailLines + 2
            End If              'Active transaction test
            If Len(QPTrim$(PO.ITEMS(i).ACCTNO)) = 0 And Len(QPTrim$(PO.ITEMS(i).Desc)) Then
               Print #PRNFile, Tab(12); Left$(PO.ITEMS(i).Desc, 20)
              Print #PRNFile, Tab(12); Right$(PO.ITEMS(i).Desc, 20)
              DetailLines = DetailLines + 2
            End If              'Active transaction test
          Next

          Print #PRNFile,
          Print #PRNFile, Tab(20); "Total Purchase Order Amount "; Tab(48); Using("$###,###,###.##", PO.POAmt)
          Print #PRNFile,
          Print #PRNFile, Tab(5); "Additional Instruction"
          Print #PRNFile, Tab(5); PO.Addinst1
          Print #PRNFile, Tab(5); PO.Addinst2
          Print #PRNFile, Tab(5); PO.Addinst3
          Print #PRNFile, ""
          Print #PRNFile, Tab(5); "This Instrument has been preaudited in the manner required by the"
          Print #PRNFile, Tab(5); "Local Government Budget and Fiscal Control Act."
          Print #PRNFile, ""  'Tab(5); "duly authorized, as required by the LOCAL GOVERNMENT ACT."
          Print #PRNFile, ""
          Print #PRNFile, Tab(5); "Purchasing Officer Signature: _____________________________________"
          Print #PRNFile, ""
          Print #PRNFile, Tab(5); "Finance Officer Signature: ________________________________________"
          Print #PRNFile, FF$
              End If
        End If
        End If
      End If
    End If
  Next
  Close
  If ActuallyPrn < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    MsgBox "No PO's To Print", vbOKOnly, "No PO's"
  Else
    ViewPrint PRNfileName$, Title$
  End If
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  Exit Sub

PrintPOHeader:
  Print #PRNFile, ""
  Header1$ = POCont(1).Header1
  Header1$ = QPTrim$(Header1$)
  Header2$ = POCont(1).Header2
  Header2$ = QPTrim$(Header2$)
  Header3$ = POCont(1).Header3
  Header3$ = QPTrim$(Header3$)
  Header4$ = POCont(1).Header4
  Header4$ = QPTrim$(Header4$)
  Print #PRNFile, Tab(33); "PURCHASE ORDER"
     Print #PRNFile, ""
  Print #PRNFile, Tab(40 - (Len(Header1$) / 2)); Header1$
  Print #PRNFile, Tab(40 - (Len(Header2$) / 2)); Header2$
  Print #PRNFile, Tab(40 - (Len(Header3$) / 2)); Header3$
  Print #PRNFile, Tab(40 - (Len(Header4$) / 2)); Header4$
  Print #PRNFile, Tab(5); "PO Date: "; Format(DateAdd("d", (PO.PODATE), "12-31-1979"), "mm/dd/yyyy"); Tab(55); "PO # "; PO.PONum
  Print #PRNFile, ""
  Print #PRNFile, Tab(5); "Vendor:"; Tab(45); "Ship To:"
  Print #PRNFile, Tab(5); PO.VNDRINF1; Tab(45); PO.SHPLINE1
  Print #PRNFile, Tab(5); PO.VNDRINF2; Tab(45); PO.SHPLINE2
  Print #PRNFile, Tab(5); PO.VNDRINF3; Tab(45); PO.SHPLINE3
  Print #PRNFile, Tab(5); PO.VNDRINF4; Tab(45); PO.SHPLINE4
  Print #PRNFile, Tab(5); ""; Tab(45); PO.SHPLINE5
  Print #PRNFile, ""
  Print #PRNFile, String$(79, "_")
  Print #PRNFile, "Ship Via: "; PO.Shipvia; Tab(35); "  FOB: "; PO.FOB; Tab(62); "Dept #"; PO.REQNUM
  Print #PRNFile, " Ship By: "; PO.SHIPON; Tab(35); "Terms: "; PO.Terms
  Print #PRNFile, String$(79, "_")
  Print #PRNFile, Tab(47); "Unit"; Tab(57); "Total"
  Print #PRNFile, "Stock #"; Tab(10); "Description"; Tab(37); "Quan"; Tab(47); "Price"; Tab(57); "Price"; Tab(65); "Acct Number"
  Print #PRNFile, String$(79, "_")
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
