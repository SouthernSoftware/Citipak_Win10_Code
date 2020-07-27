VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.OCX"
Begin VB.Form ZOldBud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7152
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   9108
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7152
   ScaleWidth      =   9108
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   288
      Left            =   96
      TabIndex        =   0
      Top             =   576
      Width           =   2916
      _Version        =   196608
      _ExtentX        =   5143
      _ExtentY        =   508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   3
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
      ColDesigner     =   "ZOldBud.frx":0000
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
      Height          =   492
      Left            =   10128
      TabIndex        =   1
      Top             =   7632
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   6852
      Width           =   9108
      _ExtentX        =   16066
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5313
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   5313
            TextSave        =   "10:55 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   5313
            TextSave        =   "4/17/02"
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
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6156
      Left            =   192
      TabIndex        =   3
      Top             =   1296
      Width           =   11868
      _Version        =   196613
      _ExtentX        =   20934
      _ExtentY        =   10859
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421504
      MaxCols         =   8
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "ZOldBud.frx":039F
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Preparation Worksheet"
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
      TabIndex        =   4
      Top             =   336
      Width           =   4812
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   3120
      Top             =   96
      Width           =   6492
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000013&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   6324
      Left            =   96
      Top             =   1200
      Width           =   12060
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   972
      Left            =   3120
      Top             =   0
      Width           =   6492
   End
End
Attribute VB_Name = "ZOldBud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLAcct As GLAcctRecType
Dim tempstr As String
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub Form_Load()
'  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, tempstr As String
'  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
'  OpenAcctFile AcctFile
'  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
'  NumAccts = LOF(AcctFile) / Len(GLAcct)
'  vaSpread1.Col = 1
'  vaSpread1.Row = 1
'  vaSpread1.CellType = CellTypeComboBox
'
'  For CntA = 1 To NumAIdxRecs
'    Get AcctIdxFileNum, CntA, GLAcctidx
'    Get AcctFile, GLAcctidx.RecNum, GLAcct
'    If GLAcct.Deleted = 0 Then
'        If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
'          vaSpread1.TypeComboBoxIndex = vaSpread1.TypeComboBoxCount + 1
'          tempstr = tempstr + QPTrim(GLAcct.Num) + "  " + (GLAcct.Title) + "    " + Str$(GLAcctidx.RecNum) + Chr(9)
'
'          'vaSpread1.TypeComboBoxList = Str$(GLAcctidx.RecNum) + "  " + QPTrim(GLAcct.Num) + "  " + QPTrim(GLAcct.Title)
'          'tempstr = tempstr + Chr(9)
'          'Exit For
'        End If
'      End If
'  Next
'  Close AcctFile
'  Close AcctIdxFileNum
'          vaSpread1.TypeComboBoxList = tempstr
  Dim cnt As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  'vaSpread1.TypeComboBoxhWnd = fpcboAcctNumNa.hwnd
 ' vaSpread1.Col = 1
  BudAcctNumName
 'Next
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  frmBudgetMaintMenu.Show
  Unload frmBudPrepMaint
End Sub

Public Function BudAcctNumName()
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  'vaSpread1.Col = Col
  'vaSpread1.Row = Row
  'vaSpread1.CellType = CellTypeComboBox
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct
    If GLAcct.DELETED = 0 Then
      If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
        'vaSpread1.TypeComboBoxIndex = vaSpread1.TypeComboBoxCount + 1
        tempstr = tempstr + QPTrim(GLAcct.Num) + "  " + (GLAcct.title) + "    " + Str$(GLAcctidx.RecNum) + Chr(9) + QPStrip(GLAcct.Num)
      End If
    End If
  Next
  Close AcctFile
  Close AcctIdxFileNum
  'vaSpread1.TypeComboBoxList = tempstr
End Function

Private Sub vaSpread1_ComboDropDown(ByVal Col As Long, ByVal Row As Long)
  vaSpread1.Col = Col
  vaSpread1.Row = Row
  vaSpread1.CellType = CellTypeComboBox
  vaSpread1.TypeComboBoxList = tempstr
End Sub


