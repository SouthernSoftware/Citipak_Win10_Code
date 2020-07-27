VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserSelect 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Maintenance"
   ClientHeight    =   8892
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12192
   Icon            =   "frmUserSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboUsers 
      Height          =   384
      Left            =   3624
      TabIndex        =   0
      Top             =   3672
      Width           =   4944
      _Version        =   196608
      _ExtentX        =   8721
      _ExtentY        =   677
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
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
      ColDesigner     =   "frmUserSelect.frx":08CA
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "F4 &Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3432
      TabIndex        =   1
      Top             =   5040
      Width           =   1356
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "F2 &New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   5496
      TabIndex        =   2
      Top             =   5040
      Width           =   1356
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
      Height          =   444
      Left            =   7560
      TabIndex        =   3
      Top             =   5040
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Top             =   8640
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   445
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
            TextSave        =   "8:46 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "1/17/2003"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select F2 to Add New User."
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
      Left            =   2496
      TabIndex        =   8
      Top             =   3096
      Width           =   7188
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "No Users - Add New or Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4518
      TabIndex        =   7
      Top             =   1992
      Visible         =   0   'False
      Width           =   3156
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the User in the list below and F4 to Edit or"
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
      Left            =   2508
      TabIndex        =   6
      Top             =   2712
      Width           =   7188
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Height          =   3252
      Left            =   2406
      Top             =   2496
      Width           =   7380
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   780
      Left            =   2580
      Top             =   816
      Width           =   7020
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3312
      TabIndex        =   4
      Top             =   1008
      Width           =   5580
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   2592
      Top             =   696
      Width           =   7020
   End
End
Attribute VB_Name = "frmUserSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim CitiPass As CitiPassType
Dim TempRec As Integer

Private Sub cmdEdit_Click()
  If fpcboUsers.ListIndex <> -1 Then
    fpcboUsers.col = 0
    TempRec = QPTrim(fpcboUsers.ColText)
    frmEnterEditPass.Rec2Form (TempRec)
    frmEnterEditPass.Show
  Else
    MsgBox "You Must First Select A User To Edit.", vbOKOnly, "Invalid Selection"
  End If
End Sub

Private Sub cmdExit_Click()
  'ClearInUse (PWcnt)
'  LevelPass = 0
'  PWcnt = 0
'  PWUser = ""
  frmMainMenu.Show
  Unload frmUserSelect
End Sub

Private Sub cmdNew_Click()
  fpcboUsers.Clear
  fpcboUsers.ListIndex = -1
  DoEvents
  frmEnterEditPass.Recnum = 0
  frmEnterEditPass.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF2:
      cmdNew_Click
      KeyCode = 0
    Case vbKeyF4:
      cmdEdit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim NumPassRecs As Integer, PassRecLen As Integer
  PassRecLen = Len(CitiPass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

If NumPassRecs > 0 Then
  FillUsers fpcboUsers
Else
  fpcboUsers.Enabled = False
  cmdEdit.Enabled = False
  Label3.Visible = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  MainLog "Out PW Entry/Edit"
  LevelPass = 0
  PWcnt = 0
  PWUser = ""
  Close CPAdminhand
  frmMainMenu.Show
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub fpcboUsers_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboUsers.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboUsers.ListIndex = -1
    fpcboUsers.Action = ActionClearSearchBuffer
  End If
  If fpcboUsers.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdEdit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Public Sub FillUsers(txt As fpCombo)
  Dim NumPassRecs As Integer, cnt As Integer, PassRecLen As Integer
  'OpenCitiPassFile CitiPassFile, NumPassRecs
  PassRecLen = Len(CitiPass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  txt.Row = -1
  For cnt = 1 To NumPassRecs
    Get CPAdminhand, cnt, CitiPass
    If Not CitiPass.DelFlag Then
      txt.InsertRow = Str$(cnt) & Chr$(9) & QPTrim(CitiPass.UserName)
    End If
  Next
  'Close CitiPassFile
End Sub
'Private Sub Cleanup()
'  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
'  Dim NPRecs As Integer, NCitiPassFile As Integer, NRecLen As Integer
'  OpenCitiPassFile CitiPassFile, NumPassRecs
'  NRecLen = Len(CitiPass)
'  If NumPassRecs = 0 Then
'    Close
'    Exit Sub
'  End If
'
'  NCitiPassFile = FreeFile
'  Open "NCitipass.dat" For Output As #NCitiPassFile
'  Close NCitiPassFile
'  Open "NCitiPass.dat" For Random Shared As NCitiPassFile Len = NRecLen
'  For cnt = 1 To NumPassRecs
'    Get CitiPassFile, cnt, CitiPass
'    If Not CitiPass.DelFlag Then
'      Put NCitiPassFile, , CitiPass
'    End If
'  Next
'  Close CitiPassFile
'  Close NCitiPassFile
'  Kill "CitiPass.dat"
'  Name "NCitiPass.dat" As "CitiPass.dat"
'End Sub

