VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch4Dup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search For Duplicate Accounts"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmSearch4Dup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Go"
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
      Left            =   4014
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4848
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
      Left            =   6846
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4848
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8508
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
            TextSave        =   "12:30 PM"
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
         Size            =   10.8
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
      Caption         =   "Select Go to Search For Account Duplications, or Escape to Exit."
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
      Height          =   828
      Left            =   3696
      TabIndex        =   4
      Top             =   3600
      Width           =   4908
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search For Duplicate Accounts"
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
      Left            =   3810
      TabIndex        =   3
      Top             =   1608
      Width           =   4572
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2748
      Left            =   3090
      Top             =   2928
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
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
Attribute VB_Name = "frmSearch4Dup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
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
      SendKeys "%G"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub cmdExit_Click()
  frmGLUtilMenu.Show
  Unload frmSearch4Dup
End Sub

Private Sub cmdGo_Click()
  frmReportOpt.Show 1
  Srch4DupeAccts
  frmGLUtilMenu.Show
  Unload frmSearch4Dup
End Sub

Private Sub Srch4DupeAccts()
  Dim Header As String, FF As String, MaxLines As Integer
  Dim AcctIdxFileNum As Integer, NumIdxRecs As Integer
  Dim PRNFile As Integer, ReportFile As String, DCnt As Integer
  Dim cnt As Integer, AccttoFind As String, SCnt As Integer
  Dim GLAcctidx As GLAcctIndexType

  '--Define vars used for printing
  Header$ = "Duplicate Account Listing"
  FF$ = Chr$(12)
  MaxLines = 55

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs

  PRNFile = FreeFile
  ReportFile$ = "DUPEACCT.PRN"
  Open ReportFile$ For Output As #PRNFile
  FrmShowPctComp.Label1 = "Searching Accounts"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdGo.Enabled = False
  Me.mnuOptions.Enabled = False
  For cnt = 1 To NumIdxRecs
    FrmShowPctComp.ShowPctComp cnt, NumIdxRecs
    Get AcctIdxFileNum, cnt, GLAcctidx
      AccttoFind$ = GLAcctidx.AcctNum
      'QPrintRC "Searching: " + AccttoFind$, 8, 17, -1
        For SCnt = 1 To NumIdxRecs
        Get AcctIdxFileNum, SCnt, GLAcctidx
        If GLAcctidx.AcctNum = AccttoFind$ And SCnt <> cnt Then
          Beep
          DCnt = DCnt + 1
          GoSub LogDupe
        End If
      Next
  Next

  If DCnt = 0 Then
    Print #PRNFile, "No duplicate account numbers found."
  End If
  Print #PRNFile, FF$
  Close
  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
  Me.mnuOptions.Enabled = True
  Load frmLoadingRpt
  If rptopt = 1 Then
    ARptErrorLog.Caption = Header$
    ARptErrorLog.GetName ReportFile$
    ARptErrorLog.startrpt
  ElseIf rptopt = 2 Then
    ViewPrint ReportFile$, Header$
    KillFile ReportFile$
  End If
Exit Sub
LogDupe:
   Print #PRNFile, AccttoFind$, GLAcctidx.RecNum
Return

End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
