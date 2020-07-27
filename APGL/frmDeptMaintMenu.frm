VERSION 5.00
Begin VB.Form frmDeptMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department Maintenance "
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmDeptMaintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDeptAddEdit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Add/Change/Delete Departments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3612
   End
   Begin VB.CommandButton cmdDeptPrintList 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Print Department Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      HelpContextID   =   50
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdDeptSort 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Department Index &Utility"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitDeptMaintMenu 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Department Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   3612
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT MAINTENANCE MENU"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   5772
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CHART OF ACCOUNTS MAINTENANCE MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   6852
   End
End
Attribute VB_Name = "frmDeptMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Private Sub cmdDeptAddEdit_Click()
  frmDeptEntryEdit.Show
  Unload frmDeptMaintMenu
End Sub
Private Sub cmdDeptPrintList_Click()
  If Exist("GLDept.DAT") Then
    frmReportOpt.Show 1
    If rptopt = 1 Then
      Call PrintDeptListReport
    ElseIf rptopt = 2 Then
      Call PrintDeptListReport2
    End If
  Else
    MsgBox "No Departments To Print", vbOKOnly, "No Depts"
  End If
End Sub
Private Sub cmdDeptSort_Click()
  If Exist("GLDept.DAT") Then
    frmSortDept.Show 1
  Else
    MsgBox "No Departments To Sort", vbOKOnly, "No Depts"
  End If
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpDepartment
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdExitDeptMaintMenu_Click()
  frmGLSetupMenu.Show
  Unload frmDeptMaintMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitDeptMaintMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub PrintDeptListReport()
  Dim DeptIdxFileNum As Integer, NumDIdxRecs As Integer
  Dim DeptFileNum As Integer, NumDepts As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
  Dim ReportFile As String, ToPrint As String, Newrp As String
  Dim Header As String
  Dim GLDeptIdx As GLDeptIndexType
  Dim GLDept As GLDeptRecType
 '  Stop
   'Define vars used for printing
  Header$ = "Master Department Listing"
'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
  OpenDeptIdx DeptIdxFileNum, NumDIdxRecs
  OpenDeptFile DeptFileNum, NumDepts
  PRNFile = FreeFile
  Newrp = "DptLst"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  For cnt = 1 To NumDIdxRecs
    Get DeptIdxFileNum, cnt, GLDeptIdx
    Get DeptFileNum, GLDeptIdx.RecNum, GLDept
    Howmany = Howmany + 1
    ToPrint$ = ""
    ToPrint$ = GLDept.DeptNum + "~" + GLDept.Title + "~"
    Print #PRNFile, ToPrint$
  Next
  Close
  Load frmLoadingRpt
  ARptListings.Label1.Caption = "Dept Number"
  ARptListings.Label2.Caption = "Title"
  ARptListings.Label3.Visible = False
  ARptListings.Label4.Caption = "Departments Listed"
  ARptListings.Total = Howmany
  ARptListings.txtDate = Now
  ARptListings.txtTown = GLUserName$
  ARptListings.Title.Caption = Header$
  ARptListings.GetName ReportFile$
  ARptListings.startrpt

  Exit Sub
End Sub
Private Sub PrintDeptListReport2()
  Dim MaxLines As Integer, DeptIdxFileNum As Integer, NumDIdxRecs As Integer
  Dim DeptFileNum As Integer, NumDepts As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
  Dim ReportFile As String, ToPrint As String, Newrp As String
  Dim FF As String, Header As String
  Dim GLDeptIdx As GLDeptIndexType
  Dim GLDept As GLDeptRecType
 '  Stop
   'Define vars used for printing
  MaxLines = 55
  FF$ = Chr$(12)
  Header$ = "Master Department Listing"
'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
  OpenDeptIdx DeptIdxFileNum, NumDIdxRecs
  OpenDeptFile DeptFileNum, NumDepts
  PRNFile = FreeFile
  Newrp = "DptLst"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  GoSub PrintDeptPageHeader
  For cnt = 1 To NumDIdxRecs
    Get DeptIdxFileNum, cnt, GLDeptIdx
    Get DeptFileNum, GLDeptIdx.RecNum, GLDept
    Howmany = Howmany + 1
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 4) = GLDept.DeptNum
    Mid$(ToPrint$, 22) = GLDept.Title
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #PRNFile, FF$
      GoSub PrintDeptPageHeader
    End If
  Next
  Print #PRNFile,
  Print #PRNFile, Howmany; "Departments listed."
  Print #PRNFile, FF$
  Close
  ViewPrint ReportFile$, "Department Listing Report"
  Kill ReportFile$
  Exit Sub
PrintDeptPageHeader:
  Print #PRNFile, Header$
  Print #PRNFile,
  Print #PRNFile, " Department Number       Title"
  Print #PRNFile, String$(80, "-")
  Linecnt = 4
  Return
End Sub

