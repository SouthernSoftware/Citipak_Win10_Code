VERSION 5.00
Begin VB.Form frmFunctionMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Function Maintenance"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmFunctionMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFunctionEdit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Add/Change/Delete Function"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
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
   Begin VB.CommandButton cmdFunctionList 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Print Function Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      HelpContextID   =   58
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdFunctionSort 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Function Index &Utility"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
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
   Begin VB.CommandButton cmdExitFunctionMenu 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Function Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCTION MAINTENANCE MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3282
      TabIndex        =   4
      Top             =   1440
      Width           =   5652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
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
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9840
      X2              =   9840
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
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
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
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
End
Attribute VB_Name = "frmFunctionMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpFunction
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdExitFunctionMenu_Click()
  frmGLSetupMenu.Show
  Unload frmFunctionMenu
End Sub
Private Sub cmdFunctionEdit_Click()
  frmFunctionEdit.Show
  Unload frmFunctionMenu
End Sub
Private Sub cmdFunctionList_Click()
  If Exist("GLFnct.DAT") Then
    frmReportOpt.Show 1
    If rptopt = 1 Then
      Call PrintFunctionReport
    ElseIf rptopt = 2 Then
      Call PrintFunctionReport2
    End If
  Else
    MsgBox "No Functions To Print", vbOKOnly, "No Functions"
  End If
End Sub
Private Sub cmdFunctionSort_Click()
  If Exist("GLFnct.DAT") Then
    frmSortFunction.Show 1
  Else
    MsgBox "No Functions To Sort", vbOKOnly, "No Functions"
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
        cmdExitFunctionMenu_Click
        KeyCode = 0
        DoEvents
      Case Else:
  End Select
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

Private Sub PrintFunctionReport()
  Dim FnctIdxFileNum As Integer, NumFIdxRecs As Integer
  Dim FnctFileNum As Integer, NumFncts As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Long, Howmany As Integer
  Dim ReportFile As String, ToPrint As String, Newrp As String
  Dim Header As String
  Dim FnctIdx As GLFNCTIndexType
  Dim Fnct As GLFNCTRecType
 '  Stop
   'Define vars used for printing
  Header$ = "Master Function Listing"
'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
  OpenFnctIdx FnctIdxFileNum, NumFIdxRecs
  OpenFnctFile FnctFileNum, NumFncts
  PRNFile = FreeFile
  Newrp = "FnctLst"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  For cnt = 1 To NumFIdxRecs
    Get FnctIdxFileNum, cnt, FnctIdx
    Get FnctFileNum, FnctIdx.RecNum, Fnct
    Howmany = Howmany + 1
    ToPrint$ = ""
    ToPrint$ = Fnct.FnctNum + "~" + Fnct.Title + "~"
    Print #PRNFile, ToPrint$
  Next
   Close
    
  Load frmLoadingRpt
  ARptListings.Label1.Caption = "Function Number"
  ARptListings.Label2.Caption = "Title"
  ARptListings.Label3.Visible = False
  ARptListings.Label4.Caption = "Functions Listed"
  ARptListings.Total = Howmany
  ARptListings.txtDate = Now
  ARptListings.txtTown = GLUserName$
  ARptListings.Title.Caption = Header$
  ARptListings.GetName ReportFile$
  ARptListings.startrpt
  Exit Sub
   
End Sub
Private Sub PrintFunctionReport2()
  Dim MaxLines As Integer, FnctIdxFileNum As Integer, NumFIdxRecs As Integer
  Dim FnctFileNum As Integer, NumFncts As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Long, Howmany As Integer
  Dim ReportFile As String, ToPrint As String, Newrp As String
  Dim FF As String, Header As String
  Dim FnctIdx As GLFNCTIndexType
  Dim Fnct As GLFNCTRecType
 '  Stop
   'Define vars used for printing
  MaxLines = 55
  FF$ = Chr$(12)
  Header$ = "Master Function Listing"
'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
  OpenFnctIdx FnctIdxFileNum, NumFIdxRecs
  OpenFnctFile FnctFileNum, NumFncts
  PRNFile = FreeFile
  Newrp = "FnctLst"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  GoSub PrintFnctPageHeader
  For cnt = 1 To NumFIdxRecs
    Get FnctIdxFileNum, cnt, FnctIdx
    Get FnctFileNum, FnctIdx.RecNum, Fnct
    Howmany = Howmany + 1
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 2) = Fnct.FnctNum
    Mid$(ToPrint$, 18) = Fnct.Title
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #PRNFile, FF$
      GoSub PrintFnctPageHeader
    End If
  Next
  Print #PRNFile,
  Print #PRNFile, Howmany; "Functions listed."
  Print #PRNFile, FF$
  Close
  ViewPrint ReportFile$, "Function Listing Report"
  Kill ReportFile$
  Exit Sub
PrintFnctPageHeader:
  Print #PRNFile, Header$
  Print #PRNFile,
  Print #PRNFile, " Function Number     Title"
  Print #PRNFile, String$(80, "-")
  Linecnt = 4
  Return
End Sub

