VERSION 5.00
Begin VB.Form frmFundMaintMenu1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fund Maintenance"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmFundMaint1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFundAddEdit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Add/Change/Delete Funds"
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
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   3480
      Width           =   3612
   End
   Begin VB.CommandButton cmdFundPrintList 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Print Fund Listing"
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
      TabIndex        =   1
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdFundSort 
      Caption         =   "Fund Index &Utility"
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
      TabIndex        =   2
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitFundMaintMenu 
      Caption         =   "E&xit Fund Maintenance"
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
      TabIndex        =   3
      Top             =   6000
      Width           =   3612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FUND MAINTENANCE MENU"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   4692
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
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
      FillColor       =   &H8000000B&
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
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
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
      FillColor       =   &H8000000B&
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
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmFundMaintMenu1"
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
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdExitFundMaintMenu_Click()
  frmGLSetupMenu.Show
  Unload frmFundMaintMenu
End Sub
Private Sub cmdFundAddEdit_Click()
  frmFundEntryEdit.Show
  Unload frmFundMaintMenu
End Sub
Private Sub cmdFundPrintList_Click()
  If Exist("GLFund.DAT") Then
    frmReportOpt.Show 1
    If rptopt = 1 Then
      Call PrintFundListReport
    ElseIf rptopt = 2 Then
      Call PrintFundListReport2
    End If
  Else
    MsgBox "No Funds To Print", vbOKOnly, "No Funds"
  End If
End Sub
Private Sub cmdFundSort_Click()
  If Exist("GLFund.DAT") Then
    frmSortFund.Show 1
  Else
    MsgBox "No Funds To Sort", vbOKOnly, "No Funds"
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
        cmdExitFundMaintMenu_Click
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

Private Sub PrintFundListReport()
  Dim FundIdxFileNum As Integer, NumFIdxRecs As Integer
  Dim FundFileNum As Integer, NumFunds As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
  Dim ReportFile As String, ToPrint As String, Newrp As String
  Dim Header As String
  Dim FundIdx As GLFundIndexType
  Dim Fund As GLFundRecType
 '  Stop
   'Define vars used for printing
  Header$ = "Master Fund Listing"
'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  OpenFundFile FundFileNum, NumFunds
  PRNFile = FreeFile
  Newrp = "FndLst"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  For cnt = 1 To NumFIdxRecs
    Get FundIdxFileNum, cnt, FundIdx
    Get FundFileNum, FundIdx.RecNum, Fund
    Howmany = Howmany + 1
    ToPrint$ = ""
    ToPrint$ = Fund.FundNum + "~" + Fund.Title + "~"
    Print #PRNFile, ToPrint$
  Next
   Close
    
  Load frmLoadingRpt
  ARptListings.Label1.Caption = "Fund Number"
  ARptListings.Label2.Caption = "Title"
  ARptListings.Label3.Visible = False
  ARptListings.Label4.Caption = "Funds Listed"
  ARptListings.Total = Howmany
  ARptListings.txtDate = Now
  ARptListings.txtTown = GLUserName$
  ARptListings.Title.Caption = Header$
  ARptListings.GetName ReportFile$
  ARptListings.startrpt
  Exit Sub
   
End Sub
Private Sub PrintFundListReport2()
  Dim MaxLines As Integer, FundIdxFileNum As Integer, NumFIdxRecs As Integer
  Dim FundFileNum As Integer, NumFunds As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
  Dim ReportFile As String, ToPrint As String, Newrp As String
  Dim FF As String, Header As String
  Dim FundIdx As GLFundIndexType
  Dim Fund As GLFundRecType
 '  Stop
   'Define vars used for printing
  MaxLines = 55
  FF$ = Chr$(12)
  Header$ = "Master Fund Listing"
'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  OpenFundFile FundFileNum, NumFunds
  PRNFile = FreeFile
  Newrp = "FndLst"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  GoSub PrintFundPageHeader
  For cnt = 1 To NumFIdxRecs
    Get FundIdxFileNum, cnt, FundIdx
    Get FundFileNum, FundIdx.RecNum, Fund
    Howmany = Howmany + 1
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 2) = Fund.FundNum
    Mid$(ToPrint$, 18) = Fund.Title
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #PRNFile, FF$
      GoSub PrintFundPageHeader
    End If
  Next
  Print #PRNFile,
  Print #PRNFile, Howmany; "Funds listed."
  Print #PRNFile, FF$
  Close
  ViewPrint ReportFile$, "Fund Listing Report"
  Kill ReportFile$
  Exit Sub
PrintFundPageHeader:
  Print #PRNFile, Header$
  Print #PRNFile,
  Print #PRNFile, " Fund Number     Title"
  Print #PRNFile, String$(80, "-")
  Linecnt = 4
  Return
End Sub

