VERSION 5.00
Begin VB.Form frmBankMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Maintenance"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmBankMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitBankMaintMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Bank Maintenance Menu"
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
      Left            =   4302
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintBankList 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Print Bank Listing"
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
      HelpContextID   =   55
      Left            =   4302
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdEditBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Edit a Bank Record"
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
      Left            =   4302
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdAddBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Add a New Bank Record"
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
      Left            =   4302
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3612
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BANK MAINTENANCE MENU"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   6852
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
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
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
End
Attribute VB_Name = "frmBankMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdAddBank_Click()
  frmBankCodeEntry.Show
  Unload frmBankMaintMenu
End Sub

Private Sub cmdEditBank_Click()
  frmBankCodeEdit.Show
  Unload frmBankMaintMenu
End Sub

Private Sub cmdExitBankMaintMenu_Click()
  frmGLSetupMenu.Show
  Unload frmBankMaintMenu
End Sub

Private Sub cmdPrintBankList_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    Call PrintBankListReport
  ElseIf rptopt = 2 Then
    PrintBankListReport2
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitBankMaintMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    End If
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpBankMaintenance
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub PrintBankListReport()
  Dim MaxLines As Integer, BankFileNum As Integer, NumBankRecs As Integer
  Dim Linecnt As Integer, Newrp As String
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
  Dim ReportFile As String, ToPrint As String
  Dim FF As String, Header As String
  Dim GLBank As GLBankRecType
  
 '  Stop
   'Define vars used for printing
   MaxLines = 55
   FF$ = Chr$(12)
   Header$ = "Master Bank Code Listing"

'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
   Newrp = "BnkLst"
   GetRPTName Newrp
   OpenBankFile BankFileNum, NumBankRecs
   PRNFile = FreeFile
   ReportFile$ = Newrp
   Open ReportFile$ For Output As #PRNFile

   'GoSub PrintBankPageHeader

   For cnt = 1 To NumBankRecs
      Get BankFileNum, cnt, GLBank
      If GLBank.Deleted = 0 Then
      
        Howmany = Howmany + 1
  
        ToPrint$ = Space$(80)
        ToPrint$ = GLBank.BankNum & "~" & QPTrim(GLBank.BankName) & "~" & QPTrim(GLBank.GLAcct)
        Print #PRNFile, ToPrint$
'        Linecnt = Linecnt + 1
'        If Linecnt > MaxLines Then
'          Print #PRNFile, FF$
'          'GoSub PrintBankPageHeader
'        End If
      End If
   Next
   'Print #PRNFile,
   'Print #PRNFile, HowMany; "Bank Codes listed."
   'Print #PRNFile, FF$
    
   Close
   Load frmLoadingRpt
   ARptBankCodeList.GetName ReportFile$
   ARptBankCodeList.txtDate = Now
   ARptBankCodeList.txtTown = GLUserName$
   ARptBankCodeList.Howmany = Howmany
   
   ARptBankCodeList.startrpt
   'ViewPrint ReportFile$, "Bank Code Listing Report"
   'Kill ReportFile$
Exit Sub

PrintBankPageHeader:
'  Print #PRNFile, Header$
'  Print #PRNFile,
'  Print #PRNFile, "Bank Code         Bank Name                  GL Account Number"
'  Print #PRNFile, String$(80, "-")
'  Linecnt = 4
Return

End Sub

Private Sub PrintBankListReport2()
  Dim MaxLines As Integer, BankFileNum As Integer, NumBankRecs As Integer
  Dim Linecnt As Integer, Newrp As String
  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
  Dim ReportFile As String, ToPrint As String
  Dim FF As String, Header As String
  Dim GLBank As GLBankRecType
  
 '  Stop
   'Define vars used for printing
   MaxLines = 55
   FF$ = Chr$(12)
   Header$ = "Master Bank Code Listing"

'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
   Newrp = "BnkLst"
   GetRPTName Newrp
   OpenBankFile BankFileNum, NumBankRecs
   PRNFile = FreeFile
   ReportFile$ = Newrp
   Open ReportFile$ For Output As #PRNFile

   GoSub PrintBankPageHeader

   For cnt = 1 To NumBankRecs
      Get BankFileNum, cnt, GLBank
      If GLBank.Deleted = 0 Then
      
        Howmany = Howmany + 1
  
        ToPrint$ = Space$(80)
        Mid$(ToPrint$, 4) = GLBank.BankNum
        Mid$(ToPrint$, 14) = GLBank.BankName
        Mid$(ToPrint$, 50) = GLBank.GLAcct
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        If Linecnt > MaxLines Then
          Print #PRNFile, FF$
          GoSub PrintBankPageHeader
        End If
      End If
   Next
   Print #PRNFile,
   Print #PRNFile, Howmany; "Bank Codes listed."
   Print #PRNFile, FF$

   Close
   
   ViewPrint ReportFile$, "Bank Code Listing Report"
   Kill ReportFile$
Exit Sub

PrintBankPageHeader:
  Print #PRNFile, Header$
  Print #PRNFile,
  Print #PRNFile, "Bank Code         Bank Name                  GL Account Number"
  Print #PRNFile, String$(80, "-")
  Linecnt = 4
Return

End Sub

