VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostGJ 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post General Journal"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   90
   ClientWidth     =   12225
   Icon            =   "frmPostGJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   2520
      Top             =   2784
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8280
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "12:04 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "6/28/2008"
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
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Ok"
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
      Left            =   8400
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7488
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
      Left            =   10080
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7488
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   324
      Left            =   4908
      TabIndex        =   7
      Top             =   2832
      Width           =   2412
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Ok to Begin Posting or Exit to Escape Posting Procedure. "
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
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   3960
      Width           =   7212
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Before You Post, Make Sure You Have Printed A General Journal Edit Report.  If You Haven't, Then Exit And Do So Now."
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
      Height          =   732
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   3240
      Width           =   6732
   End
   Begin VB.Label lblPosting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Posting to General Ledger Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   372
      Left            =   3840
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   4572
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2172
      Left            =   2400
      Top             =   2640
      Width           =   7452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Post General Journal Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4080
      TabIndex        =   3
      Top             =   1440
      Width           =   4092
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3240
      Top             =   1200
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3240
      Top             =   1080
      Width           =   5772
   End
End
Attribute VB_Name = "frmPostGJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLFundIdx As GLFundIndexType
Dim GJEdit As TrEditRecType
Dim GLTrans As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub
Private Sub Timer1_Timer()
 ' Label2.Visible = Not Label2.Visible
  '&H0080FFFF&
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Label2.ForeColor = &H80FFFF
    Shape3.BackColor = &HC0&
  Else
    Label2.ForeColor = &HFFFF&
    Shape3.BackColor = &H80&
  End If
  
End Sub

Private Sub cmdExit_Click()
'This was modal and didn't close gjmenu
  KillFileD "GJEdit.opn"
  Unload frmPostGJ
End Sub

Private Sub cmdOk_Click()
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOK.Enabled = False
  PostGJTrans
  Me.cmdExit.Enabled = True
  Me.cmdOK.Enabled = True
  EnableCloseButton Me.hwnd, True
  KillFileD "GJEdit.opn"
  Unload frmPostGJ
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"  'Arrow Down
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"   'arrow up
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"     'Esc key
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"     'alt O or f10
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
  Me.HelpContextID = hlpGJPost
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PostGJTrans()
  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer, TotDr As Double, TotCr As Double, Active As Integer, BadTrans As Integer
  Dim PRNFile As Integer
  Dim GLLogFileName As String, GLLogFile As Integer, Log As String
  Dim ReportFile As String, ToPrint As String
  Dim FundCode As String, OutofBal As Integer
  Dim CommaFmt As String, FundNum As String, strMsg As String
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer, NumFunds As Integer
'the get list of funds on gj main form
  GetFundList FundList(), NumFunds
  ReDim FundDr(1 To NumFunds) As Double
  ReDim FundCr(1 To NumFunds) As Double
  OpenGJEditFile GJEditFileNum, NumEdTrans
  TotDr = 0: TotCr = 0
  Active = 0
  OutofBal = 0
  Found = False
  For cnt = 1 To NumEdTrans
    Get GJEditFileNum, cnt, GJEdit
    If Not GJEdit.Deleted Then
      Active = Active + 1
      TotDr = Round#(TotDr + GJEdit.DrAmt)
      TotCr = Round#(TotCr + GJEdit.CrAmt)
      For Fund = 1 To NumFunds 'Get summary totals by fund
        FundNum$ = Left$(GJEdit.AcctNum, GLFundLen)
        If FundNum$ = FundList$(Fund) Then
          Found = True
          FundDr#(Fund) = Round#(FundDr#(Fund) + GJEdit.DrAmt)
          FundCr#(Fund) = Round#(FundCr#(Fund) + GJEdit.CrAmt)
          Exit For
        End If
      Next
    End If
  Next
  Close GJEditFileNum
'Give options to cancel posting and let know if out of balance
  If Active = 0 Then
    MsgBox "No Transactions To Post", vbOKOnly, "Post Canceled"
    Exit Sub
  End If
  If MsgBox("Are You Sure You Wish to Post. Once Posted You Will Not Be Able to Print a Journal.", vbOKCancel, "GJ Posting") = vbCancel Then
    Exit Sub
  Else
    If TotDr <> TotCr Then
      strMsg = "Transactions are OUT OF BALANCE, Do You Wish To Continue or Cancel?"
      OutofBal = 1
    Else
      For Fund = 1 To NumFunds 'Compare debits/credits by fund
        If FundDr#(Fund) <> 0 Or FundCr#(Fund) <> 0 Then
          If FundDr#(Fund) <> FundCr#(Fund) Then
          'if not balanced within fund Set message string
            strMsg = "Fund Totals Do Not Balance. Do You Wish To Continue or Cancel?"
            OutofBal = 2
            Exit For
          End If
        End If
      Next
    End If
    If OutofBal <> 0 Then
      If OutofBal = 1 Then
        If MsgBox(strMsg, vbOKCancel, "GJ Posting") = vbCancel Then
          Exit Sub
        End If
      Else
        If MsgBox(strMsg, vbOKCancel, "GJ Posting") = vbCancel Then
          Exit Sub
        End If
      End If
    End If
    lblPosting.Visible = True
    Active = 0                             'Reset Active counter for posting
    OpenGJEditFile GJEditFileNum, NumEdTrans
    Dim Tr2Post As GLTransRecType
    Open "GJ2POST.DAT" For Random As #2 Len = Len(Tr2Post)
    For cnt = 1 To NumEdTrans              'Assign edit file to trans format
      Get GJEditFileNum, cnt, GJEdit
      If Not GJEdit.Deleted Then
      'write to temp file
        Active = Active + 1
        Tr2Post.AcctRec = GJEdit.AcctRec
        Tr2Post.AcctNum = GJEdit.AcctNum
        Tr2Post.TRDATE = GJEdit.TRDATE
        Tr2Post.Desc = GJEdit.Desc
        Tr2Post.LDesc = GJEdit.LDesc
        Tr2Post.Ref = GJEdit.Ref
        Tr2Post.DrAmt = GJEdit.DrAmt
        Tr2Post.CrAmt = GJEdit.CrAmt
        Tr2Post.Src = "GJ" + Format$(Now, "mmddyy")
        Put #2, Active, Tr2Post
      End If
    Next
    Close
    
    Call Post2GL("GJ2POST.DAT", BadTrans, frmPostGJ, False) 'common post & link sub
    If BadTrans <> 0 Then
      KillFile "GJ2POST.DAT"
      MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
      Call MainLog("Errors Posting GJ.")
      ReportFile$ = "TempLog.PRN"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Error Log"
      End If
      frmCitiCancel.Show
      Unload frmPostGJ
      Unload frmGenJournalMenu
      Exit Sub
    End If
    Call Post2GL("GJ2POST.DAT", BadTrans, frmPostGJ, True) 'common post & link sub
    KillFile "GJEdit.DAT"                    'kill the temp files
    KillFile "GJ2POST.DAT"
    If BadTrans <> 0 Then 'posting problem
      MsgBox "Error, One or more transactions were not posted. Make sure the printer is ready and Press a Key to View Log.", vbOKOnly, "Posting Error"
      Call MainLog("Error Not All GJ Trans Posted.")
      GLLogFileName = "GLlog.dat"
      ReportFile$ = "GLlog.dat"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Posting Log"
      End If
    End If
  Call MainLog("GJ Post Complete.")
  MsgBox "Posting Procedure Completed", vbOKOnly, "GJ Posting"
  End If
  
End Sub

