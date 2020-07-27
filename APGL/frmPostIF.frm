VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostIF 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Interface Transactions"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   12195
   Icon            =   "frmPostIF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   8112
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7536
      Width           =   1356
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
      Left            =   9792
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7536
      Width           =   1356
   End
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   2592
      Top             =   2568
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8625
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   450
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
            TextSave        =   "12:04 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
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
      Left            =   2568
      TabIndex        =   7
      Top             =   3624
      Width           =   7212
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Before You Post, Make Sure You Have Printed A Journal Report.  If You Haven't, Then Exit And Do So Now."
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
      Left            =   2616
      TabIndex        =   6
      Top             =   2928
      Width           =   7116
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
      Left            =   3888
      TabIndex        =   5
      Top             =   4056
      Visible         =   0   'False
      Width           =   4572
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Post Interface Transactions"
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
      Left            =   3984
      TabIndex        =   4
      Top             =   1176
      Width           =   4476
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3288
      Top             =   936
      Width           =   5772
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
      Height          =   420
      Left            =   4728
      TabIndex        =   3
      Top             =   2496
      Width           =   2772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3288
      Top             =   816
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00404040&
      Height          =   2172
      Left            =   2448
      Top             =   2376
      Width           =   7452
   End
End
Attribute VB_Name = "frmPostIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GJEdit As TrEditRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
'Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
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
'This was modal and didn't close menu, but changed later to user % stuff
'still didn't unload menu
  Unload frmPostIF
End Sub

Private Sub cmdOk_Click()
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOK.Enabled = False
  PostIFTrans
  Me.cmdExit.Enabled = True
  Me.cmdOK.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload frmPostIF
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
  Me.HelpContextID = hlpPostInterface
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PostIFTrans()
  Dim cnt As Integer, NumIfTrans As Integer, IFEditFileNum As Integer
  Dim TotDr As Double, TotCr As Double, Active As Integer
  Dim BadTrans As Integer, ReportFile As String, GLLogFileName As String
  Dim IFEdit As TrEditRecType
  '--verify that there are transactions and they are in balance.
  OpenIFEditFile IFEditFileNum, NumIfTrans

  '--summarize the file totals
  For cnt = 1 To NumIfTrans
    Get IFEditFileNum, cnt, IFEdit
    If Not GJEdit.Deleted Then
      Active = Active + 1
      TotDr# = Round#(TotDr# + IFEdit.DrAmt)
      TotCr# = Round#(TotCr# + IFEdit.CrAmt)
    End If
  Next
  Close
  '--if no active transactions tell user and get out
  If Active = 0 Then
    MsgBox "No Transactions to Post.", vbOKOnly, "No Trans"
    If ExistD("GLUBTran.dat") And GLUBKill = 1 Then
      Kill "GLUBTran.dat"
      GLUBKill = 0
    End If
    Exit Sub
    Unload frmPostIF
  End If

  If MsgBox("Are you sure you are ready to Post?", vbYesNo, "Continue?") = vbNo Then
  'Ask user if sure ready to pos
   'If Ok = 1 Then Exit Sub       '1=No
    Exit Sub
  End If
  TotDr# = 0    'init totals to zero
  TotCr# = 0
  Active = 0    'Counter for Active Transactions

  If TotDr# <> TotCr# Then      'Transactions out of balance
    If MsgBox("Transactions are out of balance, Ok to continue with posting or Cancel?", vbOKCancel, "Continue?") = vbCancel Then     'ask user if ok to post
      Exit Sub     'No = button 1
    End If
  End If


  Active = 0    'Reset Active counter for posting
  OpenIFEditFile IFEditFileNum, NumIfTrans

  Dim Tr2Post As GLTransRecType
  Open "GJ2POST.DAT" For Random As #2 Len = Len(Tr2Post)

  For cnt = 1 To NumIfTrans     'Assign edit file to trans format
    Get IFEditFileNum, cnt, GJEdit
    If Not GJEdit.Deleted Then
      Active = Active + 1
      Tr2Post.AcctRec = GJEdit.AcctRec
      Tr2Post.AcctNum = GJEdit.AcctNum
      Tr2Post.TRDATE = GJEdit.TRDATE
      Tr2Post.Desc = GJEdit.Desc
      Tr2Post.Ref = GJEdit.Ref
      Tr2Post.DrAmt = GJEdit.DrAmt
      Tr2Post.CrAmt = GJEdit.CrAmt
      Tr2Post.Src = QPTrim$(GJEdit.Src) + Format$(Now, "mmddyy")
      Put #2, Active, Tr2Post
    End If
  Next

  Close
   Post2GL "GJ2POST.DAT", BadTrans%, frmPostIF, False
   If BadTrans <> 0 Then
      Call MainLog("Errors During IF Post.Stopped.")
      MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
      ReportFile$ = "TempLog.PRN"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Error Log"
      End If
      frmCitiCancel.Show
      Unload frmPostIF
      Unload frmGetDistMenu
      Exit Sub
    End If
    Post2GL "GJ2POST.DAT", BadTrans%, frmPostIF, True
    KillFileD "GLTRXED.DAT"            'kill the temp files
    KillFileD "GJ2POST.DAT"
    If ExistD("GLUBTran.dat") And GLUBKill = 1 Then
      Kill "GLUBTran.dat"
      GLUBKill = 0
    End If
    If BadTrans <> 0 Then                  'posting problem
      Call MainLog("Error Posting IF.")
      MsgBox "Error, One or more transactions were not posted. Make sure the printer is ready and Press a Key to View Log.", vbOKOnly, "Posting Error"
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
  Call MainLog("Posting IF - Complete.")
  MsgBox "Posting Procedure Completed", vbOKOnly, "GJ Posting"

End Sub

