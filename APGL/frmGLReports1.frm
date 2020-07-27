VERSION 5.00
Begin VB.Form frmGLReportsMenu1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger Reports"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmGLReports1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExportFiles 
      Caption         =   "&Export Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   9
      Top             =   7080
      Width           =   3612
   End
   Begin VB.CommandButton cmdQueryGLTrans 
      Caption         =   "&Query G/L Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   8
      Top             =   6600
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitGLReportsMenu 
      Caption         =   "E&xit General Ledger Reports Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   10
      Top             =   7560
      Width           =   3612
   End
   Begin VB.CommandButton cmdBudvsAct 
      Caption         =   "Bud&get vs Actual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   6
      Top             =   5640
      Width           =   3612
   End
   Begin VB.CommandButton cmdDeptBudvsAct 
      Caption         =   "&Department Budget vs Actual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   7
      Top             =   6120
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctBalSummary 
      Caption         =   "&Account Balance Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   2
      Top             =   3720
      Width           =   3612
   End
   Begin VB.CommandButton cmdBalSheet 
      Caption         =   "Balance &Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   5
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdBudgHistory 
      Caption         =   "&Budget History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctHistory 
      Caption         =   "Account &History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   3
      Top             =   4200
      Width           =   3612
   End
   Begin VB.CommandButton cmdCashBalance 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cash Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   1
      Top             =   3240
      Width           =   3612
   End
   Begin VB.CommandButton cmdTrialBalance 
      BackColor       =   &H008F8265&
      Caption         =   "&Trial Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2760
      Width           =   3612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
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
      Height          =   396
      Left            =   5094
      TabIndex        =   12
      Top             =   576
      Visible         =   0   'False
      Width           =   2028
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9840
      X2              =   9840
      Y1              =   2304
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
      Index           =   0
      X1              =   2520
      X2              =   3216
      Y1              =   8304
      Y2              =   8304
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
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL LEDGER REPORTS"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   1440
      Width           =   4692
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
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2496
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmGLReportsMenu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
'the PrYr is global that specifies this report menu was called from
'Prior Year selection screen and has changed directories to prior year.

Private Sub cmdAcctBalSummary_Click()
  frmPrnAcctBal.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdAcctHistory_Click()
  frmPrnAcctHist.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdBalSheet_Click()
  frmPrnBalSheet.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdBudgHistory_Click()
  frmPrnBudHist.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdBudvsAct_Click()
  frmPrnBudAct.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdCashBalance_Click()
  frmPrnCashBal.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdDeptBudvsAct_Click()
  frmPrnDeptBudAct.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdExportFiles_Click()
  If Exist("GLAcct.dat") Then
    ExportGL
  Else
    MsgBox "NO Account Information to Export.", vbOKOnly, "No Accounts"
  End If
End Sub

Private Sub cmdQueryGLTrans_Click()
  frmPrnQuery.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub cmdTrialBalance_Click()
  frmPrnTrialBal.Show
  If PrYr <> 1 Then
    Unload frmGLReportsMenu
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
    
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdExitGLReportsMenu_Click()
 Dim outfile As String
  If PrYr <> 0 Then
  'this is needed if folder mapped as drive
    If Right$(StartPath, 1) = ":" Then
      outfile = StartPath & "\"
      ChDir outfile
    Else
      ChDir StartPath
    End If
  End If
  frmGLMainMenu.Show
  Unload frmGLReportsMenu
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim outfile As String
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitGLReportsMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        If PrYr <> 0 Then
        'must do this if folder is mapped as drive
          If Right$(StartPath, 1) = ":" Then
            outfile = StartPath & "\"
            ChDir outfile
          Else
            ChDir StartPath
          End If
        End If
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitGLReportsMenu_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub ExportGL()
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, NumAccts As Integer, AcctPrnFile As Integer
  Dim cnt As Integer, TransFileNum As Integer, NumTrans As Long
  Dim TransPrnFile As Integer, cntl As Long
  Dim Acct As GLAcctRecType
  Dim Trans As GLTransRecType
  Dim AcctIdx As GLAcctIndexType
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumAccts
  FrmShowPctComp.Label1 = "Creating Export Files"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmGLReportsMenu
  AcctPrnFile = FreeFile
  Open "GLA.ASC" For Output As #AcctPrnFile
  Write #AcctPrnFile, "AcctNum", "Title", "Type", "PY_Act", "Bgt", "Balance"
  
  For cnt = 1 To NumGLAccts   'NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    Get AcctIdxFileNum, cnt, AcctIdx
    Get AcctFileNum, AcctIdx.RecNum, Acct
    'Done! = (cnt / NumAccts) * 100
    'LOCATE 12, 1, 0
     'Print Using; "Processing Chart of Accounts. ###% Complete."; Done!
    Get AcctFileNum, cnt, Acct
    If Not Acct.Deleted Then
      Write #AcctPrnFile, QPTrim$(Acct.Num), Acct.Title, Acct.Typ, Acct.PYAct, Acct.Bgt, Acct.Bal
    End If
  Next
  Close
  FrmShowPctComp.Label1 = "Creating Export Files"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  OpenTransFile TransFileNum, NumTrans&
   TransPrnFile = FreeFile
   Open "GLT.ASC" For Output As #TransPrnFile
   Write #TransPrnFile, "Acct Number", "Date", "Description", "Reference", "Debit", "Credit"
   For cntl& = 1 To NumTrans&
     FrmShowPctComp.ShowPctComp cntl&, NumTrans&
      'Done! = (cnt& / NumTrans&) * 100
      'Print Using; "Processing Transaction File. ###% Complete.  "; Done!
      Get TransFileNum, cntl&, Trans
      If Len(QPTrim$(Trans.Ref)) = 0 Then Trans.Ref = "~"
      If Len(QPTrim$(Trans.Desc)) = 0 Then Trans.Desc = "~"
      Write #TransPrnFile, QPTrim$(Trans.AcctNum), Format(DateAdd("d", (Trans.TRDATE), "12-31-1979"), "mm/dd/yy"), Trans.Desc, Trans.Ref, Trans.DrAmt, Trans.CrAmt
   Next
   Close

   MsgBox "Export Files have been created in the Citipak Directory.", vbOKOnly, "Export Complete"

   'SHELL "list Trans.prn"
   ActivateControls frmGLReportsMenu

End Sub

