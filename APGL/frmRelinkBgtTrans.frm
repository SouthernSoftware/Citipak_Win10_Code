VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelinkBgtTrans 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Link Budget Transactions"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   12192
   Icon            =   "frmRelinkBgtTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
      Left            =   3990
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
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
      Left            =   6870
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
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
      Caption         =   "Press F10 to RE-LINK Transactions or Escape to Exit."
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
      Left            =   3282
      TabIndex        =   6
      Top             =   4896
      Width           =   5628
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Link Budget Transactions"
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
      Left            =   4008
      TabIndex        =   5
      Top             =   1704
      Width           =   4188
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1464
      Width           =   5772
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This utility re-links Budget transaction records."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   348
      Left            =   3498
      TabIndex        =   4
      Top             =   3216
      Width           =   5196
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAKE SURE EVERYONE IS OUT OF CITIPAK BEFORE RUNNING THIS OPERATION!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   636
      Left            =   3522
      TabIndex        =   3
      Top             =   3936
      Width           =   5148
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3210
      Top             =   1344
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3276
      Left            =   2682
      Top             =   2928
      Width           =   6828
   End
End
Attribute VB_Name = "frmRelinkBgtTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim GLAcctidx As GLAcctIndexType
Dim GLTrans   As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim StartFund As String, EndFund As String, FYStartDate As Integer
Dim ActiveYear As Integer
Dim acctmsk As String, detmsk As String

Private Sub cmdExit_Click()
  frmGLUtilMenu.Show
  Unload frmRelinkBgtTrans
End Sub
Private Sub cmdGo_Click()
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdGo.Enabled = False
  Call MainLog("RelinkBgtTrans Started - Menu Option.")
  RelinkBgtTrans frmRelinkBgtTrans
  Call MainLog("RelinkBgtTrans Complet - Menu Option.")
  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
  frmGLUtilMenu.Show
  Unload frmRelinkBgtTrans
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
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

'Public Sub RelinkBgtTrans()
'  Dim BTrans As GLTransRecType
'  Dim Acct As GLAcctRecType
'  Dim TransRecLen As Integer, BgtTransFile As Integer, NumTrans As Long
'  Dim GLAcctFile As Integer, NumAccts As Integer, TCnt As Long
'  Dim LogFile As Integer, LogFileName As String, cnt As Integer
'  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
'  Dim CntA As Integer, AcctNum As String, LookFor As String
'  Dim Prev As Long, AcctRecNum As Integer, BadTran As Integer
'  Dim ToPrint As String
'  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
'  ReDim IdxAry(1 To NumAIdxRecs) As GLAcctIndexType
'  For CntA = 1 To NumAIdxRecs
'    Get AcctIdxFileNum, CntA, IdxAry(CntA)
'  Next
'    Close AcctIdxFileNum
'
'
'   TransRecLen = Len(BTrans)
'   BgtTransFile = FreeFile
'   Open "BGTTRANS.DAT" For Random As BgtTransFile Len = TransRecLen
'   NumTrans& = LOF(BgtTransFile) \ TransRecLen
'
'   OpenAcctFile GLAcctFile, NumAccts
'
'   Lock BgtTransFile
'   Lock GLAcctFile
'
'   LogFile = FreeFile
'   LogFileName$ = "GLLINK.LOG"
'   Open LogFileName$ For Append As #LogFile
'   Print #LogFile,
'   Print #LogFile, "Budget Database relink started @ " + Date$ + " @ "; Time$
'   FrmShowPctComp.Label1 = "Initializing Account Transactions."
'   FrmShowPctComp.Show , Me
'   DoEvents
'
'   '-Set the pointers in the transaction file to zero
'   For TCnt& = 1 To NumTrans&
'      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
'      Get BgtTransFile, TCnt&, BTrans
'      BTrans.NextTran = 0
'      Put BgtTransFile, TCnt&, BTrans
'   Next
'
'   FrmShowPctComp.Label1 = "Initializing Budget Transactions."
'   FrmShowPctComp.Show , Me
'   DoEvents
'
''   -Set the budget pointers in the account file to zero
'   For cnt = 1 To NumAccts
'      FrmShowPctComp.ShowPctComp cnt, NumAccts
'      Get GLAcctFile, cnt, Acct
'      Acct.FrstBTran = 0
'      Acct.Bgt = 0
'      Acct.LastBTran = 0
'      Put GLAcctFile, cnt, Acct
'   Next
'   '-Start the relink process
'   FrmShowPctComp.Label1 = "Relink Budget Transaction Database."
'   FrmShowPctComp.Show , Me
'   DoEvents
'
'   For TCnt& = 1& To NumTrans&
'
'      '-Something to look at while this is going on
'      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
'      Get BgtTransFile, TCnt&, BTrans
'
'      AcctNum$ = Trim$(BTrans.AcctNum)
''-Find the record number of the account
'      For CntA = 1 To NumAIdxRecs
'        'Here you put Jump Around Code To Speed UP MOre!!!
'        LookFor$ = Trim$(IdxAry(CntA).AcctNum)
'        If AcctNum$ = LookFor$ Then
'            'AcctRecNum = CntA
'           AcctRecNum = IdxAry(CntA).RecNum
'
'          Exit For
'        End If
'      Next
'      'AcctRecNum = AcctFind(Trim$(BTrans.AcctNum))
'      '-If we find the account
'      If AcctRecNum > 0 Then
'         Get GLAcctFile, AcctRecNum, Acct
'
'         '-Check out the pointer to the first transaction
'         Select Case Acct.FrstBTran
'
'           '-If this is the first transaction for this account
'           Case 0
'               '-Set first and last pointers to this transaction
'               Acct.FrstBTran = TCnt&
'               Acct.LastBTran = TCnt&
'               Put GLAcctFile, AcctRecNum, Acct
'
'            Case Is > 0  '-There are already transactions for this account
'               '-Remember the pointer to the last transaction.
'               Prev& = Acct.LastBTran
'               '-Set the last trans pointer to this transaction
'               Acct.LastBTran = TCnt&
'               Put GLAcctFile, AcctRecNum, Acct
'
'               '-Get the last previous transaction and set its
'               '-next tran pointer to this transaction
'               Get BgtTransFile, Prev&, BTrans
'               BTrans.NextTran = TCnt&
'               Put BgtTransFile, Prev&, BTrans
'            Case Else
'         End Select
'
'         '--update the Acct's Budget Balance
'         Select Case Acct.Typ
'            Case "A", "E"
'               Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.DrAmt) - Round#(BTrans.CrAmt)
'            Case "L", "R"
'               Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.CrAmt) - Round#(BTrans.DrAmt)
'         End Select
'         Put GLAcctFile, AcctRecNum, Acct
'
'      Else  '-could not find the account
'         BadTran = BadTran + 1
'
'         'MsgBox "Orphaned transactions: " & Using("#####", BadTran), vbOKOnly, "Errors Found"
'         GoSub LogBgtTrans '-Keep a list of orphaned transactions.
'
'      End If
'   Next
'
'   '-we're done
'   Unlock BgtTransFile
'   Unlock GLAcctFile
'
'   '-Tell user we're done.
'   If BadTran > 0 Then
'      '-Errors in trans file
'      Print #LogFile, "Relink encountered ophans. Completed @ " + Date$ + " @" + Time$
'   Else
'      '-No errors in trans file
'      Print #LogFile, "Relink of Budget Database successful. " + Date$ + " @ " + Time$
'      MsgBox "Re-link successful.", vbOKOnly, "Procedure Complete"
'
'   End If
'
'   Close
'
'Exit Sub
'
'LogBgtTrans:
'   ToPrint$ = Space$(132)
'   LSet ToPrint$ = BTrans.AcctNum
'   Mid$(ToPrint$, 18) = Format(DateAdd("d", BTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
'   Mid$(ToPrint$, 30) = Left$(BTrans.Desc, 15)
'   Mid$(ToPrint$, 50) = BTrans.Ref
'   Mid$(ToPrint$, 60) = Using("#,###,###.##", Str$(BTrans.DrAmt))
'   Mid$(ToPrint$, 70) = Using("#,###,###.##", Str$(BTrans.CrAmt))
'   Mid$(ToPrint$, 80) = "Record:" + Str$(TCnt&)
'   Print #LogFile, ToPrint$
'Return
'
'
'End Sub

