VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelinkGLTrans 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Link GL Transactions"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   12192
   Icon            =   "frmRelinkGLTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Top             =   5112
      Width           =   1332
   End
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
      Top             =   5112
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
      Left            =   3552
      TabIndex        =   6
      Top             =   3648
      Width           =   5148
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This utility re-links GL transaction records."
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
      Left            =   3618
      TabIndex        =   5
      Top             =   2928
      Width           =   4956
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1176
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Link GL Transactions"
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
      Left            =   4002
      TabIndex        =   4
      Top             =   1416
      Width           =   4188
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
      Left            =   3258
      TabIndex        =   3
      Top             =   4608
      Width           =   5676
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3210
      Top             =   1056
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3276
      Left            =   2682
      Top             =   2640
      Width           =   6828
   End
End
Attribute VB_Name = "frmRelinkGLTrans"
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
  Unload frmRelinkGLTrans
End Sub
Private Sub cmdGo_Click()
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdGo.Enabled = False
  Call MainLog("RelinkGLTrans Started - Menu Option.")
  ReLinkTrans frmRelinkGLTrans
  Call MainLog("RelinkGLTrans Complete - Menu Option.")
  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
  frmGLUtilMenu.Show
  Unload frmRelinkGLTrans
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

'''Private Sub ReLinkTrans()
'''  Dim First As Long, Last As Long, RecNo As Long, AcctRecNum As Integer
'''  Dim DrAmt As Double, CrAmt As Double, RCnt As Long, BadTran As Integer
'''  Dim Diff As Double, Bal As Double, ToPrint1 As String, ToPrint3 As String
'''  Dim CommaFmt As String, TotalFmt As String, ToPrint2 As String
'''  Dim GLTransFile As Integer, NumTrans As Long, cnt As Integer
'''  Dim GLAcctFile As Integer, NumAccts As Integer, Prev As Long
'''  Dim LogFile As Integer, LogFileName As String, TCnt As Long
'''  Dim BadDebits As Double, BadCredits As Double, ToPrint As String
'''  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
'''  Dim CntA As Integer, AcctNum As String, LookFor As String
'''  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
'''  ReDim IdxAry(1 To NumAIdxRecs) As GLAcctIndexType
'''  For CntA = 1 To NumAIdxRecs
'''    Get AcctIdxFileNum, CntA, IdxAry(CntA)
'''  Next
'''    Close AcctIdxFileNum
'''   OpenTransFile GLTransFile, NumTrans&
'''   OpenAcctFile GLAcctFile, NumAccts
'''
'''   Lock GLTransFile
'''   Lock GLAcctFile
'''
'''   LogFile = FreeFile
'''   LogFileName$ = "GLLINK.LOG"
'''   Open LogFileName$ For Append As #LogFile
'''   FrmShowPctComp.Label1 = "Initializing transaction file."
'''   FrmShowPctComp.Show , Me
'''   DoEvents
'''   EnableCloseButton Me.hwnd, False
'''   Me.cmdExit.Enabled = False
'''   Me.cmdGo.Enabled = False
'''
'''
'''   '-Set the pointers in the transaction file to zero
'''   For TCnt& = 1 To NumTrans&
'''      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
'''
'''      Get GLTransFile, TCnt&, GLTrans
'''      GLTrans.NextTran = 0
'''      Put GLTransFile, TCnt&, GLTrans
'''  Next          'Process next transaction
'''
'''   FrmShowPctComp.Label1 = "Initializing account file."
'''   FrmShowPctComp.Show , Me
'''   DoEvents
'''
'''   '-Set the pointers in the account file to zero
'''   For cnt = 1 To NumAccts
'''      FrmShowPctComp.ShowPctComp cnt, NumAccts
'''      Get GLAcctFile, cnt, GLAcct
'''      GLAcct.FrstTran = 0
'''      GLAcct.Bal = 0
'''      GLAcct.LastTran = 0
'''      Put GLAcctFile, cnt, GLAcct
'''  Next          'Process next transaction
'''   FrmShowPctComp.Label1 = "Relinking."
'''   FrmShowPctComp.Show , Me
'''   DoEvents
'''
'''    '-Start the relink process
'''   For TCnt& = 1& To NumTrans&
'''      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
'''
'''      '-Something to look at while this is going on
'''
'''      Get GLTransFile, TCnt&, GLTrans
'''      AcctNum$ = Trim$(GLTrans.AcctNum)
''''****Make The Find Faster!!!!
'''
''''-Find the record number of the account
'''      For CntA = 1 To NumAIdxRecs
'''        'Here you put Jump Around Code To Speed UP MOre!!!
'''        LookFor$ = Trim$(IdxAry(CntA).AcctNum)
'''        If AcctNum$ = LookFor$ Then
'''          AcctRecNum = CntA
'''          Exit For
'''        End If
'''      Next
'''
'''      '-If we find the account
'''      If AcctRecNum > 0 Then
'''         Get GLAcctFile, AcctRecNum, GLAcct
'''
'''         '-Check out the pointer to the first transaction
'''         Select Case GLAcct.FrstTran
'''
'''           '-If this is the first transaction for this account
'''           Case 0
'''               '-Set first and last pointers to this transaction
'''               GLAcct.FrstTran = TCnt&
'''               GLAcct.LastTran = TCnt&
'''               Put GLAcctFile, AcctRecNum, GLAcct
'''
'''            '-If there are already transactions for this account
'''            Case Is > 0
'''               '-Remember the pointer to the last transaction.
'''               Prev& = GLAcct.LastTran
'''
'''               '-Set the last trans pointer to this transaction
'''               GLAcct.LastTran = TCnt&
'''               Put GLAcctFile, AcctRecNum, GLAcct
'''
'''               '-Get the last previous transaction and set its
'''               '-next tran pointer to this transaction
'''               Get GLTransFile, Prev&, GLTrans
'''               GLTrans.NextTran = TCnt&
'''               Put GLTransFile, Prev&, GLTrans
'''
'''               'update running balance here
'''
'''         End Select
'''
'''      Else  '-could not find the account
'''         BadTran = BadTran + 1
'''
'''         'Trans.Marked = -1
'''         'PUT GLTransFile, TCnt&, Trans
'''         'Trans.Marked = 0
'''
'''         '-Keep a list of orphaned transactions.
'''         GoSub Logit
'''
'''      End If
'''
'''   Next
'''  Me.cmdExit.Enabled = True
'''  Me.cmdGo.Enabled = True
'''  EnableCloseButton Me.hwnd, True
'''
'''   '-we're done here
'''   Unlock GLTransFile
'''   Unlock GLAcctFile
'''
'''   If BadTran > 0 Then
'''      '-Errors in trans file
'''      Print #LogFile,
'''      Print #LogFile, "Orphan Transaction Totals:";
'''      Print #LogFile, Tab(58); Using("#,###,###.##", Str$(BadDebits#))
'''      Print #LogFile, Tab(70); Using("#,###,###.##", Str$(BadCredits#))
'''      Print #LogFile, "Relink completed @ " + Date$ + " @ " + Time$
'''      Print #LogFile, "Orphan transactions encountered! Call Customer Support."
'''   Else
'''      '-No errors in trans file
'''      MsgBox "Relink of Accounting Databases successful. " + Date$ + "@" + Time$, vbOKOnly, "Relink Successful"
'''   End If
'''
'''   Close
'''
'''   '-Tell user we're done.
'''   If BadTran > 0 Then
'''      '-Errors in trans file
'''      If MsgBox("Errors Encountered, Select Ok to view log or Cancel to Exit.", vbOKCancel, "Error Log") = vbOK Then
'''        ViewPrint LogFileName$, "Error Log"
'''      End If
'''   End If
'''
'''Exit Sub
'''
'''Logit:
'''   ToPrint$ = Space$(132)
'''   LSet ToPrint$ = GLTrans.AcctNum
'''   Mid$(ToPrint$, 18) = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
'''   Mid$(ToPrint$, 30) = Left$(GLTrans.Desc, 15)
'''   Mid$(ToPrint$, 46) = GLTrans.Ref
'''   Mid$(ToPrint$, 58) = Using("#'###'###.##", Str$(GLTrans.DrAmt))
'''   Mid$(ToPrint$, 70) = Using("#,###,###.##", Str$(GLTrans.CrAmt))
'''   Mid$(ToPrint$, 85) = "Record:" + Str$(TCnt&)
'''   Print #LogFile, ToPrint$
'''   BadDebits# = BadDebits# + GLTrans.DrAmt
'''   BadCredits# = BadCredits# + GLTrans.CrAmt
'''Return
'''
'''
'''End Sub
'''
