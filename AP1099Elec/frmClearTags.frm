VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClearTags 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clear Transaction Query"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmClearTags.frx":0000
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
      Left            =   3990
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4056
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
      Top             =   4056
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
            TextSave        =   "10/5/2004"
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
      BackStyle       =   0  'Transparent
      Caption         =   "Press F10 to Un-Mark All Transactions or Escape to Exit."
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
      Height          =   348
      Left            =   3216
      TabIndex        =   4
      Top             =   3456
      Width           =   6108
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1788
      Left            =   2736
      Top             =   3024
      Width           =   6828
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Un-Mark All Tagged Transactions"
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
      Left            =   3624
      TabIndex        =   3
      Top             =   1248
      Width           =   4956
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1008
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   888
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
Attribute VB_Name = "frmClearTags"
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
  Unload frmClearTags
End Sub
Private Sub cmdGo_Click()
  ClearQuery
End Sub

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


Private Sub ClearQuery()
  Dim RecNo As Long, PrintBal As String, DrCr As String
  Dim DrAmt As Double, CrAmt As Double, RCnt As Long
  Dim Diff As Double, Bal As Double, ToPrint1 As String, ToPrint3 As String
  Dim CommaFmt As String, TotalFmt As String, ToPrint2 As String
  Dim TransFileNum As Integer, NumTrans As Long
  OpenTransFile TransFileNum, NumTrans&
  DrAmt# = 0
  CrAmt# = 0
  RCnt& = 0
  FrmShowPctComp.Label1 = "Un-Marking Transactions"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdGo.Enabled = False
  Me.mnuOptions.Enabled = False
  For RecNo& = 1 To NumTrans&

    '--Set Flag to False
    Get TransFileNum, RecNo&, GLTrans
    FrmShowPctComp.ShowPctComp RecNo&, NumTrans&
    If GLTrans.Marked <> 0 Then
      GLTrans.Marked = 0
      Put TransFileNum, RecNo&, GLTrans

      '--Summarize
      RCnt& = RCnt& + 1
      DrAmt# = DrAmt# + GLTrans.DrAmt
      CrAmt# = CrAmt# + GLTrans.CrAmt
  
      '--PLAYING
      'DrAmt# = Trans.DrAmt
      'CrAmt# = Trans.CrAmt
      'CrAmt# = Trans.DrAmt
      'DrAmt# = Trans.CrAmt
    
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          Me.cmdExit.Enabled = True
          Me.cmdGo.Enabled = True
          EnableCloseButton Me.hwnd, True
          Me.mnuOptions.Enabled = True
          Unload FrmShowPctComp
          GoTo CancelExit
        End If
      End If
      
  Next          'Process next transaction

  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
  Me.mnuOptions.Enabled = True
  Close

  Diff# = Round#(DrAmt# - CrAmt#)
  Bal# = Abs(Diff#)

  If Diff# = 0 Then
    DrCr$ = " Transactions are in Balance!"
  ElseIf Diff# > 0 Then
    PrintBal = Using("###,###,###.##", Bal#)
    DrCr$ = " Debit" & Chr(13) & "Transactions are out of balance!" & Chr$(13) & PrintBal
  Else
    PrintBal = Using("###,###,###.##", Bal#)
    DrCr$ = " Credit" & Chr(13) & "Transactions are out of balance!" & Chr$(13) & PrintBal
  End If
  ToPrint1 = Using$("#####", RCnt&)
  ToPrint2 = Using$("###,###,###.##", DrAmt#)
  ToPrint3 = Using$("###,###,###.##", CrAmt#)
  Call MainLog("Clear Tags.")
  MsgBox "Records Un-Marked: " & ToPrint1 & Chr$(13) & "Total Debits: " & ToPrint2 & Chr$(13) & "Total Credits: " & ToPrint3 & Chr(13) & DrCr$, vbOKOnly, "Cleared Transactions"
CancelExit:
  Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
