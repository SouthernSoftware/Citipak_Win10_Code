VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeleteTagTrans 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Tagged Transactions"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12195
   Icon            =   "frmDeleteTagTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12195
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
      Top             =   5208
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
      Top             =   5208
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8508
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
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
            TextSave        =   "12:53 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "2/15/2013"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "BACKUP BEFORE RUNNING THIS OPERATION!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   396
      Left            =   3552
      TabIndex        =   7
      Top             =   3792
      Width           =   5148
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Transactions check in but they don't check out."
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
      Height          =   348
      Left            =   3552
      TabIndex        =   6
      Top             =   4224
      Width           =   5148
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This utility removes tagged transaction records from the history file."
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
      Height          =   588
      Left            =   3114
      TabIndex        =   5
      Top             =   3120
      Width           =   6156
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete  Tagged Transactions"
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
      Left            =   3618
      TabIndex        =   4
      Top             =   1608
      Width           =   4956
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F10 to Delete Transactions or Escape to Exit."
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
      Left            =   3312
      TabIndex        =   3
      Top             =   4704
      Width           =   5580
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   1248
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00013789&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3084
      Left            =   2688
      Top             =   2832
      Width           =   6828
   End
End
Attribute VB_Name = "frmDeleteTagTrans"
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
  Unload frmDeleteTagTrans
End Sub
Private Sub cmdGo_Click()
  DelTagTrans
  frmGLUtilMenu.Show
  Unload frmDeleteTagTrans
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

Private Sub DelTagTrans()
  Dim NTrRecLen As Integer, NGLTransFile As Integer, NumNewTrans As Long
  Dim LogFile As Integer, GLTransFile As Integer, NumTrans As Long
  Dim Tr As GLTransRecType
  Dim NTr As GLTransRecType
  Dim TRRec As Long, NewTrans As Long, ToPrint As String
  Dim ToPrint1 As String, ToPrint2 As String, ToPrint3 As String
  Dim Gone As Long, DebitsGone As Double, CreditsGone As Double
  If Exist("gltrans.old") Then
     If MsgBox("WARNING!! A Backup file from a prior operation exits. Kill it? (Y/N", vbYesNo, "Continue") = vbYes Then
        Kill "gltrans.old"
     Else
        Exit Sub
     End If
  End If
  If Exist("ngltrans.dat") Then
    Kill "ngltrans.dat"
  End If
  '--open a new GLtrans.file
  NTrRecLen = Len(NTr)
  NGLTransFile = FreeFile
  Open "ngltrans.dat" For Random Access Read Write Shared As NGLTransFile Len = NTrRecLen '85

  NumNewTrans& = LOF(NGLTransFile) \ NTrRecLen

  '--open a log file to list transactions removed
  LogFile = FreeFile
  Open "glutil.log" For Append As LogFile
  Print #LogFile,
  Print #LogFile, "Removed transactions procedure started @ " + Date$ + " " + Time$

  OpenTransFile GLTransFile, NumTrans&
''''****
''***** Put the frmshowpct here for calculations
  FrmShowPctComp.Label1 = "Processing. Please wait."
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdGo.Enabled = False
  For TRRec& = 1 To NumTrans&
    Get GLTransFile, TRRec&, Tr
    'Complete# = ( / NumTrans&) * 100
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&

    If Tr.Marked = 0 Then
      '--copy good records to the new file
      NewTrans& = NewTrans& + 1
      NTr.AcctNum = Tr.AcctNum
      NTr.TRDATE = Tr.TRDATE
      NTr.Desc = Tr.Desc
      NTr.CrAmt = Tr.CrAmt
      NTr.DrAmt = Tr.DrAmt
      NTr.Ref = Tr.Ref
      NTr.Src = Tr.Src
      NTr.NextTran = Tr.NextTran
      Put NGLTransFile, NewTrans&, NTr
    Else
      '-check em out
      ToPrint$ = Space$(80)
      LSet ToPrint$ = Tr.AcctNum
      Mid$(ToPrint$, 13) = Format(DateAdd("d", (Tr.TRDATE), "12-31-1979"), "mm/dd/yyyy")
      Mid$(ToPrint$, 24) = Left$(Tr.Desc, 19)
      Mid$(ToPrint$, 42) = Str$(Tr.CrAmt)
      Mid$(ToPrint$, 52) = Str$(Tr.DrAmt)
      Mid$(ToPrint$, 62) = Tr.Ref
      Mid$(ToPrint$, 72) = Tr.Src
      Print #LogFile, ToPrint$

      '--Keep track of what's gone
      Gone& = Gone& + 1
      DebitsGone# = DebitsGone# + Tr.DrAmt
      CreditsGone# = CreditsGone# + Tr.CrAmt

    End If

  Next
'Tell user what was deleted
  ToPrint1 = "Transactions removed: " + Using$("#####", Gone&)
  ToPrint2 = "Total Debits removed: " + Using$("###,###,###.##", DebitsGone#)
  ToPrint3 = "Total Credits removed: " + Using$("###,###,###.##", CreditsGone#)
  MsgBox ToPrint1 & Chr$(13) & ToPrint2 & Chr$(13) & ToPrint3, vbOKOnly, "Transactions Deleted"

  'log
  Print #LogFile,
  Print #LogFile, "Transactions removed: "; Gone&
  Print #LogFile, "Total Debits removed: " + Str$(DebitsGone#)
  Print #LogFile, "Total Credits removed: " + Str$(CreditsGone#)
  Print #LogFile,

  Close
  If Exist("gltrans.old") Then
    Kill "gltrans.old"
  End If
  Name "gltrans.dat" As "gltrans.old"
  Name "ngltrans.dat" As "gltrans.dat"

  MsgBox "Press any key to continue procedure.", vbOKOnly, "Continue with Relink"
  Call MainLog("DeleteTagTrans -Total Gone: " + Str$(Gone&) + " Start Relink.")
  ReLinkTrans frmDeleteTagTrans 'frmRelinkGLTrans
  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
End Sub
