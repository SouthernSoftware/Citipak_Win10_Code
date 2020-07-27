VERSION 5.00
Begin VB.Form frmSortAcct 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort Accounts"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   90
   ClientWidth     =   5445
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5445
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&XIT"
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
      Left            =   3000
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton cmdOKSort 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&OK"
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
      Left            =   1200
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sorting..............."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1032
      TabIndex        =   3
      Top             =   1176
      Visible         =   0   'False
      Width           =   3372
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmSortAcct.frx":0000
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready to Sort Account Index?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3852
   End
End
Attribute VB_Name = "frmSortAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
Dim GLAcct As GLAcctRecType
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
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub cmdExit_Click()
  Unload frmSortAcct
End Sub

Private Sub cmdOKSort_Click()
  Call MainLog("Sort Accounts Started - Menu Option.")
  Unload frmSortAcct
  frmChartAcctMaintMenu.OKFromSort
  'SortAccIdx
End Sub

Private Sub Form_Initialize()
  vWidth = Screen.Width * 0.45    ' Set width of form.
  vHeight = Screen.Height * 0.33  ' Set height of form.
  vLeft = (Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vTop = ((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  MakeWindowTopMost hwnd, True
  Me.HelpContextID = hlpAccountIndex
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Function SortAccIdx()
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, AcctFileNum As Integer
  Dim NumAccts As Integer, CntAc As Integer, GoodAccts As Integer
  Dim OutOfOrder As Boolean, TempIdxRec As GLAcctIndexType
  
  KillFileD "GLAcct.IDX"
'  FrmShowPctComp.Label1 = "Initializing Account Index."
'  FrmShowPctComp.Show , Me
'  DoEvents

  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  OpenAcctFile AcctFileNum
  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
  If NumAccts < 1 Then    'no need to sort one record
    Close AcctIdxFileNum, AcctFileNum
    Exit Function
  End If
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOKSort.Enabled = False
  Label2.Visible = True
  Screen.MousePointer = 11
  ReDim Idxbuff(1 To NumAccts) As GLAcctIndexType
  For CntAc = 1 To NumAccts
   ' FrmShowPctComp.ShowPctComp CntAc, NumAccts

    Get AcctFileNum, CntAc, GLAcct
    If GLAcct.Deleted = 0 Then
      GoodAccts = GoodAccts + 1
      Idxbuff(GoodAccts).AcctNum = GLAcct.Num
      Idxbuff(GoodAccts).RecNum = CntAc
    End If
  Next
  Close AcctFileNum
  If GoodAccts = 0 Then
    Close AcctIdxFileNum
    Exit Function
  End If
  ReDim Preserve Idxbuff(1 To GoodAccts) As GLAcctIndexType
'  FrmShowPctComp.Label1 = "Sorting...Please Wait..."
'  FrmShowPctComp.Show , Me
'  DoEvents
'  FrmShowPctComp.ShowPctComp 15, 100
  Do
    OutOfOrder = False          'assume it's sorted
    For CntAc = 1 To GoodAccts - 1
      If Idxbuff(CntAc).AcctNum > Idxbuff(CntAc + 1).AcctNum Then
        LSet TempIdxRec = Idxbuff(CntAc)
        LSet Idxbuff(CntAc) = Idxbuff(CntAc + 1)
        LSet Idxbuff(CntAc + 1) = TempIdxRec
        OutOfOrder = True       'we're not done yet
      End If
    
    Next
  Loop While OutOfOrder
  'FrmShowPctComp.ShowPctComp 55, 100
  For CntAc = 1 To GoodAccts
  'FrmShowPctComp.ShowPctComp CntAc, GoodAccts
    Put AcctIdxFileNum, CntAc, Idxbuff(CntAc)
  Next
   MsgBox "Sort is Complete, Press OK to Continue", vbOKOnly, "Sort Completed"

  Close AcctIdxFileNum
  Screen.MousePointer = 0
  Me.cmdExit.Enabled = True
  Me.cmdOKSort.Enabled = True
  EnableCloseButton Me.hwnd, True
  Label2.Visible = False
'''GetPct:
'''  PctComp = Int((cnt / TotalCnt) * 100)
'''  Label3 = PctComp
'''  ProgressBar1.Value = PctComp
'''
'''  If PctComp = 100 Then
'''    MakeWindowTopMost Me.hwnd, False
'''    Unload FrmShowPctComp
'''    DoEvents
'''  Else
'''    DoEvents
'''  End If
'''Return
End Function
