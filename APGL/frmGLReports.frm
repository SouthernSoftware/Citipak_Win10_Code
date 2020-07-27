VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmGLReportsMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger Reports"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   Icon            =   "frmGLReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdTrialBalance 
      Height          =   396
      Left            =   4308
      TabIndex        =   0
      Top             =   2352
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCashBalance 
      Height          =   405
      Left            =   4305
      TabIndex        =   1
      Top             =   2850
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   714
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":0AAF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAcctBalSummary 
      Height          =   405
      Left            =   4305
      TabIndex        =   2
      Top             =   3345
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   714
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":0C93
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAcctHistory 
      Height          =   396
      Left            =   4308
      TabIndex        =   3
      Top             =   3852
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":0E82
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBudgHistory 
      Height          =   396
      Left            =   4308
      TabIndex        =   4
      Top             =   4344
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":1069
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBalSheet 
      Height          =   405
      Left            =   4305
      TabIndex        =   5
      Top             =   4845
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   714
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":124F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBudvsAct 
      Height          =   396
      Left            =   4308
      TabIndex        =   6
      Top             =   5352
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":1434
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDeptBudvsAct 
      Height          =   396
      Left            =   4308
      TabIndex        =   7
      Top             =   5844
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":161C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdQueryGLTrans 
      Height          =   405
      Left            =   4305
      TabIndex        =   8
      Top             =   6345
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   714
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":180F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExportFiles 
      Height          =   390
      Left            =   4305
      TabIndex        =   9
      Top             =   6870
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":19FD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFNCTReports 
      Height          =   396
      Left            =   4308
      TabIndex        =   10
      Top             =   7344
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":1BE1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitGLReportsMenu 
      Height          =   405
      Left            =   4305
      TabIndex        =   11
      Top             =   7845
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   714
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmGLReports.frx":1DD4
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
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
      TabIndex        =   13
      Top             =   576
      Visible         =   0   'False
      Width           =   2028
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8304
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3240
      Y1              =   8304
      Y2              =   8304
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
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
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3720
      TabIndex        =   12
      Top             =   1440
      Width           =   4692
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
End
Attribute VB_Name = "frmGLReportsMenu"
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

Private Sub cmdFNCTReports_Click()
  If Exist("glfnct.dat") Then
    If PrYr <> 1 Then
      frmGLFunctionReports.Show
      Unload frmGLReportsMenu
    Else
      frmGLFunctionReports.Label2 = frmGLReportsMenu.Label2
      frmGLFunctionReports.Show
    End If
  Else
    MsgBox "Without GL Function File this option not available."
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
  If Right$(StartPath, 1) = "\" Then
    App.HelpFile = StartPath + "helpfiles\GL.hlp"
  Else
    App.HelpFile = StartPath + "\helpfiles\GL.hlp"
  End If
  Me.HelpContextID = hlpGLReports
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
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

