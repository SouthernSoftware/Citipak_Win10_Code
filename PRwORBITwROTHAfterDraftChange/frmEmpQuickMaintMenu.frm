VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmEmpQuickMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Employee Quick Maintenance Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmEmpQuickMaintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleMode       =   0  'User
   ScaleWidth      =   11667
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdJobDescription 
      Height          =   495
      Left            =   4005
      TabIndex        =   2
      Top             =   3690
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDirectDeposit 
      Height          =   495
      Left            =   4005
      TabIndex        =   1
      Top             =   3075
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":0AAD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPersonalData 
      Height          =   495
      Left            =   4005
      TabIndex        =   0
      Top             =   2475
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":0C94
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxWithholding 
      Height          =   495
      Left            =   4005
      TabIndex        =   3
      Top             =   4290
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":0E75
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMiscDeductions 
      Height          =   495
      Left            =   4005
      TabIndex        =   4
      Top             =   4890
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":1058
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAltEarnings 
      Height          =   495
      Left            =   4005
      TabIndex        =   5
      Top             =   5505
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":1244
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdWageDistribution 
      Height          =   495
      Left            =   4005
      TabIndex        =   6
      Top             =   6105
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":142A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBenefitSchedule 
      Height          =   495
      Left            =   4005
      TabIndex        =   7
      Top             =   6705
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":160F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4005
      TabIndex        =   8
      Top             =   7320
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmpQuickMaintMenu.frx":17F3
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
      Top             =   2102
      Width           =   971
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2208.271
      X2              =   2923.006
      Y1              =   7894.763
      Y2              =   7894.763
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8720.97
      X2              =   8720.97
      Y1              =   2154.003
      Y2              =   7888.916
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   9423.692
      X2              =   8720.97
      Y1              =   7894.763
      Y2              =   7894.763
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Height          =   1097
      Left            =   1499
      Top             =   897
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUICK EMPLOYEE MAINTENANCE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   9
      Top             =   1250
      Width           =   6012
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2222.286
      X2              =   2222.286
      Y1              =   2154.003
      Y2              =   7894.763
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   120
      Left            =   8593
      Top             =   2103
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5890
      Index           =   0
      Left            =   2217
      Top             =   2207
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5890
      Index           =   1
      Left            =   8700
      Top             =   2207
      Width           =   731
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2101
      Top             =   1971
      Width           =   975
   End
End
Attribute VB_Name = "frmEmpQuickMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAltEarnings_Click()
  frmEmpQuickMaintAltEarn.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdBenefitSchedule_Click()
  frmEmpQuickMaintBenefits.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdDirectDeposit_Click()
  Dim DHandle As Integer
  Dim DraftRec As DraftInfoFileName
  Dim FileSize As Integer
  Dim One As Integer
  Dim SHandle As Integer
   
  OpenPRDraftFile DHandle
  FileSize = LOF(DHandle) / Len(DraftRec)
  Close DHandle
  If FileSize = 0 Then
    frmMessageWOpts.Label1.Caption = "Nothing has been saved in the bank draft control file. Would you like to jump to the bank draft set up screen now?"
    frmMessageWOpts.Label1.Top = 900
    frmMessageWOpts.cmdCont.Text = "F10 Jump Now"
    frmMessageWOpts.cmdExit.Text = "ESC Don't Jump"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "continue" Then
      One = 1
      SHandle = FreeFile
      Open "quickmaintdd.dat" For Output As SHandle
      Print #SHandle, One
      Close SHandle
      Unload frmMessageWOpts
      frmACHDraftInfo.Show
      DoEvents
      Unload Me
    Else
      frmEmpQuickMaintDirDep.Show
      DoEvents
      Unload Me
    End If
  Else
    frmEmpQuickMaintDirDep.Show
    DoEvents
    Unload Me
  End If
End Sub

Private Sub cmdExit_Click()
  If Exist("quickmaintdd.dat") Then
    KillFile "quickmaintdd.dat"
  End If
  frmEmployeeMaintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdJobDescription_Click()
  frmEmpQuickMaintJobDesc.Show
  DoEvents
  Unload Me

End Sub

Private Sub cmdMiscDeductions_Click()
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim NumOfDedRecs As Integer
  
  OpenDedCodeFile DHandle
  NumOfDedRecs = LOF(DHandle) / Len(DedRec)
  Close
  If NumOfDedRecs = 0 Then
    MsgBox "No deductions have been saved. Screen load aborted."
    Exit Sub
  End If
  
  frmEmpQuickMaintDeduct.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPersonalData_Click()
  frmEmpQuickMaintPers.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdTaxWithholding_Click()
  frmEmpQuickMaintTaxWH.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdWageDistribution_Click()
  frmEmpQuickMaintWageDist.Show
  DoEvents
  Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpQuickMaintenance
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmployeeMaintMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

