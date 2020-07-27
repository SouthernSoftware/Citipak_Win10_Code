VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmControlFileMaint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Maintenance Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   FillColor       =   &H8000000B&
   Icon            =   "frmControlFileMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8651.707
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdNEWYEAR 
      Height          =   375
      Left            =   4005
      TabIndex        =   11
      Top             =   6990
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrinter 
      Height          =   375
      Left            =   4005
      TabIndex        =   10
      Top             =   6560
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":0ABD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdACHDraft 
      Height          =   375
      Left            =   4005
      TabIndex        =   9
      Top             =   6132
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":0CB0
   End
   Begin fpBtnAtlLibCtl.fpBtn EmployerFileMaintCmmd 
      Height          =   375
      Left            =   4005
      TabIndex        =   0
      Top             =   2280
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":0E9E
   End
   Begin fpBtnAtlLibCtl.fpBtn SystemFileMaintCmmd 
      Height          =   375
      Left            =   4005
      TabIndex        =   1
      Top             =   2708
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":107F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdStateTaxTbl 
      Height          =   375
      Left            =   4005
      TabIndex        =   2
      Top             =   3135
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":125E
   End
   Begin fpBtnAtlLibCtl.fpBtn FedTaxTableCmmd 
      Height          =   375
      Left            =   4005
      TabIndex        =   3
      Top             =   3564
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":1441
   End
   Begin fpBtnAtlLibCtl.fpBtn EICTableMaintCmmd 
      Height          =   375
      Left            =   4005
      TabIndex        =   4
      Top             =   3992
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":1626
   End
   Begin fpBtnAtlLibCtl.fpBtn LeaveBeneTableMaintCmmd 
      Height          =   384
      Left            =   4008
      TabIndex        =   5
      Top             =   4416
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   677
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
      ButtonDesigner  =   "frmControlFileMaint.frx":1803
   End
   Begin fpBtnAtlLibCtl.fpBtn DeductionCodeMaintCmmd 
      Height          =   375
      Left            =   4005
      TabIndex        =   6
      Top             =   4845
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":19EA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEarnings 
      Height          =   375
      Left            =   4005
      TabIndex        =   7
      Top             =   5276
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":1BCC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRetFileMaint 
      Height          =   384
      Left            =   4008
      TabIndex        =   8
      Top             =   5700
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   677
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
      ButtonDesigner  =   "frmControlFileMaint.frx":1DAD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   375
      Left            =   4005
      TabIndex        =   12
      Top             =   7416
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmControlFileMaint.frx":1F90
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Index           =   1
      Left            =   2101
      Top             =   2101
      Width           =   971
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Index           =   0
      Left            =   8593
      Top             =   2101
      Width           =   971
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1500
      Top             =   896
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONTROL MAINTENANCE MENU"
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
      TabIndex        =   13
      Top             =   1248
      Width           =   6012
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2919.248
      Y1              =   7883.965
      Y2              =   7883.965
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2148.312
      Y2              =   7882.017
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7883.965
      Y2              =   7883.965
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2151.235
      Y2              =   7883.965
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   768
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   8592
      Top             =   1968
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   2100
      Top             =   1968
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8712
      Top             =   2198
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2198
      Width           =   732
   End
End
Attribute VB_Name = "frmControlFileMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdACHDraft_Click()
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  
  If UnitRec.BankDraft = "N" Then
    frmMessage.Label1.Caption = "The 'Bank Draft Y/N?' flag on the Employer Maintenance screen is set to 'N'. The 'Bank Draft' menu is disabled when the 'Bank Draft Y/N?' flag is set to 'N'."
    frmMessage.Label1.Top = 750
    frmMessage.Show vbModal
    Exit Sub
  End If
  
  frmACHControlMenu.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub cmdEarnings_Click()
  frmEarningsCodeMaint.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub cmdExit_Click()
  frmPayrollMainMenu.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub cmdNEWYEAR_Click()
  InFileNames(1) = "PRDATA\PREMP3.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 1) = False Then
    Close
    Exit Sub
  End If
  frmWarningYearEnd.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub cmdPrinter_Click()
   InFileNames(1) = "PRDATA\PRPRNDF.DAT"
   If FilesROK(Me, InFileNames(), OutFileNames, 1) = False Then
     Exit Sub
   End If
  frmPrinterSetup.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub cmdRetFileMaint_Click()
  frmRetireFileMaint.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub cmdStateTaxTbl_Click()
  InFileNames(1) = "PRDATA\PRSTADEF.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmLoadingState.Show
  Call DeActivateControls
  DoEvents
  frmStateTaxSingle.Show
  DoEvents
  Unload frmControlFileMaint
  DoEvents
  Unload frmLoadingState

End Sub

Private Sub DeductionCodeMaintCmmd_Click()
  InFileNames(1) = "PRDATA\PRSYS.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 1) = False Then
    Close
    Exit Sub
  End If
  frmDeductionCodes.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub EICTableMaintCmmd_Click()
  frmEICMaint.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub EmployerFileMaintCmmd_Click()
  frmEmployerInfoFile.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%M"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpControlFile
  
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub LeaveBeneTableMaintCmmd_Click()
  Load frmEmpLvBeneMnt
  DoEvents
  frmEmpLvBeneMnt.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub FedTaxTableCmmd_Click()
  frmLoadingFed.Show
  Call DeActivateControls
  DoEvents
  frmFedTaxSingle.Show
  DoEvents
  Unload frmControlFileMaint
  DoEvents
  Unload frmLoadingFed
End Sub

Private Sub SystemFileMaintCmmd_Click()
  frmIntFaceMaint.Show
  DoEvents
  Unload frmControlFileMaint
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmControlFileMaint.")
      Call Terminate
      End
    End If
  End If
End Sub
Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

  EmployerFileMaintCmmd.Enabled = False
  SystemFileMaintCmmd.Enabled = False
  cmdStateTaxTbl.Enabled = False
  FedTaxTableCmmd.Enabled = False
  EICTableMaintCmmd.Enabled = False
  LeaveBeneTableMaintCmmd.Enabled = False
  DeductionCodeMaintCmmd.Enabled = False
  cmdEarnings.Enabled = False
  cmdRetFileMaint.Enabled = False
  cmdACHDraft.Enabled = False
  cmdPrinter.Enabled = False
  cmdNEWYEAR.Enabled = False
  cmdExit.Enabled = False
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = False
      End If
  Next cnt
    EnableCloseButton Me.hwnd, False
     
End Sub

Private Sub ActivateControls()
  Dim cmdButton As CommandButton
  Dim x As Control
  Dim cnt As Integer
  
  EmployerFileMaintCmmd.Enabled = True
  SystemFileMaintCmmd.Enabled = True
  cmdStateTaxTbl.Enabled = True
  FedTaxTableCmmd.Enabled = True
  EICTableMaintCmmd.Enabled = True
  LeaveBeneTableMaintCmmd.Enabled = True
  DeductionCodeMaintCmmd.Enabled = True
  cmdEarnings.Enabled = True
  cmdRetFileMaint.Enabled = True
  cmdACHDraft.Enabled = True
  cmdPrinter.Enabled = True
  cmdNEWYEAR.Enabled = True
  cmdExit.Enabled = True
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
End Sub

