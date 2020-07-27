VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmPayrollProcessingMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Processing Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmPayrollProcessingMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   585
      Top             =   900
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   360
   End
   Begin fpBtnAtlLibCtl.fpBtn PostPayTransCmmd 
      Height          =   405
      Left            =   4005
      TabIndex        =   7
      Top             =   5400
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn ACHBankDraftProcCmmd 
      Height          =   405
      Left            =   4005
      TabIndex        =   6
      Top             =   4860
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":0AB7
   End
   Begin fpBtnAtlLibCtl.fpBtn PrintPayChecksCmmd 
      Height          =   405
      Left            =   4005
      TabIndex        =   5
      Top             =   4320
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":0CA4
   End
   Begin fpBtnAtlLibCtl.fpBtn SetPayPeriodDefaultsCmmd 
      Height          =   405
      Left            =   4005
      TabIndex        =   2
      Top             =   2700
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":0E8C
   End
   Begin fpBtnAtlLibCtl.fpBtn AccrueLeaveBeneCmmd 
      Height          =   405
      Left            =   4005
      TabIndex        =   1
      Top             =   2160
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":1077
   End
   Begin fpBtnAtlLibCtl.fpBtn EnterEditPayTrans 
      Height          =   405
      Left            =   4005
      TabIndex        =   3
      Top             =   3240
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":1260
   End
   Begin fpBtnAtlLibCtl.fpBtn PrintRegisterCmmd 
      Height          =   405
      Left            =   4005
      TabIndex        =   4
      Top             =   3780
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":1453
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdVoid 
      Height          =   405
      Left            =   4005
      TabIndex        =   8
      Top             =   5940
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":1635
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdManTranEnt 
      Height          =   405
      Left            =   4005
      TabIndex        =   9
      Top             =   6480
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":1824
   End
   Begin fpBtnAtlLibCtl.fpBtn exitCmd 
      Height          =   405
      Left            =   4005
      TabIndex        =   11
      Top             =   7560
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":1A10
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   405
      Left            =   4005
      TabIndex        =   10
      Top             =   7020
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
      ButtonDesigner  =   "frmPayrollProcessingMenu.frx":1C00
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2101
      Top             =   2103
      Width           =   971
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2919.248
      Y1              =   7884
      Y2              =   7884.973
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2151.243
      Y2              =   7880.108
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2151.243
      Y2              =   7892.757
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7884.973
      Y2              =   7884.973
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   1097
      Left            =   1500
      Top             =   897
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAYROLL PROCESSING MENU"
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
      Left            =   2820
      TabIndex        =   0
      Top             =   1250
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   8592
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8712
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2220
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   2100
      Top             =   1971
      Width           =   975
   End
End
Attribute VB_Name = "frmPayrollProcessingMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub AccrueLeaveBeneCmmd_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmAccrueMenu.Show
  DoEvents
  Unload frmPayrollProcessingMenu
End Sub

Private Sub ACHBankDraftProcCmmd_Click()
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim UHandle As Integer
  Dim UnitRec As UnitFileRecType
  Dim FileHandle As Integer
  
  If Not Exist("prdata\ChecksPrinted.opn") Then
    frmWarnPrintChecksFirst.Show vbModal, Me
    Exit Sub
  End If
  
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 1) = False Then
    Close
    Exit Sub
  End If
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  
  If UnitRec.BankDraft = "N" Then
    frmMessage.Label1.Caption = "The 'Bank Draft Y/N?' flag on the Employer Maintenance screen is set to 'N'. The 'Bank Draft' menu is disabled when the 'Bank Draft Y/N?' flag is set to 'N'."
    frmMessage.Label1.Top = 750
    frmMessage.Show vbModal
    Exit Sub
  End If
  
LetEmIn:
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
  Close PHandle
  
  If PPDRec.MACTIVE = -1 Then
    frmWarnMActiveIsT.Show vbModal, Me
    Exit Sub
  End If
  
  If PPDRec.PACTIVE = 0 Then
    frmWarnNOPPD.Show vbModal, Me
    Exit Sub
  End If

  frmACHBankDraftMenu.Show
  DoEvents
  Unload frmPayrollProcessingMenu

End Sub

Private Sub cmdClear_Click()
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim FileHandle As Integer
  Dim PRType As Integer
  
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
'  If PPDRec.MACTIVE = -1 Then 'manual payroll is being used
'    frmWarnMActiveIsT.Show vbModal, Me
'    Exit Sub
'  End If
  
  If PPDRec.PACTIVE = 0 And PPDRec.MACTIVE = 0 Then  'neither manual or regular
  'is being used
     MsgBox "No payroll is currently being processed."
     Close PHandle
     Exit Sub
  End If
  
  frmPRMessageWOpts.Label1.Caption = "Clearing the current payroll resets all employee flags to 'N'. All data collection in progress is deleted. Do you wish to continue?"
  frmPRMessageWOpts.Label1.Top = 800
  frmPRMessageWOpts.Show vbModal
  If frmPRMessageWOpts.fptxtChoice.Text = "abort" Then
    Unload frmPRMessageWOpts
    Exit Sub
  Else
    Unload frmPRMessageWOpts
  End If
  
  If PPDRec.MACTIVE = -1 Then
    PRType = 2
  ElseIf PPDRec.PACTIVE = -1 Then
    PRType = 1
  Else
    PRType = 0
  End If
  
  PPDRec.MACTIVE = 0
  PPDRec.PACTIVE = 0
  Put PHandle, 1, PPDRec
  
  Close PHandle
  
  KillFile "TEMPIF.DAT"
  KillFile "prdata\ChecksPrinted.opn"
  
  Call MakeTransInactive
  
  If PRType = 1 Then
    MsgBox "The current regular payroll has been cleared successfully"
    MainLog ("Regular payroll was cleared successfully.")
  ElseIf PRType = 2 Then
    MsgBox "The current manual payroll has been cleared successfully"
    MainLog ("Manual payroll was cleared successfully.")
  Else
    MsgBox "No payroll was currently in process."
  End If
  
End Sub

Private Sub cmdManTranEnt_Click()
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  
  If Not Exist("PRDATA\PRPPDEF.DAT") Then
    frmManTransMenu.Show
    DoEvents
    Unload frmPayrollProcessingMenu
    EntryType = 2
    Exit Sub
  End If
  
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
  Close PHandle
  
  If PPDRec.PACTIVE = 0 Then
    frmManTransMenu.Show
    DoEvents
    Unload frmPayrollProcessingMenu
    EntryType = 2
  Else
    frmWarnPActiveIsT.Show vbModal, Me
  End If

End Sub

Private Sub cmdVoid_Click()
  InFileNames(1) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 1) = False Then
    Close
    Exit Sub
  End If
  frmVoidChkEmpList.Show
  DoEvents
  Unload frmPayrollProcessingMenu
  MainLog ("Void a Posted Payroll Check loaded.")
End Sub

Private Sub EnterEditPayTrans_Click()
  
  Dim PPDHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim PPDRecCnt As Integer
  
  If Exist("TEMPIF.DAT") Then 'added 12/17/2008
    If MsgBox("Payroll registers have already been processed. Saving or deleting any new data will destroy the current payroll register file which will prevent posting. Do you wish to continue anyway?", vbYesNo) = vbNo Then
      Exit Sub
    Else
      KillFile "TEMPIF.DAT"
      KillFile "PRData\ChecksPrinted.opn"
    End If
  End If
  
  OpenPPDefaultFile PPDHandle
  PPDRecCnt = LOF(PPDHandle) \ Len(PPDRec)
  If PPDRecCnt = 0 Then
    frmWarnNOPPD.Show vbModal, Me
    Exit Sub
  Else
    Get PPDHandle, 1, PPDRec
  End If
  Close PPDHandle
  
  If PPDRec.MACTIVE = -1 Then
    frmWarnMActiveIsT.Show vbModal, Me
    Exit Sub
  End If
  
  If PPDRec.MACTIVE = 0 And PPDRec.PACTIVE = 0 Then
    frmWarnNOPPD.Show vbModal, Me
    SetPayPeriodDefaultsCmmd.SetFocus
    Exit Sub
  End If
  
  InFileNames(1) = "PRDATA\PRERNCOD.DAT"
  InFileNames(2) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  
  EntryType = 1 'Normal
  frmPRTPrevEmpLookUp.Show
  DoEvents
  Unload frmPayrollProcessingMenu
End Sub

Private Sub exitCmd_Click()
   frmPayrollMainMenu.Show
   DoEvents
   Unload frmPayrollProcessingMenu
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call exitCmd_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Me.HelpContextID = hlpPayroll
  If Exist("prmain.dat") Then
    KillFile "prmain.dat"
  ElseIf Exist("paycheckmain.dat") Then
    KillFile "paycheckmain.dat"
    PrintPayChecksCmmd.Enabled = False
  End If
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  CurrCitiPath = App.Path 'must reassign CurrCitiPath here
  'because if checks are printed the check printing exe sends
  'the program here when it closes thus bypassing the main
  'menu where CurrCitiPath is assigned also '8/27/04
  
  NewListFlag = False 'this is used in the PRDefault
  'edit transaction so that when a selected employee
  'is processed the flag becomes true and the list that
  'shows up when the employee list comes up is always
  'updated even after it reloads after PRCalcScreen exits
  
  Call TaxTextLoad 'this tells the CalcPay State w/h
  'section which state to use and other state specific data
  
  'placing TaxTextLoad (originally placed in the Payroll
  'Main Menu load procedure) here allows it to be activated
  'each time Payroll Processing begins and thereby updating
  'any new changes and since PayCheck exits directly back
  'here it forces TaxTextLoad to refresh in case the user
  'wants to go back and edit another employee before
  'posting (and running paycheck again). In Payroll Main
  'it was skipped causing the program to be unable
  'to figure state tax on another employee.
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub PostPayTransCmmd_Click()
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim FileHandle As Integer
  
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
  Close PHandle
  If PPDRec.MACTIVE = -1 Then 'manual payroll is being used
    frmWarnMActiveIsT.Show vbModal, Me
    Exit Sub
  End If
  
  If PPDRec.PACTIVE = 0 Then 'neither manual or regular
  'is being used
     frmWarnNOPPD.Show
     Exit Sub
  End If
  
  InFileNames(1) = "PRDATA\PRSYS.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(4) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then
    Close
    Exit Sub
  End If
  
  'keeps posting from occurring if checks haven't been printed
  '...ChecksPrinted.opn is created after checks are written
  'and deleted during the posting process...so if it isn't
  'there then checks have not been written and posting cannot
  'occur until checks are written
  If Not Exist("prdata\ChecksPrinted.opn") Then
    frmWarnPrintChecksFirst.Show vbModal, Me
    Exit Sub
  End If
  
  EntryType = 1 'Normal posting code.
  'posting code is in frmWarningPostPayroll
  frmWarningPostPayroll.Show
End Sub

Private Sub PrintPayChecksCmmd_Click()
  Dim PPDHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim SysRec As RegDSysFileRecType
  Dim SHandle As Integer
  Dim CheckStyle As Integer
  Dim FileHandle As Integer
  Dim ShellHandle As Integer
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim One As Integer
  Dim SSHandle As Integer
  Dim AHandle As Integer
  
  If Not Exist("TEMPIF.DAT") Then
'    MsgBox "Please run payroll registers before running payroll checks. Failure to run registers will cause errors when posting to the General Ledger."
    frmWarnRunRegsFirst.Show vbModal
    MainLog ("Attempt made to run payroll checks without running payroll registers.")
    Close
    PrintRegisterCmmd.SetFocus
    Exit Sub
  End If
  
  OpenSysFile SHandle
  Get SHandle, 1, SysRec
  Close SHandle
  
  CheckStyle = SysRec.CheckStyle
  If CheckStyle = 0 Then
    MsgBox "Please select a check type in the System Interface screen"
    Exit Sub
  End If
  
  InFileNames(1) = "PRDATA\PRSYS.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PRERNCOD.DAT"
  InFileNames(4) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(5) = "PRDATA\PREMP2.DAT"
  InFileNames(6) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 6) = False Then
    Close
    Exit Sub
  End If
  
  OpenPPDefaultFile PPDHandle
  Get PPDHandle, 1, PPDRec
  Close PPDHandle
  If PPDRec.MACTIVE = -1 Then
    frmWarnMActiveIsT.Show vbModal, Me
    Exit Sub
  End If
  
  If PPDRec.PACTIVE = 0 Then
     frmWarnNOPPD.Show
     Exit Sub
  End If
  
  One = 1
  
  AHandle = FreeFile
  Open "prmain.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  If PWcnt = -3 Then 'internal use
  SSHandle = FreeFile
  Open "sosoftpw.dat" For Output As SSHandle
  Print #SSHandle, One
  Close SSHandle
  GoTo SSPW
  End If

  OpenCitiPassFile CitiPassFile, NumPassRecs 'reassign all globals
  Get CitiPassFile, PWcnt, CitiPass
  CitiPass.Flag2 = PWcnt 'set Flag2 to PWcnt denoting that we are now moving
  'to payrollcheck.exe...Flag2 is used in payrollcheck.exe in case
  'it is terminated abnormally
  Put CitiPassFile, PWcnt, CitiPass
  Close CitiPassFile
SSPW:
  Shell "Payrollcheck.exe", vbMaximizedFocus 'exiting payroll and
  Timer1.Enabled = True
'  Close 'cleans up any open files
'  Call Terminate2Shell 'closes all forms but preserves password data
'  End
  
End Sub

Private Sub PrintRegisterCmmd_Click()
  Dim SysRec(1) As RegDSysFileRecType
  Dim SysHandle As Integer
  Dim PPDHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim CITIDIR As Integer
  
  InFileNames(1) = "PRDATA\PRSYS.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PRERNCOD.DAT"
  InFileNames(4) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(5) = "PRDATA\PREMP2.DAT"
  InFileNames(6) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 6) = False Then
    Close
    Exit Sub
  End If
 
'  CITIDIR = CheckCitiDir(SysRec(1).CITIDIR)
  CITIDIR = CheckCitiDir(CurrCitiPath)

  If CITIDIR = 0 Then
    frmWarnNoCitipakPath.Show vbModal
    Exit Sub
  End If
  
  OpenPPDefaultFile PPDHandle
  Get PPDHandle, 1, PPDRec
  
  If PPDRec.MACTIVE = -1 Then
    frmWarnMActiveIsT.Show vbModal, Me
    Close
    Exit Sub
  End If
  
  If PPDRec.MACTIVE = 0 And PPDRec.PACTIVE = 0 Then
    frmWarnNOPPD.Show vbModal, Me
    SetPayPeriodDefaultsCmmd.SetFocus
    Close
    Exit Sub
  End If
  
  Close PPDHandle
  
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec(1)
  Close SysHandle
  
  Call DeActivateControls
  
  frmReportOpt.Show vbModal
  
  If SysRec(1).SplitFlag = "Y" Then
    SplitFlag = True
    If RptOpt = 2 Then
      Call PCPrintPayRegisterST(CITIDIR) 'PCPrintPayRegisterS(plit)
    ElseIf RptOpt = 1 Then
      Call PCPrintPayRegisterSG(CITIDIR) 'PCPrintPayRegisterS(plit)
    Else
      Call ActivateControls
      Exit Sub
    End If
    MainLog ("Payroll Registers processed.")
  Else
    SplitFlag = False
    If RptOpt = 2 Then
      Call PCPrintPayRegisterT(CITIDIR) 'PCPrintPayRegister
    ElseIf RptOpt = 1 Then
      Call PCPrintPayRegisterG(CITIDIR) 'PCPrintPayRegister
    Else
      Call ActivateControls
      Exit Sub
    End If
    MainLog ("Payroll Registers processed.")
  End If
  Call ActivateControls

End Sub

Private Sub SetPayPeriodDefaultsCmmd_Click()
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
  Close PHandle
  
  If PPDRec.MACTIVE = 0 Then
    frmPRDefaultSet.Show
    DoEvents
    Unload frmPayrollProcessingMenu
  Else
    frmWarnMActiveIsT.Show vbModal, Me
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If exitCmd.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via frmPayrollProcessingMenu menu bar.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

  Timer2.Enabled = False
  AccrueLeaveBeneCmmd.Enabled = False
  SetPayPeriodDefaultsCmmd.Enabled = False
  EnterEditPayTrans.Enabled = False
  PrintRegisterCmmd.Enabled = False
  PrintPayChecksCmmd.Enabled = False
  ACHBankDraftProcCmmd.Enabled = False
  PostPayTransCmmd.Enabled = False
  cmdVoid.Enabled = False
  cmdManTranEnt.Enabled = False
  exitCmd.Enabled = False
  cmdClear.Enabled = False
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
  
  Timer2.Enabled = True
  AccrueLeaveBeneCmmd.Enabled = True
  SetPayPeriodDefaultsCmmd.Enabled = True
  EnterEditPayTrans.Enabled = True
  PrintRegisterCmmd.Enabled = True
  PrintPayChecksCmmd.Enabled = True
  ACHBankDraftProcCmmd.Enabled = True
  PostPayTransCmmd.Enabled = True
  cmdVoid.Enabled = True
  cmdManTranEnt.Enabled = True
  cmdClear.Enabled = True
  exitCmd.Enabled = True
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
End Sub

Private Sub Timer1_Timer()
  Call Terminate2Shell 'closes all forms but preserves password data
End Sub

Private Sub Timer2_Timer()
  PrintPayChecksCmmd.Enabled = True
End Sub
