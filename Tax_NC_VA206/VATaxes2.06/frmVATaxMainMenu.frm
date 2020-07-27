VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Billing Main Menu"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   720
   ClientWidth     =   11655
   Icon            =   "frmVATaxMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   1566
      Top             =   3051
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1566
      Top             =   2451
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxBillingFunctions 
      Height          =   396
      Left            =   4020
      TabIndex        =   3
      Top             =   3552
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterTaxPayments 
      Height          =   396
      Left            =   4008
      TabIndex        =   2
      Top             =   3072
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":0AB3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustMaint 
      Height          =   405
      Left            =   4005
      TabIndex        =   0
      Top             =   2085
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":0C99
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAbstractMaint 
      Height          =   396
      Left            =   4008
      TabIndex        =   1
      Top             =   2580
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":0E81
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxReportingSystem 
      Height          =   396
      Left            =   4020
      TabIndex        =   4
      Top             =   4032
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":1069
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   396
      Left            =   4008
      TabIndex        =   11
      Top             =   7452
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":1251
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdManualEntry 
      Height          =   396
      Left            =   4008
      TabIndex        =   5
      Top             =   4512
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":142F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxBillAdj 
      Height          =   396
      Left            =   4008
      TabIndex        =   6
      Top             =   5004
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":1618
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxAdvertising 
      Height          =   396
      Left            =   4008
      TabIndex        =   7
      Top             =   5484
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":1803
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSetUpAndUtil 
      Height          =   405
      Left            =   4005
      TabIndex        =   8
      Top             =   5985
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":19EE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPPTRA 
      Height          =   396
      Left            =   4008
      TabIndex        =   9
      Top             =   6480
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":1BE0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDMV 
      Height          =   396
      Left            =   4008
      TabIndex        =   10
      Top             =   6960
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmVATaxMainMenu.frx":1DC1
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1098
      Index           =   1
      Left            =   1499
      Top             =   830
      Width           =   8655
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2100
      Top             =   2036
      Width           =   971
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8712
      X2              =   8712
      Y1              =   2144
      Y2              =   8045
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8592
      Top             =   2036
      Width           =   971
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8712
      X2              =   9414
      Y1              =   8037
      Y2              =   8037
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2205
      X2              =   2920
      Y1              =   8037
      Y2              =   8037
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2144
      Y2              =   8032
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX BILLING MAIN MENU"
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
      Left            =   2819
      TabIndex        =   12
      Top             =   1181
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1501
      Top             =   704
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2100
      Top             =   1903
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2219
      Top             =   2134
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8591
      Top             =   1904
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8712
      Top             =   2134
      Width           =   732
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuLinkPins 
         Caption         =   "Build Dummy Transactions for Link To Prop Pin"
      End
      Begin VB.Menu mnuLateNotice2Y 
         Caption         =   "Make All  Late Notice Flags 'Y'"
      End
      Begin VB.Menu mnuLateNotice2N 
         Caption         =   "Make All Late Notice Flags 'N'"
      End
      Begin VB.Menu mnuClipZip 
         Caption         =   "Clip Off Zip Hyphens"
      End
      Begin VB.Menu mnuMakeAllActive 
         Caption         =   "Make All Customers Active"
      End
      Begin VB.Menu mnuDateRepair 
         Caption         =   "Open Data Repair Screen"
      End
      Begin VB.Menu mnuCountyEdit 
         Caption         =   "Allow County Number Edit On/Off"
      End
   End
End
Attribute VB_Name = "frmVATaxMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAbstractMaint_Click()
  If LevelPass = 1 Then
    DelAbs = True
    AddCust = False
    EditCust = False
    THistRpt = False
    RPayEntry = False
    PPayEntry = False
    frmVATaxAbsMaint.Show
    DoEvents
    Unload Me
  ElseIf LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdCustMaint_Click()
'  Call ClearNegBalances
  If LevelPass = 1 Then
    If Not Exist(TaxSetupName) Then
      Call TaxMsg(900, "Please save data in the 'Tax System Setup' screen before continuing.")
      Exit Sub
    End If
    frmVATaxCustMaintMenu.Show
    DoEvents
    Unload Me
  ElseIf LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
  
End Sub

Private Sub cmdDMV_Click()
  If LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
  
  frmVATaxDMVMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPPTRA_Click()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisDate As Integer
  Dim ThatDate As Integer
  Dim ChangeDate$
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim x As Long
  
  If LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
  
  If Exist(TaxBillPostDateFile) Then
    OpenBillPostDateFile PostHandle, NumOfPostRecs
    For x = 1 To NumOfPostRecs
      Get PostHandle, x, PostRec
      If PostRec.BillType = "P" And PostRec.PPTRAPosted <> "Y" Then
        Exit For
      End If
    Next x
    Close
    If x > NumOfPostRecs Then
      Call TaxMsg(800, "There are no WINDOWS personal billing files saved. PPTRA Removal access denied.")
      Exit Sub
    End If
  Else
    Call TaxMsg(800, "There are no WINDOWS billing files saved. PPTRA Removal access denied.")
    Exit Sub
  End If
      
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle

  ThisDate = Date2Num(Date)
  ThatDate = TaxMasterRec.LawChngDate
  ChangeDate = MakeRegDate(ThatDate)
  
  If ThisDate < ThatDate Then
    Call TaxMsg(700, "The 'Date the Delinquent/Discount Reg Changes' date field on the System Setup screen indicates a date of " + ChangeDate + ". PPTRA removals cannot take place until that date. Access denied.")
    Exit Sub
  End If
  
  frmVATaxPPTRAMenu.Show
  DoEvents
  Unload Me
End Sub
Private Sub cmdEnterTaxPayments_Click()
  
  If Not Exist("TAXCUST.DAT") Then
    frmVATaxMsg.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("TAXSETUP.DAT") Then
    frmVATaxMsg.Label1.Caption = "Please complete the Tax Setup data before continuing."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Exit Sub
  End If
  
  'start here on 5/9
  If LevelPass = 1 Or LevelPass = 3 Then
    frmVATaxPayOperEntry.Show
    DoEvents
    Unload Me
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
  
End Sub

Private Sub cmdExit_Click()
  Call KillWaste
  MainLog ("Taxes.exe terminated via normal exit in Tax Billing Main Menu.")
  DoEvents
  Call Ready4others(PWcnt)
  DoEvents
  If Exist(QPTrim$(StartPath) + "\" + "Citipak.exe") Then
    Shell QPTrim$(StartPath) + "\" + "Citipak.exe", vbMaximizedFocus
  End If
  
'  Dim SubRptHandle3 As Integer
'  SubRptHandle3 = FreeFile

  Timer1.Enabled = True

End Sub

Private Sub cmdManualEntry_Click()
  If LevelPass = 1 Then
    frmVATaxManualBillMenu.Show
    DoEvents
    Unload Me
  ElseIf LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdSetUpAndUtil_Click()
  If LevelPass = 1 Then
    frmVATaxBillSetUpMenu.Show
    DoEvents
    Unload Me
  ElseIf LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdTaxAdvertising_Click()
  If LevelPass = 1 Then
    If Check4PayBatch("R") = True Then
      frmVATaxUnpostedPaylist.BillType = "R"
      frmVATaxUnpostedPaylist.Show vbModal
      DoEvents
      Exit Sub
    End If
    frmVATaxAdvColMenu.Show
    DoEvents
    Unload Me
  ElseIf LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdTaxBillAdj_Click()
  If LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    Exit Sub
  End If
  
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
    frmVATaxAdjustments.Show
    DoEvents
    Unload Me
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
    frmVATaxPAdjustments.Show
    DoEvents
    Unload Me
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    DoEvents
    Unload frmVATaxBillPostOpt
    Exit Sub
  End If
  
End Sub

Private Sub cmdTaxBillingFunctions_Click()
  If LevelPass = 1 Then
    frmVATaxBillingMenu.Show
    DoEvents
    Unload Me
  ElseIf LevelPass = 3 Then
    MsgBox "Your Password Does Not Allow Access To This Area.", vbOKOnly, "Access Denied"
  ElseIf LevelPass = 2 Then
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdTaxReportingSystem_Click()
  frmVATaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
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
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim cnt&, dl&
  Dim ThisDir$
  Dim CERec As AllowCountyEdit
  Dim CEHandle As Integer
  
'  frmVATaxDataRepair.Show 'pulls up DataRepair screen
  AddCust = False
  EditCust = False
  DelAbs = False
  THistRpt = False
  RPayEntry = False
  PPayEntry = False
  
  CurrCitiPath = App.Path
  
  If Mid(CurrCitiPath, Len(CurrCitiPath), 1) <> "\" Then
    CurrCitiPath = CurrCitiPath + "\"
  End If
  
  If PWcnt > 0 Then
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, CitiPass
    Close CitiPassFile
  End If
  
'  If CitiPass.Administ = True Or PWcnt = -3 Then
  If PWcnt = -3 Then
    mnuOptions.Visible = True
  Else
    mnuOptions.Visible = False
  End If
  
  Clipboard.Clear
  If App.PrevInstance Then
    ActivatePrevInstance 'don't want two payroll
    'programs open at once
  End If
  'the next series of code is used to get the
  'identity of the current clerk using payroll
  'and recorded anytime MainLog is accessed
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  
  If FromTX = False Then
    FromTX = True
    cmdExit.Enabled = False
  End If
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  
  'this saves the current path
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  Me.HelpContextID = hlpTaxBillingMain

  'only use these next two lines when working in the environment
  'comment out the rest of the time
'  LevelPass = 1
'  PWcnt = 6
'  OperNum = 12
  If Exist(CntyEditFile) Then 'added 7/11/07
    OpenCountyEditFile CEHandle
    Get CEHandle, 1, CERec
    Close CEHandle
    If Date2Num(Date) > CERec.AllowCountyEditXDate Then
      KillFile CntyEditFile
    End If
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxMainMenu.")
      Call Terminate
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If

End Sub

Private Sub mnuCountyEdit_Click()
  frmVATaxCountyNumEdit.Show vbModal
  DoEvents
End Sub

Private Sub mnuDateRepair_Click()
  frmVATaxDataRepair.Show
End Sub

Private Sub Timer1_Timer()
  Call Terminate2Shell 'closes all forms but does not clear password data
End Sub

Private Sub Timer2_Timer()
  cmdExit.Enabled = True
End Sub

Private Sub mnuLinkPins_Click()
  Dim TaxTranRec As TaxTransactionType
  Dim TaxTranHandle As Integer
  Dim NumOfTaxTranRecs As Long
  Dim x As Long, y As Long
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim ThisNumOfTaxTrans As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim Notified1 As Boolean
  Dim Notified2 As Boolean
  
  If TaxMsgWOpts(800, "Are you sure you want to link up real pins using dummy transactions?", "F10 Continue", "ESC Escape") = "abort" Then
    Exit Sub
  End If
  Notified1 = False
  Notified2 = False
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TaxTranHandle, NumOfTaxTranRecs
  frmVATaxShowPctComp.Label1 = "Linking Real Pins Procedure #1"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To NumOfTaxTranRecs
    Get TaxTranHandle, x, TaxTranRec
    If QPTrim$(TaxTranRec.Description) = "SoSoft: Link Real Pin" Then
        If Notified1 = True Then GoTo SkipIt
        frmVATaxShowPctComp.Hide
        If TaxMsgWOpts(800, "The linking of real pins to dummy transactions has already taken place. To continue anyway press F10. Otherwise, press ESC to escape.", "F10 Continue", "ESC Escape") = "abort" Then
          Unload frmVATaxShowPctComp
          Close
          Exit Sub
        Else
          Notified1 = True
          frmVATaxShowPctComp.Show
          MainLog ("User warned that the linking of real pins to dummy transactions has already taken place but elected to continue anyway.")
        End If
      ElseIf QPTrim$(TaxTranRec.RealPin) <> "0" Then
        If Notified2 = True Then GoTo SkipIt
        frmVATaxShowPctComp.Hide
        If TaxMsgWOpts(800, "Please be advised that transactions have been posted with real pin numbers stored.", "F10 Continue", "ESC Escape") = "abort" Then
          Unload frmVATaxShowPctComp
          Close
          Exit Sub
        Else
          Notified2 = True
          frmVATaxShowPctComp.Show
          MainLog ("User advised that transactions have been posted with real pin numbers stored and elected to continue anyway.")
        End If
      End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTaxTranRecs
  Next x
  Unload frmVATaxShowPctComp
  
  frmVATaxShowPctComp.Label1 = "Linking Real Pins Procedure #2"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  ThisNumOfTaxTrans = NumOfTaxTranRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  For x = 1 To NumOfRealRecs
    Get RHandle, x, RealPropRec
    If RealPropRec.Deleted = 0 Then
      If QPTrim$(RealPropRec.RealPin) <> "" Then
        Get TCHandle, RealPropRec.CustPin, TaxCust
        ThisNumOfTaxTrans = ThisNumOfTaxTrans + 1
        TaxTranRec.CustPin = RealPropRec.CustPin
        TaxTranRec.Posted2GL = "Y"
        TaxTranRec.RealPin = RealPropRec.RealPin
        TaxTranRec.Description = "SoSoft: Link Real Pin"
        TaxTranRec.Amount = 0
        TaxTranRec.BelongTo = 0
        TaxTranRec.BillType = "R"
        TaxTranRec.CntyPara = ""
        TaxTranRec.CustomerRec = RealPropRec.CustPin
        TaxTranRec.CyclPara = ""
        TaxTranRec.DiscAmt = 0
        TaxTranRec.DiscXDate = 0
        TaxTranRec.DMVBatch = 0
        TaxTranRec.DMVSubmitted = "Y"
        TaxTranRec.FromPrePay = "N"
        TaxTranRec.InternalPin = 0
        TaxTranRec.LastTrans = TaxCust.LastTrans
        TaxCust.LastTrans = ThisNumOfTaxTrans
        Put TCHandle, RealPropRec.CustPin, TaxCust
        TaxTranRec.OperNum = 0
        TaxTranRec.Padding = ""
        TaxTranRec.PersPin = 0
        TaxTranRec.Revenue.Collection = 0
        TaxTranRec.Revenue.CollectionPd = 0
        TaxTranRec.Revenue.Future1 = 0
        TaxTranRec.Revenue.Future1Pd = 0
        TaxTranRec.Revenue.Future2 = 0
        TaxTranRec.Revenue.Future2Pd = 0
        TaxTranRec.Revenue.Interest = 0
        TaxTranRec.Revenue.InterestPd = 0
        TaxTranRec.Revenue.LateList = 0
        TaxTranRec.Revenue.LateListPd = 0
        TaxTranRec.Revenue.pad = ""
        TaxTranRec.Revenue.Penalty = 0
        TaxTranRec.Revenue.PenaltyPd = 0
        TaxTranRec.Revenue.PrePaidAmt = 0
        TaxTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(RealPropRec.CustPin, "R"))
        TaxTranRec.Revenue.PrePaidUsed = 0
        TaxTranRec.Revenue.Principle1 = 0
        TaxTranRec.Revenue.Principle1Pd = 0
        TaxTranRec.Revenue.Principle2 = 0
        TaxTranRec.Revenue.Principle2Pd = 0
        TaxTranRec.Revenue.Principle3 = 0
        TaxTranRec.Revenue.Principle3Pd = 0
        TaxTranRec.Revenue.Principle4 = 0
        TaxTranRec.Revenue.Principle4Pd = 0
        TaxTranRec.Revenue.Principle5 = 0
        TaxTranRec.Revenue.Principle5Pd = 0
        TaxTranRec.Revenue.RevOpt1 = 0
        TaxTranRec.Revenue.RevOpt1Pd = 0
        TaxTranRec.Revenue.RevOpt2 = 0
        TaxTranRec.Revenue.RevOpt2Pd = 0
        TaxTranRec.Revenue.RevOpt3 = 0
        TaxTranRec.Revenue.RevOpt3Pd = 0
        TaxTranRec.TaxYear = 0
        TaxTranRec.TransDate = Date2Num(Date)
        TaxTranRec.TranType = 31
        TaxTranRec.TShpPara = ""
        TaxTranRec.PPTRADisc = 0
        TaxTranRec.PPTRARmvl = 0
        TaxTranRec.PPTRARmvlDate = 0
        TaxTranRec.PPTRAVal = 0
        Put TaxTranHandle, ThisNumOfTaxTrans, TaxTranRec
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfRealRecs
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  
  Call Savemsg(900, "The linking has completed successfully")
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMainMenu", "mnuLinkPins_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Unload frmVATaxShowPctComp
    EnableCloseButton Me.hwnd, True
    Close
  
End Sub

Private Sub mnuLateNotice2Y_Click()
  Dim CustRec As TaxCustType
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim CHandle As Integer

  OpenTaxCustFile CHandle, NumOfCRecs
  For x = 1 To NumOfCRecs
    Get CHandle, x, CustRec
    CustRec.LateNotice = "Y"
    Put CHandle, x, CustRec
  Next x
  Close
  Call Savemsg(900, "All late notice flags have been set to 'Y'.")
End Sub
Private Sub mnuLateNotice2N_Click()
  Dim CustRec As TaxCustType
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim CHandle As Integer

  OpenTaxCustFile CHandle, NumOfCRecs
  For x = 1 To NumOfCRecs
    Get CHandle, x, CustRec
    CustRec.LateNotice = "N"
    Put CHandle, x, CustRec
  Next x
  Close
  Call Savemsg(900, "All late notice flags have been set to 'N'.")
End Sub

Private Sub mnuRepairTaxYears_Click()
  Dim x As Long
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxYear As Integer
  Dim YrCnt As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  If NumOfTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TranType <> 1 Then
      If TransRec.TaxYear = 0 Then
        If TransRec.BelongTo > 0 Then
          Get THandle, TransRec.BelongTo, TransRec
          TaxYear = TransRec.TaxYear
          Get THandle, x, TransRec
          TransRec.TaxYear = TaxYear
          Put THandle, x, TransRec
          YrCnt = YrCnt + 1
        End If
      End If
    End If
  Next x
  
  Close
  Call Savemsg(900, "A total of " + CStr(YrCnt) + " errant tax years were corrected successfully.")
  
End Sub

Private Sub mnuReconstructHistory_Click()
  Call ClearNegBalances
End Sub

Private Sub mnuMakeAllActive_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    TaxCust.Active = "Y"
    Put TCHandle, x, TaxCust
  Next x
  
  Close
  
  Call TaxMsg(900, "Finished.")
  
End Sub

Private Sub mnuClipZip_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim ThisZip$, ZipCnt As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "No customer records saved.")
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    ThisZip = QPTrim$(TaxCust.Zip)
    If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
      ThisZip = Mid(ThisZip, 1, 5)
      TaxCust.Zip = ThisZip
      Put TCHandle, x, TaxCust
      ZipCnt = ZipCnt + 1
    End If
  Next x
  
  Close
  Call TaxMsg(900, "A total of " + CStr(ZipCnt) + " zip codes have been clipped successfully.")
End Sub

