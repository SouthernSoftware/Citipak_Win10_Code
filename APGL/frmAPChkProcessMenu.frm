VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmAPChkProcessMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Check Processing"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   Icon            =   "frmAPChkProcessMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdRunOpenPays 
      Height          =   435
      Left            =   4305
      TabIndex        =   0
      Top             =   2355
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSelInvPay 
      Height          =   435
      Left            =   4305
      TabIndex        =   1
      Top             =   2910
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":0AB5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdChkPreRpt 
      Height          =   444
      Left            =   4308
      TabIndex        =   2
      Top             =   3456
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":0CA8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPostAPChks 
      Height          =   420
      Left            =   4308
      TabIndex        =   8
      Top             =   6648
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":0E96
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrnChksDet 
      Height          =   420
      Left            =   4308
      TabIndex        =   7
      Top             =   6120
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":107D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintChkRegister 
      Height          =   420
      Left            =   4308
      TabIndex        =   6
      Top             =   5592
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":126F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCancelPrintChk 
      Height          =   420
      Left            =   4308
      TabIndex        =   5
      Top             =   5076
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":145C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRestartChks 
      Height          =   420
      Left            =   4308
      TabIndex        =   4
      Top             =   4548
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":164E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintAPChecks 
      Height          =   444
      Left            =   4308
      TabIndex        =   3
      Top             =   3996
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":1838
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitAPChkMenu 
      Height          =   420
      Left            =   4308
      TabIndex        =   10
      Top             =   7704
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":1A20
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdVoidPostChk 
      Height          =   420
      Left            =   4308
      TabIndex        =   9
      Top             =   7176
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPChkProcessMenu.frx":1C11
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   4
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   5
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   9840
      X2              =   9840
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A/P CHECK PROCESSING MENU"
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
      Left            =   3120
      TabIndex        =   11
      Top             =   1440
      Width           =   5964
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3384
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3384
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3384
      X2              =   3384
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3216
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9696
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   996
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
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
      Height          =   5940
      Index           =   0
      Left            =   2520
      Top             =   2352
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5940
      Index           =   2
      Left            =   9000
      Top             =   2352
      Width           =   732
   End
End
Attribute VB_Name = "frmAPChkProcessMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim APLedgerRec As APLedger81RecType
Dim TPayList As TPayListType
Dim TPayList2 As TPayListType
Dim TPayListD As TPayListType
Dim Vendor As VendorRecType
Dim TPayNot As TPayNotListType
Dim T2Pay As TPayNotListType
Dim GLFundIdx As GLFundIndexType
Dim Over As clsTextBoxOverRider
Dim APCheck As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class

Private Sub cmdChkPreRpt_Click()
  If Exist("TPAYLIST.LST") Then
    frmPreAuditOption.Show 1, frmAPChkProcessMenu
    'cmdPrintAPChecks.SetFocus
  Else
    MsgBox "No Invoices Have Been Selected For Payment", vbOKOnly, "No Selection"
  End If
End Sub


Private Sub cmdPrnChksDet_Click()
  If Exist("APCHKINF.DAT") Then
    frmChkListOpt.Show 1, frmAPChkProcessMenu
  Else
    MsgBox "Checks Have Not Been Printed.", vbOKOnly, "No Selection"
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitAPChkMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbYes Then
        KillFile "APChk.opn"
        Call MainLog("Close via AP Chk Menu.")
        ClearInUse PWcnt
      Else
        Cancel = True
      End If
    End If
  End If
End Sub

Private Sub cmdExitAPChkMenu_Click()
  KillFile "APChk.opn"
  Call MainLog("Exit AP Chk Menu.")
  frmAPMainMenu.Show
  Unload frmAPChkProcessMenu
End Sub

Private Sub cmdPostAPChks_Click()
  If Exist("APCHKINF.DAT") Then
    frmPostMsg.Show
  Else
    MsgBox "There Are NO Checks To Post.", vbOKOnly, "No Checks"
  End If
End Sub

Private Sub cmdPrintAPChecks_Click()
  GetAPCheck APCheck
  If Exist("TPAYLIST.LST") Then
    If Exist("TPayList2.lst") Then
      If APCheck > 0 Then
        frmPrnAPChecks.Show
        Unload frmAPChkProcessMenu
      Else
        MsgBox "You Must Set Up the AP Check Code on GL User Setup Screen First.", vbOKOnly, "Missing Setup Info"
      End If
    Else
      MsgBox "You MUST Run the Check PreAudit Report before printing Checks.", vbOKOnly, "No PreAudit Report"
    End If
  Else
    MsgBox "No Invoices Have Been Selected For Payment", vbOKOnly, "No Selection"
  End If
End Sub

Private Sub cmdPrintChkRegister_Click()
  If Exist("APCHKINF.DAT") Then
      frmReportOpt.Show 1
      If rptopt = 1 Then
        PrintCheckListing
      ElseIf rptopt = 2 Then
        PrintCheckListing2
      End If
    Else
    MsgBox "Checks Have Not Been Printed.", vbOKOnly, "No Selection"
  End If
End Sub

Private Sub cmdRestartChks_Click()
  If Exist("APCHKINF.DAT") Then
    frmRePrnAPChecks.Show
    Unload frmAPChkProcessMenu
  Else
    MsgBox "There Are NO Checks To Reprint.", vbOKOnly, "No Checks"
  End If
End Sub

Private Sub cmdRunOpenPays_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    PrnOpenPays frmAPChkProcessMenu
  ElseIf rptopt = 2 Then
    PrnOpenPays2 frmAPChkProcessMenu
  End If
End Sub

Private Sub cmdSelInvPay_Click()
  If Exist("APCHKINF.DAT") Then
    If Exist("TPAYLIST.LST") Then
      frmWarning.Label1 = "Invoices Have Been Selected and Checks"
      frmWarning.Label6 = "Have Been Printed, But NOT POSTED!!"
      frmWarning.Label5.FontSize = 12
      frmWarning.Label5 = "If you continue this file will be Destroyed!!"
      frmWarning.Label4 = "If you are unsure what to do Please STOP!"
      frmWarning.Label2 = "Call Software Support for instructions."
      frmWarning.Show 1, Me
      Select Case frmWarning.nogo
      Case True  'ok=1 then no don't continue
        Exit Sub
      Case False
        Kill "APCHKINF.DAT"
        Kill "TPAYLIST.LST"
        If Exist("TPaylist2.lst") Then Kill "Tpaylist2.lst"
        If Exist("Tpaylistd.lst") Then Kill "TPaylistd.lst"
        If Exist("TPaynot.lst") Then Kill "TPaynot.lst"
        If Exist("T2Pay.lst") Then Kill "T2Pay.lst"
        frmSelectOPays.Show
        Unload frmAPChkProcessMenu
      End Select
    Else
      frmSelectOPays.Show
      Unload frmAPChkProcessMenu
    End If
  Else
    frmSelectOPays.Show
    Unload frmAPChkProcessMenu
  End If
End Sub

Private Sub cmdCancelPrintChk_Click()
  If Exist("APCHKINF.DAT") Then
    frmChkPrnCancel.Show
    Unload frmAPChkProcessMenu
    'frmVoidPrnMsg.Show 1
  Else
    MsgBox "There Are NO Checks To Cancel.", vbOKOnly, "No Checks"
  End If
End Sub

Private Sub cmdVoidPostChk_Click()
  frmAPCkVoid.Show
  Unload frmAPChkProcessMenu
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpChkProcess
'''  cmdRunOpenPays.HelpContextID = hlpOpenPayRep
'''  cmdPrintChkRegister.HelpContextID = hlpChkReg
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    'Me.Visible = True
    'Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitAPChkMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub

Public Sub PreAuditRpt(Pagechk As Boolean)
  Dim TPayListFile As Integer, PayListRecLen As Integer, P As String
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, cnt As Integer, FF As String, MaxLines As Integer
  Dim Dash As String, PrintFile2 As Integer, Dash2 As String, PageNum As Integer
  Dim NumFunds As Integer, APDistRecLen As Integer, VendorFile As Integer
  Dim PrintFile  As Integer, TPayCnt As Integer, NumVRecs As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VRecNum As Long
  Dim ChkCnt As Integer, Linecnt As Integer, Title As String
  Dim Page As String, TotalChkAmt As Double, VendTotal As Double
  Dim NextDist As Long, ThisFund As String, FundCnt As Integer
  Dim ToPrint As String, TPVend As String, TPInv As String, TPDist As String
  Dim TmpRecno As Long, TPayListdFile As Integer, PayNotRecLen As Integer
  Dim TNotcnt As Integer
  FF$ = Chr$(12)
  MaxLines = 55
  Dash$ = String$(80, "-")
  Dash2$ = String$(61, "-")
  ToPrint$ = ""
  Mid$(Dash2$, 1, 7) = Space$(7)
  PageNum = 0
  ')*&&(&(&(&(&(
  getnoPays
  '&^%&%&%&%&%&%&
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub
  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType2
  ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType
  DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  PayListRecLen = Len(TPayList2)
  TPayListFile = FreeFile
  Open "TPAYLIST2.LST" For Random Shared As TPayListFile Len = PayListRecLen
  TPayCnt = LOF(TPayListFile) \ 6
  If TPayCnt = 0 Then
    ActivateControls frmAPChkProcessMenu
    Close
    MsgBox "No Check will be generated for Payables selected.", vbOKOnly, "View Select Invoices"
    Exit Sub
  End If
  
  If TPayCnt > 0 Then
    FrmShowPctComp.Label1 = "Creating Pre-Audit Report"
    FrmShowPctComp.Show , Me
    DoEvents
    DeActivateControls frmAPChkProcessMenu
  End If

  ReDim TPayArray(1 To TPayCnt) As TPayListType
  For Pcnt = 1 To TPayCnt
    Get TPayListFile, Pcnt, TPayList2
    FrmShowPctComp.ShowPctComp Pcnt, TPayCnt
    TPayArray(Pcnt).VendorRecNum = TPayList2.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayList2.LedgerRecNum
  Next
  Close TPayListFile
  
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  PrintFile = FreeFile
  Open "APCHKREG.PRN" For Output As PrintFile
  PrintFile2 = FreeFile
  Open "APCHKFund.PRN" For Output As PrintFile2

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  
  Get VendorFile, TPayArray(1).VendorRecNum, Vendor
  Get APLedgerFile, TPayArray(1).LedgerRecNum, APLedgerRec(1)

  'GoSub PADoChkRegHeader
  TmpRecno = TPayArray(1).LedgerRecNum
  GoSub PrintVendHeader
  'GoSub PAInvHeader
  GoSub PrintDist
  VRecNum& = TPayArray(1).VendorRecNum
  ChkCnt = 1
  For cnt = 2 To TPayCnt
  FrmShowPctComp.ShowPctComp cnt, TPayCnt
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    ActivateControls frmAPChkProcessMenu
    Unload FrmShowPctComp
    GoTo CancelExit
  End If
    If VRecNum& <> TPayArray(cnt).VendorRecNum Then
      GoSub FinishVendor
'      If Linecnt > MaxLines Then
'        Print #PrintFile, FF$
'        GoSub PADoChkRegHeader
'      End If
      Get VendorFile, TPayArray(cnt).VendorRecNum, Vendor
      GoSub PrintVendHeader
      'GoSub PAInvHeader
      VRecNum& = TPayArray(cnt).VendorRecNum
      ChkCnt = ChkCnt + 1
    End If
    Get APLedgerFile, TPayArray(cnt).LedgerRecNum, APLedgerRec(1)
    TmpRecno = TPayArray(cnt).LedgerRecNum
    GoSub PrintDist
'    If Linecnt > MaxLines Then
'      Print #PrintFile, FF$
'      GoSub PADoChkRegHeader
'    End If
  Next
  GoSub FinishVendor
  
 ' Print #PrintFile, FF$
  'print fund stuff
  'Linecnt = 1
  'GoSub PADoFundHeader
  For cnt = 1 To NumFunds
    If FundAmts(cnt) <> 0 Then
'      If Linecnt > MaxLines Then
'        Print #PrintFile, FF$
'        GoSub PADoFundHeader
'      End If
      ToPrint$ = Using("#,###", Str$(Val(FundList$(cnt)))) + "~" + Using("##,###,###.##", Str$(FundAmts(cnt)))
      Print #PrintFile2, ToPrint$
      'Linecnt = Linecnt + 1
    End If
  Next
  'Print #PrintFile, FF$
  'Close
  
 ' Erase FundList$, FundAmts, APLedgerRec, APDistRec
  'Erase LedInfo, DistInfo, ChkRegInfo           ', ChkInfo
  PayListRecLen = Len(TPayListD)
  TPayListdFile = FreeFile
  Open "TPayListD.lst" For Random Shared As TPayListdFile Len = PayListRecLen
  TNotcnt = LOF(TPayListdFile) \ PayListRecLen
  If TNotcnt = 0 Then
    Close TPayListdFile
  Else

  ReDim TPayArray(1 To TNotcnt) As TPayListType
  For Pcnt = 1 To TNotcnt
    Get TPayListdFile, Pcnt, TPayListD
    TPayArray(Pcnt).VendorRecNum = TPayListD.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayListD.LedgerRecNum
  Next
  Close TPayListdFile
    
  Get VendorFile, TPayArray(1).VendorRecNum, Vendor
  Get APLedgerFile, TPayArray(1).LedgerRecNum, APLedgerRec(1)
  'GoSub PADoChkRegHeader
  TmpRecno = TPayArray(1).LedgerRecNum
  TPVend$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME) + "-No Check Will Be Generated"
  'GoSub PAInvHeader
  GoSub PrintDist2
  VRecNum& = TPayArray(1).VendorRecNum
  'ChkCnt = 1
  For cnt = 2 To TNotcnt
  FrmShowPctComp.ShowPctComp cnt, TNotcnt
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    ActivateControls frmAPChkProcessMenu
    Unload FrmShowPctComp
    GoTo CancelExit
  End If
    If VRecNum& <> TPayArray(cnt).VendorRecNum Then
      GoSub FinishVendor
'      If Linecnt > MaxLines Then
'        Print #PrintFile, FF$
'        GoSub PADoChkRegHeader
'      End If
      Get VendorFile, TPayArray(cnt).VendorRecNum, Vendor
      TPVend$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME) + "-No Check Will Be Generated"
      'GoSub PAInvHeader
      VRecNum& = TPayArray(cnt).VendorRecNum
      'ChkCnt = ChkCnt + 1
    End If
    Get APLedgerFile, TPayArray(cnt).LedgerRecNum, APLedgerRec(1)
    TmpRecno = TPayArray(cnt).LedgerRecNum
    GoSub PrintDist2
'    If Linecnt > MaxLines Then
'      Print #PrintFile, FF$
'      GoSub PADoChkRegHeader
'    End If
  Next
  GoSub FinishVendor
  
 ' Print #PrintFile, FF$
  'print fund stuff
  'Linecnt = 1
  'GoSub PADoFundHeader
'  For cnt = 1 To NumFunds
'    If FundAmts(cnt) <> 0 Then
''      If Linecnt > MaxLines Then
''        Print #PrintFile, FF$
''        GoSub PADoFundHeader
''      End If
'      ToPrint$ = Using("#,###", Str$(Val(FundList$(cnt)))) + "~" + Using("##,###,###.##", Str$(FundAmts(cnt)))
'      Print #PrintFile2, ToPrint$
'      'Linecnt = Linecnt + 1
'    End If
'  Next
  'Print #PrintFile, FF$
  End If
  Close
  Erase FundList$, FundAmts, APLedgerRec, APDistRec
  Erase LedInfo, DistInfo, ChkRegInfo           ', ChkInfo
 
  Title$ = "Check Pre-Audit Report"
  Call MainLog("ChkPreAudit Rpt.")
  ActivateControls frmAPChkProcessMenu
  Load frmLoadingRpt
  If Pagechk = True Then
    ARptPreAudit.GroupFooter1.NewPage = ddNPAfter
  End If
  ARptPreAudit.totvends = ChkCnt
  ARptPreAudit.GetName "APCHKREG.PRN", "APCHKFund.PRN"
  ARptPreAudit.txtTown.Caption = GLUserName$
  ARptPreAudit.txtDate.Caption = Now
  ARptPreAudit.RptVendTot.DataValue = TotalChkAmt#
  ARptPreAudit.Label1.Caption = Title$
  ARptPreAudit.startrpt

  'ViewPrint "APCHKREG.PRN", title$
ExitPreAudit:
  Exit Sub

'PADoFundHeader:
'  PageNum = PageNum + 1
'  Page$ = Using("###", Str$(PageNum))
'  Print #PrintFile, "Check Pre-Audit Report Summary                                  Page:" + Page$
'  Print #PrintFile, Dash$
'  Print #PrintFile,
'  Print #PrintFile, "    Checks to Print        "; Using("#,###", Str$(ChkCnt))
'  Print #PrintFile, "           Totaling  "; Using("$###,###,###.##", Str$(TotalChkAmt#))
'  Print #PrintFile,
'  Print #PrintFile, " By Fund:"
'  Linecnt = 7
'  Return

'PADoChkRegHeader:
'  PageNum = PageNum + 1
'  Page$ = Using("###", Str$(PageNum))
'  Print #PrintFile, "A/P Check Pre-Audit Report                                      Page:" + Page$
'  Print #PrintFile, "Run Date: " + Date$
'  Print #PrintFile,
'  Print #PrintFile, Dash$
'  Linecnt = 4
'  'GoSub PrintVendHeader
'  'GoSub PAInvHeader
'
'Return

'PAInvHeader:
'  Print #PrintFile,
'  Print #PrintFile, "Inv Date    Due Date    Inv Num                     PO            Amount"
'  Print #PrintFile, "----------  ----------  -------------------------  ------------ ----------------"
'  Linecnt = Linecnt + 2
'Return
PrintDist:
  TPInv$ = ""
  LSet ChkRegInfo(1).VendName = Vendor.VNAME
  
  LSet LedInfo(1).InvDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).DueDate = Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).InvNum = QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim$(APLedgerRec(1).Comment)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).PONum, 10)
  Else
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).MPONum, 10)
  End If
  P$ = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
 'RSet LedInfo(1).Amt = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  TPInv$ = Str(TmpRecno) + "~" + LedInfo(1).InvNum + "~" + LedInfo(1).InvDate + "~" + LedInfo(1).DueDate + "~"
  TPInv$ = TPInv$ + LedInfo(1).PONum + "~" + P$

  'Linecnt = Linecnt + 2
  NextDist& = APLedgerRec(1).FrstDist

'  Print #PrintFile, Tab(50); "Dist Acct        Dist Amount"
'  Linecnt = Linecnt + 1

  Do Until NextDist& = 0
    Get APDistFile, NextDist&, APDistRec(1)
'    If Linecnt > MaxLines Then
'      Print #PrintFile, FF$
'      GoSub PADoChkRegHeader
'    End If
    LSet LedInfo(1).InvDate = ""
    LSet LedInfo(1).DueDate = ""
    LSet LedInfo(1).InvNum = ""
    LSet LedInfo(1).PONum = ""
    RSet LedInfo(1).Amt = ""

    LSet LedInfo(1).DistAcct = APDistRec(1).DistAcctNum
    'LSet LedInfo(1).DistAmt = Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))
    P$ = Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))

    'PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; DistInfo(1).Di
    'PRINT #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvN
    TPDist$ = ""
    'Linecnt = Linecnt + 1
    ThisFund$ = Left$(APDistRec(1).DistAcctNum, GLFundLen)
    TPDist$ = LedInfo(1).DistAcct + "~" + P$ + "~" + ThisFund$
    ToPrint$ = TPVend$ + "~" + TPInv$ + "~" + TPDist$
    Print #PrintFile, ToPrint$
    ToPrint$ = ""
'    ThisFund$ = ""
'    TPDist$ = ""
'    P$ = ""
    For FundCnt = 1 To NumFunds
      If ThisFund$ = FundList$(FundCnt) Then
        FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
        Exit For
      End If
    Next
    NextDist& = APDistRec(1).NextDist
  Loop
'  Print #PrintFile, ""
'  Linecnt = Linecnt + 1
Return
PrintDist2:
  TPInv$ = ""
  LSet ChkRegInfo(1).VendName = Vendor.VNAME
  
  LSet LedInfo(1).InvDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).DueDate = Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).InvNum = QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim$(APLedgerRec(1).Comment)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).PONum, 10)
  Else
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).MPONum, 10)
  End If
  P$ = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
 'RSet LedInfo(1).Amt = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  'TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  TPInv$ = Str(TmpRecno) + "~" + LedInfo(1).InvNum + "~" + LedInfo(1).InvDate + "~" + LedInfo(1).DueDate + "~"
  TPInv$ = TPInv$ + LedInfo(1).PONum + "~" + P$

  'Linecnt = Linecnt + 2
  NextDist& = APLedgerRec(1).FrstDist

'  Print #PrintFile, Tab(50); "Dist Acct        Dist Amount"
'  Linecnt = Linecnt + 1

  Do Until NextDist& = 0
    Get APDistFile, NextDist&, APDistRec(1)
'    If Linecnt > MaxLines Then
'      Print #PrintFile, FF$
'      GoSub PADoChkRegHeader
'    End If
    LSet LedInfo(1).InvDate = ""
    LSet LedInfo(1).DueDate = ""
    LSet LedInfo(1).InvNum = ""
    LSet LedInfo(1).PONum = ""
    RSet LedInfo(1).Amt = ""

    LSet LedInfo(1).DistAcct = APDistRec(1).DistAcctNum
    'LSet LedInfo(1).DistAmt = Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))
    P$ = Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))

    'PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; DistInfo(1).Di
    'PRINT #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvN
    TPDist$ = ""
    'Linecnt = Linecnt + 1
    ThisFund$ = Left$(APDistRec(1).DistAcctNum, GLFundLen)
    TPDist$ = LedInfo(1).DistAcct + "~" + P$ + "~" + ThisFund$
    ToPrint$ = TPVend$ + "~" + TPInv$ + "~" + TPDist$
    Print #PrintFile, ToPrint$
    ToPrint$ = ""
'    ThisFund$ = ""
'    TPDist$ = ""
'    P$ = ""

'    For FundCnt = 1 To NumFunds
'      If ThisFund$ = FundList$(FundCnt) Then
'        FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
'        Exit For
'      End If
'    Next
    NextDist& = APDistRec(1).NextDist
  Loop
'  Print #PrintFile, ""
'  Linecnt = Linecnt + 1
Return

PrintVendHeader:
  'PRINT #PrintFile, Dash$
  TPVend$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME)
  'Linecnt = Linecnt + 1
Return

FinishVendor:
'  Print #PrintFile, Dash$
'  Print #PrintFile, Vendor.vnum; Vendor.VNAME; Tab(50); "Vendor Total: "; Tab(65); Using("$###,###,###.##", Str$(VendTotal#))
'  Print #PrintFile, Dash$
'  Print #PrintFile, ""
  VendTotal# = 0
  TPVend$ = ""
'  If Pagechk = True Then
'    Linecnt = 56
'  Else
'    Linecnt = Linecnt + 4
'  End If
Return
CancelExit:
  Exit Sub
End Sub


Private Sub PrintCheckListing()
  Dim Dash As String, FF As String, ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer
  Dim VendorFile As Integer, NumVRecs As Long, PrintFile As Integer
  Dim Cnt2 As Long, TCheckAmt As Double, Title As String, Temp As Integer
  Dim low As Long, High As Long, ToPrint As String
  If Not Exist("APCHKINF.DAT") Then Exit Sub
  FrmShowPctComp.Label1 = "Creating Check Listing Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmAPChkProcessMenu
  Dash$ = String$(78, "-")
  FF$ = Chr$(12)
  ChkinfoFile = FreeFile
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  ChkInfoRecLen = Len(CHKinfo(1))
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  ReDim CHKinfo(1 To VCnt) As CheckInfoType3
  'FGetAH "APCHKINF.DAT", Chkinfo(1), ChkInfoRecLen, VCnt
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For Temp = 1 To VCnt
    Get ChkinfoFile, Temp, CHKinfo(Temp)
  Next
  low = LBound(CHKinfo)
  High = UBound(CHKinfo)
  QCkSort CHKinfo(), low, High
  OpenVendorFile VendorFile, NumVRecs
  PrintFile = FreeFile
  Open "APCHKLST.PRN" For Output As #PrintFile

'  Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).ChkDate), "12-31-1979"), "mm/dd/yyyy")
'  Print #PrintFile,
'  Print #PrintFile, " Check No.         Description"
'  Print #PrintFile, Dash$

  For cnt = 1 To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPChkProcessMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2& = CHKinfo(cnt).StartChk To CHKinfo(cnt).LastChk
      'Print #PrintFile, Using("#######", Str$(Cnt2&));
      ToPrint$ = Using("#######", Str$(Cnt2&)) + "~"
      If CHKinfo(cnt).VoidFlag Then
        'Print #PrintFile, Tab(20); Vendor.VNAME; Tab(50); "  CANCELED BY USER"
        ToPrint$ = ToPrint$ + QPTrim(Vendor.VNAME) + "~  CANCELED BY USER"
      ElseIf Cnt2& < CHKinfo(cnt).LastChk Then
        'Print #PrintFile, Tab(20); Vendor.VNAME; Tab(52); "            VOID"
        ToPrint$ = ToPrint$ + QPTrim(Vendor.VNAME) + "~            VOID"
      Else
        TCheckAmt# = Round(TCheckAmt# + CHKinfo(cnt).ChkAmt)
        'Print #PrintFile, Tab(20); Vendor.VNAME; Tab(54); Using("$##,###,###.##", Str$(CHKinfo(cnt).ChkAmt))
        ToPrint$ = ToPrint$ + QPTrim(Vendor.VNAME) + "~" + Using("$##,###,###.##", Str$(CHKinfo(cnt).ChkAmt))
      End If
      Print #PrintFile, ToPrint$
    Next
  Next
  
'  Print #PrintFile, Dash$
'  Print #PrintFile, VCnt & " Checks Totaling"; Tab(54); Using("$##,###,###.##", Str$(TCheckAmt#))
'  Print #PrintFile, FF$
  Close
  Title$ = "A/P Check Listing"
  Call MainLog("Print AP Chk List.")
  ActivateControls frmAPChkProcessMenu
  'ViewPrint "APCHKLST.PRN", title$
  Load frmLoadingRpt
  ARptAPChkListing.GetName "APCHKLST.PRN"
  ARptAPChkListing.txtTown.Caption = GLUserName$
  ARptAPChkListing.txtDate.Caption = Now
  ARptAPChkListing.Label1.Caption = Title$
  ARptAPChkListing.txtTot = VCnt
  ARptAPChkListing.startrpt

CancelExit:
  Exit Sub
ExitCheckListing:

End Sub
Public Sub PrintChkListDist(RptBreak As Boolean)
  Dim TPDist As String, ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer
  Dim VendorFile As Integer, NumVRecs As Long, PrintFile As Integer
  Dim Cnt2 As Long, TCheckAmt As Double, Title As String, Temp As Integer
  Dim low As Long, High As Long, cntven As Integer
  Dim TPayListFile As Integer, PayListRecLen As Integer, P As String
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, ToPrint As String, TPSum As String
  Dim NumFunds As Integer, APDistRecLen As Integer
  Dim TPayCnt As Integer, TPCk As String, TPInv As String
  Dim APDistFile As Integer, NumDistRecs As Long, VRecNum As Long
  Dim ChkCnt As Integer, Linecnt As Integer
  Dim Page As String, TotalChkAmt As Double, VendTotal As Double
  Dim NextDist As Long, ThisFund As String, FundCnt As Integer
  'FF$ = Chr$(12)
  'MaxLines = 55
  'Dash$ = String$(80, "-")
  'PageNum = 0
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub
  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType2
  ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType
  DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  PayListRecLen = Len(TPayList)
  TPayListFile = FreeFile
  Open "TPAYLIST2.LST" For Random Shared As TPayListFile Len = PayListRecLen
  TPayCnt = LOF(TPayListFile) \ 6
  If TPayCnt = 0 Then
    Exit Sub
  End If
  ReDim TPayArray(1 To TPayCnt) As TPayListType
  For Pcnt = 1 To TPayCnt
    Get TPayListFile, Pcnt, TPayList
    TPayArray(Pcnt).VendorRecNum = TPayList.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayList.LedgerRecNum
  Next
  Close TPayListFile
  If Not Exist("APCHKINF.DAT") Then Exit Sub
  FrmShowPctComp.Label1 = "Creating Check Listing Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmAPChkProcessMenu
  ChkinfoFile = FreeFile
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  ChkInfoRecLen = Len(CHKinfo(1))
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  ReDim CHKinfo(1 To VCnt) As CheckInfoType3
  'FGetAH "APCHKINF.DAT", Chkinfo(1), ChkInfoRecLen, VCnt
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For Temp = 1 To VCnt
    Get ChkinfoFile, Temp, CHKinfo(Temp)
  Next
  low = LBound(CHKinfo)
  High = UBound(CHKinfo)
  QCkSort CHKinfo(), low, High
  OpenVendorFile VendorFile, NumVRecs
  PrintFile = FreeFile
  Open "APCHKDet.PRN" For Output As #PrintFile
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
'  If RptBreak = False Then
'    Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).ChkDate), "12-31-1979"), "mm/dd/yyyy")
'    Linecnt = 1
'  End If
  For cnt = 1 To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPChkProcessMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    TPCk$ = ""
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2& = CHKinfo(cnt).StartChk To CHKinfo(cnt).LastChk
'      'GoSub Printheadstuff
'        Print #PrintFile, " Check No.         Description                                Check Amount"
'        Print #PrintFile, Dash$
'        Linecnt = Linecnt + 2
'      Print #PrintFile, ;
      TPSum$ = " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).chkdate), "12-31-1979"), "mm/dd/yyyy")
      TPCk$ = Using("#######", Str$(Cnt2&)) + "~"
      If CHKinfo(cnt).VoidFlag Then
        TPCk$ = TPCk$ + QPTrim(Vendor.VNAME) + "~" + "  CANCELED BY USER"
      ElseIf Cnt2& < CHKinfo(cnt).LastChk Then
        TPCk$ = TPCk$ + QPTrim(Vendor.VNAME) + "~" + "             VOID"
      Else
        
        TCheckAmt# = Round(TCheckAmt# + CHKinfo(cnt).ChkAmt)
        TPCk$ = TPCk$ + QPTrim(Vendor.VNAME) + "~" + Using("$##,###,###.##", Str$(CHKinfo(cnt).ChkAmt))
      End If
    Next
      'GoSub PrintInvHeader
      For cntven = 1 To TPayCnt
      VRecNum& = TPayArray(cntven).VendorRecNum

        If VRecNum& = CHKinfo(cnt).VendorRecNum Then
        
          Get APLedgerFile, TPayArray(cntven).LedgerRecNum, APLedgerRec(1)
'          If Linecnt >= MaxLines Then
'            GoSub Printheadstuff
'            GoSub PrintInvHeader
'          End If

          GoSub PrintDist
        End If
      Next
    'Print #PrintFile, Tab(30); "-------**********-------"
  Next
'******Here we go totals and close*************
'  If RptBreak = True Or Linecnt >= MaxLines - 4 Then
'    Print #PrintFile, FF$
'  End If
'  Print #PrintFile, ""
'  Print #PrintFile, Dash$
'  Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).ChkDate), "12-31-1979"), "mm/dd/yyyy")
'  Print #PrintFile, VCnt; "Checks Totaling - "; Tab(40); Using("$##,###,###.##", Str$(TCheckAmt#))
'  Print #PrintFile, FF$
  Close
  Erase FundList$, FundAmts, APLedgerRec, APDistRec
  Erase LedInfo, DistInfo, ChkRegInfo           ', ChkInfo

  Title$ = "A/P Check Report"
  Call MainLog("Print AP Chk List w/dist.")
  ActivateControls frmAPChkProcessMenu
  Load frmLoadingRpt
  'ViewPrint "APCHKDet.PRN", title$
  If RptBreak = True Then
    ARptCkListDist.GroupFooter1.NewPage = ddNPAfter
  End If
    
  ARptCkListDist.GetName "APCHKDet.PRN"
  ARptCkListDist.txtTown.Caption = GLUserName$
  ARptCkListDist.txtDate.Caption = Now
  ARptCkListDist.Label1.Caption = Title$
  ARptCkListDist.Label15.Caption = TPSum$
  ARptCkListDist.totchks = Using("$##,###,###.##", Str$(TCheckAmt#))
  ARptCkListDist.txtTot = VCnt
  ARptCkListDist.startrpt

  Exit Sub
PrintDist:
  TPInv$ = ""
  'LSet ChkRegInfo(1).VENDNAME = Vendor.VNAME
  TPInv$ = Str(TPayArray(cntven).LedgerRecNum) + "~" + Left$(APLedgerRec(1).DOCNum, 20) + "~"
  TPInv$ = TPInv$ + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  TPInv$ = TPInv$ + "~" + Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    TPInv$ = TPInv$ + "~" + Left$(APLedgerRec(1).PONum, 10)
  Else
    TPInv$ = TPInv$ + "~" + Left$(APLedgerRec(1).MPONum, 10)
  End If
  TPInv$ = TPInv$ + "~" + Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
  
 'RSet LedInfo(1).Amt = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  'Print #PrintFile, ToPrint$

  'Linecnt = Linecnt + 2
  NextDist& = APLedgerRec(1).FrstDist
'    If Linecnt > MaxLines - 2 Then
'      GoSub Printheadstuff
'    End If

  'Print #PrintFile, Tab(44); "Distributions"
  'Linecnt = Linecnt + 1

  Do Until NextDist& = 0
    ToPrint$ = Space(80)
    TPDist$ = ""
    Get APDistFile, NextDist&, APDistRec(1)
    TPDist$ = QPTrim(APDistRec(1).DistAcctNum) + "~"
    TPDist$ = TPDist$ + Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))
    ToPrint$ = TPCk$ + "~" + TPInv$ + "~" + TPDist$
    Print #PrintFile, ToPrint$
    'Linecnt = Linecnt + 1
    NextDist& = APDistRec(1).NextDist
  Loop
 ' Print #PrintFile, ""
 ' Linecnt = Linecnt + 1
Return
'Printheadstuff:
'If RptBreak = True Or Linecnt >= MaxLines - 2 Then
'  If cnt <> 1 Then
'  Print #PrintFile, FF$
'  End If
'  Linecnt = 1
'  Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).ChkDate), "12-31-1979"), "mm/dd/yyyy")
'End If
'  'Print #PrintFile,
'Return
'PrintInvHeader:
'  If Linecnt >= MaxLines - 3 Then
'    GoSub Printheadstuff
'  End If
'  Print #PrintFile,
'  Print #PrintFile, "   Inv Date    Due Date    Inv Num                  PO           Inv Amount"
'  Print #PrintFile, "   ----------  ----------  -------------------- ------------ ---------------"
'  Linecnt = Linecnt + 2
'Return

CancelExit:
  Exit Sub
ExitCheckListing:

End Sub
Public Sub PreAuditRpt2(Pagechk As Boolean)
  Dim TPayListFile As Integer, PayListRecLen As Integer, P As String
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, cnt As Integer, FF As String, MaxLines As Integer
  Dim Dash As String, Dash2 As String, PageNum As Integer
  Dim NumFunds As Integer, APDistRecLen As Integer, VendorFile As Integer
  Dim PrintFile  As Integer, TPayCnt As Integer, NumVRecs As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VRecNum As Long
  Dim ChkCnt As Integer, Linecnt As Integer, Title As String
  Dim Page As String, TotalChkAmt As Double, VendTotal As Double
  Dim NextDist As Long, ThisFund As String, FundCnt As Integer
  Dim TmpRecno As Long, TPayNotFile As Integer, PayNotRecLen As Integer
  Dim TNotcnt As Integer, onlyneg As Boolean
  Dim TPayListdFile As Integer
  FF$ = Chr$(12)
  MaxLines = 55
  Dash$ = String$(80, "-")
  Dash2$ = String$(61, "-")
  Mid$(Dash2$, 1, 7) = Space$(7)
  PageNum = 0
  getnoPays
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub
  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType2
  ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType
  DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  PayListRecLen = Len(TPayList)
  TPayListFile = FreeFile
  Open "TPAYLIST2.LST" For Random Shared As TPayListFile Len = PayListRecLen
  TPayCnt = LOF(TPayListFile) \ 6
  If TPayCnt = 0 Then
    Close
    MsgBox "No Check will be generated for Payables selected.", vbOKOnly, "View Select Invoices"
    Exit Sub
  End If
  If TPayCnt > 0 Then
    FrmShowPctComp.Label1 = "Creating Pre-Audit Report"
    FrmShowPctComp.Show , Me
    DoEvents
    DeActivateControls frmAPChkProcessMenu
  End If

  ReDim TPayArray(1 To TPayCnt) As TPayListType
  For Pcnt = 1 To TPayCnt
    Get TPayListFile, Pcnt, TPayList2
    TPayArray(Pcnt).VendorRecNum = TPayList2.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayList2.LedgerRecNum
  Next
  Close TPayListFile
  
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  PrintFile = FreeFile
  Open "APCHKREG.PRN" For Output As PrintFile

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  
  Get VendorFile, TPayArray(1).VendorRecNum, Vendor
  Get APLedgerFile, TPayArray(1).LedgerRecNum, APLedgerRec(1)

  GoSub PADoChkRegHeader
  GoSub PrintVendHeader
  GoSub PAInvHeader
  GoSub PrintDist
  VRecNum& = TPayArray(1).VendorRecNum
  ChkCnt = 1
  FrmShowPctComp.ShowPctComp ChkCnt, TPayCnt
  For cnt = 2 To TPayCnt
  FrmShowPctComp.ShowPctComp cnt, TPayCnt
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    ActivateControls frmAPChkProcessMenu
    Unload FrmShowPctComp
    GoTo CancelExit
  End If
    If VRecNum& <> TPayArray(cnt).VendorRecNum Then
      GoSub FinishVendor
      If Linecnt > MaxLines Then
        Print #PrintFile, FF$
        GoSub PADoChkRegHeader
      End If
      Get VendorFile, TPayArray(cnt).VendorRecNum, Vendor
      GoSub PrintVendHeader
      GoSub PAInvHeader
      VRecNum& = TPayArray(cnt).VendorRecNum
      ChkCnt = ChkCnt + 1
    End If
    Get APLedgerFile, TPayArray(cnt).LedgerRecNum, APLedgerRec(1)
    GoSub PrintDist
    If Linecnt > MaxLines Then
      Print #PrintFile, FF$
      GoSub PADoChkRegHeader
    End If
  Next
  GoSub FinishVendor
  Print #PrintFile, FF$
  'print fund stuff
  Linecnt = 1
'  GoSub PADoFundHeader
'  For cnt = 1 To NumFunds
'    If FundAmts(cnt) <> 0 Then
'      If Linecnt > MaxLines Then
'        Print #PrintFile, FF$
'        GoSub PADoFundHeader
'      End If
'      Print #PrintFile, "      Fund "; Using("#,###", Str$(Val(FundList$(cnt)))); "  Amt  "; Using("##,###,###.##", Str$(FundAmts(cnt)))
'      Linecnt = Linecnt + 1
'    End If
'  Next
 ' Print #PrintFile, FF$
  'Close
  PayListRecLen = Len(TPayListD)
  TPayListdFile = FreeFile
  Open "TPayListD.lst" For Random Shared As TPayListdFile Len = PayListRecLen
  TNotcnt = LOF(TPayListdFile) \ PayListRecLen
  If TNotcnt = 0 Then
    Close TPayListdFile
  Else
  onlyneg = True
  'Erase TPayArray
  ReDim TPayArray(1 To TNotcnt) As TPayListType
  For Pcnt = 1 To TNotcnt
    Get TPayListdFile, Pcnt, TPayListD
    TPayArray(Pcnt).VendorRecNum = TPayListD.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayListD.LedgerRecNum
  Next
  Close TPayListdFile
    
  Get VendorFile, TPayArray(1).VendorRecNum, Vendor
  Get APLedgerFile, TPayArray(1).LedgerRecNum, APLedgerRec(1)
 ' GoSub PADoChkRegHeader
  GoSub PrintVendHeader
  GoSub PAInvHeader
  GoSub PrintDist
  VRecNum& = TPayArray(1).VendorRecNum
  'ChkCnt = 1
  For cnt = 2 To TNotcnt
  FrmShowPctComp.ShowPctComp cnt, TNotcnt
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    ActivateControls frmAPChkProcessMenu
    Unload FrmShowPctComp
    GoTo CancelExit
  End If
    If VRecNum& <> TPayArray(cnt).VendorRecNum Then
      GoSub FinishVendor
      If Linecnt > MaxLines Then
        Print #PrintFile, FF$
        GoSub PADoChkRegHeader
      End If
      Get VendorFile, TPayArray(cnt).VendorRecNum, Vendor
      GoSub PrintVendHeader
      GoSub PAInvHeader
      VRecNum& = TPayArray(cnt).VendorRecNum
      'ChkCnt = ChkCnt + 1
    End If
    Get APLedgerFile, TPayArray(cnt).LedgerRecNum, APLedgerRec(1)
    GoSub PrintDist
    If Linecnt > MaxLines Then
      Print #PrintFile, FF$
      GoSub PADoChkRegHeader
    End If
  Next
  GoSub FinishVendor
  Print #PrintFile, FF$
  'print fund stuff
  Linecnt = 1
  End If
  GoSub PADoFundHeader
  For cnt = 1 To NumFunds
    If FundAmts(cnt) <> 0 Then
      If Linecnt > MaxLines Then
        Print #PrintFile, FF$
        GoSub PADoFundHeader
      End If
      Print #PrintFile, "      Fund "; Using("#,###", Str$(Val(FundList$(cnt)))); "  Amt  "; Using("##,###,###.##", Str$(FundAmts(cnt)))
      Linecnt = Linecnt + 1
    End If
  Next
  If onlyneg Then
    Print #PrintFile,
    Print #PrintFile, " ** Indicates A Check will not be generated and Invoices will remain open."
  End If
  Print #PrintFile, FF$
  
  Close
  
  Erase FundList$, FundAmts, APLedgerRec, APDistRec
  Erase LedInfo, DistInfo, ChkRegInfo           ', ChkInfo
  Title$ = "Check Pre-Audit Report"
  Call MainLog("ChkPreAudit Rpt.")
  ActivateControls frmAPChkProcessMenu
  ViewPrint "APCHKREG.PRN", Title$
ExitPreAudit:
  Exit Sub

PADoFundHeader:
  PageNum = PageNum + 1
  Page$ = Using("###", Str$(PageNum))
  Print #PrintFile, "Check Pre-Audit Report Summary                                  Page:" + Page$
  Print #PrintFile, Dash$
  Print #PrintFile,
  Print #PrintFile, "    Checks to Print        "; Using("#,###", Str$(ChkCnt))
  Print #PrintFile, "           Totaling  "; Using("$###,###,###.##", Str$(TotalChkAmt#))
  Print #PrintFile,
  Print #PrintFile, " By Fund:"
  Linecnt = 7
  Return

PADoChkRegHeader:
  PageNum = PageNum + 1
  Page$ = Using("###", Str$(PageNum))
  Print #PrintFile, "A/P Check Pre-Audit Report                                      Page:" + Page$
  Print #PrintFile, "Run Date: " + Date$
  Print #PrintFile,
  Print #PrintFile, Dash$
  Linecnt = 4
  'GoSub PrintVendHeader
  'GoSub PAInvHeader

Return

PAInvHeader:
  Print #PrintFile,
  Print #PrintFile, "Inv Date   Due Date    Inv Num/Desc                   PO            Amount"
  Print #PrintFile, "---------- ----------  -------------------------  ------------ ----------------"
  Linecnt = Linecnt + 2
Return
PrintDist:
  LSet ChkRegInfo(1).VendName = QPTrim$(Vendor.VNAME)
  LSet LedInfo(1).InvDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).DueDate = Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).InvNum = QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim$(APLedgerRec(1).Comment)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    LSet LedInfo(1).PONum = Left$(QPTrim$(APLedgerRec(1).PONum), 10)
  Else
    LSet LedInfo(1).PONum = Left$(QPTrim$(APLedgerRec(1).MPONum), 10)
  End If
  P$ = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
 'RSet LedInfo(1).Amt = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  If Not onlyneg Then
    TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  End If
  Print #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvNum; LedInfo(1).PONum; Tab(67); P$

  Linecnt = Linecnt + 2
  NextDist& = APLedgerRec(1).FrstDist

  Print #PrintFile, Tab(50); "Dist Acct        Dist Amount"
  Linecnt = Linecnt + 1

  Do Until NextDist& = 0
    Get APDistFile, NextDist&, APDistRec(1)
    If Linecnt > MaxLines Then
      Print #PrintFile, FF$
      GoSub PADoChkRegHeader
    End If
    LSet LedInfo(1).InvDate = ""
    LSet LedInfo(1).DueDate = ""
    LSet LedInfo(1).InvNum = ""
    LSet LedInfo(1).PONum = ""
    RSet LedInfo(1).Amt = ""

    LSet LedInfo(1).DistAcct = APDistRec(1).DistAcctNum
    'LSet LedInfo(1).DistAmt = Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))
    P$ = Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))

    'PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; DistInfo(1).Di
    'PRINT #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvN
    Print #PrintFile, Tab(50); LedInfo(1).DistAcct; Tab(67); P$

    Linecnt = Linecnt + 1
    If Not onlyneg Then
      ThisFund$ = Left$(APDistRec(1).DistAcctNum, GLFundLen)
      For FundCnt = 1 To NumFunds
        If ThisFund$ = FundList$(FundCnt) Then
          FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
          Exit For
        End If
      Next
    End If
    NextDist& = APDistRec(1).NextDist
  Loop
  Print #PrintFile, ""
  Linecnt = Linecnt + 1
Return

PrintVendHeader:
  'PRINT #PrintFile, Dash$
  If Not onlyneg Then
    Print #PrintFile, Vendor.vnum; Vendor.VNAME
  Else
    Print #PrintFile, Vendor.vnum; Vendor.VNAME; "-No Check Will Be Generated"
  End If
  Linecnt = Linecnt + 1
Return

FinishVendor:
  Print #PrintFile, Dash$
  Print #PrintFile, Vendor.vnum; Vendor.VNAME;
  If onlyneg Then
    Print #PrintFile, "-No Check Will Be Generated";
  End If
  Print #PrintFile, Tab(50); "Vendor Total: ";
  If onlyneg Then
    Print #PrintFile, Tab(65); Using("$###,###,###.##", Str$(VendTotal#)); "**"
  Else
    Print #PrintFile, Tab(65); Using("$###,###,###.##", Str$(VendTotal#))
  End If
  Print #PrintFile, Dash$
  Print #PrintFile, ""
  VendTotal# = 0
  If Pagechk = True Then
    Linecnt = 56
  Else
    Linecnt = Linecnt + 4
  End If
Return
CancelExit:
  Exit Sub
End Sub

Private Sub PrintCheckListing2()
  Dim Dash As String, FF As String, ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer
  Dim VendorFile As Integer, NumVRecs As Long, PrintFile As Integer
  Dim Cnt2 As Long, TCheckAmt As Double, Title As String, Temp As Integer
  Dim low As Long, High As Long
  If Not Exist("APCHKINF.DAT") Then Exit Sub
  FrmShowPctComp.Label1 = "Creating Check Listing Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmAPChkProcessMenu
  Dash$ = String$(78, "-")
  FF$ = Chr$(12)
  ChkinfoFile = FreeFile
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  ChkInfoRecLen = Len(CHKinfo(1))
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  ReDim CHKinfo(1 To VCnt) As CheckInfoType3
  'FGetAH "APCHKINF.DAT", Chkinfo(1), ChkInfoRecLen, VCnt
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For Temp = 1 To VCnt
    Get ChkinfoFile, Temp, CHKinfo(Temp)
  Next
  low = LBound(CHKinfo)
  High = UBound(CHKinfo)
  QCkSort CHKinfo(), low, High
  OpenVendorFile VendorFile, NumVRecs
  PrintFile = FreeFile
  Open "APCHKLST.PRN" For Output As #PrintFile

  Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).chkdate), "12-31-1979"), "mm/dd/yyyy")
  Print #PrintFile,
  Print #PrintFile, " Check No.         Description"
  Print #PrintFile, Dash$

  For cnt = 1 To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPChkProcessMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2& = CHKinfo(cnt).StartChk To CHKinfo(cnt).LastChk
      Print #PrintFile, Using("#######", Str$(Cnt2&));
      If CHKinfo(cnt).VoidFlag Then
        Print #PrintFile, Tab(20); Vendor.VNAME; Tab(50); "  CANCELED BY USER"
      ElseIf Cnt2& < CHKinfo(cnt).LastChk Then
        Print #PrintFile, Tab(20); Vendor.VNAME; Tab(52); "            VOID"
      Else
        TCheckAmt# = Round(TCheckAmt# + CHKinfo(cnt).ChkAmt)
        Print #PrintFile, Tab(20); Vendor.VNAME; Tab(54); Using("$##,###,###.##", Str$(CHKinfo(cnt).ChkAmt))
      End If
    Next
  Next
  
  Print #PrintFile, Dash$
  Print #PrintFile, VCnt & " Checks Totaling"; Tab(54); Using("$##,###,###.##", Str$(TCheckAmt#))
  Print #PrintFile, FF$
  Close
  Title$ = "A/P Check Listing"
  Call MainLog("Print AP Chk List.")
  ActivateControls frmAPChkProcessMenu
  ViewPrint "APCHKLST.PRN", Title$
CancelExit:
  Exit Sub
ExitCheckListing:

End Sub
Public Sub PrintChkListDist2(RptBreak As Boolean)
  Dim Dash As String, FF As String, ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer
  Dim VendorFile As Integer, NumVRecs As Long, PrintFile As Integer
  Dim Cnt2 As Long, TCheckAmt As Double, Title As String, Temp As Integer
  Dim low As Long, High As Long, cntven As Integer
  Dim TPayListFile As Integer, PayListRecLen As Integer, P As String
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, MaxLines As Integer, ToPrint As String
  Dim NumFunds As Integer, APDistRecLen As Integer
  Dim TPayCnt As Integer, PageNum As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VRecNum As Long
  Dim ChkCnt As Integer, Linecnt As Integer
  Dim Page As String, TotalChkAmt As Double, VendTotal As Double
  Dim NextDist As Long, ThisFund As String, FundCnt As Integer
  FF$ = Chr$(12)
  MaxLines = 52
  Dash$ = String$(80, "-")
  PageNum = 0
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub
  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType2
  ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType
  DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  PayListRecLen = Len(TPayList)
  TPayListFile = FreeFile
  Open "TPAYLIST2.LST" For Random Shared As TPayListFile Len = PayListRecLen
  TPayCnt = LOF(TPayListFile) \ 6
  If TPayCnt = 0 Then
    Exit Sub
  End If
  ReDim TPayArray(1 To TPayCnt) As TPayListType
  For Pcnt = 1 To TPayCnt
    Get TPayListFile, Pcnt, TPayList
    TPayArray(Pcnt).VendorRecNum = TPayList.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayList.LedgerRecNum
  Next
  Close TPayListFile
  If Not Exist("APCHKINF.DAT") Then Exit Sub
  FrmShowPctComp.Label1 = "Creating Check Listing Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmAPChkProcessMenu
  ChkinfoFile = FreeFile
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  ChkInfoRecLen = Len(CHKinfo(1))
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  ReDim CHKinfo(1 To VCnt) As CheckInfoType3
  'FGetAH "APCHKINF.DAT", Chkinfo(1), ChkInfoRecLen, VCnt
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For Temp = 1 To VCnt
    Get ChkinfoFile, Temp, CHKinfo(Temp)
  Next
  low = LBound(CHKinfo)
  High = UBound(CHKinfo)
  QCkSort CHKinfo(), low, High
  OpenVendorFile VendorFile, NumVRecs
  PrintFile = FreeFile
  Open "APCHKDet.PRN" For Output As #PrintFile
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  If RptBreak = False Then
    Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).chkdate), "12-31-1979"), "mm/dd/yyyy")
    Linecnt = 1
  End If
  For cnt = 1 To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPChkProcessMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2& = CHKinfo(cnt).StartChk To CHKinfo(cnt).LastChk
      GoSub Printheadstuff
        Print #PrintFile, " Check No.         Description                                Check Amount"
        Print #PrintFile, Dash$
        Linecnt = Linecnt + 2
      Print #PrintFile, Using("#######", Str$(Cnt2&));
      If CHKinfo(cnt).VoidFlag Then
        Print #PrintFile, Tab(20); Vendor.VNAME; Tab(50); "  CANCELED BY USER"
      ElseIf Cnt2& < CHKinfo(cnt).LastChk Then
        Print #PrintFile, Tab(20); Vendor.VNAME; Tab(52); "            VOID"
      Else
        
        TCheckAmt# = Round(TCheckAmt# + CHKinfo(cnt).ChkAmt)
        Print #PrintFile, Tab(20); Vendor.VNAME; Tab(65); Using("$##,###,###.##", Str$(CHKinfo(cnt).ChkAmt))
      End If
    Next
      GoSub PrintInvHeader
      For cntven = 1 To TPayCnt
      VRecNum& = TPayArray(cntven).VendorRecNum

        If VRecNum& = CHKinfo(cnt).VendorRecNum Then
        
          Get APLedgerFile, TPayArray(cntven).LedgerRecNum, APLedgerRec(1)
          If Linecnt >= MaxLines Then
            GoSub Printheadstuff
            GoSub PrintInvHeader
          End If

          GoSub PrintDist
        End If
      Next
    Print #PrintFile, Tab(30); "-------**********-------"
  Next
'******Here we go totals and close*************
  If RptBreak = True Or Linecnt >= MaxLines - 4 Then
    Print #PrintFile, FF$
  End If
  Print #PrintFile, ""
  Print #PrintFile, Dash$
  Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).chkdate), "12-31-1979"), "mm/dd/yyyy")
  Print #PrintFile, VCnt; "Checks Totaling - "; Tab(40); Using("$##,###,###.##", Str$(TCheckAmt#))
  Print #PrintFile, FF$
  Close
  Erase FundList$, FundAmts, APLedgerRec, APDistRec
  Erase LedInfo, DistInfo, ChkRegInfo           ', ChkInfo

  Title$ = "A/P Check Report"
  Call MainLog("Print AP Chk List w/dist.")
  ActivateControls frmAPChkProcessMenu
  ViewPrint "APCHKDet.PRN", Title$
  Exit Sub
PrintDist:
  ToPrint$ = Space(80)
  'LSet ChkRegInfo(1).VENDNAME = Vendor.VNAME
  Mid$(ToPrint$, 4) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 16) = Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 28) = Left$(APLedgerRec(1).DOCNum, 20)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    Mid$(ToPrint$, 50) = Left$(APLedgerRec(1).PONum, 10)
  Else
    Mid$(ToPrint$, 50) = Left$(APLedgerRec(1).MPONum, 10)
  End If
  Mid$(ToPrint$, 63) = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
 'RSet LedInfo(1).Amt = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  Print #PrintFile, ToPrint$

  Linecnt = Linecnt + 2
  NextDist& = APLedgerRec(1).FrstDist
    If Linecnt > MaxLines - 2 Then
      GoSub Printheadstuff
    End If

  Print #PrintFile, Tab(44); "Distributions"
  Linecnt = Linecnt + 1

  Do Until NextDist& = 0
    Get APDistFile, NextDist&, APDistRec(1)
    If Linecnt > MaxLines - 2 Then
      GoSub Printheadstuff
    End If
    ToPrint$ = Space(80)
    Mid$(ToPrint$, 35) = QPTrim(APDistRec(1).DistAcctNum)
    Mid$(ToPrint$, 50) = Using$("##,###,###.##", Str$(APDistRec(1).DistAmt))

    Print #PrintFile, ToPrint$
    Linecnt = Linecnt + 1
    NextDist& = APDistRec(1).NextDist
  Loop
  Print #PrintFile, ""
  Linecnt = Linecnt + 1
Return
Printheadstuff:
If RptBreak = True Or Linecnt >= MaxLines - 2 Then
  If cnt <> 1 Then
  Print #PrintFile, FF$
  End If
  Linecnt = 1
  Print #PrintFile, " A/P Check Listing:  " + Format(DateAdd("d", (CHKinfo(1).chkdate), "12-31-1979"), "mm/dd/yyyy")
End If
  'Print #PrintFile,
Return
PrintInvHeader:
  If Linecnt >= MaxLines - 3 Then
    GoSub Printheadstuff
  End If
  Print #PrintFile,
  Print #PrintFile, "   Inv Date    Due Date    Inv Num                  PO           Inv Amount"
  Print #PrintFile, "   ----------  ----------  -------------------- ------------ ---------------"
  Linecnt = Linecnt + 2
Return

CancelExit:
  Exit Sub
ExitCheckListing:

End Sub
Private Sub getnoPays()
  Dim TPayListFile As Integer, PayListRecLen As Integer, P As String
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, cnt As Integer, FF As String, MaxLines As Integer
  Dim Dash As String, PrintFile2 As Integer, Dash2 As String, PageNum As Integer
  Dim NumFunds As Integer, APDistRecLen As Integer, VendorFile As Integer
  Dim PrintFile  As Integer, TPayCnt As Integer, NumVRecs As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VRecNum As Long
  Dim ChkCnt As Integer, Linecnt As Integer, Title As String
  Dim Page As String, TotalChkAmt As Double, VendTotal As Double
  Dim NextDist As Long, ThisFund As String, FundCnt As Integer
  Dim ToPrint As String, TPVend As String, TPInv As String, TPDist As String
  Dim TmpRecno As Long, TPayNotFile As Integer, PayNotRecLen As Integer
  Dim TNotcnt As Integer, TPaylist2file As Integer, TPayListdFile As Integer
  Dim Cnt2 As Integer, ccnt As Integer, ccnt2 As Integer, thiscnt As Integer
  Dim T2cnt As Integer, T2PayFile As Integer, Pay2RecLen As Integer
  Dim thisTcnt As Integer
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub
  KillFile "TPayNot.lst"
  KillFile "T2Pay.lst"
  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType2
  ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType
  thiscnt = 0
  thisTcnt = 0
  DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  PayListRecLen = Len(TPayList)
  TPayListFile = FreeFile
  Open "TPAYLIST.LST" For Random Shared As TPayListFile Len = PayListRecLen
  TPayCnt = LOF(TPayListFile) \ 6
  If TPayCnt = 0 Then
    Close
    Exit Sub
  End If
  
  If TPayCnt > 0 Then
    FrmShowPctComp.Label1 = "Creating Pre-Audit Report"
    FrmShowPctComp.Show , Me
    DoEvents
    DeActivateControls frmAPChkProcessMenu
  End If
'Ones with neg amt
  PayNotRecLen = Len(TPayNot)
  TPayNotFile = FreeFile
  Open "TPayNot.lst" For Random Shared As TPayNotFile Len = PayNotRecLen
'Ones with pos amt
  Pay2RecLen = Len(T2Pay)
  T2PayFile = FreeFile
  Open "T2Pay.lst" For Random Shared As T2PayFile Len = Pay2RecLen

  ReDim TPayArray(1 To TPayCnt) As TPayListType
  For Pcnt = 1 To TPayCnt
    FrmShowPctComp.ShowPctComp Pcnt, TPayCnt
    Get TPayListFile, Pcnt, TPayList
    TPayArray(Pcnt).VendorRecNum = TPayList.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayList.LedgerRecNum
  Next
  Close TPayListFile

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  
  Get VendorFile, TPayArray(1).VendorRecNum, Vendor
  Get APLedgerFile, TPayArray(1).LedgerRecNum, APLedgerRec(1)
  TmpRecno = TPayArray(1).LedgerRecNum
  GoSub PrintDist
  VRecNum& = TPayArray(1).VendorRecNum
  ChkCnt = 1
  For cnt = 2 To TPayCnt
    FrmShowPctComp.ShowPctComp cnt, TPayCnt
    If VRecNum& <> TPayArray(cnt).VendorRecNum Then
      GoSub FinishVendor
      Get VendorFile, TPayArray(cnt).VendorRecNum, Vendor
      VRecNum& = TPayArray(cnt).VendorRecNum
      
    End If
    Get APLedgerFile, TPayArray(cnt).LedgerRecNum, APLedgerRec(1)
    TmpRecno = TPayArray(cnt).LedgerRecNum
    GoSub PrintDist
  Next
  GoSub FinishVendor
  Close
  KillFile "TPayList2.lst"
  KillFile "TPayListd.lst"
  If Exist("TPayNot.lst") Then
    PayNotRecLen = Len(TPayNot)
    TPayNotFile = FreeFile
    Open "TPayNot.lst" For Random Shared As TPayNotFile Len = PayNotRecLen
    TNotcnt = LOF(TPayNotFile) \ 6
    ReDim tpaynarray(1 To TNotcnt) As TPayNotListType
    For Pcnt = 1 To TNotcnt
      Get TPayNotFile, Pcnt, TPayNot
      tpaynarray(Pcnt).VendorRecNum = TPayNot.VendorRecNum
      tpaynarray(Pcnt).Amt = TPayNot.Amt
    Next
    Close TPayNotFile
    TPayListdFile = FreeFile
    Open "TPAYLISTD.LST" For Random Shared As TPayListdFile Len = PayListRecLen
    ccnt = 0
    ccnt2 = 0
    For cnt = 1 To TPayCnt
      For Cnt2 = 1 To TNotcnt
        If TPayArray(cnt).VendorRecNum = tpaynarray(Cnt2).VendorRecNum Then
          TPayListD.LedgerRecNum = TPayArray(cnt).LedgerRecNum
          TPayListD.VendorRecNum = TPayArray(cnt).VendorRecNum
          ccnt = ccnt + 1
          Put TPayListdFile, ccnt, TPayListD
        End If
      Next
     Next

    If Exist("T2Pay.lst") Then
      Pay2RecLen = Len(T2Pay)
      T2PayFile = FreeFile
      Open "T2Pay.lst" For Random Shared As T2PayFile Len = Pay2RecLen
      T2cnt = LOF(T2PayFile) \ 6
      ReDim t2paynarray(1 To T2cnt) As TPayNotListType
      For Pcnt = 1 To T2cnt
        Get T2PayFile, Pcnt, T2Pay
        t2paynarray(Pcnt).VendorRecNum = T2Pay.VendorRecNum
        t2paynarray(Pcnt).Amt = T2Pay.Amt
      Next
      Close T2PayFile
      TPaylist2file = FreeFile
      Open "TPAYLIST2.LST" For Random Shared As TPaylist2file Len = PayListRecLen
      ccnt = 0
      ccnt2 = 0
      For cnt = 1 To TPayCnt
        For Cnt2 = 1 To T2cnt
          If TPayArray(cnt).VendorRecNum = t2paynarray(Cnt2).VendorRecNum Then
            TPayList2.LedgerRecNum = TPayArray(cnt).LedgerRecNum
            TPayList2.VendorRecNum = TPayArray(cnt).VendorRecNum
            ccnt2 = ccnt2 + 1
            Put TPaylist2file, ccnt2, TPayList2
          End If
        Next
      Next
     ' Close T2Paylist
    End If
  Else
    TPaylist2file = FreeFile
    Open "TPAYLIST2.LST" For Random Shared As TPaylist2file Len = PayListRecLen
    ccnt = 0
    ccnt2 = 0
    For cnt = 1 To TPayCnt
      TPayList2.LedgerRecNum = TPayArray(cnt).LedgerRecNum
      TPayList2.VendorRecNum = TPayArray(cnt).VendorRecNum
      ccnt2 = ccnt2 + 1
      Put TPaylist2file, ccnt2, TPayList2
    Next

  End If
  Close
'  KillFile "TPayList.old"
'  Name "TPayList.lst" As "TPayList.old"
'  KillFile "TPayList.lst"
'  Name "TPayList2.lst" As "TPayList.lst"
  Erase TPayArray, tpaynarray
  Erase FundList$, FundAmts, APLedgerRec, APDistRec
  Erase LedInfo, DistInfo, ChkRegInfo           ', ChkInfo
ExitPreAudit:
  Exit Sub
PrintDist:
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  NextDist& = APLedgerRec(1).FrstDist
Return

FinishVendor:
  If Not VendTotal# > 0 Then
    TPayNot.VendorRecNum = VRecNum&
    TPayNot.Amt = VendTotal#
    thiscnt = thiscnt + 1
    Put TPayNotFile, thiscnt, TPayNot
  Else
    T2Pay.VendorRecNum = VRecNum&
    T2Pay.Amt = VendTotal#
    thisTcnt = thisTcnt + 1
    Put T2PayFile, thisTcnt, T2Pay
  End If
  
  VendTotal# = 0
  TPVend$ = ""
Return
CancelExit:
  Exit Sub
End Sub

