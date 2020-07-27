VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Taxes vs 2.05 Main Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11640
   Icon            =   "frmTaxMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1560
      Top             =   2434
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   1560
      Top             =   3034
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxBillingFunctions 
      Height          =   444
      Left            =   4008
      TabIndex        =   3
      Top             =   3936
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterTaxPayments 
      Height          =   435
      Left            =   4005
      TabIndex        =   2
      Top             =   3396
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":0AB3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustMaint 
      Height          =   435
      Left            =   4005
      TabIndex        =   0
      Top             =   2310
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":0C99
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAbstractMaint 
      Height          =   435
      Left            =   4005
      TabIndex        =   1
      Top             =   2853
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":0E81
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxReportingSystem 
      Height          =   444
      Left            =   4008
      TabIndex        =   4
      Top             =   4488
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":1069
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4005
      TabIndex        =   9
      Top             =   7230
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":1251
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdManualEntry 
      Height          =   435
      Left            =   4005
      TabIndex        =   5
      Top             =   5040
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":142F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxBillAdj 
      Height          =   450
      Left            =   4005
      TabIndex        =   6
      Top             =   5583
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":1618
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxAdvertising 
      Height          =   435
      Left            =   4005
      TabIndex        =   7
      Top             =   6141
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":1803
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSetUpAndUtil 
      Height          =   435
      Left            =   4005
      TabIndex        =   8
      Top             =   6684
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
      ButtonDesigner  =   "frmTaxMainMenu.frx":19EE
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
      Left            =   2813
      TabIndex        =   10
      Top             =   1164
      Width           =   6012
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2214
      X2              =   2214
      Y1              =   2127
      Y2              =   8015
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199
      X2              =   2914
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8706
      X2              =   9408
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8586
      Top             =   2019
      Width           =   971
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8706
      X2              =   8706
      Y1              =   2127
      Y2              =   8028
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
      Top             =   2019
      Width           =   971
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1098
      Index           =   1
      Left            =   1493
      Top             =   803
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1495
      Top             =   687
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2094
      Top             =   1886
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2213
      Top             =   2117
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8585
      Top             =   1887
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8706
      Top             =   2117
      Width           =   732
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuLinkPins 
         Caption         =   "Build Dummy Transactions for Link To Prop Pin"
      End
      Begin VB.Menu mnuLateNotice2Y 
         Caption         =   "Change Late Notice Flags to 'Y'"
      End
      Begin VB.Menu mnuLateNotice2N 
         Caption         =   "Change Late Notice Flags to 'N'"
      End
      Begin VB.Menu mnuMakeAllActive 
         Caption         =   "Make All Customers Active"
      End
      Begin VB.Menu mnuClipZip 
         Caption         =   "Clip Off Zip Hypens"
      End
      Begin VB.Menu mnuFixPersPropDates 
         Caption         =   "Fix Personal Prop Dates"
      End
      Begin VB.Menu mnuRepairNegAndFutureVals 
         Caption         =   "Repair Screen"
      End
      Begin VB.Menu mnuCountyEdit 
         Caption         =   "Allow County Number Edit On/Off"
      End
   End
End
Attribute VB_Name = "frmTaxMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAbstractMaint_Click()
  DelAbs = True
  AddCust = False
  EditCust = False
  THistRpt = False
  PayEntry = False
  frmTaxAbsMaint.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdCustMaint_Click()
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim x As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  
  If Not Exist(TaxSetupName) Then
    Call TaxMsg(900, "Please save data in the 'Tax System Setup' screen before continuing.")
    Exit Sub
  End If
  
  frmTaxCustMaintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdEnterTaxPayments_Click()
  If Not Exist("TAXCUST.DAT") Then
    frmTaxMsg.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("TAXSETUP.DAT") Then
    frmTaxMsg.Label1.Caption = "Please complete the Tax Setup data before continuing."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Exit Sub
  End If
  
  If LevelPass = 1 Or LevelPass = 3 Then
    frmTaxPayOperEntry.Show
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
  
  Timer1.Enabled = True

End Sub

Private Sub cmdManualEntry_Click()
  frmTaxManualBillMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSetUpAndUtil_Click()
  frmTaxBillSetUpMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdTaxAdvertising_Click()
  If Check4PayBatch = True Then
'    Call TaxMsg(800, "An unposted payment file is ready for posting. Advertising calculations cannot be conducted until these payments are posted.")
    frmTaxUnpostedPayList.Show vbModal
    DoEvents
    Exit Sub
  End If
 
  frmTaxAdvColMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdTaxBillAdj_Click()
  frmTaxAdjustments.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdTaxBillingFunctions_Click()
  frmTaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdTaxReportingSystem_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
'      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim Cnt&, dl&
  Dim ThisDir$
  Dim CERec As AllowCountyEdit
  Dim CEHandle As Integer
  
  AddCust = False
  EditCust = False
  DelAbs = False
  THistRpt = False
  PayEntry = False
  
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
  Cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, Cnt)
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
'  frmTaxDataRepair.Show
'  LevelPass = 1
'  PWcnt = 13
'  OperNum = 13

  If Exist(CntyEditFile) Then 'added 7/11/07
    OpenCountyEditFile CEHandle
    Get CEHandle, 1, CERec
    Close CEHandle
    If Date2Num(Date) > CERec.AllowCountyEditXDate Then
      KillFile CntyEditFile
    End If
  End If

  
  'Unload frmTaxMainMenu
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxMainMenu.")
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
  frmTaxCountyNumEdit.Show vbModal
  DoEvents
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
  frmTaxShowPctComp.Label1 = "Linking Real Pins Procedure #1"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To NumOfTaxTranRecs
    Get TaxTranHandle, x, TaxTranRec
    If QPTrim$(TaxTranRec.Description) = "SoSoft: Link Real Pin" Then
        If Notified1 = True Then GoTo SkipIt
        frmTaxShowPctComp.Hide
        If TaxMsgWOpts(800, "The linking of real pins to dummy transactions has already taken place. To continue anyway press F10. Otherwise, press ESC to escape.", "F10 Continue", "ESC Escape") = "abort" Then
          Unload frmTaxShowPctComp
          Close
          Exit Sub
        Else
          Notified1 = True
          frmTaxShowPctComp.Show
          MainLog ("User warned that the linking of real pins to dummy transactions has already taken place but elected to continue anyway.")
        End If
      ElseIf QPTrim$(TaxTranRec.RealPin) <> "0" Then
        If Notified2 = True Then GoTo SkipIt
        frmTaxShowPctComp.Hide
        If TaxMsgWOpts(800, "Please be advised that transactions have been posted with real pin numbers stored.", "F10 Continue", "ESC Escape") = "abort" Then
          Unload frmTaxShowPctComp
          Close
          Exit Sub
        Else
          Notified2 = True
          frmTaxShowPctComp.Show
          MainLog ("User advised that transactions have been posted with real pin numbers stored and elected to continue anyway.")
        End If
      End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTaxTranRecs
  Next x
  Unload frmTaxShowPctComp
  
  frmTaxShowPctComp.Label1 = "Linking Real Pins Procedure #2"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
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
        TaxTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(RealPropRec.CustPin))
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
        Put TaxTranHandle, ThisNumOfTaxTrans, TaxTranRec
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfRealRecs
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  
  Call Savemsg(900, "The linking has completed successfully")
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMainMenu", "mnuLinkPins_Click", Erl)
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
    Unload frmTaxShowPctComp
    EnableCloseButton Me.hwnd, True
    Close

End Sub

Private Sub Timer1_Timer()
  Call Terminate2Shell 'closes all forms but does not clear password data
End Sub

Private Sub Timer2_Timer()
  cmdExit.Enabled = True
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

Private Sub mnuRepairNegAndFutureVals_Click()
  frmTaxDataRepair.Show
  DoEvents
  Me.Hide
End Sub

Private Sub mnuFixPersPropDates_Click()
  Dim PersRec As PersonalRecType
  Dim x As Long
  Dim NumOfPRecs As Long
  Dim PHandle As Integer
  Dim PCnt As Long
  
  frmTaxShowPctComp.Label1 = "Fixing Personal Property Dates"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  OpenPersPropFile PHandle, NumOfPRecs
  For x = 1 To NumOfPRecs
    Get PHandle, x, PersRec
    If PersRec.PROPDATE < 0 Then
      PersRec.PROPDATE = 0
      Put PHandle, x, PersRec
      PCnt = PCnt + 1
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfPRecs
  Next x
  Unload frmTaxShowPctComp
  
  Close PHandle
  
  If PCnt > 0 Then
    Call TaxMsg(900, "A total of " + CStr(PCnt) + " dates were corrected successfully.")
  Else
    Call TaxMsg(900, "All dates were OK...none were updated.")
  End If

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

