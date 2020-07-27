VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDCSetupMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decals Setup Maintenance Menu"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmDCSetupMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   324
      Left            =   336
      TabIndex        =   12
      Top             =   1728
      Visible         =   0   'False
      Width           =   924
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   372
      Left            =   312
      TabIndex        =   11
      Top             =   1008
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   564
      Left            =   888
      TabIndex        =   10
      Top             =   2568
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "10:27 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "11/15/2007"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdSetupInfo 
      Height          =   492
      Left            =   3846
      TabIndex        =   0
      Top             =   2256
      Width           =   4524
      _Version        =   131072
      _ExtentX        =   7980
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRelink 
      Height          =   492
      Left            =   3846
      TabIndex        =   4
      Top             =   5092
      Width           =   4524
      _Version        =   131072
      _ExtentX        =   7980
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":0AF4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReindexName 
      Height          =   492
      Left            =   3846
      TabIndex        =   2
      Top             =   3674
      Width           =   4524
      _Version        =   131072
      _ExtentX        =   7980
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":0D13
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdApplicationDef 
      Height          =   480
      Left            =   3840
      TabIndex        =   1
      Top             =   2970
      Width           =   4530
      _Version        =   131072
      _ExtentX        =   7990
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":0F38
   End
   Begin fpBtnAtlLibCtl.fpBtn Command1 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   4410
      Width           =   4530
      _Version        =   131072
      _ExtentX        =   7980
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":1163
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitMenu 
      Height          =   480
      Left            =   3840
      TabIndex        =   7
      Top             =   7230
      Width           =   4530
      _Version        =   131072
      _ExtentX        =   7990
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":134E
   End
   Begin fpBtnAtlLibCtl.fpBtn Command2 
      Height          =   492
      Left            =   3846
      TabIndex        =   5
      Top             =   5801
      Width           =   4524
      _Version        =   131072
      _ExtentX        =   7980
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":152F
   End
   Begin fpBtnAtlLibCtl.fpBtn fpbtnClearExp 
      Height          =   492
      Left            =   3846
      TabIndex        =   6
      Top             =   6510
      Width           =   4524
      _Version        =   131072
      _ExtentX        =   7980
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCSetupMenu.frx":171D
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Decal Setup Maintenance Menu"
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
      Left            =   3540
      TabIndex        =   9
      Top             =   1104
      Width           =   5148
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmDCSetupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class


Private Sub cmdReindexName_Click()
  SortDCNameIndex frmDCSetupMenu
End Sub

Private Sub cmdApplicationDef_Click()
  Load frmApplicationLetter
  DoEvents
  frmApplicationLetter.Show
  Unload Me
End Sub

Private Sub cmdExitMenu_Click()
  Load frmDCMainMenu
  DoEvents
  frmDCMainMenu.Show
  Unload Me
End Sub


Private Sub cmdRelink_Click()
  If OK4Secure = True Then
    RelinkDCStuff frmDCSetupMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub Command3_Click()
'For Pocahontas  will need to fix the cust file with DFD first then
  'run the cleanupcust
  'rename the dcdustn.dat to dccust.dat
  'then run the savemastertoveh
  'also need to delete the trans file!!!!!!!!!!!!!!
'For Pennington Gap convert trans then clean blank recs out of cust then do the rest
'CleanupCust
'Fixthenums
'SaveMastertoVeh
'FixTransLink
''fixuptheveh


'
End Sub

Private Sub Command4_Click()
frmExpCustomerInfo.ExpDecalCust
End Sub

Private Sub Command5_Click()
frmExpCustomerInfo.ExpDecalVeh
End Sub

Private Sub fpbtnClearExp_Click()
  If OK4Secure = True Then
    Load frmDeletebyExpire
    DoEvents
    frmDeletebyExpire.Show
    Unload Me
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdSetupInfo_Click()
  Load frmSystemSetup
  DoEvents
  frmSystemSetup.Show
  Unload Me
End Sub

'Private Sub cmdReprint_Click()
'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'  If Not Exist("UBFBILLS.PRN") Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO PRN FILE. Reprint Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL PRINT FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'
'  If Not Exist(UBFinBillsFile) Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO BILL FILE! Reprint Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'  frmBillPrinting.REPRN True, True
'  Load frmBillPrinting
'  DoEvents
'  frmBillPrinting.Show
'  Unload frmUBFinalBillPrintMenu
'
'End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Me.HelpContextID = hlpDecalSetup
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via DCSEtupMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

'Private Sub cmdPrnAllUBBills_Click()
'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'
'  If Not Exist(UBFinBillsFile) Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO BILL FILE! Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'  frmBillPrinting.REPRN False, True
'  Load frmBillPrinting
'  DoEvents
'  frmBillPrinting.Show
'  Unload frmUBFinalBillPrintMenu
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitMenu_Click
      KeyCode = 0
    Case vbKeyHome
      cmdSetupInfo.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub Command1_Click()
'SaveMastertoVeh
'  If OK4Secure = True Then
'  'Fixthenums
'    If MsgBox("Continue with transaction conversion?", vbYesNo, "Continue") = vbYes Then
'ConvertTrans
'CleanupCust
'Fixthenums
'SaveMastertoVeh
'FixTransLink '    End If
'  Else
'    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'  End If
End Sub
Private Sub Command2_Click()
  If OK4Secure = True Then
    If MsgBox("Continue with Running Balance recalc?", vbYesNo, "Continue") = vbYes Then
      RecalcBal
    End If
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub ConvertTrans()
Dim DCTranLen  As Integer, NumOfRecs As Long
Dim DCTran  As Integer, olD As Integer
Dim cnt As Long
If Exist(DCPath$ + "DCSetup.dat") Then
  If MsgBox("Setup file exist continue with trasaction conversion?", vbYesNo, "Continue") = vbNo Then
    GoTo donealready
  End If
End If
'CleanupCust
ReDim olDCTrans(1) As OldDCTransRecType
ReDim DCTrans(1) As DCTransRecType
If Exist(DCPath$ + "DCTrans.dat") Then
  DCTranLen = Len(DCTrans(1))
  DCTran = FreeFile
  NumOfRecs = FileSize(DCPath$ + "DCTrans.DAT") \ DCTranLen
  Open DCPath$ + "DCTrans.DAT" For Random Shared As DCTran Len = DCTranLen
  For cnt = 1 To 5
    Get DCTran, cnt, DCTrans(1)
    If DCTrans(1).ChkByte = Chr$(1) Then
      MsgBox "Already Converted.", vbOKOnly, "Converted"
      Close
      GoTo donealready
    End If
  Next
  Close
 '' GoTo donealready
  
  olD = FreeFile
  Open DCPath$ + "DCTrans.DAT" For Random Shared As olD Len = DCTranLen
  For cnt = 1 To NumOfRecs
    Get olD, cnt, olDCTrans(1)
    DCTrans(1).CustomerNumber = olDCTrans(1).CustomerNumber
    DCTrans(1).TransDate = olDCTrans(1).TransDate
    DCTrans(1).TransAmount = olDCTrans(1).TransAmount
    DCTrans(1).TransType = olDCTrans(1).TransType
    DCTrans(1).TRVinDesc = olDCTrans(1).TRVinDesc
    DCTrans(1).CashAmount = olDCTrans(1).CashAmount
    DCTrans(1).ChkAmount = olDCTrans(1).ChkAmount
    DCTrans(1).BalanceAfterTrans = olDCTrans(1).BalanceAfterTrans
    DCTrans(1).makemodel = olDCTrans(1).makemodel
    DCTrans(1).StateTag = olDCTrans(1).StateTag
    DCTrans(1).ExpireDate = olDCTrans(1).ExpireDate
    DCTrans(1).Sticker = olDCTrans(1).Sticker
    DCTrans(1).NextTrans = olDCTrans(1).NextTrans
    DCTrans(1).OperNum = olDCTrans(1).OperNum
    DCTrans(1).GLInterfaced = olDCTrans(1).GLInterfaced
    DCTrans(1).DecalCat = olDCTrans(1).DecalCat
    'I  special for Independence
'    If Len(QPTrim$(olDCTrans(1).DecalCat)) = 0 And (DCTrans(1).CustomerNumber = 944 Or DCTrans(1).CustomerNumber = 563 Or DCTrans(1).CustomerNumber = 927 Or DCTrans(1).CustomerNumber = 326 Or DCTrans(1).CustomerNumber = 949) Then
'      DCTrans(1).DecalCat = "AUT"
'    End If
    'I
    DCTrans(1).ChkByte = Chr$(1)
'''    DCTrans(1).TransType = DCTrans(1).TransType
'''    If DCTrans(1).TransType < 1 Then Stop
    DCTrans(1).ExtraDesc = ""
    DCTrans(1).VoidFlag = "N"
    DCTrans(1).VehRecord = 0
    If (DCTrans(1).CashAmount <> 0) And (DCTrans(1).ChkAmount <> 0) Then
      DCTrans(1).TransTender = 3
    ElseIf DCTrans(1).CashAmount <> 0 Then
      DCTrans(1).TransTender = 1
    ElseIf DCTrans(1).ChkAmount <> 0 Then
      DCTrans(1).TransTender = 2
    Else
      DCTrans(1).TransTender = 0
    End If
    DCTrans(1).ExtraRoom = ""
    Put #DCTran, , DCTrans(1)
  Next
  MsgBox "Transaction File Converted", vbOKOnly, "Completed"
donealready:
  Close '#DCTran
Else
  MsgBox "Transaction File Missing", vbOKOnly, "Nothing Converted"
End If 'file already exist do nothing
End Sub
Private Sub RecalcBal()
Dim DCTranLen  As Integer, NumOfRecs As Long, DCCustRecLen As Integer
Dim DCFile As Integer, PrevTranBal As Double, TrHandle As Integer
Dim cnt As Long, CntT As Long, PrevTranRec As Long
ReDim DCCustRec(1) As DCCustRecType
ReDim DCTransRec(1) As DCTransRecType
If Exist(DCPath$ + "DCCust.dat") And Exist(DCPath$ + "DCTrans.dat") Then
  DCCustRecLen = Len(DCCustRec(1))
  TrHandle = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = DCCustRecLen
  NumOfRecs = FileSize(DCPath$ + "DCCust.DAT") \ DCCustRecLen

  DCTranLen = Len(DCTransRec(1))
  DCFile = FreeFile
  Open DCPath$ + "DCTrans.DAT" For Random Shared As DCFile Len = DCTranLen
  For cnt = 1 To NumOfRecs
    Get TrHandle, cnt, DCCustRec(1)
    PrevTranRec& = DCCustRec(1).FirstTrans
    PrevTranBal = 0
    If PrevTranRec& > 0 Then
      Do While PrevTranRec& > 0
        CntT& = PrevTranRec&
        Get DCFile, CntT&, DCTransRec(1)
        If DCTransRec(1).TransType = 1 Or DCTransRec(1).TransType = 4 Then
          DCTransRec(1).BalanceAfterTrans = PrevTranBal + DCTransRec(1).TransAmount
        ElseIf DCTransRec(1).TransType = 2 Or DCTransRec(1).TransType = 3 Then
          DCTransRec(1).BalanceAfterTrans = PrevTranBal - DCTransRec(1).TransAmount
        End If
        PrevTranRec& = DCTransRec(1).NextTrans
        PrevTranBal = DCTransRec(1).BalanceAfterTrans
        Put #DCFile, CntT&, DCTransRec(1)
      Loop
    End If
  Next
  Close #DCFile
  Close #TrHandle
Else
  MsgBox "Files Missing", vbOKOnly, "Nothing Recalculated"
End If 'file already exist do nothing

End Sub

Private Sub SaveMastertoVeh()
Dim NumOfRecs As Long, DCCustRecLen As Integer
Dim DCFile As Integer, PrevTranBal As Double, TrHandle As Integer
Dim cnt As Long, CntT As Long, PrevVehRec As Long
ReDim DCCustRec(1) As DCCustRecType
ReDim DCTransRec(1) As DCTransRecType
If Exist(DCPath$ + "DCCust.dat") And Exist(DCPath$ + "DCVeh.dat") Then
  DCCustRecLen = Len(DCCustRec(1))
  TrHandle = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = DCCustRecLen
  NumOfRecs = FileSize(DCPath$ + "DCCust.DAT") \ DCCustRecLen
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen

  For cnt = 1 To NumOfRecs
    Get TrHandle, cnt, DCCustRec(1)
  ''If cnt = 533 Then Stop
    PrevVehRec& = DCCustRec(1).FirstCar
    If PrevVehRec& > 0 Then
      Do While PrevVehRec& > 0
        CntT = PrevVehRec&
        Get DCvFile, CntT, DCVRec(1)
 '''       If CntT = 803 Then Stop
 '       DCVRec(1).NextRec = 0
'          PrevVehRec& = DCVRec(1).NextRec
'          DCVRec(1).Active = "N"
'         Else
          
          PrevVehRec& = DCVRec(1).NextRec
          DCVRec(1).Active = "Y"
'         End If
          DCVRec(1).MasterRecord = cnt
          Put #DCvFile, CntT&, DCVRec(1)
      Loop
    End If
  Next
  Close #DCvFile
  Close #TrHandle
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  For cnt = 1 To NumOfVRecs
    Get DCvFile, cnt, DCVRec(1)
    If DCVRec(1).MasterRecord <= 0 Then
      DCVRec(1).Active = "N"
      Put #DCvFile, cnt&, DCVRec(1)
    End If
  Next
  MsgBox "MasterRec Nums Restored to Veh Recs.", vbOKOnly, "Rec's updated"
Else
  MsgBox "MasterRecord Num not updated in Vehicle.", vbOKOnly, "Missing Files"
End If 'file already exist do nothing

End Sub
Private Sub CleanupCust()   'to renumber cust '''and 0 out the trans links
  Dim DCCustRecLen As Integer, NumOfDCRecs As Long
  Dim DCFile As Integer, cnt As Long, ncnt As Long, DCFile2 As Integer
  ReDim DCCustRec(1) As DCCustRecType
  DCCustRecLen = Len(DCCustRec(1))
  FrmShowPctComp.Label1 = "Cleaning up Customers."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  DCFile2 = FreeFile
  Open "DCCUSTN.dat" For Random Access Read Write Shared As DCFile2 Len = DCCustRecLen
  DCFile = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As DCFile Len = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  For cnt = 1 To NumOfDCRecs
    Get DCFile, cnt, DCCustRec(1)
    '''''If cnt = 4515 Then Stop
    FrmShowPctComp.ShowPctComp cnt, NumOfDCRecs
    If (Len(QPTrim$(DCCustRec(1).CUSTNUMB)) > 0) Then
      ncnt = ncnt + 1
      DCCustRec(1).CUSTNUMB = ncnt
      DCCustRec(1).FirstTrans = 0
      Put DCFile2, ncnt, DCCustRec(1)
    End If
'    If (Len(QPTrim$(DCCustRec(1).CUSTNUMB)) <= 0) Then
'        DCCustRec(1).Deleted = "Y"
'        DCCustRec(1).CUSTNUMB = cnt
'        Put DCFile, cnt, DCCustRec(1)
'    End If
  Next cnt
  Close DCFile
End Sub
'Private Sub Fixthenums()  'if the cust num has been blanked out or -1
'    ''also if had to clean up file will renumber the acct's
'Dim NumOfRecs As Long, DCCustRecLen As Integer
'Dim DCFile As Integer, TrHandle As Integer
'Dim cnt As Long, numnum As Long
'ReDim DCCustRec(1) As DCCustRecType
'If Exist(DCPath$ + "DCCust.dat") Then
'  DCCustRecLen = Len(DCCustRec(1))
'  TrHandle = FreeFile
'  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = DCCustRecLen
'  NumOfRecs = FileSize(DCPath$ + "DCCust.DAT") \ DCCustRecLen
'  For cnt = 1 To NumOfRecs
'    Get TrHandle, cnt, DCCustRec(1)
'    If (Len(QPTrim$(DCCustRec(1).CUSTNUMB)) <= 0) Then
'        DCCustRec(1).Deleted = "Y"
'    Else
'        DCCustRec(1).Deleted = "N"
'    End If
'    DCCustRec(1).CUSTNUMB = cnt
'    Put #TrHandle, cnt, DCCustRec(1)
'  Next
'  Close #TrHandle
'End If
'End Sub

Private Sub FixTransLink()
Dim DCTranLen  As Integer, NumOfRecs As Long, DCCustRecLen As Integer
Dim DCFile As Integer, PrevTranBal As Double, TrHandle As Integer
Dim cnt As Long, CntT As Long, PrevTranRec As Long
ReDim DCCustRec(1) As DCCustRecType
ReDim DCTransRec(1) As DCTransRecType
If Exist(DCPath$ + "DCCust.dat") And Exist(DCPath$ + "DCTrans.dat") Then
  DCCustRecLen = Len(DCCustRec(1))
  TrHandle = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = DCCustRecLen
  NumOfRecs = FileSize(DCPath$ + "DCCust.DAT") \ DCCustRecLen

  DCTranLen = Len(DCTransRec(1))
  DCFile = FreeFile
  Open DCPath$ + "DCTrans.DAT" For Random Shared As DCFile Len = DCTranLen
  For cnt = 1 To NumOfRecs
    Get TrHandle, cnt, DCCustRec(1)
    PrevTranRec& = DCCustRec(1).FirstTrans
    PrevTranBal = 0
    If PrevTranRec& > 0 Then
      Do While PrevTranRec& > 0
        CntT& = PrevTranRec&
        Get DCFile, CntT&, DCTransRec(1)
        DCTransRec(1).CustomerNumber = DCCustRec(1).CUSTNUMB
        PrevTranRec& = DCTransRec(1).NextTrans
        Put #DCFile, CntT&, DCTransRec(1)
        'DCTransRec(1).NextTrans = 0
      Loop
    End If
  Next
  Close #DCFile
  Close #TrHandle
Else
  MsgBox "Problem", vbOKOnly, "Nothing fixed"
End If

End Sub
Private Sub fixuptheveh()
  Dim DCVehLen As Integer, foundit As Boolean
  Dim CFile As Integer, VFile As Integer
  Dim NumOfCust As Long, NumOfVeh As Long
  Dim NumOfRecs As Long, DCCustLen As Integer
  Dim cnt As Long, numnum As Long, CustRec As Long
  ReDim DCCustRec(1) As DCCustRecType
  ReDim DCVehRec(1 To 2) As DCVehType
  DCCustLen = Len(DCCustRec(1))
  DCVehLen = Len(DCVehRec(1))
  foundit = False
  If Exist(DCPath$ + "DCCust.dat") Then
    CFile = FreeFile
    Open "DCCust.dat" For Random Shared As CFile Len = DCCustLen
    NumOfCust& = LOF(CFile) / DCCustLen
    VFile = FreeFile
    Open "DCVEH.dat" For Random Shared As VFile Len = DCVehLen
    NumOfVeh& = LOF(VFile) / DCVehLen
    FrmShowPctComp.Label1 = "Relinking Vehicles"
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show , Me
    For cnt& = 1 To NumOfVeh&
      FrmShowPctComp.ShowPctComp cnt, NumOfVeh
      Get VFile, cnt&, DCVehRec(1)
      ''If cnt& = 2348 Then Stop
  '    CustRec& = DCVehRec(1).MasterRecord
'''      'just for shanendoah
'''      If CustRec& = 3411 Or CustRec& = 3847 Then
'''          DCVehRec(1).MasterRecord = 0
'''          DCVehRec(1).Active = "N"
'''          Put VFile, cnt&, DCVehRec(1)
'''          GoTo SkipVeh
'''      End If
  '    If CustRec& > 0 Then  'cant use greater than numofcust cause numbers still off
        For numnum = 1 To NumOfCust&
          Get CFile, numnum, DCCustRec(1)
          If CustRec& = Val(DCCustRec(1).CUSTNUMB) Then
            If DCCustRec(1).FirstCar = 0 Then  'if the first car
              DCCustRec(1).FirstCar = cnt&
              DCCustRec(1).LastCar = cnt&
              Put CFile, numnum, DCCustRec(1)
            Else                               'nope not first
              DCCustRec(1).LastCar = cnt&      'set new last car in cust
              Put CFile, numnum, DCCustRec(1) 'put cust back
            End If
            DCVehRec(1).MasterRecord = numnum
            Put VFile, cnt&, DCVehRec(1)
            foundit = True  'not used yet
            GoTo SkipVeh
          End If
        Next
'      Else
'        DCVehRec(1).Active = "N"
'        Put VFile, cnt&, DCVehRec(1)
'    End If
SkipVeh:
    Next
    Close VFile
    Close CFile
  End If
End Sub

