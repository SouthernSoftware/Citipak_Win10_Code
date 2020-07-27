VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxPenaltyMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Calculations Menu"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPenaltyMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   450
      Left            =   4020
      TabIndex        =   4
      Top             =   5760
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
      ButtonDesigner  =   "frmVATaxPenaltyMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   450
      Left            =   4020
      TabIndex        =   3
      Top             =   5100
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
      ButtonDesigner  =   "frmVATaxPenaltyMenu.frx":0ABB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintTrans 
      Height          =   435
      Left            =   4020
      TabIndex        =   2
      Top             =   4455
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
      ButtonDesigner  =   "frmVATaxPenaltyMenu.frx":0CA7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditTrans 
      Height          =   435
      Left            =   4020
      TabIndex        =   1
      Top             =   3810
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
      ButtonDesigner  =   "frmVATaxPenaltyMenu.frx":0E94
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCalcPen 
      Height          =   435
      Left            =   4020
      TabIndex        =   0
      Top             =   3150
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
      ButtonDesigner  =   "frmVATaxPenaltyMenu.frx":1080
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4020
      TabIndex        =   5
      Top             =   6435
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
      ButtonDesigner  =   "frmVATaxPenaltyMenu.frx":1265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX PENALTY BILLING MENU"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   6012
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2160
      Y2              =   8048
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2208
      X2              =   2923
      Y1              =   8052
      Y2              =   8052
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8712
      X2              =   9414
      Y1              =   8052
      Y2              =   8052
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   8592
      Top             =   2054
      Width           =   972
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8712
      X2              =   8712
      Y1              =   2160
      Y2              =   8061
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   2100
      Top             =   2052
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1104
      Index           =   1
      Left            =   1500
      Top             =   842
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   720
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   2100
      Top             =   1920
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2220
      Top             =   2148
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1920
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   1
      Left            =   8712
      Top             =   2148
      Width           =   732
   End
End
Attribute VB_Name = "frmVATaxPenaltyMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim PrincInt As Boolean
  Dim IntInt As Boolean
  Dim AdvColInt As Boolean
  Dim LateListInt As Boolean
  Dim Opt1Int As Boolean
  Dim Opt2Int As Boolean
  Dim Opt3Int As Boolean
  Dim Years() As Integer
  Dim YrCnt As Integer

Private Sub cmdCalcPen_Click()
  Dim One As Integer
  Dim AHandle As Integer
  Dim ThisAns$
  Dim Message$
  If Exist(TaxPenRateTblFile) Then
    frmVATaxCalcPenalty.Show
    DoEvents
    Unload Me
  Else
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Jump To Real"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump To Pers"
    Message = "No penalty rate tables have been set up. Would you like to jump to the penalty set up screen now?"
    ThisAns = TaxMsgW3Opts(800, Message, "F5 Jump To Pers", "F10 Jump To Real", "ESC Exit")
    Unload frmVATaxMsgW3Opts
    If ThisAns = "abort" Then
      Exit Sub
    ElseIf ThisAns = "continue" Then
      One = 1
      AHandle = FreeFile
      Open "C:\CPWork\penmenu.dat" For Output As AHandle
      Print #AHandle, One
      Close AHandle
      frmVATaxPenRateSetUpTbl.Show
      DoEvents
      Unload Me
    ElseIf ThisAns = "option" Then
      One = 1
      AHandle = FreeFile
      Open "C:\CPWork\penmenu.dat" For Output As AHandle
      Print #AHandle, One
      Close AHandle
      frmVATaxPPenRateSetUpTbl.Show
      DoEvents
      Unload Me
    End If
  End If
  
End Sub

Private Sub cmdClear_Click()
  Dim IntTrans As PenaltyRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim PIntTrans As PenaltyRecType
  Dim NumOfPIRRecs As Long
  Dim PIRHandle As Integer
  Dim x As Long, y As Long
  
  If Exist(TaxRPenFile) And Exist(TaxPPenFile) Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      GoSub DeletePersonal
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      GoSub DeleteReal
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
  ElseIf Exist(TaxPPenFile) And Not Exist(TaxRPenFile) Then
    GoSub DeletePersonal
  ElseIf Not Exist(TaxPPenFile) And Exist(TaxRPenFile) Then
    GoSub DeleteReal
  ElseIf Not Exist(TaxPPenFile) And Not Exist(TaxRPenFile) Then
    Call TaxMsg(800, "No unposted penalty calculation files currently exist. Delete attempt aborted.")
  End If
  
  Exit Sub
  
DeletePersonal:
  If TaxMsgWOpts(600, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED PERSONAL PENALTY BILLING FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED PERSONAL PENALTY BILLING FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    Close
    KillFile TaxPPenFile
    MainLog ("User deleted unposted personal penalty calculation files after being warned about the consequences.")
    Call TaxMsg(800, "All unposted personal penalty calculation files have been deleted successfully.")
  End If
  
  Return
  
DeleteReal:
  If TaxMsgWOpts(600, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED REAL PENALTY BILLING FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED REAL PENALTY BILLING FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    Close
    KillFile TaxRPenFile
    MainLog ("User deleted unposted real penalty calculation files after being warned about the consequences.")
    Call TaxMsg(800, "All unposted real penalty calculation files have been deleted successfully.")
  End If
  
  Return

End Sub

Private Sub cmdEditTrans_Click()
  Dim NumOfRINRecs As Long
  Dim NumOfPINRecs As Long
  Dim x As Long, y As Long
  Dim PenRec As PenaltyRecType
  Dim RINHandle As Integer
  Dim PINHandle As Integer
  
  OpenRPenRecFile RINHandle, NumOfRINRecs
  OpenPPenRecFile PINHandle, NumOfPINRecs
  
  If NumOfRINRecs = 0 And NumOfPINRecs = 0 Then
    Call TaxMsg(900, "There are no penalty calculation records saved.")
    Close
    Exit Sub
  Else
    For x = 1 To NumOfRINRecs
      Get RINHandle, x, PenRec
      If PenRec.DelFlag = False Then
        Exit For
      End If
    Next x
    For y = 1 To NumOfPINRecs
      Get PINHandle, y, PenRec
      If PenRec.DelFlag = False Then
        Exit For
      End If
    Next y
  End If
  If x > NumOfRINRecs And y > NumOfPINRecs Then
    Call TaxMsg(900, "There are no penalty calculation records saved.")
    Close
    Exit Sub
  End If
  
  Close
  frmVATaxEditPen.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  KillFile "C:\CPWork\penmenu.dat"
  frmVATaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim PenTrans As PenaltyRecType
  Dim NumOfRPNRecs As Long
  Dim RPNHandle As Integer
  Dim NumOfPPNRecs As Long
  Dim PPNHandle As Integer
  Dim x As Long, y As Long
  
  NumOfRPNRecs = 0
  NumOfPPNRecs = 0
  If Exist(TaxRPenFile) Then
    OpenRPenRecFile RPNHandle, NumOfRPNRecs
  End If
  If Exist(TaxPPenFile) Then
    OpenPPenRecFile PPNHandle, NumOfPPNRecs
  End If
  
  If NumOfRPNRecs = 0 And NumOfPPNRecs = 0 Then
    Close RPNHandle
    Close PPNHandle
    Call TaxMsg(900, "There are no penalty calculation records saved.")
    Exit Sub
  Else
    For x = 1 To NumOfRPNRecs
      Get RPNHandle, x, PenTrans
      If PenTrans.DelFlag = False Then
        Exit For
      End If
    Next x
    For y = 1 To NumOfPPNRecs
      Get PPNHandle, y, PenTrans
      If PenTrans.DelFlag = False Then
        Exit For
      End If
    Next y
  End If
  If x > NumOfRPNRecs And y > NumOfPPNRecs Then
    Call TaxMsg(900, "There are no penalty calculation records saved.")
    Close RPNHandle
    Close PPNHandle
    Exit Sub
  End If

  Close RPNHandle
  Close PPNHandle
  
  frmVATaxPenaltyPost.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintTrans_Click()
  Dim PenTrans As PenaltyRecType
  Dim NumOfRPNRecs As Long
  Dim RPNHandle As Integer
  Dim NumOfPPNRecs As Long
  Dim PPNHandle As Integer
  Dim x As Long, y As Long
  
  NumOfRPNRecs = 0
  NumOfPPNRecs = 0
  If Exist(TaxRPenFile) Then
    OpenRPenRecFile RPNHandle, NumOfRPNRecs
  End If
  If Exist(TaxPPenFile) Then
    OpenPPenRecFile PPNHandle, NumOfPPNRecs
  End If
  
  If NumOfRPNRecs = 0 And NumOfPPNRecs = 0 Then
    Close RPNHandle
    Close PPNHandle
    Call TaxMsg(900, "There are no penalty calculation records saved.")
    Exit Sub
  Else
    For x = 1 To NumOfRPNRecs
      Get RPNHandle, x, PenTrans
      If PenTrans.DelFlag = False Then
        Exit For
      End If
    Next x
    For y = 1 To NumOfPPNRecs
      Get PPNHandle, y, PenTrans
      If PenTrans.DelFlag = False Then
        Exit For
      End If
    Next y
  End If
  If x > NumOfRPNRecs And y > NumOfPPNRecs Then
    Call TaxMsg(900, "There are no penalty calculation records saved.")
    Close RPNHandle
    Close PPNHandle
    Exit Sub
  End If

  Close RPNHandle
  Close PPNHandle
  frmVATaxPenaltyTransRpt.Show
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
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxPenaltyBilling
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxInterestMenu.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

'Private Sub PrintGraphics()
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim IntRec As InterestRecType
'  Dim NumOfIRRecs As Long
'  Dim IRHandle As Integer
'  Dim x As Long, y As Integer
'  Dim Town$
'  Dim dlm$
'  Dim RptHandle As Integer
'  Dim RptFile$
'  Dim SubRptHandle As Integer
'  Dim SubRptFile$
'  Dim TotInt As Double
'  Dim TotCurrInt As Double
'  Dim TotPastInt As Double
'  Dim TCnt As Long
'
'  dlm$ = "~"
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'
'  Town$ = QPTrim$(TaxMasterRec.Name)
'
'  RptFile$ = "TAXRPTS\TAXINT.RPT"     'Report File Name
'  RptHandle = FreeFile
'  Open RptFile$ For Output As #RptHandle
'
'  If
'  Call GetYears
'  ReDim YearAmts(1 To YrCnt) As Double
'
'  OpenRInterestRecFile IRHandle, NumOfIRRecs
'  For x = 1 To NumOfIRRecs
'    Get IRHandle, x, IntRec
'    '                   0               1                    2
'    Print #RptHandle, Town; dlm; IntRec.CurYear; dlm; IntRec.CustRec; dlm;
'    '                            3                           4
'    Print #RptHandle, QPTrim$(IntRec.CustName); dlm; IntRec.BillNumber; dlm;
'    If IntRec.DelFlag <> 0 Then
'      '                        5                  6
'      Print #RptHandle, IntRec.TaxYear; dlm; "Deleted"; dlm;
'    Else
'      '                        5                  6
'      Print #RptHandle, IntRec.TaxYear; dlm; IntRec.Amount; dlm;
'    End If
'    TotInt = OldRound(TotInt + IntRec.Amount)
'    If IntRec.TaxYear = TaxMasterRec.RTaxYear Then
'      TotCurrInt = OldRound(TotCurrInt + IntRec.Amount)
'    Else
'      TotPastInt = OldRound(TotPastInt + IntRec.Amount)
'    End If
'    TCnt = TCnt + 1
'    '                    7             8                9
'    Print #RptHandle, TotInt; dlm; TotCurrInt; dlm; TotPastInt; dlm; TCnt
'    For y = 1 To YrCnt
'      If IntRec.TaxYear = Years(y) Then
'        YearAmts(y) = OldRound(YearAmts(y) + IntRec.Amount)
'        Exit For
'      End If
'    Next y
'  Next x
'
'  Close
'
'  SubRptFile$ = "TAXRPTS\SUBTAXINT.RPT"     'Report File Name
'  SubRptHandle = FreeFile
'  Open SubRptFile$ For Output As #SubRptHandle
'
'  For x = 1 To YrCnt
'    Print #SubRptHandle, Years(x); dlm; YearAmts(x)
'  Next x
'
'  Close
'
'  arVATaxInterestRpt.Show
'
'End Sub

Private Sub GetYears(TType As String)
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Integer
  Dim BigNum As Integer
  Dim HoldNum As Integer
  Dim Thisx As Integer
  Dim Nextx As Integer
  
  If TType = "R" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
    ReDim Years(1 To 1) As Integer
    YrCnt = 0
    For x = 1 To NumOfIRRecs
      Get IRHandle, x, IntRec
      If x = 1 Then
        YrCnt = 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = IntRec.TaxYear
      Else
        For y = 1 To YrCnt
          If IntRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > YrCnt Then
          YrCnt = YrCnt + 1
          ReDim Preserve Years(1 To YrCnt) As Integer
          Years(YrCnt) = IntRec.TaxYear
        End If
      End If
    Next x
  
    Close IRHandle
  
    BigNum = -1
    Nextx = 1
    Do
      For x = Nextx To YrCnt
        If Years(x) > BigNum Then
          BigNum = Years(x)
          Thisx = x
        End If
      Next x
      HoldNum = Years(Nextx)
      Years(Nextx) = Years(Thisx)
      Years(Thisx) = HoldNum
      Nextx = Nextx + 1
      If Nextx > YrCnt Then Exit Do
      BigNum = -1
    Loop
  ElseIf TType = "P" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
    ReDim Years(1 To 1) As Integer
    YrCnt = 0
    For x = 1 To NumOfIRRecs
      Get IRHandle, x, IntRec
      If x = 1 Then
        YrCnt = 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = IntRec.TaxYear
      Else
        For y = 1 To YrCnt
          If IntRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > YrCnt Then
          YrCnt = YrCnt + 1
          ReDim Preserve Years(1 To YrCnt) As Integer
          Years(YrCnt) = IntRec.TaxYear
        End If
      End If
    Next x
  
    Close IRHandle
  
    BigNum = -1
    Nextx = 1
    Do
      For x = Nextx To YrCnt
        If Years(x) > BigNum Then
          BigNum = Years(x)
          Thisx = x
        End If
      Next x
      HoldNum = Years(Nextx)
      Years(Nextx) = Years(Thisx)
      Years(Thisx) = HoldNum
      Nextx = Nextx + 1
      If Nextx > YrCnt Then Exit Do
      BigNum = -1
    Loop
  End If
    
End Sub

'Private Sub PrintText()
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim IntRec As InterestRecType
'  Dim NumOfIRRecs As Long
'  Dim IRHandle As Integer
'  Dim x As Long, y As Integer
'  Dim Town$
'  Dim Page As Integer
'  Dim LineCnt As Integer
'  Dim MaxLines As Integer
'  Dim RptHandle As Integer
'  Dim RptFile$, FF$
'  Dim TotInt As Double
'  Dim TotCurrInt As Double
'  Dim TotPastInt As Double
'  Dim ThisYear As String
'  Dim TCnt As Long
'
'  MaxLines = 56
'  FF$ = Chr(12)
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'
'  Town$ = QPTrim$(TaxMasterRec.Name)
'
'  RptFile$ = "TAXRPTS\TAXINT.PRN"     'Report File Name
'  RptHandle = FreeFile
'  Open RptFile$ For Output As #RptHandle
'
'  Call GetYears
'  ReDim YearAmts(1 To YrCnt) As Double
'
'  OpenRInterestRecFile IRHandle, NumOfIRRecs
'  Get IRHandle, 1, IntRec
'  ThisYear = CStr(IntRec.CurYear)
'  GoSub PrintHeader
'  For x = 1 To NumOfIRRecs
'    Get IRHandle, x, IntRec
'    ThisYear = CStr(IntRec.CurYear)
'    If QPTrim$(IntRec.BillNumber) = "" Then IntRec.BillNumber = "UNKNOWN"
'    Print #RptHandle, Using$("####0", IntRec.CustRec); Tab(8); QPTrim$(IntRec.CustName);
'    Print #RptHandle, Tab(50); Using$("####", IntRec.TaxYear); Tab(56); QPTrim$(IntRec.BillNumber);
'    If IntRec.DelFlag <> 0 Then
'      Print #RptHandle, Tab(70); "    Deleted"
'    Else
'      Print #RptHandle, Tab(70); Using$("$###,##0.00", IntRec.Amount)
'    End If
'    TCnt = TCnt + 1
'    LineCnt = LineCnt + 1
'    If LineCnt >= MaxLines Then
'      Print #RptHandle, FF$
'      GoSub PrintHeader
'    End If
'    TotInt = OldRound(TotInt + IntRec.Amount)
'    If IntRec.TaxYear = TaxMasterRec.RTaxYear Then
'      TotCurrInt = OldRound(TotCurrInt + IntRec.Amount)
'    Else
'      TotPastInt = OldRound(TotPastInt + IntRec.Amount)
'    End If
'    For y = 1 To YrCnt
'      If IntRec.TaxYear = Years(y) Then
'        YearAmts(y) = OldRound(YearAmts(y) + IntRec.Amount)
'        Exit For
'      End If
'    Next y
'  Next x
'
'  Print #RptHandle, FF$
'  Page = Page + 1
'  Print #RptHandle, Tab(15); "Property Tax Billing: Interest Calculation Register"
'  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
'  Print #RptHandle, "Date: " + CStr(Date)
'  Print #RptHandle, "Current Tax Year: " + ThisYear
'  Print #RptHandle, String(80, "-")
'  Print #RptHandle, Tab(2); "Total Transactions:     "; Tab(27); Using$("#####0", TCnt)
'  Print #RptHandle, Tab(2); "Total Interest Charged: "; Tab(27); Using$("$###,###,##0.00", TotInt)
'  Print #RptHandle, Tab(2); "Total Current Interest: "; Tab(27); Using$("$###,###,##0.00", TotCurrInt)
'  Print #RptHandle, Tab(2); "Total Past Interest:    "; Tab(27); Using("$###,###,##0.00", TotPastInt)
'  Print #RptHandle,
'  Print #RptHandle, Tab(2); "Interest Breakdown by Year:"
'  Print #RptHandle, Tab(4); "Year"; Tab(12); "Interest Calculation"
'  For x = 1 To YrCnt
'    Print #RptHandle, Tab(4); Using$("###0", Years(x)); Tab(17); Using$("$###,###,##0.00", YearAmts(x))
'  Next x
'
'  Print #RptHandle, FF$
'  Close
'
'  ViewPrint RptFile, "Interest Calculations", True
'
'  Exit Sub
'
'PrintHeader:
'  Page = Page + 1
'  Print #RptHandle, Tab(15); "Property Tax Billing: Interest Calculation Register"
'  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
'  Print #RptHandle, "Date: " + CStr(Date)
'  Print #RptHandle, "Current Tax Year: " + ThisYear
'  Print #RptHandle, "Acct #:"; Tab(8); "Customer Name"; Tab(48); "Tax Yr"; Tab(57); "Bill #"; Tab(73); "Interest"
'  Print #RptHandle, String(80, "-")
'  LineCnt = 6
'  Return
'
'End Sub



