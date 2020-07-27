VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxInterestMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Interest Billing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxInterestMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   435
      Left            =   4020
      TabIndex        =   4
      Top             =   5820
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
      ButtonDesigner  =   "frmVATaxInterestMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   435
      Left            =   4020
      TabIndex        =   3
      Top             =   5145
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
      ButtonDesigner  =   "frmVATaxInterestMenu.frx":0ABC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintTrans 
      Height          =   435
      Left            =   4020
      TabIndex        =   2
      Top             =   4470
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
      ButtonDesigner  =   "frmVATaxInterestMenu.frx":0CA9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditTrans 
      Height          =   435
      Left            =   4020
      TabIndex        =   1
      Top             =   3795
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
      ButtonDesigner  =   "frmVATaxInterestMenu.frx":0E97
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCalcInt 
      Height          =   435
      Left            =   4020
      TabIndex        =   0
      Top             =   3120
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
      ButtonDesigner  =   "frmVATaxInterestMenu.frx":1084
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4020
      TabIndex        =   5
      Top             =   6510
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
      ButtonDesigner  =   "frmVATaxInterestMenu.frx":126A
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
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
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
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8586
      Top             =   2007
      Width           =   971
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8706
      X2              =   9408
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199
      X2              =   2914
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2214
      X2              =   2214
      Y1              =   2127
      Y2              =   8015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX INTEREST BILLING MENU"
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
      TabIndex        =   6
      Top             =   1164
      Width           =   6012
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
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8585
      Top             =   1887
      Width           =   972
   End
End
Attribute VB_Name = "frmVATaxInterestMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim PrincInt As Boolean
  Dim IntInt As Boolean
  Dim AdvColInt As Boolean
  Dim LateListInt As Boolean
  Dim Opt1Int As Boolean
  Dim Opt2Int As Boolean
  Dim Opt3Int As Boolean
  Dim Years() As Integer
  Dim YrCnt As Integer

Private Sub cmdCalcInt_Click()
  frmVATaxCalcInterest.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdClear_Click()
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim PIntTrans As InterestRecType
  Dim NumOfPIRRecs As Long
  Dim PIRHandle As Integer
  Dim x As Long, y As Long
  
  If Exist(TaxRIntFile) And Exist(TaxPIntFile) Then
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
  ElseIf Exist(TaxPIntFile) And Not Exist(TaxRIntFile) Then
    GoSub DeletePersonal
  ElseIf Not Exist(TaxPIntFile) And Exist(TaxRIntFile) Then
    GoSub DeleteReal
  ElseIf Not Exist(TaxPIntFile) And Not Exist(TaxRIntFile) Then
    Call TaxMsg(800, "No unposted interest calculation files currently exist. Delete attempt aborted.")
  End If
  
  Exit Sub
  
DeletePersonal:
  If TaxMsgWOpts(600, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED PERSONAL INTEREST BILLING FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED PERSONAL INTEREST BILLING FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    Close
    KillFile TaxPIntFile
    MainLog ("User deleted unposted personal interest calculation files after being warned about the consequences.")
    Call TaxMsg(800, "All unposted personal interest calculation files have been deleted successfully.")
  End If
  
  Return
  
DeleteReal:
  If TaxMsgWOpts(600, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED REAL INTEREST BILLING FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED REAL INTEREST BILLING FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    Close
    KillFile TaxRIntFile
    MainLog ("User deleted unposted real interest calculation files after being warned about the consequences.")
    Call TaxMsg(800, "All unposted real interest calculation files have been deleted successfully.")
  End If
  
  Return

End Sub

Private Sub cmdEditTrans_Click()
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long
  Dim RealInt As Boolean
  Dim PersInt As Boolean
  
'  If Exist("TAXRINT.DAT") And Exist("TAXPINT.DAT") Then
'    frmVATaxBillPostOpt.Show vbModal
'    If frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
'      GoSub FigurePers
'    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
'      GoSub FigureReal
'    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
'      DoEvents
'      Exit Sub
'    End If
  If Not Exist("TAXRINT.DAT") And Not Exist("TAXPINT.DAT") Then
    Call TaxMsg(900, "There are no real or personal interest calculation records saved.")
  End If
'  ElseIf Exist("TAXRINT.DAT") And Not Exist("TAXPIINT.DAT") Then
'    OpenRInterestRecFile IRHandle, NumOfIRRecs
'    For x = 1 To NumOfIRRecs
'      Get IRHandle, x, IntTrans
'      If IntTrans.DelFlag = False Then
'        Exit For
'      End If
'    Next x
'    If x > NumOfIRRecs Then
'      Call TaxMsg(900, "All real interest calculation records have been deleted.")
'      Close IRHandle
'      Exit Sub
'    End If
'    Close IRHandle
'    frmVATaxEditInt.Show
'    DoEvents
'    Unload Me
'  ElseIf Not Exist("TAXRINT.DAT") And Exist("TAXPINT.DAT") Then
'    OpenPInterestRecFile IRHandle, NumOfIRRecs
'    For x = 1 To NumOfIRRecs
'      Get IRHandle, x, IntTrans
'      If IntTrans.DelFlag = False Then
'        Exit For
'      End If
'    Next x
'    If x > NumOfIRRecs Then
'      Call TaxMsg(900, "All personal interest calculation records have been deleted.")
'      Close IRHandle
'      Exit Sub
'    End If
'    Close IRHandle
    frmVATaxEditInt.Show
    DoEvents
    Unload Me
'  End If
  
  
End Sub

Private Sub cmdExit_Click()
  frmVATaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim PIntTrans As InterestRecType
  Dim NumOfPIRRecs As Long
  Dim PIRHandle As Integer
  Dim x As Long, y As Long
  
  If Exist(TaxRIntFile) Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  End If
  If Exist(TaxPIntFile) Then
    OpenPInterestRecFile PIRHandle, NumOfPIRRecs
  End If
  
  If NumOfIRRecs = 0 And NumOfPIRRecs = 0 Then
    Call TaxMsg(900, "There are no interest calculation records saved.")
    Close IRHandle
    Exit Sub
  Else
    For x = 1 To NumOfIRRecs
      Get IRHandle, x, IntTrans
      If IntTrans.DelFlag = False Then
        Exit For
      End If
    Next x
    For y = 1 To NumOfPIRRecs
      Get PIRHandle, y, PIntTrans
      If PIntTrans.DelFlag = False Then
        Exit For
      End If
    Next y
  End If
  If x > NumOfIRRecs And y > NumOfPIRRecs Then
    Call TaxMsg(900, "There are no interest calculation records saved.")
    Close IRHandle
    Exit Sub
  End If
  
  Close
  frmVATaxInterestPost.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintTrans_Click()
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Long
  Dim PIntTrans As InterestRecType
  Dim NumOfPIRRecs As Long
  Dim PIRHandle As Integer
  
  If Exist(TaxRIntFile) Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  End If
  If Exist(TaxPIntFile) Then
    OpenPInterestRecFile PIRHandle, NumOfPIRRecs
  End If
  
  If NumOfIRRecs = 0 And NumOfPIRRecs = 0 Then
    Call TaxMsg(900, "There are no interest calculation records saved.")
    Close IRHandle
    Close PIRHandle
    Exit Sub
  End If
  
  If NumOfIRRecs > 0 Then
    For x = 1 To NumOfIRRecs
      Get IRHandle, x, IntTrans
      If IntTrans.DelFlag = False Then
        Exit For
      End If
    Next x
    If x > NumOfIRRecs Then
      Call TaxMsg(900, "There are no real interest calculation records saved.")
      Close IRHandle
      Exit Sub
    End If
  End If
  
  If NumOfPIRRecs > 0 Then
    For y = 1 To NumOfPIRRecs
      Get PIRHandle, y, PIntTrans
      If PIntTrans.DelFlag = False Then
        Exit For
      End If
    Next y
    If y > NumOfPIRRecs Then
      Call TaxMsg(900, "There are no personal interest calculation records saved.")
      Close PIRHandle
      Exit Sub
    End If
  End If
  
  Close
  
  frmVATaxReportOpt.Show vbModal
  If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
    Unload frmVATaxReportOpt
    Call PrintGraphics(NumOfIRRecs, NumOfPIRRecs)
  ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
    frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Unload frmVATaxReportOpt
    Call PrintText(NumOfIRRecs, NumOfPIRRecs)
  End If
  
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
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxInterestBilling
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
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub PrintGraphics(RCnt As Long, PCnt As Long)
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Integer
  Dim Town$
  Dim dlm$
  Dim RptHandle As Integer
  Dim RptFile$
  Dim SubRptHandle As Integer
  Dim SubRptFile$
  Dim TotInt As Double
  Dim TotCurrInt As Double
  Dim TotPastInt As Double
  Dim TCnt As Long
  Dim UseR As Boolean
  Dim UseP As Boolean
  
  UseR = False
  UseP = False
  
  If RCnt And PCnt > 0 Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      UseP = True
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      UseR = True
      Unload frmVATaxBillPostOpt
    End If
  ElseIf RCnt > 0 And PCnt = 0 Then
    UseR = True
  ElseIf RCnt = 0 And PCnt > 0 Then
    UseP = True
  Else
    Call TaxMsg(800, "ERROR: There is a problem determining the type on which to report. Please try again.")
    Exit Sub
  End If
  
  dlm$ = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TAXINT.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  If UseP = True Then
    Call GetYears("P")
  ElseIf UseR = True Then
    Call GetYears("R")
  End If
  ReDim YearAmts(1 To YrCnt) As Double
  
  If UseR = True Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  ElseIf UseP = True Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
  End If
  
  For x = 1 To NumOfIRRecs
    Get IRHandle, x, IntRec
    '                   0               1                    2
    Print #RptHandle, Town; dlm; IntRec.CurYear; dlm; IntRec.CustRec; dlm;
    '                            3                           4
    Print #RptHandle, QPTrim$(IntRec.CustName); dlm; IntRec.BillNumber; dlm;
    If IntRec.DelFlag <> 0 Then
      '                        5                  6
      Print #RptHandle, IntRec.TaxYear; dlm; "Deleted"; dlm;
    Else
      '                        5                  6
      Print #RptHandle, IntRec.TaxYear; dlm; IntRec.Amount; dlm;
    End If
    TotInt = OldRound(TotInt + IntRec.Amount)
    If IntRec.TaxYear = TaxMasterRec.RTaxYear Then
      TotCurrInt = OldRound(TotCurrInt + IntRec.Amount)
    Else
      TotPastInt = OldRound(TotPastInt + IntRec.Amount)
    End If
    TCnt = TCnt + 1
    '                    7             8                9
    Print #RptHandle, TotInt; dlm; TotCurrInt; dlm; TotPastInt; dlm; TCnt; dlm;
    
    If UseR = True Then
      '                   10
      Print #RptHandle, "REAL"
    ElseIf UseP = True Then
      '                     10
      Print #RptHandle, "PERSONAL"
    Else
      '                    10
      Print #RptHandle, "UNKNOWN"
    End If
    
    For y = 1 To YrCnt
      If IntRec.TaxYear = Years(y) Then
        YearAmts(y) = OldRound(YearAmts(y) + IntRec.Amount)
        Exit For
      End If
    Next y
  Next x
  
  Close

  SubRptFile$ = "TAXRPTS\SUBTAXINT.RPT"     'Report File Name
  SubRptHandle = FreeFile
  Open SubRptFile$ For Output As #SubRptHandle
  
  For x = 1 To YrCnt
    Print #SubRptHandle, Years(x); dlm; YearAmts(x)
  Next x
  
  Close

  arVATaxInterestRpt.Show

End Sub

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

Private Sub PrintText(RCnt As Long, PCnt As Long)
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Integer
  Dim Town$
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim RptHandle As Integer
  Dim RptFile$, FF$
  Dim TotInt As Double
  Dim TotCurrInt As Double
  Dim TotPastInt As Double
  Dim ThisYear As String
  Dim TCnt As Long
  Dim UseR As Boolean
  Dim UseP As Boolean
  
  UseR = False
  UseP = False
  
  If RCnt And PCnt > 0 Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      UseP = True
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      UseR = True
      Unload frmVATaxBillPostOpt
    End If
  ElseIf RCnt > 0 And PCnt = 0 Then
    UseR = True
  ElseIf RCnt = 0 And PCnt > 0 Then
    UseP = True
  Else
    Call TaxMsg(800, "ERROR: There is a problem determining the type on which to report. Please try again.")
    Exit Sub
  End If
  
  
  MaxLines = 56
  FF$ = Chr(12)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TAXINT.PRN"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  If UseP = True Then
    Call GetYears("P")
  ElseIf UseR = True Then
    Call GetYears("R")
  End If
  ReDim YearAmts(1 To YrCnt) As Double
  
  If UseR = True Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  ElseIf UseP = True Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
  End If
  
  Get IRHandle, 1, IntRec
  ThisYear = CStr(IntRec.CurYear)
  GoSub PrintHeader
  For x = 1 To NumOfIRRecs
    Get IRHandle, x, IntRec
    ThisYear = CStr(IntRec.CurYear)
    If QPTrim$(IntRec.BillNumber) = "" Then IntRec.BillNumber = "UNKNOWN"
    Print #RptHandle, Using$("####0", IntRec.CustRec); Tab(8); QPTrim$(IntRec.CustName);
    Print #RptHandle, Tab(50); Using$("####", IntRec.TaxYear); Tab(56); QPTrim$(IntRec.BillNumber);
    If IntRec.DelFlag <> 0 Then
      Print #RptHandle, Tab(70); "    Deleted"
    Else
      Print #RptHandle, Tab(70); Using$("$###,##0.00", IntRec.Amount)
    End If
    TCnt = TCnt + 1
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    TotInt = OldRound(TotInt + IntRec.Amount)
    If IntRec.TaxYear = TaxMasterRec.RTaxYear Then
      TotCurrInt = OldRound(TotCurrInt + IntRec.Amount)
    Else
      TotPastInt = OldRound(TotPastInt + IntRec.Amount)
    End If
    For y = 1 To YrCnt
      If IntRec.TaxYear = Years(y) Then
        YearAmts(y) = OldRound(YearAmts(y) + IntRec.Amount)
        Exit For
      End If
    Next y
  Next x
  
  Print #RptHandle, FF$
  Page = Page + 1
  Print #RptHandle, Tab(15); "Property Tax Billing: Interest Calculation Register"
  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
  Print #RptHandle, "Date: " + CStr(Date)
  Print #RptHandle, "Current Tax Year: " + ThisYear
  Print #RptHandle, String(80, "-")
  Print #RptHandle, Tab(2); "Total Transactions:     "; Tab(27); Using$("#####0", TCnt)
  Print #RptHandle, Tab(2); "Total Interest Charged: "; Tab(27); Using$("$###,###,##0.00", TotInt)
  Print #RptHandle, Tab(2); "Total Current Interest: "; Tab(27); Using$("$###,###,##0.00", TotCurrInt)
  Print #RptHandle, Tab(2); "Total Past Interest:    "; Tab(27); Using("$###,###,##0.00", TotPastInt)
  Print #RptHandle,
  Print #RptHandle, Tab(2); "Interest Breakdown by Year:"
  Print #RptHandle, Tab(4); "Year"; Tab(12); "Interest Calculation"
  For x = 1 To YrCnt
    Print #RptHandle, Tab(4); Using$("###0", Years(x)); Tab(17); Using$("$###,###,##0.00", YearAmts(x))
  Next x
  
  Print #RptHandle, FF$
  Close

  ViewPrint RptFile, "Interest Calculations", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  If UseR = True Then
    Print #RptHandle, Tab(14); "Property Tax Billing: Real Interest Calculation Register"
  ElseIf UseP = True Then
    Print #RptHandle, Tab(10); "Property Tax Billing: Personal Interest Calculation Register"
  Else
    Print #RptHandle, Tab(18); "Property Tax Billing: Interest Calculation Register"
  End If
  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
  Print #RptHandle, "Date: " + CStr(Date)
  Print #RptHandle, "Current Tax Year: " + ThisYear
  Print #RptHandle, "Acct #:"; Tab(8); "Customer Name"; Tab(48); "Tax Yr"; Tab(57); "Bill #"; Tab(73); "Interest"
  Print #RptHandle, String(80, "-")
  LineCnt = 6
  Return
  
End Sub


