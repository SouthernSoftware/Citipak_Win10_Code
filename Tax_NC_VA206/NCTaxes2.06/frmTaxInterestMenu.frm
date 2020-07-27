VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxInterestMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Interest Billing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxInterestMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   435
      Left            =   4005
      TabIndex        =   3
      Top             =   5010
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
      ButtonDesigner  =   "frmTaxInterestMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintTrans 
      Height          =   435
      Left            =   4005
      TabIndex        =   2
      Top             =   4440
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
      ButtonDesigner  =   "frmTaxInterestMenu.frx":0AB7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditTrans 
      Height          =   435
      Left            =   4005
      TabIndex        =   1
      Top             =   3885
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
      ButtonDesigner  =   "frmTaxInterestMenu.frx":0CA5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCalcInt 
      Height          =   435
      Left            =   4005
      TabIndex        =   0
      Top             =   3330
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
      ButtonDesigner  =   "frmTaxInterestMenu.frx":0E92
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4005
      TabIndex        =   4
      Top             =   6120
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
      ButtonDesigner  =   "frmTaxInterestMenu.frx":1078
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   435
      Left            =   4005
      TabIndex        =   6
      Top             =   5565
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
      ButtonDesigner  =   "frmTaxInterestMenu.frx":1255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1098
      Index           =   1
      Left            =   1493
      Top             =   813
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
      Top             =   2027
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
      TabIndex        =   5
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
Attribute VB_Name = "frmTaxInterestMenu"
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
  frmTaxCalcInterest.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdClear_Click()
  
  If Not Exist("TAXINT.DAT") Then
    Call TaxMsg(900, "No interest calc files currently exist. Delete attempt aborted.")
    Exit Sub
  End If
  
  If TaxMsgWOpts(600, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED INTEREST CALCULATION FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED INTEREST CALCULATION FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    KillFile "TAXINT.DAT"
    KillFile "TAXINTCK.DAT"
    MainLog ("User deleted unposted interest calculations files after being warned about the consequences.")
    Call TaxMsg(900, "All unposted interest calculations files have been deleted successfully.")
  End If

End Sub

Private Sub cmdEditTrans_Click()
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
  
  If NumOfIRRecs = 0 Then
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
  End If
  If x > NumOfIRRecs Then
    Call TaxMsg(900, "There are no interest calculation records saved.")
    Close IRHandle
    Exit Sub
  End If
  
  Close IRHandle
  frmTaxEditInt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  frmTaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxIntRec As InterestRecType
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
  
  If NumOfIRRecs = 0 Then
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
  End If
  If x > NumOfIRRecs Then
    Call TaxMsg(900, "There are no interest calculation records saved.")
    Close IRHandle
    Exit Sub
  End If
  
  frmTaxInterestPost.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintTrans_Click()
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long
  OpenInterestRecFile IRHandle, NumOfIRRecs
  
  If NumOfIRRecs = 0 Then
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
  End If
  If x > NumOfIRRecs Then
    Call TaxMsg(900, "There are no interest calculation records saved.")
    Close IRHandle
    Exit Sub
  End If
  
  Close IRHandle
  frmTaxReportOpt.Show vbModal
  If frmTaxReportOpt.fptxtPrintType.Text = "Graphical" Then
    Unload frmTaxReportOpt
    Call PrintGraphics
  ElseIf frmTaxReportOpt.fptxtPrintType.Text = "Text" Then
    frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Unload frmTaxReportOpt
    Call PrintText
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxInterestMenu.")
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

Private Sub PrintGraphics()
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
  
  dlm$ = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TAXINT.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  Call GetYears
  ReDim YearAmts(1 To YrCnt) As Double
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
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
    If IntRec.TaxYear = TaxMasterRec.TaxYear Then
      TotCurrInt = OldRound(TotCurrInt + IntRec.Amount)
    Else
      TotPastInt = OldRound(TotPastInt + IntRec.Amount)
    End If
    TCnt = TCnt + 1
    '                    7             8                9
    Print #RptHandle, TotInt; dlm; TotCurrInt; dlm; TotPastInt; dlm; TCnt
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

  arTaxInterestRpt.Show

End Sub

Private Sub GetYears()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Integer
  Dim BigNum As Integer
  Dim HoldNum As Integer
  Dim Thisx As Integer
  Dim Nextx As Integer
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
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
    
End Sub

Private Sub PrintText()
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
  
  MaxLines = 56
  FF$ = Chr(12)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TAXINT.PRN"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  Call GetYears
  ReDim YearAmts(1 To YrCnt) As Double
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
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
    If IntRec.TaxYear = TaxMasterRec.TaxYear Then
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
  Print #RptHandle, Tab(15); "Property Tax Billing: Interest Calculation Register"
  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
  Print #RptHandle, "Date: " + CStr(Date)
  Print #RptHandle, "Current Tax Year: " + ThisYear
  Print #RptHandle, "Acct #:"; Tab(8); "Customer Name"; Tab(48); "Tax Yr"; Tab(57); "Bill #"; Tab(73); "Interest"
  Print #RptHandle, String(80, "-")
  LineCnt = 6
  Return
  
End Sub


