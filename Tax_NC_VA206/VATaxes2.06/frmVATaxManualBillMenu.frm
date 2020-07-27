VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxManualBillMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Tax Billing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxManualBillMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Top             =   5355
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
      ButtonDesigner  =   "frmVATaxManualBillMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintTrans 
      Height          =   435
      Left            =   3960
      TabIndex        =   2
      Top             =   4815
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
      ButtonDesigner  =   "frmVATaxManualBillMenu.frx":0AAE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditTrans 
      Height          =   432
      Left            =   3960
      TabIndex        =   1
      Top             =   4260
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxManualBillMenu.frx":0CA0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterBill 
      Height          =   432
      Left            =   3960
      TabIndex        =   0
      Top             =   3720
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxManualBillMenu.frx":0E84
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   432
      Left            =   3960
      TabIndex        =   4
      Top             =   5916
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxManualBillMenu.frx":106F
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
      Caption         =   "MANUAL TAX BILLING MENU"
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
End
Attribute VB_Name = "frmVATaxManualBillMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim PrintOpt$

Private Sub cmdEditTrans_Click()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  If NumOfTMRecs = 0 Then
    Close
    Call TaxMsg(900, "There are no manual tax bills saved.")
    Exit Sub
  End If
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.Deleted <> True Then
      Exit For
    End If
  Next x
  
  Close TMHandle
  
  If x > NumOfTMRecs Then
    Close
    Call TaxMsg(900, "There are no manual tax bills saved.")
    Exit Sub
  End If
  
  frmVATaxManualBillEdit.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdEnterBill_Click()
  Dim One As Integer
  Dim AHandle As Integer
  
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
'    If Check4PayBatch("R") = True Then
'      frmVATaxUnpostedPaylist.BillType = "R"
'      frmVATaxUnpostedPaylist.Show vbModal
'      Call TaxMsg(800, "An unposted real payment file is ready for posting. Manual billing cannot be conducted until these real payments are posted.")
'      Exit Sub
'    End If
    One = 1
    AHandle = FreeFile
    Open "C:\CPWork\rmanualbill.dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
    frmVATaxManualBillEntry.Show
    DoEvents
    Unload Me
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
'    If Check4PayBatch("P") = True Then
'      frmVATaxUnpostedPaylist.BillType = "P"
'      frmVATaxUnpostedPaylist.Show vbModal
'      Call TaxMsg(800, "An unposted personal payment file is ready for posting. Manual billing cannot be conducted until these personal payments are posted.")
'      Exit Sub
'    End If
    One = 1
    AHandle = FreeFile
    Open "C:\CPWork\pmanualbill.dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
    frmVATaxPManualBillEntry.Show
    DoEvents
    Unload Me
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    DoEvents
    Unload frmVATaxBillPostOpt
    Exit Sub
  End If
End Sub

Private Sub cmdExit_Click()
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer
  
'  If Check4PayBatch("R") = True Then
'    frmVATaxUnpostedPaylist.BillType = "R"
'    frmVATaxUnpostedPaylist.Show vbModal
'    Exit Sub
'  End If
'
'  If Check4PayBatch("P") = True Then
'    frmVATaxUnpostedPaylist.BillType = "P"
'    frmVATaxUnpostedPaylist.Show vbModal
'    Exit Sub
'  End If
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  If NumOfTMRecs = 0 Then
    Close
    Call TaxMsg(900, "There are no manual tax bills saved.")
    Exit Sub
  End If
  
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.Deleted <> True Then
      Exit For
    End If
  Next x
  
  Close TMHandle
  
  If x > NumOfTMRecs Then
    Close
    Call TaxMsg(900, "There are no manual tax bills saved.")
    Exit Sub
  End If
  
  frmVATaxManualPost.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintTrans_Click()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.Deleted = True Then
      GoTo Deleted
    Else
      Exit For
    End If
Deleted:
  Next x
  
  Close TMHandle
  
  If x > NumOfTMRecs Then
    Call TaxMsg(900, "There are no manual bill records saved.")
    Exit Sub
  End If
  
  frmVATaxReportOptWOpt.Show vbModal
  If frmVATaxReportOptWOpt.fptxtPrintType.Text = "Graphical Name" Then
    Unload frmVATaxReportOptWOpt
    PrintOpt = "Graphical Name"
    Call PrintGraphics
  ElseIf frmVATaxReportOptWOpt.fptxtPrintType.Text = "Graphical Number" Then
    Unload frmVATaxReportOptWOpt
    PrintOpt = "Graphical Number"
    Call PrintGraphics
  ElseIf frmVATaxReportOptWOpt.fptxtPrintType.Text = "Text Name" Then
    Unload frmVATaxReportOptWOpt
    PrintOpt = "Text Name"
    Call TaxMsg(900, "Pitch 12 is recommended for this report.")
    Call PrintText
  ElseIf frmVATaxReportOptWOpt.fptxtPrintType.Text = "Text Number" Then
    Unload frmVATaxReportOptWOpt
    PrintOpt = "Text Number"
    Call TaxMsg(900, "Pitch 12 is recommended for this report.")
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
  PrintOpt$ = "None"
  Me.HelpContextID = hlpManualTaxBilling
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxManualBillMenu.")
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
  Dim RptFile$
  Dim RptHandle As Integer
  Dim Sub1RptFile$
  Dim Sub1RptHandle As Integer
  Dim dlm$
  Dim Town$
  Dim TaxSURec As TaxMasterType
  Dim TXMHandle As Integer
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer, y As Integer
  Dim ECnt As Integer
  Dim ThisLong As Long
  Dim ThisLongS$
  Dim ThisRec$
  Dim Nextx As Integer
  Dim BigName$, SmallName$
  Dim BigNum As Long, SmallNum As Long
  Dim LilNum$
  Dim HoldNum$
  Dim HoldLong As Long
  Dim Thisx As Integer
  Dim HoldInt As Integer
  Dim Opt1Desc As String * 13
  Dim Opt2Desc As String * 13
  Dim Opt3Desc As String * 13
  Dim POpt1Desc As String * 13
  Dim POpt2Desc As String * 13
  Dim POpt3Desc As String * 13
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfSortRealRecs As Long
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfSortPersRecs As Long
  Dim ThisClass$
  Dim BillCnt As Integer
  Dim TotCustAmt As Double
  Dim GTotPrinc As Double
  Dim GTotRInt As Double
  Dim GTotAdvCol As Double
  Dim GTotLateList As Double
  Dim GTotRPen As Double
  Dim GTotOpt1 As Double
  Dim GTotOpt2 As Double
  Dim GTotOpt3 As Double
  Dim GTotPOpt1 As Double
  Dim GTotPOpt2 As Double
  Dim GTotPOpt3 As Double
  Dim GTotPers As Double
  Dim GTotPInt As Double
  Dim GTotMT As Double
  Dim GTotMC As Double
  Dim GTotFE As Double
  Dim GTotMH As Double
  Dim GTotPPen As Double
  Dim GTotal As Double
  Dim RGTotal As Double
  Dim PGTotal As Double
  Dim CombCnt As Integer
  Dim RealCnt As Integer
  Dim PersCnt As Integer
  Dim CombNewYear As Integer
  Dim RealNewYear As Integer
  Dim PersNewYear As Integer
  Dim CombYrCnt As Integer
  Dim RealYrCnt As Integer
  Dim PersYrCnt As Integer
  Dim LilYear As Integer
  Dim HoldRec As Integer
  Dim YrTotPrinc As Double
  Dim YrTotRInt As Double
  Dim YrTotAdvCol As Double
  Dim YrTotLateList As Double
  Dim YrTotRPen As Double
  Dim YrTotOpt1 As Double
  Dim YrTotOpt2 As Double
  Dim YrTotOpt3 As Double
  Dim YrTotPers As Double
  Dim YrTotPInt As Double
  Dim YrTotMT As Double
  Dim YrTotMC As Double
  Dim YrTotFE As Double
  Dim YrTotMH As Double
  Dim YrTotPPen As Double
  Dim YrTotPOpt1 As Double
  Dim YrTotPOpt2 As Double
  Dim YrTotPOpt3 As Double
  Dim YrGTot As Double
  Dim YrRGTot As Double
  Dim YrPGTot As Double
  Dim CombYear As Integer
  Dim ThisECnt As Integer
  Dim ThisCombCnt As Integer
  Dim ThisPersCnt As Integer
  Dim ThisRealCnt As Integer
  Dim RealYear As Integer
  Dim PersYear As Integer
  
  OpenTaxSetUpFile TXMHandle
  Get TXMHandle, 1, TaxSURec
  Close TXMHandle
  
  RSet Opt1Desc$ = QPTrim$(TaxSURec.OptRev1)
  RSet Opt2Desc$ = QPTrim$(TaxSURec.OptRev2)
  RSet Opt3Desc$ = QPTrim$(TaxSURec.OptRev3)
  RSet POpt1Desc$ = QPTrim$(TaxSURec.POptRev1)
  RSet POpt2Desc$ = QPTrim$(TaxSURec.POptRev2)
  RSet POpt3Desc$ = QPTrim$(TaxSURec.POptRev3)
  
  Town = QPTrim$(TaxSURec.Name)
  dlm = "~"
  RptFile$ = "TAXRPTS\TXMANEDT.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  ECnt = 0
  ReDim PrintList(1 To 1) As String
  ReDim PrintLong(1 To 1) As Long
  ReDim PrintRec(1 To 1) As Integer
  ReDim SortRealRec(1 To 1) As Integer
  ReDim SortPersRec(1 To 1) As Integer
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  
  BigName = ""
  BigNum = 0
  RealCnt = 0
  PersCnt = 0
  RealNewYear = 0
  PersNewYear = 0
  
  If PrintOpt = "Graphical Name" Then
    For x = 1 To NumOfTMRecs
      Get TMHandle, x, TaxMRec
      If TaxMRec.Deleted = True Then
        GoTo Deleted
      Else
        ECnt = ECnt + 1
        ThisLongS = CStr(TaxMRec.Account)
        ThisRec = QPTrim$(TaxMRec.SName) + ThisLongS
        ReDim Preserve PrintList(1 To ECnt) As String
        PrintList(ECnt) = ThisRec
        ReDim Preserve PrintRec(1 To ECnt) As Integer
        PrintRec(ECnt) = x
        If ThisRec > BigName Then
          BigName = ThisRec
        End If
      End If
Deleted:
    Next x
    GoSub SortByName
  Else
    For x = 1 To NumOfTMRecs
      Get TMHandle, x, TaxMRec
      If TaxMRec.Deleted = True Then
        GoTo Deleted2
      Else
        ECnt = ECnt + 1
        ThisLong = TaxMRec.Account
        ReDim Preserve PrintLong(1 To ECnt) As Long
        PrintLong(ECnt) = ThisLong
        ReDim Preserve PrintRec(1 To ECnt) As Integer
        PrintRec(ECnt) = x
        If ThisLong > BigNum Then
          BigNum = ThisLong
        End If
      End If
Deleted2:
    Next x
    GoSub SortByAcctNum
  End If
 
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
      If Mid(TaxMRec.BillType, 1, 1) = "R" Then
        If TaxMRec.TaxYear > RealNewYear Then
          RealNewYear = TaxMRec.TaxYear
        End If
        RealCnt = RealCnt + 1
        ReDim Preserve SortRealRec(1 To RealCnt) As Integer
        SortRealRec(RealCnt) = PrintRec(x)
      ElseIf Mid(TaxMRec.BillType, 1, 1) = "P" Then
        If TaxMRec.TaxYear > PersNewYear Then
          PersNewYear = TaxMRec.TaxYear
        End If
        PersCnt = PersCnt + 1
        ReDim Preserve SortPersRec(1 To PersCnt) As Integer
        SortPersRec(PersCnt) = PrintRec(x)
      End If
  Next x
 
  OpenRealPropFile RHandle, NumOfSortRealRecs
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
    If TaxMRec.Class = "P" Then GoTo Skip1
    TotCustAmt = OldRound(TaxMRec.TaxAmount + TaxMRec.IntAmount + TaxMRec.AdColAmount + TaxMRec.LateList)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3 + TaxMRec.Penalty)
    GTotPrinc = OldRound(GTotPrinc + TaxMRec.TaxAmount)
    GTotRInt = OldRound(GTotRInt + TaxMRec.IntAmount)
    GTotAdvCol = OldRound(GTotAdvCol + TaxMRec.AdColAmount)
    GTotLateList = OldRound(GTotLateList + TaxMRec.LateList)
    GTotOpt1 = OldRound(GTotOpt1 + TaxMRec.OptRev1)
    GTotOpt2 = OldRound(GTotOpt2 + TaxMRec.OptRev2)
    GTotOpt3 = OldRound(GTotOpt3 + TaxMRec.OptRev3)
    GTotRPen = OldRound(GTotRPen + TaxMRec.Penalty)
    RGTotal = OldRound(RGTotal + TaxMRec.TaxAmount + TaxMRec.IntAmount + TaxMRec.AdColAmount + TaxMRec.LateList + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3 + TaxMRec.Penalty)
    GTotal = OldRound(GTotal + TaxMRec.TaxAmount + TaxMRec.IntAmount + TaxMRec.AdColAmount + TaxMRec.LateList + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3 + TaxMRec.Penalty)
    If TaxMRec.Class = "M" Then
      ThisClass = "MOCK"
    ElseIf TaxMRec.Class = "R" Then
      ThisClass = "REAL"
    Else
      ThisClass = "NA"
    End If
    '                   0                1                          2
    Print #RptHandle, Town$; dlm; TaxMRec.Account; dlm; QPTrim$(TaxMRec.SName); dlm;
    '                        3                   4                  5              6
    Print #RptHandle, TaxMRec.BillNum; dlm; ThisClass; dlm; Opt1Desc; dlm; Opt2Desc; dlm;
    '                     7                    8                      9                        10
    Print #RptHandle, Opt3Desc$; dlm; TaxMRec.TaxAmount; dlm; TaxMRec.IntAmount; dlm; TaxMRec.AdColAmount; dlm;
    '                        11                    12                    13
    Print #RptHandle, TaxMRec.LateList; dlm; TaxMRec.OptRev1; dlm; TaxMRec.OptRev2; dlm;
    '                        14                   15                      16
    Print #RptHandle, TaxMRec.OptRev3; dlm; TaxMRec.TaxYear; dlm; MakeRegDate(TaxMRec.TransDate); dlm;
   
    If TaxMRec.RealRec > 0 Then
      Get RHandle, TaxMRec.RealRec, RealPropRec
      '                              17
      Print #RptHandle, QPTrim$(RealPropRec.RealPin); dlm;
    Else
      '                  17
      Print #RptHandle, "NA"; dlm;
    End If
    '                      18              19             20              21
    Print #RptHandle, TotCustAmt; dlm; GTotPrinc; dlm; GTotRInt; dlm; GTotAdvCol; dlm;
    '                      22               23             24             25            26          27
    Print #RptHandle, GTotLateList; dlm; GTotOpt1; dlm; GTotOpt2; dlm; GTotOpt3; dlm; GTotal; dlm; ECnt; dlm;
    '                        28                 29                30                31            32
    Print #RptHandle, TaxMRec.FarmEquip; dlm; GTotFE; dlm; TaxMRec.MobHomes; dlm; GTotMH; dlm; PGTotal; dlm;
    '                   33            34             35              36              37              38
    Print #RptHandle, RealCnt; dlm; PersCnt; dlm; GTotPOpt1; dlm; GTotPOpt2; dlm; GTotPOpt3; dlm; GTotPers; dlm;
    '                   39             40            41             42            43              44
    Print #RptHandle, GTotPInt; dlm; GTotMT; dlm; GTotPPen; dlm; GTotRPen; dlm; RGTotal; dlm; POpt1Desc; dlm;
    '                     45               46             47              48
    Print #RptHandle, POpt2Desc$; dlm; POpt3Desc$; dlm; GTotMC; dlm; TaxMRec.Penalty
Skip1:
  Next x
 
  Close RHandle
  
  OpenPersPropFile PHandle, NumOfSortPersRecs
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
    If TaxMRec.Class <> "P" Then GoTo Skip2
    TotCustAmt = OldRound(TaxMRec.Personal + TaxMRec.IntAmount + TaxMRec.MachTools + TaxMRec.MerchCap)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.FarmEquip + TaxMRec.MobHomes + TaxMRec.Penalty)
    GTotPers = OldRound(GTotPers + TaxMRec.Personal)
    GTotPInt = OldRound(GTotPInt + TaxMRec.IntAmount)
    GTotMT = OldRound(GTotMT + TaxMRec.MachTools)
    GTotMC = OldRound(GTotMC + TaxMRec.MerchCap)
    GTotFE = OldRound(GTotFE + TaxMRec.FarmEquip)
    GTotMH = OldRound(GTotMH + TaxMRec.MobHomes)
    GTotPOpt1 = OldRound(GTotPOpt1 + TaxMRec.OptRev1)
    GTotPOpt2 = OldRound(GTotPOpt2 + TaxMRec.OptRev2)
    GTotPOpt3 = OldRound(GTotPOpt3 + TaxMRec.OptRev3)
    GTotPPen = OldRound(GTotPPen + TaxMRec.Penalty)
    PGTotal = OldRound(PGTotal + TaxMRec.Personal + TaxMRec.IntAmount + TaxMRec.MachTools + TaxMRec.MerchCap + TaxMRec.FarmEquip + TaxMRec.MobHomes + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3 + TaxMRec.Penalty)
    GTotal = OldRound(GTotal + TaxMRec.Personal + TaxMRec.IntAmount + TaxMRec.MachTools + TaxMRec.MerchCap + TaxMRec.FarmEquip + TaxMRec.MobHomes + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3 + TaxMRec.Penalty)
    
    If TaxMRec.Class = "P" Then
      ThisClass = "PERSONAL"
    Else
      ThisClass = "NA"
    End If
    '                   0                1                          2
    Print #RptHandle, Town$; dlm; TaxMRec.Account; dlm; QPTrim$(TaxMRec.SName); dlm;
    '                        3                  4               5               6
    Print #RptHandle, TaxMRec.BillNum; dlm; ThisClass; dlm; Opt1Desc; dlm; Opt2Desc; dlm;
    '                     7                    8                      9                        10
    Print #RptHandle, Opt3Desc$; dlm; TaxMRec.Personal; dlm; TaxMRec.IntAmount; dlm; TaxMRec.MachTools; dlm;
    '                        11                    12                    13
    Print #RptHandle, TaxMRec.MerchCap; dlm; TaxMRec.OptRev1; dlm; TaxMRec.OptRev2; dlm;
    '                        14                   15                             16
    Print #RptHandle, TaxMRec.OptRev3; dlm; TaxMRec.TaxYear; dlm; MakeRegDate(TaxMRec.TransDate); dlm;
   
    If TaxMRec.PersRec > 0 Then
      Get PHandle, TaxMRec.PersRec, PersPropRec
      '                             17
      Print #RptHandle, QPTrim$(PersPropRec.PropPin); dlm;
    Else
      '                  17
      Print #RptHandle, "NA"; dlm;
    End If
    '                     18              19             20              21
    Print #RptHandle, TotCustAmt; dlm; GTotPrinc; dlm; GTotRInt; dlm; GTotAdvCol; dlm;
    '                      22               23             24            25             26          27
    Print #RptHandle, GTotLateList; dlm; GTotOpt1; dlm; GTotOpt2; dlm; GTotOpt3; dlm; GTotal; dlm; ECnt; dlm;
    '                        28                 29                30                31            32
    Print #RptHandle, TaxMRec.FarmEquip; dlm; GTotFE; dlm; TaxMRec.MobHomes; dlm; GTotMH; dlm; PGTotal; dlm;
    '                   33            34             35              36              37              38
    Print #RptHandle, RealCnt; dlm; PersCnt; dlm; GTotPOpt1; dlm; GTotPOpt2; dlm; GTotPOpt3; dlm; GTotPers; dlm;
    '                   39             40            41             42            43              44
    Print #RptHandle, GTotPInt; dlm; GTotMT; dlm; GTotPPen; dlm; GTotRPen; dlm; RGTotal; dlm; POpt1Desc; dlm;
    '                     45               46             47               48
    Print #RptHandle, POpt2Desc$; dlm; POpt3Desc$; dlm; GTotMC; dlm; TaxMRec.Penalty
Skip2:
  Next x
 
  Close PHandle
  Close RptHandle
  
  GoSub PrintSub1
  
  Close TMHandle
  Close Sub1RptHandle
  
  arVATaxManEditRpt.Show
 
  Exit Sub
 
PrintSub1:
  Sub1RptFile$ = "TAXRPTS\TXMANSUB1.RPT"     'Report File Name
  Sub1RptHandle = FreeFile
  Open Sub1RptFile$ For Output As #Sub1RptHandle
 
  If RealCnt = 0 Then GoTo NoRealCnt

  LilYear = RealNewYear + 1
  Nextx = 1
  Do
    For x = Nextx To RealCnt
      Get TMHandle, SortRealRec(x), TaxMRec
      If TaxMRec.TaxYear < LilYear Then
        LilYear = TaxMRec.TaxYear
        Thisx = x
      End If
    Next x
    HoldRec = SortRealRec(Nextx)
    SortRealRec(Nextx) = SortRealRec(Thisx)
    SortRealRec(Thisx) = HoldRec
    If Nextx = 1 Then RealYear = LilYear
    Nextx = Nextx + 1
    If Nextx > RealCnt Then
      Exit Do
    End If
    LilYear = RealNewYear + 1
  Loop

  ReDim RealRec(1 To RealCnt) As Integer
  For x = 1 To RealCnt
    RealRec(x) = SortRealRec(x)
  Next x

  ReDim Years(1 To 1) As Integer

  For x = 1 To RealCnt
    Get TMHandle, RealRec(x), TaxMRec
      If x = 1 Then
        RealYrCnt = RealYrCnt + 1
        ReDim Preserve Years(1 To RealYrCnt) As Integer
        Years(RealYrCnt) = TaxMRec.TaxYear
      Else
        For y = 1 To RealYrCnt
          If TaxMRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > RealYrCnt Then
          RealYrCnt = RealYrCnt + 1
          ReDim Preserve Years(1 To RealYrCnt) As Integer
          Years(RealYrCnt) = TaxMRec.TaxYear
        End If
      End If
  Next x

NoRealCnt:

  If PersCnt = 0 Then GoTo NoPersCnt
 
  LilYear = PersNewYear + 1
  Nextx = 1
  Do
    For x = Nextx To PersCnt
      Get TMHandle, SortPersRec(x), TaxMRec
      If TaxMRec.TaxYear < LilYear Then
        LilYear = TaxMRec.TaxYear
        Thisx = x
      End If
    Next x
    HoldRec = SortPersRec(Nextx)
    SortPersRec(Nextx) = SortPersRec(Thisx)
    SortPersRec(Thisx) = HoldRec
    If Nextx = 1 Then PersYear = LilYear
    Nextx = Nextx + 1
    If Nextx > PersCnt Then
      Exit Do
    End If
    LilYear = PersNewYear + 1
  Loop
 
  ReDim PersRec(1 To PersCnt) As Integer
  For x = 1 To PersCnt
    PersRec(x) = SortPersRec(x)
  Next x
  
  ReDim Years(1 To 1) As Integer

  For x = 1 To PersCnt
    Get TMHandle, PersRec(x), TaxMRec
      If x = 1 Then
        PersYrCnt = PersYrCnt + 1
        ReDim Preserve Years(1 To PersYrCnt) As Integer
        Years(PersYrCnt) = TaxMRec.TaxYear
      Else
        For y = 1 To PersYrCnt
          If TaxMRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > PersYrCnt Then
          PersYrCnt = PersYrCnt + 1
          ReDim Preserve Years(1 To PersYrCnt) As Integer
          Years(PersYrCnt) = TaxMRec.TaxYear
        End If
      End If
  Next x
  
NoPersCnt:
  
  YrTotPrinc = 0
  YrTotRInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotRPen = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrGTot = 0
  YrRGTot = 0
  For x = 1 To RealCnt
    Get TMHandle, RealRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotRInt = OldRound(YrTotRInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotRPen = OldRound(YrTotRPen + TaxMRec.Penalty)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrRGTot = OldRound(YrTotPrinc + YrTotRInt + YrTotAdvCol + YrTotLateList + YrTotRPen + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    YrGTot = OldRound(YrTotPrinc + YrTotRInt + YrTotAdvCol + YrTotLateList + YrTotRPen + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    If RealYrCnt > 1 And x < RealCnt Then
      Get TMHandle, RealRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> RealYear Then
        '                        0                 1              2               3                  4
        Print #Sub1RptHandle, RealYear; dlm; YrTotPrinc; dlm; YrTotRInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
        '                        5                 6              7              8          9
        Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrRGTot; dlm; "R"; dlm;
        '                        10             11             12             13          14      15
        Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; YrTotRPen; dlm; 0; dlm; 0
        RealYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotRInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotRPen = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrRGTot = 0
      End If
    ElseIf x = RealCnt Then
      '                        0                 1              2               3                  4
      Print #Sub1RptHandle, RealYear; dlm; YrTotPrinc; dlm; YrTotRInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
      '                        5                 6              7              8          9
      Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrRGTot; dlm; "R"; dlm;
        '                      10             11             12              13          14      15
      Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; YrTotRPen; dlm; 0; dlm; 0
  End If
  Next x
  
  YrTotPInt = 0
  YrTotMT = 0
  YrTotMC = 0
  YrTotFE = 0
  YrTotMH = 0
  YrTotPPen = 0
  YrTotPOpt1 = 0
  YrTotPOpt2 = 0
  YrTotPOpt3 = 0
  YrPGTot = 0
  For x = 1 To PersCnt
    Get TMHandle, PersRec(x), TaxMRec
    YrTotPers = OldRound(YrTotPers + TaxMRec.Personal)
    YrTotPInt = OldRound(YrTotPInt + TaxMRec.IntAmount)
    YrTotMT = OldRound(YrTotMT + TaxMRec.MachTools)
    YrTotMC = OldRound(YrTotMC + TaxMRec.MerchCap)
    YrTotPPen = OldRound(YrTotPPen + TaxMRec.Penalty)
    YrTotPOpt1 = OldRound(YrTotPOpt1 + TaxMRec.OptRev1)
    YrTotPOpt2 = OldRound(YrTotPOpt2 + TaxMRec.OptRev2)
    YrTotPOpt3 = OldRound(YrTotPOpt3 + TaxMRec.OptRev3)
    YrTotFE = OldRound(YrTotFE + TaxMRec.FarmEquip)
    YrTotMH = OldRound(YrTotMH + TaxMRec.MobHomes)
    YrPGTot = OldRound(YrTotPers + YrTotPInt + YrTotMT + YrTotMC + YrTotFE + YrTotMH + YrTotPOpt1 + YrTotPOpt2 + YrTotPOpt3 + YrTotPPen)
    YrGTot = OldRound(YrTotPers + YrTotPInt + YrTotMT + YrTotMC + YrTotFE + YrTotMH + YrTotPOpt1 + YrTotPOpt2 + YrTotPOpt3 + YrTotPPen)
    If PersYrCnt > 1 And x < PersCnt Then
      Get TMHandle, PersRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> PersYear And x < PersCnt Then
        '                        0                 1              2             3             4
        Print #Sub1RptHandle, PersYear; dlm; YrTotPers; dlm; YrTotPInt; dlm; YrTotMT; dlm; YrTotMC; dlm;
        '                        5                 6               7               8            9
        Print #Sub1RptHandle, YrTotPOpt1; dlm; YrTotPOpt2; dlm; YrTotPOpt3; dlm; YrPGTot; dlm; "P"; dlm;
        '                        10               11             12              13              14           15
        Print #Sub1RptHandle, POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm; YrTotPPen; dlm; YrTotFE; dlm; YrTotMH
        PersYear = TaxMRec.TaxYear
        YrTotPers = 0
        YrTotPInt = 0
        YrTotMT = 0
        YrTotMC = 0
        YrTotFE = 0
        YrTotMH = 0
        YrTotPPen = 0
        YrTotPOpt1 = 0
        YrTotPOpt2 = 0
        YrTotPOpt3 = 0
        YrPGTot = 0
      End If
    ElseIf x = PersCnt Then
      '                        0                 1              2             3             4
      Print #Sub1RptHandle, PersYear; dlm; YrTotPers; dlm; YrTotPInt; dlm; YrTotMT; dlm; YrTotMC; dlm;
      '                        5                 6                7               8           9
      Print #Sub1RptHandle, YrTotPOpt1; dlm; YrTotPOpt2; dlm; YrTotPOpt3; dlm; YrPGTot; dlm; "P"; dlm;
      '                         10             11               12             13              14            15
      Print #Sub1RptHandle, POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm; YrTotPPen; dlm; YrTotFE; dlm; YrTotMH
    End If
  Next x
  
  Return
  
SortByName:
  BigName = BigName + "z"
  SmallName = BigName
  Nextx = 1
  Thisx = 0
  Do
    For x = Nextx To ECnt
      If PrintList(x) < SmallName Then
        SmallName = PrintList(x)
        Thisx = x
      End If
    Next x
    
    HoldNum = PrintList(Nextx)
    HoldInt = PrintRec(Nextx)
    PrintList(Nextx) = PrintList(Thisx)
    PrintRec(Nextx) = PrintRec(Thisx)
    PrintList(Thisx) = HoldNum
    PrintRec(Thisx) = HoldInt
    Nextx = Nextx + 1
    If Nextx > ECnt Then Exit Do
    Thisx = 0
    SmallName = BigName
  Loop

  Return
  
SortByAcctNum:
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  Thisx = 0
  Do
    For x = Nextx To ECnt
      If PrintLong(x) < SmallNum Then
        SmallNum = PrintLong(x)
        Thisx = x
      End If
    Next x
    HoldLong = PrintLong(Nextx)
    HoldInt = PrintRec(Nextx)
    PrintLong(Nextx) = PrintLong(Thisx)
    PrintRec(Nextx) = PrintRec(Thisx)
    PrintLong(Thisx) = HoldLong
    PrintRec(Thisx) = HoldInt
    Nextx = Nextx + 1
    If Nextx > ECnt Then Exit Do
    Thisx = 0
    SmallNum = BigNum
  Loop
  
  Return
  
End Sub
Private Sub PrintText()
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim TaxSURec As TaxMasterType
  Dim TXMHandle As Integer
  Dim Town$
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer, y As Integer
  Dim ECnt As Integer
  Dim ThisLong As Long
  Dim ThisLongS$
  Dim ThisRec$
  Dim Nextx As Integer
  Dim BigName$, SmallName$
  Dim BigNum As Long, SmallNum As Long
  Dim LilNum$
  Dim HoldNum$
  Dim HoldLong As Long
  Dim Thisx As Integer
  Dim HoldInt As Integer
  Dim Opt1Desc As String * 10
  Dim Opt2Desc As String * 10
  Dim Opt3Desc As String * 10
  Dim POpt1Desc As String * 10
  Dim POpt2Desc As String * 10
  Dim POpt3Desc As String * 10
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfSortRealRecs As Long
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfSortPersRecs As Long
  Dim ThisClass$
  Dim BillCnt As Integer
  Dim TotCustAmt As Double
  Dim GTotPrinc As Double
  Dim GTotRInt As Double
  Dim GTotPInt As Double
  Dim GTotRPen As Double
  Dim GTotPPen As Double
  Dim GTotAdvCol As Double
  Dim GTotLateList As Double
  Dim GTotPers As Double
  Dim GTotMT As Double
  Dim GTotMC As Double
  Dim GTotFE As Double
  Dim GTotMH As Double
  Dim RGTotOpt1 As Double
  Dim RGTotOpt2 As Double
  Dim RGTotOpt3 As Double
  Dim PGTotOpt1 As Double
  Dim PGTotOpt2 As Double
  Dim PGTotOpt3 As Double
  Dim PGTotal As Double
  Dim RGTotal As Double
  Dim GTotal As Double
  Dim RealCnt As Integer
  Dim PersCnt As Integer
  Dim RealNewYear As Integer
  Dim PersNewYear As Integer
  Dim RealYrCnt As Integer
  Dim PersYrCnt As Integer
  Dim LilYear As Integer
  Dim HoldRec As Integer
  Dim YrTotPrinc As Double
  Dim YrTotRInt As Double
  Dim YrTotAdvCol As Double
  Dim YrTotLateList As Double
  Dim YrTotOpt1 As Double
  Dim YrTotOpt2 As Double
  Dim YrTotOpt3 As Double
  Dim YrTotPOpt1 As Double
  Dim YrTotPOpt2 As Double
  Dim YrTotPOpt3 As Double
  Dim YrGTot As Double
  Dim YrPGTot As Double
  Dim YrRGTot As Double
  Dim ThisECnt As Integer
  Dim ThisPersCnt As Integer
  Dim ThisRealCnt As Integer
  Dim RealYear As Integer
  Dim PersYear As Integer
  Dim Page As Integer
  Dim ThisAcct As Long
  Dim ROverAllPrinted As Boolean
  Dim POverAllPrinted As Boolean
  Dim YrTotRPen As Double
  Dim YrTotPPen As Double
  Dim YrTotPers As Double
  Dim YrTotMT As Double
  Dim YrTotMC As Double
  Dim YrTotFE As Double
  Dim YrTotMH As Double
  Dim YrTotPInt As Double
  Dim YTotPPen As Double
  
  ROverAllPrinted = False
  POverAllPrinted = False
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  
  OpenTaxSetUpFile TXMHandle
  Get TXMHandle, 1, TaxSURec
  Close TXMHandle
  
  Town = QPTrim$(TaxSURec.Name)
  Opt1Desc$ = QPTrim$(TaxSURec.OptRev1)
  Opt2Desc$ = QPTrim$(TaxSURec.OptRev2)
  Opt3Desc$ = QPTrim$(TaxSURec.OptRev3)
  POpt1Desc$ = QPTrim$(TaxSURec.POptRev1)
  POpt2Desc$ = QPTrim$(TaxSURec.POptRev2)
  POpt3Desc$ = QPTrim$(TaxSURec.POptRev3)
  
  RptFile$ = "TaxManEdit.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  ReDim PrintList(1 To 1) As String
  ReDim PrintLong(1 To 1) As Long
  ReDim PrintRec(1 To 1) As Integer
  ReDim SortRealRec(1 To 1) As Integer
  ReDim SortPersRec(1 To 1) As Integer
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  
  BigName = ""
  BigNum = 0
  RealCnt = 0
  PersCnt = 0
  RealNewYear = 0
  PersNewYear = 0
  
  If PrintOpt = "Text Name" Then
    For x = 1 To NumOfTMRecs
      Get TMHandle, x, TaxMRec
      If TaxMRec.Deleted = True Then
        GoTo Deleted
      Else
        ECnt = ECnt + 1
        ThisLongS = CStr(TaxMRec.Account)
        ThisRec = QPTrim$(TaxMRec.SName) + ThisLongS
        ReDim Preserve PrintList(1 To ECnt) As String
        PrintList(ECnt) = ThisRec
        ReDim Preserve PrintRec(1 To ECnt) As Integer
        PrintRec(ECnt) = x
        If ThisRec > BigName Then
          BigName = ThisRec
        End If
      End If
Deleted:
    Next x
    GoSub SortByName
  Else
    For x = 1 To NumOfTMRecs
      Get TMHandle, x, TaxMRec
      If TaxMRec.Deleted = True Then
        GoTo Deleted2
      Else
        ECnt = ECnt + 1
        ThisLong = TaxMRec.Account
        ReDim Preserve PrintLong(1 To ECnt) As Long
        PrintLong(ECnt) = ThisLong
        ReDim Preserve PrintRec(1 To ECnt) As Integer
        PrintRec(ECnt) = x
        If ThisLong > BigNum Then
          BigNum = ThisLong
        End If
      End If
Deleted2:
    Next x
    GoSub SortByAcctNum
  End If
  
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
      If Mid(TaxMRec.BillType, 1, 1) = "R" Then
        If TaxMRec.TaxYear > RealNewYear Then
          RealNewYear = TaxMRec.TaxYear
        End If
        RealCnt = RealCnt + 1
        ReDim Preserve SortRealRec(1 To RealCnt) As Integer
        SortRealRec(RealCnt) = PrintRec(x)
      ElseIf Mid(TaxMRec.BillType, 1, 1) = "P" Then
        If TaxMRec.TaxYear > PersNewYear Then
          PersNewYear = TaxMRec.TaxYear
        End If
        PersCnt = PersCnt + 1
        ReDim Preserve SortPersRec(1 To PersCnt) As Integer
        SortPersRec(PersCnt) = PrintRec(x)
      End If
  Next x
  If RealCnt = 0 Then GoTo PrintPersOnly
  GoSub PrintRHeader
 
  OpenRealPropFile RHandle, NumOfSortRealRecs
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
    If TaxMRec.BillType = "P" Then GoTo Skip1
    If ThisAcct <> TaxMRec.Account Then
      ThisAcct = TaxMRec.Account
      If LineCnt > MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintRHeader
      End If
      GoSub PrintCustHeader
    End If
    TotCustAmt = OldRound(TaxMRec.TaxAmount + TaxMRec.IntAmount + TaxMRec.AdColAmount + TaxMRec.LateList)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3 + TaxMRec.Penalty)
    GTotPrinc = OldRound(GTotPrinc + TaxMRec.TaxAmount)
    GTotRInt = OldRound(GTotRInt + TaxMRec.IntAmount)
    GTotRPen = OldRound(GTotRPen + TaxMRec.Penalty)
    GTotAdvCol = OldRound(GTotAdvCol + TaxMRec.AdColAmount)
    GTotLateList = OldRound(GTotLateList + TaxMRec.LateList)
    RGTotOpt1 = OldRound(RGTotOpt1 + TaxMRec.OptRev1)
    RGTotOpt2 = OldRound(RGTotOpt2 + TaxMRec.OptRev2)
    RGTotOpt3 = OldRound(RGTotOpt3 + TaxMRec.OptRev3)
    RGTotal = OldRound(GTotPrinc + GTotRInt + GTotPPen + GTotAdvCol + GTotLateList + RGTotOpt1 + RGTotOpt2 + RGTotOpt3)
    GTotal = OldRound(GTotal + RGTotal)
    If TaxMRec.Class = "M" Then
      ThisClass = "    MOCK"
    ElseIf TaxMRec.Class = "R" Then
      ThisClass = "    REAL"
    Else
      ThisClass = "      NA"
    End If
    
    Print #RptHandle, MakeRegDate(TaxMRec.TransDate); Tab(15); Using$("####0", TaxMRec.BillNum); Tab(29); ThisClass;
    If TaxMRec.RealRec > 0 Then
      Get RHandle, TaxMRec.RealRec, RealPropRec
      Print #RptHandle, Tab(40); QPTrim$(RealPropRec.RealPin)
    Else
      Print #RptHandle, Tab(40); "   NA"
    End If
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintRHeader
      GoSub PrintCustHeader
    End If
    Print #RptHandle, Using$("$#,###,##0.00", TaxMRec.TaxAmount); Tab(20); Using$("$##,##0.00", TaxMRec.IntAmount); Tab(35); Using$("$##,##0.00", TaxMRec.AdColAmount);
    Print #RptHandle, Tab(50); Using$("$##,##0.00", TaxMRec.LateList); Tab(61); Using$("$##,##0.00", TaxMRec.Penalty)
    Print #RptHandle, Tab(4); Using$("$##,##0.00", TaxMRec.OptRev1); Tab(20); Using$("$##,##0.00", TaxMRec.OptRev2); Tab(35); Using$("$##,##0.00", TaxMRec.OptRev3); Tab(86); Using$("$###,##0.00", TotCustAmt)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintRHeader
      GoSub PrintCustHeader
    End If
Skip1:
  Next x
  
  Print #RptHandle, FF$
  
  Close RHandle
  
  If PersCnt = 0 Then GoTo NoPers2Print
  
PrintPersOnly:
  GoSub PrintPHeader
  OpenPersPropFile PHandle, NumOfSortPersRecs
  TotCustAmt = 0
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
    If TaxMRec.BillType <> "P" Then GoTo Skip2
    If ThisAcct <> TaxMRec.Account Then
      ThisAcct = TaxMRec.Account
      If LineCnt > MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintPHeader
      End If
      GoSub PrintCustHeader
    End If
    TotCustAmt = OldRound(TaxMRec.Personal + TaxMRec.IntAmount + TaxMRec.MachTools + TaxMRec.MerchCap)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.FarmEquip + TaxMRec.MobHomes + TaxMRec.Penalty)
    GTotPers = OldRound(GTotPers + TaxMRec.Personal)
    GTotPInt = OldRound(GTotPInt + TaxMRec.IntAmount)
    GTotPPen = OldRound(GTotPPen + TaxMRec.Penalty)
    GTotMT = OldRound(GTotMT + TaxMRec.MachTools)
    GTotFE = OldRound(GTotFE + TaxMRec.FarmEquip)
    GTotMH = OldRound(GTotMH + TaxMRec.MobHomes)
    GTotMC = OldRound(GTotMC + TaxMRec.MerchCap)
    PGTotOpt1 = OldRound(PGTotOpt1 + TaxMRec.OptRev1)
    PGTotOpt2 = OldRound(PGTotOpt2 + TaxMRec.OptRev2)
    PGTotOpt3 = OldRound(PGTotOpt3 + TaxMRec.OptRev3)
    PGTotal = OldRound(GTotPers + GTotPInt + GTotPPen + GTotMT + GTotMC + GTotFE + GTotMH + PGTotOpt1 + PGTotOpt2 + PGTotOpt3)
    GTotal = OldRound(GTotal + PGTotal)
    If TaxMRec.Class = "P" Then
      ThisClass = "PERSONAL"
    Else
      ThisClass = "      NA"
    End If
    
    Print #RptHandle, MakeRegDate(TaxMRec.TransDate); Tab(15); Using$("####0", TaxMRec.BillNum); Tab(29); ThisClass;
    If TaxMRec.PersRec > 0 Then
      Get PHandle, TaxMRec.PersRec, PersPropRec
      Print #RptHandle, Tab(45); QPTrim$(PersPropRec.PropPin)
    Else
      Print #RptHandle, Tab(45); "   NA"
    End If
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintPHeader
      GoSub PrintCustHeader
    End If
    Print #RptHandle, Using$("$#,###,##0.00", TaxMRec.Personal); Tab(18); Using$("$##,##0.00", TaxMRec.MachTools); Tab(29); Using$("$##,##0.00", TaxMRec.MerchCap);
    Print #RptHandle, Tab(41); Using$("$##,##0.00", TaxMRec.FarmEquip); Tab(53); Using$("$##,##0.00", TaxMRec.MobHomes); Tab(63); Using$("$##,##0.00", TaxMRec.IntAmount);
    Print #RptHandle, Tab(74); Using$("$##,##0.00", TaxMRec.Penalty)
    Print #RptHandle, Using$("$#,###,##0.00", TaxMRec.OptRev1); Tab(18); Using$("$##,##0.00", TaxMRec.OptRev2); Tab(29); Using$("$##,##0.00", TaxMRec.OptRev3); Tab(86); Using$("$###,##0.00", TotCustAmt)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintPHeader
      GoSub PrintCustHeader
    End If
Skip2:
  Next x
  
  Close PHandle
  Print #RptHandle, FF$
  
NoPers2Print:
  If RealCnt > 0 Then
    GoSub RPrintSummary
  End If
  
  If PersCnt > 0 Then
    GoSub PPrintSummary
  End If
  
  GoSub PrintSub1
  
  Print #RptHandle, FF$
  Close RptHandle
  Close TMHandle
 
  Close

  ViewPrint RptFile$, "Tax Manual Bill Edit Report", True
  
  KillFile RptFile$
 
  Exit Sub
 
PrintSub1:
  If RealCnt = 0 Then GoTo NoRealCnt
 
  LilYear = RealNewYear + 1
  Nextx = 1
  Do
    For x = Nextx To RealCnt
      Get TMHandle, SortRealRec(x), TaxMRec
      If TaxMRec.TaxYear < LilYear Then
        LilYear = TaxMRec.TaxYear
        Thisx = x
      End If
    Next x
    HoldRec = SortRealRec(Nextx)
    SortRealRec(Nextx) = SortRealRec(Thisx)
    SortRealRec(Thisx) = HoldRec
    If Nextx = 1 Then RealYear = LilYear
    Nextx = Nextx + 1
    If Nextx > RealCnt Then
      Exit Do
    End If
    LilYear = RealNewYear + 1
  Loop
 
  ReDim RealRec(1 To RealCnt) As Integer
  For x = 1 To RealCnt
    RealRec(x) = SortRealRec(x)
  Next x
  
  ReDim Years(1 To 1) As Integer

  For x = 1 To RealCnt
    Get TMHandle, RealRec(x), TaxMRec
      If x = 1 Then
        RealYrCnt = RealYrCnt + 1
        ReDim Preserve Years(1 To RealYrCnt) As Integer
        Years(RealYrCnt) = TaxMRec.TaxYear
      Else
        For y = 1 To RealYrCnt
          If TaxMRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > RealYrCnt Then
          RealYrCnt = RealYrCnt + 1
          ReDim Preserve Years(1 To RealYrCnt) As Integer
          Years(RealYrCnt) = TaxMRec.TaxYear
        End If
      End If
  Next x

NoRealCnt:

  If PersCnt = 0 Then GoTo NoPersCnt
 
  LilYear = PersNewYear + 1
  Nextx = 1
  Do
    For x = Nextx To PersCnt
      Get TMHandle, SortPersRec(x), TaxMRec
      If TaxMRec.TaxYear < LilYear Then
        LilYear = TaxMRec.TaxYear
        Thisx = x
      End If
    Next x
    HoldRec = SortPersRec(Nextx)
    SortPersRec(Nextx) = SortPersRec(Thisx)
    SortPersRec(Thisx) = HoldRec
    If Nextx = 1 Then PersYear = LilYear
    Nextx = Nextx + 1
    If Nextx > PersCnt Then
      Exit Do
    End If
    LilYear = PersNewYear + 1
  Loop
 
  ReDim PersRec(1 To PersCnt) As Integer
  For x = 1 To PersCnt
    PersRec(x) = SortPersRec(x)
  Next x
  
  ReDim Years(1 To 1) As Integer

  For x = 1 To PersCnt
    Get TMHandle, PersRec(x), TaxMRec
      If x = 1 Then
        PersYrCnt = PersYrCnt + 1
        ReDim Preserve Years(1 To PersYrCnt) As Integer
        Years(PersYrCnt) = TaxMRec.TaxYear
      Else
        For y = 1 To PersYrCnt
          If TaxMRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > PersYrCnt Then
          PersYrCnt = PersYrCnt + 1
          ReDim Preserve Years(1 To PersYrCnt) As Integer
          Years(PersYrCnt) = TaxMRec.TaxYear
        End If
      End If
  Next x
  
NoPersCnt:
  
  YrTotPrinc = 0
  YrTotRInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotRPen = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrRGTot = 0
  YrGTot = 0
  If RealCnt > 0 Then
    Print #RptHandle, Tab(5); "Bill Results By Year: REAL ONLY"
    Print #RptHandle, Tab(5); String$(92, ".")
    LineCnt = LineCnt + 2
    If LineCnt > MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintYearHeader
      GoSub RPrintSummary
    End If
  End If
  For x = 1 To RealCnt
    Get TMHandle, RealRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotRInt = OldRound(YrTotRInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotRPen = OldRound(YrTotRPen + TaxMRec.Penalty)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrRGTot = OldRound(YrTotPrinc + YrTotRInt + YrTotAdvCol + YrTotLateList + YrTotOpt1 + YrTotOpt2 + YrTotOpt3 + YrTotRPen)
    YrGTot = OldRound(YrGTot + YrRGTot)
    If RealYrCnt > 1 And x < RealCnt Then
      Get TMHandle, RealRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> RealYear Then
        Print #RptHandle, "    Year: " + Using$("####", RealYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrRGTot)
        Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(65); "Penalty"
        Print #RptHandle, Tab(6); Opt1Desc; Tab(20); Opt2Desc; Tab(32); Opt3Desc
        Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotRInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList); Tab(61); Using$("$###,##0.00", GTotRPen)
        Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotOpt1); Tab(19); Using$("$###,##0.00", YrTotOpt2); Tab(31); Using$("$###,##0.00", YrTotOpt3)
        Print #RptHandle, Tab(5); String$(92, ".")
        LineCnt = LineCnt + 6
        If LineCnt > MaxLines - 6 Then
          Print #RptHandle, FF$
          GoSub PrintYearHeader
          GoSub RPrintSummary
        End If
        RealYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotRInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotRPen = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrGTot = 0
      End If
    ElseIf x = RealCnt Then
      Print #RptHandle, "    Year: " + Using$("####", RealYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrRGTot)
      Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(65); "Penalty"
      Print #RptHandle, Tab(6); Opt1Desc; Tab(20); Opt2Desc; Tab(32); Opt3Desc
      Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotRInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList); Tab(61); Using$("$###,##0.00", GTotRPen)
      Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotOpt1); Tab(19); Using$("$###,##0.00", YrTotOpt2); Tab(31); Using$("$###,##0.00", YrTotOpt3)
      LineCnt = LineCnt + 4
      If LineCnt = MaxLines - 5 Then
        Print #RptHandle, FF$
        GoSub PrintYearHeader
        GoSub RPrintSummary
      End If
  End If
  Next x
  
  If RealCnt > 0 And PersCnt > 0 Then
    Print #RptHandle, FF$
    GoSub PrintYearHeader
    GoSub PPrintSummary
  End If
  
  Print #RptHandle,
  
  YrTotPers = 0
  YrTotPInt = 0
  YrTotMT = 0
  YrTotMC = 0
  YrTotFE = 0
  YrTotMH = 0
  YrTotPPen = 0
  YrTotPOpt1 = 0
  YrTotPOpt2 = 0
  YrTotPOpt3 = 0
  YrPGTot = 0
  If PersCnt > 0 Then
    Print #RptHandle, Tab(5); "Bill Results By Year: PERSONAL ONLY"
    Print #RptHandle, Tab(5); String$(92, ".")
    LineCnt = LineCnt + 2
    If LineCnt > MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintYearHeader
      GoSub PPrintSummary
    End If
  End If
  
  For x = 1 To PersCnt
    Get TMHandle, PersRec(x), TaxMRec
    YrTotPers = OldRound(YrTotPers + TaxMRec.Personal)
    YrTotPInt = OldRound(YrTotPInt + TaxMRec.IntAmount)
    YrTotMT = OldRound(YrTotMT + TaxMRec.MachTools)
    YrTotMC = OldRound(YrTotMC + TaxMRec.MerchCap)
    YrTotFE = OldRound(YrTotFE + TaxMRec.FarmEquip)
    YrTotMH = OldRound(YrTotMH + TaxMRec.MobHomes)
    YrTotPPen = OldRound(YrTotPPen + TaxMRec.Penalty)
    YrTotPOpt1 = OldRound(YrTotPOpt1 + TaxMRec.OptRev1)
    YrTotPOpt2 = OldRound(YrTotPOpt2 + TaxMRec.OptRev2)
    YrTotPOpt3 = OldRound(YrTotPOpt3 + TaxMRec.OptRev3)
    YrPGTot = OldRound(YrTotPers + YrTotPInt + YrTotMT + YrTotMC + YrTotFE + YrTotMH + YrTotPOpt1 + YrTotPOpt2 + YrTotPOpt3 + YrTotPPen)
    YrGTot = OldRound(YrGTot + YrPGTot)
    If PersYrCnt > 1 And x < PersCnt Then
      Get TMHandle, PersRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> PersYear And x < PersCnt Then
        Print #RptHandle, "    Year: " + Using$("####", PersYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrPGTot)
        Print #RptHandle, Tab(8); "Personal"; Tab(20); "Mach Tools"; Tab(33); "Merch Cap"; Tab(48); "Farm Equip"; Tab(63); "Mob Homes"; Tab(76); "Interest"; Tab(90); "Penalty"
        Print #RptHandle, Tab(6); POpt1Desc; Tab(20); POpt2Desc; Tab(32); POpt3Desc
        Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPers); Tab(19); Using$("$###,##0.00", YrTotMT); Tab(31); Using$("$###,##0.00", YrTotMC); Tab(47); Using$("$###,##0.00", YrTotFE);
        Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotMH); Tab(73); Using$("$###,##0.00", YrTotPInt); Tab(86); Using("$###,##0.00", YrTotPPen)
        Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPOpt1); Tab(19); Using$("$###,##0.00", YrTotPOpt2); Tab(31); Using$("$###,##0.00", YrTotPOpt3)
        Print #RptHandle, Tab(5); String$(92, ".")
        LineCnt = LineCnt + 5
        If LineCnt > MaxLines - 6 Then
          Print #RptHandle, FF$
          GoSub PrintYearHeader
          GoSub PPrintSummary
        End If
        PersYear = TaxMRec.TaxYear
        YrTotPers = 0
        YrTotPInt = 0
        YrTotMT = 0
        YrTotMC = 0
        YrTotFE = 0
        YrTotMH = 0
        YrTotPPen = 0
        YrTotPOpt1 = 0
        YrTotPOpt2 = 0
        YrTotPOpt3 = 0
        YrPGTot = 0
      End If
    ElseIf x = PersCnt Then
      Print #RptHandle, "    Year: " + Using$("####", PersYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrPGTot)
      Print #RptHandle, Tab(8); "Personal"; Tab(20); "Mach Tools"; Tab(33); "Merch Cap"; Tab(48); "Farm Equip"; Tab(63); "Mob Homes"; Tab(76); "Interest"; Tab(90); "Penalty"
      Print #RptHandle, Tab(6); POpt1Desc; Tab(20); POpt2Desc; Tab(32); POpt3Desc
      Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPers); Tab(19); Using$("$###,##0.00", YrTotMT); Tab(31); Using$("$###,##0.00", YrTotMC); Tab(47); Using$("$###,##0.00", YrTotFE);
      Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotMH); Tab(73); Using$("$###,##0.00", YrTotPInt); Tab(86); Using$("$###,##0.00", YrTotPPen)
      Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPOpt1); Tab(19); Using$("$###,##0.00", YrTotPOpt2); Tab(31); Using$("$###,##0.00", YrTotPOpt3)
      LineCnt = LineCnt + 5
      If LineCnt = MaxLines - 6 Then
        Print #RptHandle, FF$
        GoSub PrintYearHeader
        GoSub PPrintSummary
      End If
    End If
  Next x
  
  Return
  
SortByName:
  BigName = BigName + "z"
  SmallName = BigName
  Nextx = 1
  Thisx = 0
  Do
    For x = Nextx To ECnt
      If PrintList(x) < SmallName Then
        SmallName = PrintList(x)
        Thisx = x
      End If
    Next x
    
    HoldNum = PrintList(Nextx)
    HoldInt = PrintRec(Nextx)
    PrintList(Nextx) = PrintList(Thisx)
    PrintRec(Nextx) = PrintRec(Thisx)
    PrintList(Thisx) = HoldNum
    PrintRec(Thisx) = HoldInt
    Nextx = Nextx + 1
    If Nextx > ECnt Then Exit Do
    Thisx = 0
    SmallName = BigName
  Loop

  Return
  
SortByAcctNum:
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  Thisx = 0
  Do
    For x = Nextx To ECnt
      If PrintLong(x) < SmallNum Then
        SmallNum = PrintLong(x)
        Thisx = x
      End If
    Next x
    HoldLong = PrintLong(Nextx)
    HoldInt = PrintRec(Nextx)
    PrintLong(Nextx) = PrintLong(Thisx)
    PrintRec(Nextx) = PrintRec(Thisx)
    PrintLong(Thisx) = HoldLong
    PrintRec(Thisx) = HoldInt
    Nextx = Nextx + 1
    If Nextx > ECnt Then Exit Do
    Thisx = 0
    SmallNum = BigNum
  Loop
  
  Return
  
PrintRHeader:
  Page = Page + 1
  RSet Opt1Desc = QPTrim$(Opt1Desc)
  RSet Opt2Desc = QPTrim$(Opt2Desc)
  RSet Opt3Desc = QPTrim$(Opt3Desc)
  Print #RptHandle, Tab(30); "Tax Manual Bill Edit Report"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Str(Page)
  Print #RptHandle,
  Print #RptHandle, "Acct Num"; Tab(13); "Customer Name"
  Print #RptHandle, "Trans Date"; Tab(15); "Bill #"; Tab(28); "Property Type"; Tab(43); "PIN #"
  Print #RptHandle, Tab(5); "Principle"; Tab(22); "Interest"; Tab(38); "Adv/Col"; Tab(51); "Late List"; Tab(64); "Penalty"
  Print #RptHandle, Tab(4); Opt1Desc; Tab(20); Opt2Desc; Tab(35); Opt3Desc; Tab(87); "Bill Total"
  Print #RptHandle, String$(96, "=")
  LineCnt = 8
  Return
  
PrintPHeader:
  Page = Page + 1
  RSet POpt1Desc = QPTrim$(POpt1Desc)
  RSet POpt2Desc = QPTrim$(POpt2Desc)
  RSet POpt3Desc = QPTrim$(POpt3Desc)
  Print #RptHandle, Tab(30); "Tax Manual Bill Edit Report"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Str(Page)
  Print #RptHandle,
  Print #RptHandle, "Acct Num"; Tab(13); "Customer Name"
  Print #RptHandle, "Trans Date"; Tab(15); "Bill #"; Tab(28); "Property Type"; Tab(43); "PIN #"
  Print #RptHandle, Tab(6); "Personal"; Tab(18); "Mach Tools"; Tab(30); "Merch Cap"; Tab(41); "Farm Equip"; Tab(54); "Mob Homes"; Tab(65); "Interest"; Tab(77); "Penalty"
  Print #RptHandle, Tab(4); POpt1Desc; Tab(18); POpt2Desc; Tab(29); POpt3Desc; Tab(87); "Bill Total"
  Print #RptHandle, String$(96, "=")
  LineCnt = 7
  Return
   
PrintYearHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Manual Bill Edit Report"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Str(Page)
  Print #RptHandle, String$(96, "=")
  LineCnt = 4
  Return
   
PrintCustHeader:
  If LineCnt <> 7 Then
    Print #RptHandle, String$(96, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Using$("#######0", TaxMRec.Account); Tab(13); QPTrim$(TaxMRec.SName)
  Print #RptHandle, String$(96, ".")
  LineCnt = LineCnt + 2
  Return

RPrintSummary:
  Page = Page + 1
  RSet Opt1Desc = QPTrim$(Opt1Desc)
  RSet Opt2Desc = QPTrim$(Opt2Desc)
  RSet Opt3Desc = QPTrim$(Opt3Desc)
  Print #RptHandle, Tab(28); "Tax Manual Bill Edit Report Summary"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Str(Page)
  Print #RptHandle, String$(96, "=")
  
  If ROverAllPrinted = False Then
    GoSub RPrintOverallAmts
  End If
  
  LineCnt = 4
  Return
  
PPrintSummary:
  Page = Page + 1
  RSet Opt1Desc = QPTrim$(Opt1Desc)
  RSet Opt2Desc = QPTrim$(Opt2Desc)
  RSet Opt3Desc = QPTrim$(Opt3Desc)
  
  If POverAllPrinted = False Then
    GoSub PPrintOverallAmts
  End If
  
  LineCnt = 4
  Return
  
RPrintOverallAmts:
  RGTotal = OldRound(GTotPrinc + GTotRInt + GTotAdvCol + GTotLateList + RGTotOpt1 + RGTotOpt2 + RGTotOpt3 + GTotRPen)
  ROverAllPrinted = True
  Print #RptHandle, Tab(5); "TOTAL OVERALL REAL AMOUNTS:"
  Print #RptHandle, Tab(5); String$(92, "-")
  Print #RptHandle, Tab(5); "Total Real Bills Entered: " + CStr(RealCnt); Tab(65); "Total Amount: " + Using$("$###,###,##0.00", RGTotal)
  Print #RptHandle,
  Print #RptHandle, Tab(5); "Total Principle Amount: "; Tab(32); Using$("$###,###,##0.00", GTotPrinc); Tab(56); "Total Interest Amount: "; Tab(80); Using$("$###,##0.00", GTotRInt)
  Print #RptHandle, Tab(7); "Total Adv/Col Amount: "; Tab(36); Using("$###,##0.00", GTotAdvCol); Tab(55); "Total Late List Amount: "; Tab(80); Using$("$###,##0.00", GTotLateList)
  Print #RptHandle, Tab(14); "Total Penalty: "; Tab(36); Using$("$###,##0.00", GTotRPen)
  Print #RptHandle, Tab(11); "Total " + Opt1Desc + ": "; Tab(36); Using("$###,##0.00", RGTotOpt1); Tab(61); "Total " + Opt2Desc + ": "; Tab(80); Using$("$###,##0.00", RGTotOpt2)
  Print #RptHandle, Tab(11); "Total " + Opt3Desc + ": "; Tab(36); Using("$###,##0.00", RGTotOpt3)
  Print #RptHandle, Tab(5); String$(92, "-")
  Print #RptHandle,
  LineCnt = LineCnt + 10
    
  Return
  
PPrintOverallAmts:
  PGTotal = OldRound(GTotPers + GTotPInt + GTotMT + GTotMC + PGTotOpt1 + PGTotOpt2 + PGTotOpt3 + GTotFE + GTotMH + GTotPPen)
  POverAllPrinted = True
  Print #RptHandle, Tab(5); "TOTAL OVERALL PERSONAL AMOUNTS:"
  Print #RptHandle, Tab(5); String$(92, "-")
  Print #RptHandle, Tab(5); "Total Personal Bills Entered: " + CStr(PersCnt); Tab(65); "Total Amount: " + Using$("$###,###,##0.00", PGTotal)
  Print #RptHandle,
  Print #RptHandle, Tab(6); "Total Personal Amount: "; Tab(32); Using$("$###,###,##0.00", GTotPers); Tab(54); "Total Mach Tools Amount: "; Tab(80); Using$("$###,##0.00", GTotMT)
  Print #RptHandle, Tab(5); "Total Merch Cap Amount: "; Tab(36); Using("$###,##0.00", GTotMC); Tab(54); "Total Farm Equip Amount: "; Tab(80); Using$("$###,##0.00", GTotFE)
  Print #RptHandle, Tab(5); "Total Mob Homes Amount: "; Tab(36); Using$("$###,##0.00", GTotMH); Tab(63); "Total Interest: "; Tab(80); Using$("$###,##0.00", GTotPInt)
  Print #RptHandle, Tab(14); "Total Penalty: "; Tab(36); Using$("$###,##0.00", GTotPPen);
  Print #RptHandle, Tab(11); "Total " + POpt1Desc + ": "; Tab(36); Using("$###,##0.00", PGTotOpt1); Tab(61); "Total " + POpt2Desc + ": "; Tab(80); Using$("$###,##0.00", PGTotOpt2)
  Print #RptHandle, Tab(11); "Total " + POpt3Desc + ": "; Tab(36); Using("$###,##0.00", PGTotOpt3)
  Print #RptHandle, Tab(5); String$(92, "-")
  Print #RptHandle,
  LineCnt = LineCnt + 11

  Return
End Sub
