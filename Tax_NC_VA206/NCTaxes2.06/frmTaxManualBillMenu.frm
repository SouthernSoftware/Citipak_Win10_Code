VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxManualBillMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Tax Billing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxManualBillMenu.frx":0000
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
      Top             =   5205
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
      ButtonDesigner  =   "frmTaxManualBillMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintTrans 
      Height          =   435
      Left            =   4005
      TabIndex        =   2
      Top             =   4665
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
      ButtonDesigner  =   "frmTaxManualBillMenu.frx":0AAE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditTrans 
      Height          =   435
      Left            =   4005
      TabIndex        =   1
      Top             =   4110
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
      ButtonDesigner  =   "frmTaxManualBillMenu.frx":0CA0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterBill 
      Height          =   435
      Left            =   4005
      TabIndex        =   0
      Top             =   3570
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
      ButtonDesigner  =   "frmTaxManualBillMenu.frx":0E84
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4005
      TabIndex        =   4
      Top             =   5760
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
      ButtonDesigner  =   "frmTaxManualBillMenu.frx":106F
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
Attribute VB_Name = "frmTaxManualBillMenu"
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
  
  frmTaxManualBillEdit.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdEnterBill_Click()
  Dim One As Integer
  Dim AHandle As Integer
  
'  If Check4PayBatch = True Then
'    Call TaxMsg(800, "An unposted payment file is ready for posting. Manual billing cannot be conducted until these payments are posted.")
'    Exit Sub
'  End If
  
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\manualbill.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  frmTaxManualBillEntry.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  frmTaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer
  
'  If Check4PayBatch = True Then
'    frmTaxUnpostedPayList.Show vbModal
'    Call TaxMsg(800, "An unposted payment file is ready for posting. Manual bill posting cannot be conducted until these payments are posted.")
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
  
  frmTaxManualPost.Show
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
  
  frmTaxReportOptWOpt.Show vbModal
  If frmTaxReportOptWOpt.fptxtPrintType.Text = "Graphical Name" Then
    Unload frmTaxReportOptWOpt
    PrintOpt = "Graphical Name"
    Call PrintGraphics
  ElseIf frmTaxReportOptWOpt.fptxtPrintType.Text = "Graphical Number" Then
    Unload frmTaxReportOptWOpt
    PrintOpt = "Graphical Number"
    Call PrintGraphics
  ElseIf frmTaxReportOptWOpt.fptxtPrintType.Text = "Text Name" Then
    Unload frmTaxReportOptWOpt
    PrintOpt = "Text Name"
    Call TaxMsg(900, "Pitch 12 is recommended for this report.")
    Call PrintText
  ElseIf frmTaxReportOptWOpt.fptxtPrintType.Text = "Text Number" Then
    Unload frmTaxReportOptWOpt
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxManualBillMenu.")
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
  Dim Sub2RptFile$
  Dim Sub2RptHandle As Integer
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
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
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
  Dim GTotInt As Double
  Dim GTotAdvCol As Double
  Dim GTotLateList As Double
  Dim GTotOpt1 As Double
  Dim GTotOpt2 As Double
  Dim GTotOpt3 As Double
  Dim GTotal As Double
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
  Dim YrTotInt As Double
  Dim YrTotAdvCol As Double
  Dim YrTotLateList As Double
  Dim YrTotOpt1 As Double
  Dim YrTotOpt2 As Double
  Dim YrTotOpt3 As Double
  Dim YrGTot As Double
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
  
  Opt1Desc$ = QPTrim$(TaxSURec.OptRev1)
  Opt2Desc$ = QPTrim$(TaxSURec.OptRev2)
  Opt3Desc$ = QPTrim$(TaxSURec.OptRev3)
  
  Town = QPTrim$(TaxSURec.Name)
  dlm = "~"
  RptFile$ = "TAXRPTS\TXMANEDT.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  ECnt = 0
  ReDim PrintList(1 To 1) As String
  ReDim PrintLong(1 To 1) As Long
  ReDim PrintRec(1 To 1) As Integer
  ReDim SortCombRec(1 To 1) As Integer
  ReDim SortRealRec(1 To 1) As Integer
  ReDim SortPersRec(1 To 1) As Integer
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  
  BigName = ""
  BigNum = 0
  CombCnt = 0
  RealCnt = 0
  PersCnt = 0
  CombNewYear = 0
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
      If Mid(TaxMRec.BillType, 1, 1) = "C" Then
        If TaxMRec.TaxYear > CombNewYear Then
          CombNewYear = TaxMRec.TaxYear
        End If
        CombCnt = CombCnt + 1
        ReDim Preserve SortCombRec(1 To CombCnt) As Integer
        SortCombRec(CombCnt) = PrintRec(x)
      ElseIf Mid(TaxMRec.BillType, 1, 1) = "R" Then
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
  OpenPersPropFile PHandle, NumOfSortPersRecs
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
    TotCustAmt = OldRound(TaxMRec.TaxAmount + TaxMRec.IntAmount + TaxMRec.AdColAmount + TaxMRec.LateList)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3)
    GTotPrinc = OldRound(GTotPrinc + TaxMRec.TaxAmount)
    GTotInt = OldRound(GTotInt + TaxMRec.IntAmount)
    GTotAdvCol = OldRound(GTotAdvCol + TaxMRec.AdColAmount)
    GTotLateList = OldRound(GTotLateList + TaxMRec.LateList)
    GTotOpt1 = OldRound(GTotOpt1 + TaxMRec.OptRev1)
    GTotOpt2 = OldRound(GTotOpt2 + TaxMRec.OptRev2)
    GTotOpt3 = OldRound(GTotOpt3 + TaxMRec.OptRev3)
    GTotal = OldRound(GTotPrinc + GTotInt + GTotAdvCol + GTotLateList + GTotOpt1 + GTotOpt2 + GTotOpt3)
   
    If TaxMRec.Class = "M" Then
      ThisClass = "MOCK"
    ElseIf TaxMRec.Class = "R" Then
      ThisClass = "REAL"
    ElseIf TaxMRec.Class = "P" Then
     ThisClass = "PERSONAL"
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
    ElseIf TaxMRec.PersRec > 0 Then
      Get PHandle, TaxMRec.PersRec, PersPropRec
      '                             17
      Print #RptHandle, QPTrim$(PersPropRec.PropPin); dlm;
    Else
      '                  17
      Print #RptHandle, "NA"; dlm;
    End If
    '                      18              19             20              21
    Print #RptHandle, TotCustAmt; dlm; GTotPrinc; dlm; GTotInt; dlm; GTotAdvCol; dlm;
    '                      22               23             24             25            26          27
    Print #RptHandle, GTotLateList; dlm; GTotOpt1; dlm; GTotOpt2; dlm; GTotOpt3; dlm; GTotal; dlm; ECnt
  Next x
 
  Close RHandle
  Close PHandle
  Close RptHandle
  
  GoSub PrintSub1
  
  Close TMHandle
  Close Sub1RptHandle
  Close Sub2RptHandle
  
  arTaxManEditRpt.Show
 
  Exit Sub
 
PrintSub1:
  Sub1RptFile$ = "TAXRPTS\TXMANSUB1.RPT"     'Report File Name
  Sub1RptHandle = FreeFile
  Open Sub1RptFile$ For Output As #Sub1RptHandle
 
  If CombCnt = 0 Then GoTo NoCombCnt
 
  LilYear = CombNewYear + 1
  Nextx = 1
  Do
    For x = Nextx To CombCnt
      Get TMHandle, SortCombRec(x), TaxMRec
      If TaxMRec.TaxYear < LilYear Then
        LilYear = TaxMRec.TaxYear
        Thisx = x
      End If
    Next x
    HoldRec = SortCombRec(Nextx)
    SortCombRec(Nextx) = SortCombRec(Thisx)
    SortCombRec(Thisx) = HoldRec
    If Nextx = 1 Then CombYear = LilYear
    Nextx = Nextx + 1
    If Nextx > CombCnt Then
      Exit Do
    End If
    LilYear = CombNewYear + 1
  Loop
  

  ReDim CombRec(1 To CombCnt) As Integer
  For x = 1 To CombCnt
    CombRec(x) = SortCombRec(x)
  Next x
    
  ReDim Years(1 To 1) As Integer

  For x = 1 To CombCnt
    Get TMHandle, CombRec(x), TaxMRec
      If x = 1 Then
        CombYrCnt = CombYrCnt + 1
        ReDim Preserve Years(1 To CombYrCnt) As Integer
        Years(CombYrCnt) = TaxMRec.TaxYear
      Else
        For y = 1 To CombYrCnt
          If TaxMRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > CombYrCnt Then
          CombYrCnt = CombYrCnt + 1
          ReDim Preserve Years(1 To CombYrCnt) As Integer
          Years(CombYrCnt) = TaxMRec.TaxYear
        End If
      End If
  Next x
      
NoCombCnt:
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
  YrTotInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrGTot = 0
  
  For x = 1 To CombCnt
    Get TMHandle, CombRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotInt = OldRound(YrTotInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrGTot = OldRound(YrTotPrinc + YrTotInt + YrTotAdvCol + YrTotLateList + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    If CombYrCnt > 1 And x < CombCnt Then
      Get TMHandle, CombRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> CombYear And x < CombCnt Then
        '                        0                 1              2               3                  4
        Print #Sub1RptHandle, CombYear; dlm; YrTotPrinc; dlm; YrTotInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
        '                        5                 6              7              8          9
        Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrGTot; dlm; "C"; dlm;
        '                        10             11             12
        Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc
       CombYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrGTot = 0
      End If
    ElseIf x = CombCnt Then
      '                        0                 1              2               3                  4
      Print #Sub1RptHandle, CombYear; dlm; YrTotPrinc; dlm; YrTotInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
      '                        5                 6              7              8          9
      Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrGTot; dlm; "C"; dlm;
      '                        10             11             12
      Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc
    End If
  Next x
  
  YrTotPrinc = 0
  YrTotInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrGTot = 0
  For x = 1 To RealCnt
    Get TMHandle, RealRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotInt = OldRound(YrTotInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrGTot = OldRound(YrTotPrinc + YrTotInt + YrTotAdvCol + YrTotLateList + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    If RealYrCnt > 1 And x < RealCnt Then
      Get TMHandle, RealRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> RealYear Then
        '                        0                 1              2               3                  4
        Print #Sub1RptHandle, RealYear; dlm; YrTotPrinc; dlm; YrTotInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
        '                        5                 6              7              8          9
        Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrGTot; dlm; "R"; dlm;
        '                        10             11             12
        Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc
        RealYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrGTot = 0
      End If
    ElseIf x = RealCnt Then
      '                        0                 1              2               3                  4
      Print #Sub1RptHandle, RealYear; dlm; YrTotPrinc; dlm; YrTotInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
      '                        5                 6              7              8          9
      Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrGTot; dlm; "R"; dlm;
        '                        10           11             12
      Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc
  End If
  Next x
  
  YrTotPrinc = 0
  YrTotInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrGTot = 0
  For x = 1 To PersCnt
    Get TMHandle, PersRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotInt = OldRound(YrTotInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrGTot = OldRound(YrTotPrinc + YrTotInt + YrTotAdvCol + YrTotLateList + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    If PersYrCnt > 1 And x < PersCnt Then
      Get TMHandle, PersRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> PersYear And x < PersCnt Then
        '                        0                 1              2               3                  4
        Print #Sub1RptHandle, PersYear; dlm; YrTotPrinc; dlm; YrTotInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
        '                        5                 6              7              8          9
        Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrGTot; dlm; "P"; dlm;
        '                        10             11             12
        Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc
        PersYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrGTot = 0
      End If
    ElseIf x = PersCnt Then
      '                        0                 1              2               3                  4
      Print #Sub1RptHandle, PersYear; dlm; YrTotPrinc; dlm; YrTotInt; dlm; YrTotAdvCol; dlm; YrTotLateList; dlm;
      '                        5                 6              7              8          9
      Print #Sub1RptHandle, YrTotOpt1; dlm; YrTotOpt2; dlm; YrTotOpt3; dlm; YrGTot; dlm; "P"; dlm;
      '                        10             11             12
      Print #Sub1RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc
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
  Dim GTotInt As Double
  Dim GTotAdvCol As Double
  Dim GTotLateList As Double
  Dim GTotOpt1 As Double
  Dim GTotOpt2 As Double
  Dim GTotOpt3 As Double
  Dim GTotal As Double
  Dim TGTotOpt1 As Double
  Dim TGTotOpt2 As Double
  Dim TGTotOpt3 As Double
  Dim TGTotal As Double
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
  Dim YrTotInt As Double
  Dim YrTotAdvCol As Double
  Dim YrTotLateList As Double
  Dim YrTotOpt1 As Double
  Dim YrTotOpt2 As Double
  Dim YrTotOpt3 As Double
  Dim YrGTot As Double
  Dim CombYear As Integer
  Dim ThisECnt As Integer
  Dim ThisCombCnt As Integer
  Dim ThisPersCnt As Integer
  Dim ThisRealCnt As Integer
  Dim RealYear As Integer
  Dim PersYear As Integer
  Dim Page As Integer
  Dim ThisAcct As Long
  Dim OverAllPrinted As Boolean
  
  OverAllPrinted = False
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
  
  RptFile$ = "TaxManEdit.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  GoSub PrintHeader
  
  ReDim PrintList(1 To 1) As String
  ReDim PrintLong(1 To 1) As Long
  ReDim PrintRec(1 To 1) As Integer
  ReDim SortCombRec(1 To 1) As Integer
  ReDim SortRealRec(1 To 1) As Integer
  ReDim SortPersRec(1 To 1) As Integer
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  
  BigName = ""
  BigNum = 0
  CombCnt = 0
  RealCnt = 0
  PersCnt = 0
  CombNewYear = 0
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
      If Mid(TaxMRec.BillType, 1, 1) = "C" Then
        If TaxMRec.TaxYear > CombNewYear Then
          CombNewYear = TaxMRec.TaxYear
        End If
        CombCnt = CombCnt + 1
        ReDim Preserve SortCombRec(1 To CombCnt) As Integer
        SortCombRec(CombCnt) = PrintRec(x)
      ElseIf Mid(TaxMRec.BillType, 1, 1) = "R" Then
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
  OpenPersPropFile PHandle, NumOfSortPersRecs
  For x = 1 To ECnt
    Get TMHandle, PrintRec(x), TaxMRec
    If ThisAcct <> TaxMRec.Account Then
      ThisAcct = TaxMRec.Account
      If LineCnt > MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      GoSub PrintCustHeader
    End If
    TotCustAmt = OldRound(TaxMRec.TaxAmount + TaxMRec.IntAmount + TaxMRec.AdColAmount + TaxMRec.LateList)
    TotCustAmt = OldRound(TotCustAmt + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3)
    GTotPrinc = OldRound(GTotPrinc + TaxMRec.TaxAmount)
    GTotInt = OldRound(GTotInt + TaxMRec.IntAmount)
    GTotAdvCol = OldRound(GTotAdvCol + TaxMRec.AdColAmount)
    GTotLateList = OldRound(GTotLateList + TaxMRec.LateList)
    GTotOpt1 = OldRound(GTotOpt1 + TaxMRec.OptRev1)
    GTotOpt2 = OldRound(GTotOpt2 + TaxMRec.OptRev2)
    GTotOpt3 = OldRound(GTotOpt3 + TaxMRec.OptRev3)
    GTotal = OldRound(GTotPrinc + GTotInt + GTotAdvCol + GTotLateList + GTotOpt1 + GTotOpt2 + GTotOpt3)
   
    If TaxMRec.Class = "M" Then
      ThisClass = "    MOCK"
    ElseIf TaxMRec.Class = "R" Then
      ThisClass = "    REAL"
    ElseIf TaxMRec.Class = "P" Then
      ThisClass = "PERSONAL"
    Else
      ThisClass = "      NA"
    End If
    
    Print #RptHandle, MakeRegDate(TaxMRec.TransDate); Tab(15); Using$("####0", TaxMRec.BillNum); Tab(29); ThisClass;
    If TaxMRec.RealRec > 0 Then
      Get RHandle, TaxMRec.RealRec, RealPropRec
      Print #RptHandle, Tab(40); QPTrim$(RealPropRec.RealPin)
    ElseIf TaxMRec.PersRec > 0 Then
      Get PHandle, TaxMRec.PersRec, PersPropRec
      Print #RptHandle, Tab(40); QPTrim$(PersPropRec.PropPin)
    Else
      Print #RptHandle, Tab(40); "   NA"
    End If
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
    Print #RptHandle, Using$("$#,###,##0.00", TaxMRec.TaxAmount); Tab(15); Using$("$##,##0.00", TaxMRec.IntAmount); Tab(25); Using$("$##,##0.00", TaxMRec.AdColAmount);
    Print #RptHandle, Tab(35); Using$("$##,##0.00", TaxMRec.LateList); Tab(50); Using$("$##,##0.00", TaxMRec.OptRev1); Tab(61); Using$("$##,##0.00", TaxMRec.OptRev2);
    Print #RptHandle, Tab(73); Using$("$##,##0.00", TaxMRec.OptRev3); Tab(86); Using$("$###,##0.00", TotCustAmt)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
    
  Next x
  
  Print #RptHandle, FF$
  
  
  GoSub PrintSummary
  
  Print #RptHandle, FF$
  Close RHandle
  Close PHandle
  Close RptHandle
  Close TMHandle
 
  Close

  ViewPrint RptFile$, "Tax Manual Bill Edit Report", True
  
  KillFile RptFile$
 
  Exit Sub
 
PrintSub1:
  If CombCnt = 0 Then GoTo NoCombCnt
 
  LilYear = CombNewYear + 1
  Nextx = 1
  Do
    For x = Nextx To CombCnt
      Get TMHandle, SortCombRec(x), TaxMRec
      If TaxMRec.TaxYear < LilYear Then
        LilYear = TaxMRec.TaxYear
        Thisx = x
      End If
    Next x
    HoldRec = SortCombRec(Nextx)
    SortCombRec(Nextx) = SortCombRec(Thisx)
    SortCombRec(Thisx) = HoldRec
    If Nextx = 1 Then CombYear = LilYear
    Nextx = Nextx + 1
    If Nextx > CombCnt Then
      Exit Do
    End If
    LilYear = CombNewYear + 1
  Loop
  
  ReDim CombRec(1 To CombCnt) As Integer
  For x = 1 To CombCnt
    CombRec(x) = SortCombRec(x)
  Next x
    
  ReDim Years(1 To 1) As Integer

  For x = 1 To CombCnt
    Get TMHandle, CombRec(x), TaxMRec
      If x = 1 Then
        CombYrCnt = CombYrCnt + 1
        ReDim Preserve Years(1 To CombYrCnt) As Integer
        Years(CombYrCnt) = TaxMRec.TaxYear
      Else
        For y = 1 To CombYrCnt
          If TaxMRec.TaxYear = Years(y) Then
            Exit For
          End If
        Next y
        If y > CombYrCnt Then
          CombYrCnt = CombYrCnt + 1
          ReDim Preserve Years(1 To CombYrCnt) As Integer
          Years(CombYrCnt) = TaxMRec.TaxYear
        End If
      End If
  Next x
      
NoCombCnt:
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
  YrTotInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrGTot = 0
  
  Print #RptHandle, "Bill Type: COMBINED"
  Print #RptHandle, Tab(5); String$(92, ".")
  LineCnt = LineCnt + 2
  If LineCnt > MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintSummary
  End If
  For x = 1 To CombCnt
    Get TMHandle, CombRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotInt = OldRound(YrTotInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrGTot = OldRound(YrTotPrinc + YrTotInt + YrTotAdvCol + YrTotLateList + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    If CombYrCnt > 1 And x < CombCnt Then
      Get TMHandle, CombRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> CombYear And x < CombCnt Then
        Print #RptHandle, "    Year: " + Using$("####", CombYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrGTot)
        Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(62); Opt1Desc; Tab(74); Opt2Desc; Tab(87); Opt3Desc
        Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList);
        Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotOpt1); Tab(73); Using$("$###,##0.00", YrTotOpt2); Tab(86); Using$("$###,##0.00", YrTotOpt3)
        Print #RptHandle, Tab(5); String$(92, ".")
        LineCnt = LineCnt + 4
        If LineCnt > MaxLines - 4 Then
          Print #RptHandle, FF$
          GoSub PrintSummary
        End If
        CombYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrGTot = 0
      End If
    ElseIf x = CombCnt Then
      Print #RptHandle, "    Year: " + Using$("####", CombYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrGTot)
      Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(62); Opt1Desc; Tab(74); Opt2Desc; Tab(87); Opt3Desc
      Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList);
      Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotOpt1); Tab(73); Using$("$###,##0.00", YrTotOpt2); Tab(86); Using$("$###,##0.00", YrTotOpt3)
      LineCnt = LineCnt + 3
      If LineCnt = MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintSummary
      End If
    End If
  Next x
  Print #RptHandle,
  
  YrTotPrinc = 0
  YrTotInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrGTot = 0
  Print #RptHandle, "Bill Type: REAL ONLY"
  Print #RptHandle, Tab(5); String$(92, ".")
  LineCnt = LineCnt + 2
  If LineCnt > MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintSummary
  End If
  For x = 1 To RealCnt
    Get TMHandle, RealRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotInt = OldRound(YrTotInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrGTot = OldRound(YrTotPrinc + YrTotInt + YrTotAdvCol + YrTotLateList + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    If RealYrCnt > 1 And x < RealCnt Then
      Get TMHandle, RealRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> RealYear Then
        Print #RptHandle, "    Year: " + Using$("####", RealYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrGTot)
        Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(62); Opt1Desc; Tab(74); Opt2Desc; Tab(87); Opt3Desc
        Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList);
        Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotOpt1); Tab(73); Using$("$###,##0.00", YrTotOpt2); Tab(86); Using$("$###,##0.00", YrTotOpt3)
        Print #RptHandle, Tab(5); String$(92, ".")
        LineCnt = LineCnt + 4
        If LineCnt > MaxLines - 4 Then
          Print #RptHandle, FF$
          GoSub PrintSummary
        End If
        RealYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrGTot = 0
      End If
    ElseIf x = RealCnt Then
      Print #RptHandle, "    Year: " + Using$("####", RealYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrGTot)
      Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(62); Opt1Desc; Tab(74); Opt2Desc; Tab(87); Opt3Desc
      Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList);
      Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotOpt1); Tab(73); Using$("$###,##0.00", YrTotOpt2); Tab(86); Using$("$###,##0.00", YrTotOpt3)
      LineCnt = LineCnt + 3
      If LineCnt = MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintSummary
      End If
  End If
  Next x
  
  Print #RptHandle,
  YrTotPrinc = 0
  YrTotInt = 0
  YrTotAdvCol = 0
  YrTotLateList = 0
  YrTotOpt1 = 0
  YrTotOpt2 = 0
  YrTotOpt3 = 0
  YrGTot = 0
  Print #RptHandle, "Bill Type: PERSONAL ONLY"
  Print #RptHandle, Tab(5); String$(92, ".")
  LineCnt = LineCnt + 2
  If LineCnt > MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintSummary
  End If
  For x = 1 To PersCnt
    Get TMHandle, PersRec(x), TaxMRec
    YrTotPrinc = OldRound(YrTotPrinc + TaxMRec.TaxAmount)
    YrTotInt = OldRound(YrTotInt + TaxMRec.IntAmount)
    YrTotAdvCol = OldRound(YrTotAdvCol + TaxMRec.AdColAmount)
    YrTotLateList = OldRound(YrTotLateList + TaxMRec.LateList)
    YrTotOpt1 = OldRound(YrTotOpt1 + TaxMRec.OptRev1)
    YrTotOpt2 = OldRound(YrTotOpt2 + TaxMRec.OptRev2)
    YrTotOpt3 = OldRound(YrTotOpt3 + TaxMRec.OptRev3)
    YrGTot = OldRound(YrTotPrinc + YrTotInt + YrTotAdvCol + YrTotLateList + YrTotOpt1 + YrTotOpt2 + YrTotOpt3)
    If PersYrCnt > 1 And x < PersCnt Then
      Get TMHandle, PersRec(x + 1), TaxMRec
      If TaxMRec.TaxYear <> PersYear And x < PersCnt Then
        Print #RptHandle, "    Year: " + Using$("####", PersYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrGTot)
        Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(62); Opt1Desc; Tab(74); Opt2Desc; Tab(87); Opt3Desc
        Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList);
        Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotOpt1); Tab(73); Using$("$###,##0.00", YrTotOpt2); Tab(86); Using$("$###,##0.00", YrTotOpt3)
        Print #RptHandle, Tab(5); String$(92, ".")
        LineCnt = LineCnt + 4
        If LineCnt > MaxLines - 4 Then
          Print #RptHandle, FF$
          GoSub PrintSummary
        End If
        PersYear = TaxMRec.TaxYear
        YrTotPrinc = 0
        YrTotInt = 0
        YrTotAdvCol = 0
        YrTotLateList = 0
        YrTotOpt1 = 0
        YrTotOpt2 = 0
        YrTotOpt3 = 0
        YrGTot = 0
      End If
    ElseIf x = PersCnt Then
      Print #RptHandle, "    Year: " + Using$("####", PersYear); Tab(20); "Year Total: " + Using$("$###,###,##0.00", YrGTot)
      Print #RptHandle, Tab(7); "Principle"; Tab(22); "Interest"; Tab(35); "Adv/Col"; Tab(49); "Late List"; Tab(62); Opt1Desc; Tab(74); Opt2Desc; Tab(87); Opt3Desc
      Print #RptHandle, Tab(5); Using$("$###,##0.00", YrTotPrinc); Tab(19); Using$("$###,##0.00", YrTotInt); Tab(31); Using$("$###,##0.00", YrTotAdvCol); Tab(47); Using$("$###,##0.00", YrTotLateList);
      Print #RptHandle, Tab(61); Using$("$###,##0.00", YrTotOpt1); Tab(73); Using$("$###,##0.00", YrTotOpt2); Tab(86); Using$("$###,##0.00", YrTotOpt3)
      LineCnt = LineCnt + 3
      If LineCnt = MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintSummary
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
  
PrintHeader:
  Page = Page + 1
  RSet Opt1Desc = QPTrim$(Opt1Desc)
  RSet Opt2Desc = QPTrim$(Opt2Desc)
  RSet Opt3Desc = QPTrim$(Opt3Desc)
  Print #RptHandle, Tab(30); "Tax Manual Bill Edit Report"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Str(Page)
  Print #RptHandle,
  Print #RptHandle, "Acct Num"; Tab(13); "Customer Name"
  Print #RptHandle, "Trans Date"; Tab(15); "Bill #"; Tab(25); "Property Type"; Tab(40); "PIN #"
  Print #RptHandle, Tab(5); "Principal"; Tab(17); "Interest"; Tab(28); "Adv/Col"; Tab(37); "LateList"; Tab(50); Opt1Desc; Tab(61); Opt2Desc; Tab(73); Opt3Desc; Tab(87); "Bill Total"
  Print #RptHandle, String$(96, "=")
  LineCnt = 7
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

PrintSummary:
  Page = Page + 1
  RSet Opt1Desc = QPTrim$(Opt1Desc)
  RSet Opt2Desc = QPTrim$(Opt2Desc)
  RSet Opt3Desc = QPTrim$(Opt3Desc)
  Print #RptHandle, Tab(23); "Tax Manual Bill Edit Report Summary"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Str(Page)
  Print #RptHandle, String$(96, "=")
  If OverAllPrinted = False Then
    GoSub PrintOverallAmts
  End If
  LineCnt = 4
  GoSub PrintSub1
  Return
  
PrintOverallAmts:
  TGTotOpt1 = OldRound(TGTotOpt1 + TaxMRec.OptRev1)
  TGTotOpt2 = OldRound(TGTotOpt2 + TaxMRec.OptRev2)
  TGTotOpt3 = OldRound(TGTotOpt3 + TaxMRec.OptRev3)
  TGTotal = OldRound(GTotPrinc + GTotInt + GTotAdvCol + GTotLateList + TGTotOpt1 + TGTotOpt2 + TGTotOpt3)
  OverAllPrinted = True
  Print #RptHandle, "TOTAL OVERALL AMOUNTS:"
  Print #RptHandle, Tab(5); String$(92, "-")
  Print #RptHandle, Tab(5); "Total Bills Entered: " + CStr(ECnt); Tab(35); "Total Amount: " + Using$("$###,###,##0.00", TGTotal)
  Print #RptHandle,
  Print #RptHandle, Tab(5); "Total Principal Amount: "; Tab(32); Using$("$###,###,##0.00", GTotPrinc); Tab(56); "Total Interest Amount: "; Tab(80); Using$("$###,##0.00", GTotInt)
  Print #RptHandle, Tab(7); "Total Adv/Col Amount: "; Tab(36); Using("$###,##0.00", GTotAdvCol); Tab(55); "Total Late List Amount: "; Tab(80); Using$("$###,##0.00", GTotLateList)
  Print #RptHandle, Tab(11); "Total " + Opt1Desc + ": "; Tab(36); Using("$###,##0.00", TGTotOpt1); Tab(61); "Total " + Opt2Desc + ": "; Tab(80); Using$("$###,##0.00", TGTotOpt2)
  Print #RptHandle, Tab(11); "Total " + Opt3Desc + ": "; Tab(36); Using("$###,##0.00", TGTotOpt3)
  Print #RptHandle, Tab(5); String$(92, "-")
  Print #RptHandle,
  LineCnt = LineCnt + 10
  Return
End Sub
