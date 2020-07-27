VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxBillPrintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Printing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxBillPrintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdBillReport 
      Height          =   432
      Left            =   4008
      TabIndex        =   2
      Top             =   4860
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
      ButtonDesigner  =   "frmTaxBillPrintMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprintBills 
      Height          =   435
      Left            =   4005
      TabIndex        =   1
      Tag             =   "0"
      Top             =   4275
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
      ButtonDesigner  =   "frmTaxBillPrintMenu.frx":0AB6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintBills 
      Height          =   432
      Left            =   4008
      TabIndex        =   0
      Top             =   3696
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
      ButtonDesigner  =   "frmTaxBillPrintMenu.frx":0C9B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   432
      Left            =   4008
      TabIndex        =   3
      Top             =   5436
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
      ButtonDesigner  =   "frmTaxBillPrintMenu.frx":0E7E
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
      Caption         =   "TAX BILL PRINT MENU"
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
      TabIndex        =   4
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
Attribute VB_Name = "frmTaxBillPrintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdBillReport_Click()
  If Not Exist(TaxBillFile) Then
    Call TaxMsg(900, "No pre-billing records could be found. Make sure pre-billing is completed before continuing.")
    Exit Sub
  End If
  
  If Not Exist("txblsprn.dat") Then
    Call TaxMsg(900, "ERROR: Tax bills have not been printed.")
    Close
    Exit Sub
  End If
  
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

Private Sub cmdExit_Click()
  frmTaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintBills_Click()
  frmTaxBillPrinting.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdReprintBills_Click()
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  If Not Exist("txblsprn.dat") Then
    Call TaxMsg(900, "Tax bills have not yet been printed. Please process tax bills.")
    Exit Sub
  End If
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Select Case TaxMasterRec.TaxForm
    Case 29999, 20000, 20001
      Call TaxMsg(900, "The current tax form saved is in an export format. With this format reprints are unnecessary.")
      Exit Sub
    Case Else
  End Select
  
  frmTaxBillReprinting.Show
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
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxBillPrint

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxBillPrintMenu.")
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
  Dim x As Long
  Dim dlm$
  Dim Town$
  Dim TBillRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TotReal As Double
  Dim TotBillCnt As Long
  Dim TotPers As Double
  Dim Total As Double
  Dim TotCredit As Double
  Dim TotOwed As Double
  Dim UseMinBill As Integer '12/7/06
  Dim MinBillAmt As Double '12/7/06
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  'on error goto ERRORSTUFF
  
  dlm$ = "~"
  OpenBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  UseMinBill = TaxMasterRec.MinTxOpt
  MinBillAmt = TaxMasterRec.MinBill
  
  RptFile$ = "TAXRPTS\TXBLRPT.RPT"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  frmTaxShowPctComp.Label1 = "Creating Tax Billing Report"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("MORTIDX.DAT") = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("ZIPIDX.DAT") = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBillRec
    If TBillRec.BillNumber <= 0 Then GoTo SkipEr
    If UseMinBill = 1 Then '12/7/06
      If OldRound(TBillRec.RealTaxDue + TBillRec.PersTaxDue + TBillRec.LateTaxDue) < MinBillAmt Then
        GoTo SkipEr
      End If
    End If
    TotReal = TotReal + TBillRec.RealTaxDue + TBillRec.LateTaxDue
    TotPers = TotPers + TBillRec.PersTaxDue
    TotCredit = TotCredit + TBillRec.OverPayAmt
    TotBillCnt = TotBillCnt + 1
    Total = Total + TBillRec.RealTaxDue + TBillRec.PersTaxDue + TBillRec.LateTaxDue
    TotOwed = Total - TotCredit
    '                   0                     1                                      2
    Print #RptHandle, Town; dlm; QPTrim$(BillInfo.CountyPara); dlm; QPTrim$(BillInfo.CyclePara); dlm;
    '                                3                                  4                         5
    Print #RptHandle, QPTrim$(BillInfo.TwnShpPara); dlm; QPTrim$(BillInfo.SplitPara); dlm; BillInfo.TaxYear; dlm;
    '                         6                           7                              8
    Print #RptHandle, TBillRec.BillNumber; dlm; QPTrim$(TBillRec.CustName); dlm; TBillRec.RealTaxDue + TBillRec.LateTaxDue; dlm;
    '                          9                        10                      11              12
    Print #RptHandle, TBillRec.PersTaxDue; dlm; TBillRec.TotalBillDue; dlm; TotBillCnt; dlm; TotReal; dlm;
    '                    13           14               15                                    16
    Print #RptHandle, TotPers; dlm; Total; dlm; TBillRec.OverPayAmt; dlm; OldRound(TBillRec.TotalBillDue - TBillRec.OverPayAmt); dlm;
    '                     17             18
    Print #RptHandle, TotCredit; dlm; TotOwed
    
SkipEr:
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  Close
   
  arTaxBillingRpt.Show
  frmTaxLoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrintMenu", "PrintGraphics", Erl)
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
    Close
 
  
End Sub

Private Sub PrintText()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Long
  Dim Town$
  Dim TBillRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TotReal As Double
  Dim TotBillCnt As Long
  Dim TotPers As Double
  Dim Total As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Line$
  Dim Page As Integer, FF$
  Dim TotCredit As Double
  Dim TotOwed As Double
  Dim UseMinBill As Integer '12/7/06
  Dim MinBillAmt As Double '12/7/06
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  'on error goto ERRORSTUFF
  
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  Line = String$(80, "-")
  OpenBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TXBLRPT.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  frmTaxShowPctComp.Label1 = "Creating Tax Billing Report"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("MORTIDX.DAT") = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("ZIPIDX.DAT") = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBillRec
    If TBillRec.BillNumber <= 0 Then GoTo SkipEr
    If UseMinBill = 1 Then '12/7/06
      If OldRound(TBillRec.RealTaxDue + TBillRec.PersTaxDue + TBillRec.LateTaxDue) < MinBillAmt Then
        GoTo SkipEr
      End If
    End If
    TotReal = TotReal + TBillRec.RealTaxDue + TBillRec.LateTaxDue
    TotPers = TotPers + TBillRec.PersTaxDue
    TotCredit = TotCredit + TBillRec.OverPayAmt
    TotBillCnt = TotBillCnt + 1
    Total = Total + TBillRec.RealTaxDue + TBillRec.PersTaxDue + TBillRec.LateTaxDue
    TotOwed = Total - TotCredit
    Print #RptHandle, Using("####0", TBillRec.BillNumber);
    Print #RptHandle, Tab(12); QPTrim$(Left$(TBillRec.CustName, 32));
    Print #RptHandle, Tab(45); Using("##,##0.00", TBillRec.RealTaxDue + TBillRec.LateTaxDue); Tab(59); Using("##,##0.00", TBillRec.PersTaxDue);
    Print #RptHandle, Tab(69); Using("#,###,##0.00", TBillRec.TotalBillDue)
    LineCnt = LineCnt + 1
    If TBillRec.OverPayAmt > 0 Then
      Print #RptHandle, Tab(12); "Credit Applied to Bill: " + QPTrim$(Using$("$##,##0.00", TBillRec.OverPayAmt)); Tab(50); "Total Owed: " + QPTrim$(Using$("$##,##0.00", TBillRec.TotalBillDue - TBillRec.OverPayAmt))
      Print #RptHandle, String(80, ".")
      LineCnt = LineCnt + 2
    End If
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
SkipEr:
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  If LineCnt >= MaxLines - 2 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, Line
  Print #RptHandle, "Billing Totals:"; Tab(20); CStr(TotBillCnt); Tab(41); Using("$#,###,##0.00", TotReal); Tab(55); Using("$#,###,##0.00", TotPers); Tab(68); Using("$#,###,##0.00", Total)
  If TotCredit > 0 Then
    Print #RptHandle, Tab(5); "Total Credit Applied: "; Tab(68); Using$("$#,###,##0.00", -TotCredit)
    Print #RptHandle, Tab(5); "Total Balance: "; Tab(68); Using$("$#,###,##0.00", Total - TotCredit)
  End If
  Print #RptHandle, FF$
  
  Close

  ViewPrint RptFile$, "Tax Bills Printed Report", True
  
  KillFile RptFile$
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(14); "Property Tax Billing : Bills Printed Report For Tax Year "; CStr(BillInfo.TaxYear)
  Print #RptHandle, "Town: "; Town
  Print #RptHandle, "County: "; QPTrim$(BillInfo.CountyPara); Tab(50); "Cycle: "; QPTrim$(BillInfo.CyclePara)
  Print #RptHandle, "Township: "; QPTrim$(BillInfo.TwnShpPara); Tab(50); "Split Real/Pers: "; QPTrim$(BillInfo.SplitPara)
  Print #RptHandle, "Date: "; CStr(Date); Tab(74); "Page #"; CStr(Page)
  Print #RptHandle,
  Print #RptHandle, "Bill No."; Tab(12); "Customer Name                     Real Due      Pers Due        Total"
  Print #RptHandle, Line$
  LineCnt = 8

  Return
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrintMenu", "PrintText", Erl)
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
    Close
 
  
End Sub

