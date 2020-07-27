VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxBillPrintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Printing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxBillPrintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdBillReport 
      Height          =   420
      Left            =   3960
      TabIndex        =   2
      Top             =   4890
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmVATaxBillPrintMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprintBills 
      Height          =   444
      Left            =   3960
      TabIndex        =   1
      Tag             =   "0"
      Top             =   4296
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
      ButtonDesigner  =   "frmVATaxBillPrintMenu.frx":0AB6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintBills 
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
      ButtonDesigner  =   "frmVATaxBillPrintMenu.frx":0C9B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   432
      Left            =   3960
      TabIndex        =   3
      Top             =   5460
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
      ButtonDesigner  =   "frmVATaxBillPrintMenu.frx":0E7E
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
Attribute VB_Name = "frmVATaxBillPrintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class

Private Sub cmdBillReport_Click()
  If Not Exist(RealTaxBillFile) And Not Exist(PersTaxBillFile) Then
    Call TaxMsg(800, "No real or personal pre-billing records could be found. Make sure either personal or real pre-billing is completed before continuing.")
    Exit Sub
  End If
  
  If Exist(RealTaxBillFile) And Not Exist(PersTaxBillFile) Then
    frmVATaxReportOpt.Show vbModal
    If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
      Unload frmVATaxReportOpt
      Call PrintRealGraphics
    ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
      frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Unload frmVATaxReportOpt
      Call PrintRealText
    End If
  ElseIf Exist(PersTaxBillFile) And Not Exist(RealTaxBillFile) Then
    frmVATaxReportOpt.Show vbModal
    If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
      Unload frmVATaxReportOpt
      Call PrintPersGraphics
    ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
      frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Unload frmVATaxReportOpt
      Call PrintPersText
    End If
  Else
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      Unload frmVATaxBillPostOpt
      frmVATaxReportOpt.Show vbModal
      If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
        Unload frmVATaxReportOpt
        Call PrintRealGraphics
      ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
        frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
        frmVATaxMsg.Label1.Top = 900
        frmVATaxMsg.Show vbModal
        Unload frmVATaxReportOpt
        Call PrintRealText
      End If
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      Unload frmVATaxBillPostOpt
      frmVATaxReportOpt.Show vbModal
      If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
        Unload frmVATaxReportOpt
        Call PrintPersGraphics
      ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
        frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
        frmVATaxMsg.Label1.Top = 900
        frmVATaxMsg.Show vbModal
        Unload frmVATaxReportOpt
        Call PrintPersText
      End If
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmVATaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintBills_Click()
  If Not Exist(RealTaxBillInfoFile) And Not Exist(PersTaxBillInfoFile) Then
    Call TaxMsg(900, "Please process tax prebilling before printing tax bills.")
    Exit Sub
  End If
  
  frmVATaxBillPrinting.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdReprintBills_Click()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisForm As Form
  
  Set ThisForm = New frmVATaxBillReprinting
  
  If Not Exist(RealTaxBillInfoFile) And Not Exist(PersTaxBillInfoFile) Then
    Call TaxMsg(900, "Tax bills have not yet been printed. Please process tax bills.")
    Exit Sub
  End If
  
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
    If Not Exist("txrblsprn.dat") Then
      Call TaxMsg(900, "Real tax bills have not yet been printed. Please process tax bills.")
      Exit Sub
    Else
      frmVATaxBillReprinting.fpcmbType.Text = "REAL"
    End If
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
    If Not Exist("txpblsprn.dat") Then
      Call TaxMsg(900, "Personal tax bills have not yet been printed. Please process tax bills.")
      Exit Sub
    Else
      frmVATaxBillReprinting.fpcmbType.Text = "PERSONAL"
    End If
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    DoEvents
    Unload frmVATaxBillPostOpt
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
  
  frmVATaxBillReprinting.Show
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
  Me.HelpContextID = hlpTaxBillPrint

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxBillPrintMenu.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub PrintRealGraphics()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Long
  Dim dlm$
  Dim Town$
  Dim TBillRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TotReal As Double
  Dim TotBillCnt As Long
  Dim TotPers As Double
  Dim Total As Double
  Dim TotCredit As Double
  Dim TotOwed As Double
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintRZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim AHandle As Integer
  
'  On Error GoTo ERRORSTUFF
'  AHandle = FreeFile
'  Open ("billreport.dat") For Output As AHandle
  dlm$ = "~"
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TXBLRPT.RPT"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("MORTIDX.DAT") = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("RZipIdx.Dat") = True Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  frmVATaxShowPctComp.Label1 = "Creating Real Tax Billing Report"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  
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
    If TBillRec.BillPrinted = False Then GoTo SkipEr
    If TBillRec.BillNumber <= 0 Then GoTo SkipEr
    TotReal = OldRound(TotReal + TBillRec.RealTaxDue + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    TotCredit = TotCredit + TBillRec.OverPayAmt
    TotBillCnt = TotBillCnt + 1
    Total = OldRound(Total + TBillRec.RealTaxDue + TBillRec.LateTaxDue)
    TotOwed = Total - TotCredit
'    Print #AHandle, CStr(TBillRec.CustRec) + "~" + Using$("###,###,##0.00", TBillRec.TotalBillDue)
    '                   0                     1                                      2
    Print #RptHandle, Town; dlm; QPTrim$(BillInfo.CountyPara); dlm; QPTrim$(BillInfo.CyclePara); dlm;
    '                                3                                  4                         5
    Print #RptHandle, QPTrim$(BillInfo.TwnShpPara); dlm; QPTrim$(BillInfo.SplitPara); dlm; BillInfo.TaxYear; dlm;
    '                         6                           7                              8
    Print #RptHandle, TBillRec.BillNumber; dlm; QPTrim$(TBillRec.CustName); dlm; TBillRec.RealTaxDue + TBillRec.LateTaxDue; dlm;
    '                          9                        10                      11              12
    Print #RptHandle, TBillRec.CustRec; dlm; TBillRec.TotalBillDue; dlm; TotBillCnt; dlm; TotReal; dlm;
    '                    13           14               15                                    16
    Print #RptHandle, TotPers; dlm; Total; dlm; TBillRec.OverPayAmt; dlm; OldRound(TBillRec.TotalBillDue - TBillRec.OverPayAmt); dlm;
    '                     17             18
    Print #RptHandle, TotCredit; dlm; TotOwed
    
SkipEr:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  Close
   
  arVATaxBillingRpt.Show
  frmVATaxLoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrintMenu", "PrintRealGraphics", Erl)
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

Private Sub PrintRealText()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Long
  Dim Town$
  Dim TBillRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TotReal As Double
  Dim TotBillCnt As Long
  Dim ThisOth As Double
  Dim TotOth As Double
  Dim Total As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Line$
  Dim Page As Integer, FF$
  Dim TotCredit As Double
  Dim TotOwed As Double
  Dim HeadLen As Integer
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintRZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  Line = String$(80, "-")
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TXRBLRPT.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  frmVATaxShowPctComp.Label1 = "Creating Real Tax Billing Report"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("MORTIDX.DAT") = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("RZipIdx.Dat") = True Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
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
    If TBillRec.BillPrinted = False Then GoTo SkipEr
    If TBillRec.BillNumber <= 0 Then GoTo SkipEr
    TotReal = TotReal + TBillRec.RealTaxDue + TBillRec.LateTaxDue
    TotCredit = TotCredit + TBillRec.OverPayAmt
    TotBillCnt = TotBillCnt + 1
    Total = Total + TBillRec.RealTaxDue + TBillRec.LateTaxDue ' + TBillRec.PersTaxDue
    ThisOth = OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    TotOth = TotOth + ThisOth
    TotOwed = Total - TotCredit
    If LineCnt >= MaxLines - 2 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    Print #RptHandle, Using("####0", TBillRec.BillNumber);
    Print #RptHandle, Tab(8); Using$("####0", TBillRec.CustRec);
    Print #RptHandle, Tab(16); QPTrim$(Left$(TBillRec.CustName, 30));
    Print #RptHandle, Tab(47); Using("##,##0.00", TBillRec.RealTaxDue + TBillRec.LateTaxDue);
    Print #RptHandle, Tab(58); Using("##,##0.00", TBillRec.OverPayAmt);
    Print #RptHandle, Tab(69); Using("#,###,##0.00", TBillRec.TotalBillDue - TBillRec.OverPayAmt)
    LineCnt = LineCnt + 1
'    If ThisOth > 0 Then
'      Print #RptHandle, Tab(8); "Other Tax:" + Using$("$#,##0.00", ThisOth)
'      LineCnt = LineCnt + 1
'    End If
SkipEr:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  If LineCnt >= MaxLines - 2 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, Line
  Print #RptHandle, "Billing Totals:"; Tab(20); CStr(TotBillCnt); Tab(43); Using("$#,###,##0.00", TotReal); Tab(56); Using("$###,##0.00", TotCredit); Tab(68); Using("$#,###,##0.00", OldRound(Total - TotCredit))
'  If TotOth > 0 Then
'    Print #RptHandle, "Other Taxes: "; Tab(20); Using$("$###,##0.00", TotOth)
'  End If
  Print #RptHandle, FF$
  
  Close

  ViewPrint RptFile$, "Real Tax Bills Printed Report", True
  
  KillFile RptFile$
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  
  Print #RptHandle, Tab(12); "Property Tax Billing : Bills Printed Report For Tax Year "; CStr(BillInfo.TaxYear)
  Print #RptHandle, "Town: "; Town
  HeadLen = Len(QPTrim$(BillInfo.CyclePara))
  HeadLen = 81 - (HeadLen + 7)
  Print #RptHandle, "County: "; QPTrim$(BillInfo.CountyPara); Tab(HeadLen); "Cycle: "; QPTrim$(BillInfo.CyclePara)
  HeadLen = Len(QPTrim$(BillInfo.SplitPara))
  HeadLen = 81 - (HeadLen + 17)
  Print #RptHandle, "Township: "; QPTrim$(BillInfo.TwnShpPara); Tab(HeadLen); "Split Real/Pers: "; QPTrim$(BillInfo.SplitPara)
  Print #RptHandle, "Date: "; CStr(Date); Tab(74); "Page #"; CStr(Page)
  Print #RptHandle,
  Print #RptHandle, "Bill #"; Tab(9); "Acct #"; Tab(16); "Customer Name"; Tab(49); "Tax Due"; Tab(60); "Credits"; Tab(76); "Total"
  Print #RptHandle, Line$
  LineCnt = 8

  Return
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrintMenu", "PrintRealText", Erl)
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

Private Sub PrintPersGraphics()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Long
  Dim dlm$
  Dim Town$
  Dim TBillRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TotBillCnt As Long
  Dim TotPers As Double
  Dim TotMT As Double
  Dim TotMC As Double
  Dim TotFE As Double
  Dim TotMH As Double
  Dim TotOth As Double
  Dim Total As Double
  Dim TotCredit As Double
  Dim TotOwed As Double
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim AHandle As Integer
  
  On Error GoTo ERRORSTUFF
'  AHandle = FreeFile
'  Open ("pbillreport.dat") For Output As AHandle
  
  dlm$ = "~"
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TXPBLRPT.RPT"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  frmVATaxShowPctComp.Label1 = "Creating Personal Tax Billing Report"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  If Exist("PZipIdx.Dat") = True Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBillRec
    If TBillRec.BillPrinted = False Then GoTo SkipEr
    If TBillRec.BillNumber <= 0 Then GoTo SkipEr
    TotPers = OldRound(TotPers + TBillRec.PersTaxDue - TBillRec.PPTRADiscnt)
    TotMT = OldRound(TotMT + TBillRec.MTTaxDue)
    TotMC = OldRound(TotMC + TBillRec.MCTaxDue)
    TotFE = OldRound(TotFE + TBillRec.FETaxDue)
    TotMH = OldRound(TotMH + TBillRec.MHTaxDue)
    TotOth = OldRound(TotOth + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    TotCredit = TotCredit + TBillRec.OverPayAmt
    TotBillCnt = TotBillCnt + 1
    Total = OldRound(Total + TBillRec.PersTaxDue + TBillRec.MTTaxDue + TBillRec.MCTaxDue + TBillRec.FETaxDue + TBillRec.MHTaxDue - TBillRec.PPTRADiscnt)
    Total = OldRound(Total + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    TotOwed = Total - TotCredit
'    Print #AHandle, CStr(TBillRec.CustRec) + "~" + Using$("###,###,##0.00", TBillRec.TotalBillDue)
    '                   0                     1                                      2
    Print #RptHandle, Town; dlm; QPTrim$(BillInfo.CountyPara); dlm; QPTrim$(BillInfo.CyclePara); dlm;
    '                                3                                  4                         5
    Print #RptHandle, QPTrim$(BillInfo.TwnShpPara); dlm; QPTrim$(BillInfo.SplitPara); dlm; BillInfo.TaxYear; dlm;
    '                         6                           7                                8
    Print #RptHandle, TBillRec.BillNumber; dlm; QPTrim$(TBillRec.CustName); dlm; OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt); dlm;
    '                          9                        10                      11              12
    Print #RptHandle, TBillRec.CustRec; dlm; TBillRec.TotalBillDue; dlm; TotBillCnt; dlm; TotPers; dlm;
    '                    13           14               15                                    16
    Print #RptHandle, TotPers; dlm; Total; dlm; TBillRec.OverPayAmt; dlm; OldRound(TBillRec.TotalBillDue - TBillRec.OverPayAmt); dlm;
    '                     17             18           19          20         21          22           23                24
    Print #RptHandle, TotCredit; dlm; TotOwed; dlm; TotMT; dlm; TotMC; dlm; TotFE; dlm; TotMH; dlm; TotOth; dlm; TBillRec.MTTaxDue; dlm;
    '                        25                       26                      27
    Print #RptHandle, TBillRec.MCTaxDue; dlm; TBillRec.FETaxDue; dlm; TBillRec.MHTaxDue; dlm;
    '                                                    28
    Print #RptHandle, OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    
SkipEr:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  Close
   
  arVATaxPersBillingRpt.Show
  frmVATaxLoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrintMenu", "PrintPersGraphics", Erl)
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

Private Sub PrintPersText()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Long
  Dim Town$
  Dim TBillRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TotBillCnt As Long
  Dim TotPers As Double
  Dim TotMT As Double
  Dim TotMC As Double
  Dim TotFE As Double
  Dim TotMH As Double
  Dim TotOth As Double
  Dim Total As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Line$
  Dim Page As Integer, FF$
  Dim TotCredit As Double
  Dim TotOwed As Double
  Dim HeadLen As Integer
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  Line = String$(80, "-")
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TXPBLRPT.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  frmVATaxShowPctComp.Label1 = "Creating Personal Tax Billing Report"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("PZipIdx.Dat") = True Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBillRec
    If TBillRec.BillPrinted = False Then GoTo SkipEr
    If TBillRec.BillNumber <= 0 Then GoTo SkipEr
    TotPers = OldRound(TotPers + TBillRec.PersTaxDue - TBillRec.PPTRADiscnt)
    TotMT = OldRound(TotMT + TBillRec.MTTaxDue)
    TotMC = OldRound(TotMC + TBillRec.MCTaxDue)
    TotFE = OldRound(TotFE + TBillRec.FETaxDue)
    TotMH = OldRound(TotMH + TBillRec.MHTaxDue)
    TotOth = OldRound(TotOth + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    TotCredit = OldRound(TotCredit + TBillRec.OverPayAmt)
    TotBillCnt = TotBillCnt + 1
    Total = Total + TBillRec.TotalBillDue
    TotOwed = Total - TotCredit
    If LineCnt >= MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    Print #RptHandle, Using("####0", TBillRec.BillNumber);
    Print #RptHandle, Tab(8); Using("####0", TBillRec.CustRec);
    Print #RptHandle, Tab(18); QPTrim$(Left$(TBillRec.CustName, 30));
    Print #RptHandle, Tab(49); Using("##,##0.00", TBillRec.TotalBillDue);
    Print #RptHandle, Tab(60); Using("##,##0.00", TBillRec.OverPayAmt);
    Print #RptHandle, Tab(69); Using("#,###,##0.00", TBillRec.TotalBillDue - TBillRec.OverPayAmt)
    Print #RptHandle, Tab(10); Using("$##,##0.00", TBillRec.PersTaxDue - TBillRec.PPTRADiscnt);
    Print #RptHandle, Tab(21); Using("$##,##0.00", TBillRec.MTTaxDue); Tab(32); Using("$##,##0.00", TBillRec.MCTaxDue);
    Print #RptHandle, Tab(43); Using("$##,##0.00", TBillRec.FETaxDue); Tab(54); Using("$##,##0.00", TBillRec.MHTaxDue);
    Print #RptHandle, Tab(65); Using("$##,##0.00", OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
    Print #RptHandle, Line
    LineCnt = LineCnt + 3
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    
SkipEr:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  If LineCnt >= MaxLines - 8 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, "Billing Totals:"
  Print #RptHandle, "Bill Count: " + CStr(TotBillCnt)
  Print #RptHandle, Tab(8); "Tot Personal: "; Tab(28); Using("$#,###,##0.00", TotPers); Tab(48); "Total Other: "; Tab(68); Using$("$#,###,##0.00", TotOth)
  Print #RptHandle, Tab(8); "Total Mach Tools: "; Tab(28); Using("$#,###,##0.00", TotMT); Tab(48); "Total Owed: "; Tab(68); Using$("$#,###,##0.00", Total)
  Print #RptHandle, Tab(8); "Total Merch Cap: "; Tab(28); Using("$#,###,##0.00", TotMC); Tab(48); "Credit Applied: "; Tab(68); Using$("$#,###,##0.00", TotCredit)
  Print #RptHandle, Tab(8); "Total Farm Equip: "; Tab(28); Using("$#,###,##0.00", TotFE); Tab(48); "Total Billed: "; Tab(68); Using$("$#,###,##0.00", OldRound(Total - TotCredit))
  Print #RptHandle, Tab(8); "Total Mbl Homes: "; Tab(28); Using("$#,###,##0.00", TotMH)
  Print #RptHandle, FF$
  
  Close

  ViewPrint RptFile$, "Personal Tax Bills Printed Report", True
  
  KillFile RptFile$
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(6); "Personal Property Tax Billing : Bills Printed Report For Tax Year "; CStr(BillInfo.TaxYear)
  Print #RptHandle, "Town: "; Town
  HeadLen = Len(QPTrim$(BillInfo.CyclePara))
  HeadLen = 81 - (HeadLen + 7)
  Print #RptHandle, "County: "; QPTrim$(BillInfo.CountyPara); Tab(HeadLen); "Cycle: "; QPTrim$(BillInfo.CyclePara)
  HeadLen = Len(QPTrim$(BillInfo.SplitPara))
  HeadLen = 81 - (HeadLen + 17)
  Print #RptHandle, "Township: "; QPTrim$(BillInfo.TwnShpPara); Tab(HeadLen); "Split Real/Pers: "; QPTrim$(BillInfo.SplitPara)
  Print #RptHandle, "Date: "; CStr(Date); Tab(74); "Page #"; CStr(Page)
  Print #RptHandle,
  Print #RptHandle, "Bill #"; Tab(10); "Acct #"; Tab(18); "Customer Name"; Tab(51); "Tax Due"; Tab(62); "Credits"; Tab(76); "Total"
  Print #RptHandle, Tab(12); "Pers Tax"; Tab(25); "MT Tax"; Tab(36); "MC Tax"; Tab(47); "FE Tax"; Tab(58); "MH Tax"; Tab(70); "Other"
  Print #RptHandle, Line$
  LineCnt = 8

  Return
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrintMenu", "PrintPersText", Erl)
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

