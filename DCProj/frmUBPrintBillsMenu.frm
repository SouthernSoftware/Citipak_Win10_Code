VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBPrintBillsMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Bill Printing Menu"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmUBPrintBillsMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPostalRtRpt 
      Caption         =   "Print Postal &Route Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3864
      TabIndex        =   4
      Top             =   5304
      Width           =   4524
   End
   Begin VB.CommandButton cmdPostedBillReprint 
      Caption         =   "&Posted Bill Reprinting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3864
      TabIndex        =   5
      Top             =   6024
      Width           =   4524
   End
   Begin VB.CommandButton cmdPrnAllUBBills 
      BackColor       =   &H008F8265&
      Caption         =   "Print &All Utility Bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2424
      Width           =   4524
   End
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reprint &Selected Bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3864
      TabIndex        =   1
      Top             =   3144
      Width           =   4524
   End
   Begin VB.CommandButton cmdZipcodeReport 
      Caption         =   "Print &Zipcode Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3864
      TabIndex        =   3
      Top             =   4584
      Width           =   4524
   End
   Begin VB.CommandButton cmdBStatusBills 
      Caption         =   "Print ""&B"" Status Bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3864
      TabIndex        =   2
      Top             =   3864
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitPrnBillsMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3864
      TabIndex        =   6
      Top             =   6744
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4:41 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "12/13/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Utility Bill Printing Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3540
      TabIndex        =   8
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
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
End
Attribute VB_Name = "frmUBPrintBillsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdBStatusBills_Click()
  Load frmBillBalanceDuePrint
  DoEvents
  frmBillBalanceDuePrint.Show
  Unload frmUBPrintBillsMenu
End Sub

Private Sub cmdExitPrnBillsMenu_Click()
  Load frmUBBillingMenu
  DoEvents
  frmUBBillingMenu.Show
  Unload frmUBPrintBillsMenu
End Sub

Private Sub cmdPostalRtRpt_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String

  If Not ChkBillFile% Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO BILL FILE!"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "No Billing Information File Found!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
  DeActivateControls Me
  frmReportOpt.Show 1
  If rptopt = 1 Then
    'do the graphics
    PrintPostalRtReport True
  ElseIf rptopt = 2 Then
    'do the text
   PrintPostalRtReport False
   ActivateControls Me
  Else
    ActivateControls Me
  End If
  
End Sub

Private Sub cmdPostedBillReprint_Click()
  frmCustEditLookUP.Caption = "Posted Bill Reprint Search"
  frmCustEditLookUP.Label1.Caption = "Posted Bill Reprint Search"
  frmCustEditLookUP.Wheretogo frmUBPrintBillsMenu, frmBillPostedRePrn
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBPrintBillsMenu
End Sub

Private Sub cmdPrnAllUBBills_Click()
  frmBillPrinting.REPRN False, False
  Load frmBillPrinting
  DoEvents
  frmBillPrinting.Show
  Unload frmUBPrintBillsMenu
End Sub

Private Sub cmdReprint_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBBILLS.PRN") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO PRN FILE. Reprint"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO BILL PRINT FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If

  If Not ChkBillFile% Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO BILL FILE! Reprint"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO BILL FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
  frmBillPrinting.REPRN True, False
  Load frmBillPrinting
  DoEvents
  frmBillPrinting.Show
  Unload frmUBPrintBillsMenu
End Sub

Private Sub cmdZipcodeReport_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String

  If Not ChkBillFile% Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO BILL FILE!"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "No Billing Information File Found!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
  DeActivateControls Me
  frmReportOpt.Show 1

  If rptopt = 1 Then
    'do the graphics
    PrintZipReport True
  ElseIf rptopt = 2 Then
    'do the text
   PrintZipReport False
   ActivateControls Me
  Else
    ActivateControls Me
  End If
''PrintZipLabel
End Sub

'Private Sub cmdExitUBBillingProcess_Click()
'  Load frmUBMainMenu
'  DoEvents
'  frmUBMainMenu.Show
'  Unload frmUBBillingMenu
'End Sub


'Private Sub cmdUBAdjustments_Click()
'  Load frmAdjustmentEntry
'  DoEvents
'  frmAdjustmentEntry.Show
'  Unload frmUBBillingMenu
'End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitPrnBillsMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via PrintBillsMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
  '  Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdPrnAllUBBills.SetFocus
    Case vbKeyEnd
      cmdExitPrnBillsMenu.SetFocus
    Case Else:
  End Select
End Sub
Private Sub PrintPostalRtReport(Grpt As Boolean)
  Dim RptName As String, MaxLine As Integer, Dash80 As String
  Dim CustRecLen As Integer, UBBillRecLen As Integer, CHandle As Integer
  Dim BHandle As Integer, NumBillRecs As Long, IdxRecLen As Integer
  Dim cnt As Long, PostalCnt As Long, Low As Long, High As Long
  Dim PrnCnt As Long, ZCnt As Long, HadIt As Boolean, PCnt As Long
  Dim UBRpt As Integer
  UBLog " IN: Postal Route Report."
  PageNo = 0
  RptName$ = UBPath$ + "UBPOSTRT.RPT"

  MaxLine = 50
  FrmShowPctComp.Label1 = "Creating Report"
  FrmShowPctComp.Show , Me

  'ShowProcessingScrn "Reading Billing Information."
  Dash80$ = String$(80, "-")
  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  ReDim UBBillRec(1) As UBTransRecType
  UBBillRecLen = Len(UBBillRec(1))

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen

  BHandle = FreeFile
  Open UBPath$ + "UBBILLS.DAT" For Random Shared As BHandle Len = UBBillRecLen
  NumBillRecs& = LOF(BHandle) \ UBBillRecLen

  ReDim PostalIdx(1 To NumBillRecs&) As UBPostalIndexType
  IdxRecLen = Len(PostalIdx(1))
  ReDim PrnArray(1 To 1) As UBPostalIndexType
  For cnt& = 1 To NumBillRecs&
    Get BHandle, cnt&, UBBillRec(1)
    If UBBillRec(1).ActiveFlag Then
      Get CHandle, UBBillRec(1).CustAcctNo, UBCustRec(1)
      PostalCnt = PostalCnt + 1
      PostalIdx(PostalCnt).Route = QPTrim$(UBCustRec(1).POSTRTE)
    End If
    'ShowPctComp cnt&, NumBillRecs&
    FrmShowPctComp.ShowPctComp cnt, NumBillRecs&
    If FrmShowPctComp.Out = True Then
      Close
      UBLog "ABORTED: Printing PostalRt Rpt."
      FrmShowPctComp.Out = False
      Exit Sub
    End If

  Next
  Close
  If PostalCnt <= 0 Then
    MsgBox "No Information to Print.", vbOKOnly, "No Info"
    ActivateControls Me
    GoTo ExitPostalReport
  End If

  ReDim Preserve PostalIdx(1 To PostalCnt) As UBPostalIndexType
  Low = LBound(PostalIdx)
  High = UBound(PostalIdx)
  PostalQSort PostalIdx(), Low, High

  PrnCnt = 1
  ReDim Preserve PrnArray(1 To PrnCnt) As UBPostalIndexType
  PrnArray(PrnCnt).Route = PostalIdx(1).Route
  PrnArray(PrnCnt).RecNum = 1

  For ZCnt = 2 To PostalCnt
    HadIt = False
    For PCnt = 1 To PrnCnt
      If PrnArray(PCnt).Route = PostalIdx(ZCnt).Route Then
        PrnArray(PCnt).RecNum = PrnArray(PCnt).RecNum + 1
        HadIt = True
        Exit For
      End If
    Next
    If Not HadIt Then
      PrnCnt = PrnCnt + 1
      ReDim Preserve PrnArray(1 To PrnCnt) As UBPostalIndexType
      PrnArray(PrnCnt).Route = PostalIdx(ZCnt).Route
      PrnArray(PrnCnt).RecNum = 1
    End If
  Next
  Erase PostalIdx, UBBillRec, UBCustRec

  'ShowProcessingScrn "Processing Zipcode Report."
  FrmShowPctComp.Label1 = "Calculating Report"
  FrmShowPctComp.Show , Me

  UBRpt = FreeFile
  Open RptName$ For Output As UBRpt
  GoSub PostalHeader

  For cnt = 1 To PrnCnt
    FrmShowPctComp.ShowPctComp cnt, PrnCnt
    If FrmShowPctComp.Out = True Then
      Close
      UBLog "ABORTED: Printing PostalRt Rpt."
      FrmShowPctComp.Out = False
      Exit Sub
    End If

    Print #UBRpt, QPTrim(PrnArray(cnt).Route), Tab(30); Using("#####", PrnArray(cnt).RecNum)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLine Then
      Print #UBRpt, Chr$(12)
      GoSub PostalHeader
    End If

  Next

  GoSub PostalFooter

  Close
  If Not Grpt Then
    ViewPrint RptName$, "Billing Postal Route Report."
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBPrintBillsMenu
    ARptLineRpt.GetName RptName$
    ARptLineRpt.startrpt
  End If
ExitPostalReport:
  UBLog "OUT: Postal Route Report."
  Exit Sub

PostalHeader:
  PageNo = PageNo + 1
  If Grpt Then
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
  End If
  Print #UBRpt, Tab(24); "Billing Postal Route Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, Tab(4); "Postal Route"; Tab(30); "Count"
  Print #UBRpt, Dash80$
  If Grpt Then
    LineCnt = 7
  Else
    LineCnt = 4
  End If
  Return

PostalFooter:
  Print #UBRpt, Dash80$
  Print #UBRpt, " Unique Postal Routes:"; PrnCnt; "   Bills Printed:"; PostalCnt
  Print #UBRpt, Chr$(12)
  Return

  End Sub

Private Sub PrintZipReport(Grpt As Boolean)
  Dim RptName As String, MaxLine As Integer, Dash80 As String
  Dim CustRecLen As Integer, UBBillRecLen As Integer, CHandle As Integer
  Dim BHandle As Integer, NumBillRecs As Long, IdxRecLen As Integer
  Dim cnt As Long, ZipCnt As Long, Low As Long, High As Long
  Dim PrnCnt As Long, ZCnt As Long, HadIt As Boolean, PCnt As Long
  Dim UBRpt As Integer
  UBLog " IN: Zipcode Report."
  PageNo = 0
  RptName$ = UBPath$ + "UBZIPRPT.RPT"
  MaxLine = 50
  FrmShowPctComp.Label1 = "Creating Report"
  FrmShowPctComp.Show , Me

  'ShowProcessingScrn "Reading Billing Information."
  Dash80$ = String$(80, "-")

  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  ReDim UBBillRec(1) As UBTransRecType
  UBBillRecLen = Len(UBBillRec(1))

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen

  BHandle = FreeFile
  Open UBPath$ + "UBBILLS.DAT" For Random Shared As BHandle Len = UBBillRecLen
  NumBillRecs& = LOF(BHandle) \ UBBillRecLen

  ReDim ZipIndex(1 To NumBillRecs&) As UBPostalIndexType
  IdxRecLen = Len(ZipIndex(1))
  ReDim PrnArray(1 To 1) As UBPostalIndexType
  For cnt& = 1 To NumBillRecs&
    Get BHandle, cnt&, UBBillRec(1)
    If UBBillRec(1).ActiveFlag Then
      Get CHandle, UBBillRec(1).CustAcctNo, UBCustRec(1)
      ZipCnt = ZipCnt + 1
      'IF HamFlag THEN
        ZipIndex(ZipCnt).ZIPCODE = Left$(QPTrim(UBCustRec(1).ZIPCODE), 5)
      'ELSE
      '  RSET ZipIndex(ZipCnt).ZipCode = QPTrim$(UBCustRec(1).ZipCode)
      'END IF
    End If
    'ShowPctComp cnt&, NumBillRecs&
    FrmShowPctComp.ShowPctComp cnt, NumBillRecs&
    If FrmShowPctComp.Out = True Then
      Close
      UBLog "ABORTED: Printing Zip Rpt."
      FrmShowPctComp.Out = False
      Exit Sub
    End If

  Next
  Close
  If ZipCnt <= 0 Then
    MsgBox "No Information to Print.", vbOKOnly, "No Info"
    ActivateControls Me
    GoTo ExitZipReport
  End If
  ReDim Preserve ZipIndex(1 To ZipCnt) As UBPostalIndexType
  Low = LBound(ZipIndex)
  High = UBound(ZipIndex)
  PostZipQSort ZipIndex(), Low, High
  PrnCnt = 1
  ReDim Preserve PrnArray(1 To PrnCnt) As UBPostalIndexType
  PrnArray(PrnCnt).ZIPCODE = ZipIndex(1).ZIPCODE
  PrnArray(PrnCnt).RecNum = 1

  For ZCnt = 2 To ZipCnt
    HadIt = False
    For PCnt = 1 To PrnCnt
      If PrnArray(PCnt).ZIPCODE = ZipIndex(ZCnt).ZIPCODE Then
        PrnArray(PCnt).RecNum = PrnArray(PCnt).RecNum + 1
        HadIt = True
        Exit For
      End If
    Next
    If Not HadIt Then
      PrnCnt = PrnCnt + 1
      ReDim Preserve PrnArray(1 To PrnCnt) As UBPostalIndexType
      PrnArray(PrnCnt).ZIPCODE = ZipIndex(ZCnt).ZIPCODE
      PrnArray(PrnCnt).RecNum = 1
    End If
   ' ShowPctComp ZCnt, ZipCnt
  Next
  Erase ZipIndex, UBBillRec, UBCustRec

  'ShowProcessingScrn "Processing Zipcode Report."
  FrmShowPctComp.Label1 = "Calculating Report"
  FrmShowPctComp.Show , Me
  UBRpt = FreeFile
  Open RptName$ For Output As UBRpt
  GoSub ZipHeader
  For cnt = 1 To PrnCnt
    Print #UBRpt, QPTrim(PrnArray(cnt).ZIPCODE), Tab(30); Using("#####", PrnArray(cnt).RecNum)
    FrmShowPctComp.ShowPctComp cnt, PrnCnt
    If FrmShowPctComp.Out = True Then
      Close
      UBLog "ABORTED: Printing Zip Rpt."
      FrmShowPctComp.Out = False
      Exit Sub
    End If
    LineCnt = LineCnt + 1
    If LineCnt > MaxLine Then
      Print #UBRpt, Chr$(12)
      GoSub ZipHeader
    End If

    'ShowPctComp cnt, PrnCnt
  Next

  GoSub ZipFooter
  Close
  
  If Not Grpt Then
     ViewPrint RptName$, "Billing Zipcode Report."
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBPrintBillsMenu
    ARptLineRpt.GetName RptName$
    ARptLineRpt.startrpt
  End If

   
ExitZipReport:
  UBLog "OUT: Zipcode Report."
  Exit Sub

ZipHeader:
  PageNo = PageNo + 1
  If Grpt Then
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
  End If
  Print #UBRpt, Tab(24); "Billing Zipcode Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, Tab(4); "Zip Code"; Tab(30); "Count"
  Print #UBRpt, Dash80$
  If Grpt Then
    LineCnt = 7
  Else
    LineCnt = 4
  End If
  Return

ZipFooter:
  Print #UBRpt, Dash80$
  Print #UBRpt, " Unique Zipcodes:"; PrnCnt; "   Bills Printed:"; ZipCnt
  Print #UBRpt, Chr$(12)
  Return

End Sub
Private Sub PrintZipLabel()
  Dim RptName As String, cnt As Integer
  Dim UBRpt As Integer, ToPrint As String
  RptName$ = UBPath$ + "UBZIPLab.RPT"
'  FrmShowPctComp.Label1 = "Creating Report"
'  FrmShowPctComp.Show , Me
  ToPrint$ = ""

  'ShowProcessingScrn "Processing Zipcode Report."
'  FrmShowPctComp.Label1 = "Calculating Report"
'  FrmShowPctComp.Show , Me
  UBRpt = FreeFile
  Open RptName$ For Output As UBRpt
  ToPrint$ = " ~ ~MXD GREENSBORO NC~FCM LTR BC WKG"
  ToPrint$ = ToPrint$ + "~ ~270~2700024601~HOLLY SPRINGS NC 27540"
  For cnt = 1 To 4
    Print #UBRpt, ToPrint$
  Next
'    FrmShowPctComp.ShowPctComp cnt, PrnCnt
'    If FrmShowPctComp.Out = True Then
'      Close
'      UBLog "ABORTED: Printing Zip Rpt."
'      FrmShowPctComp.Out = False
'      Exit Sub
'    End If
'    LineCnt = LineCnt + 1
'    If LineCnt > MaxLine Then
'      Print #UBRpt, Chr$(12)
'      GoSub ZipHeader
'    End If

    'ShowPctComp cnt, PrnCnt
'  Next

'  GoSub ZipFooter
  Close
  
   'Load frmLoadingRpt
   ' frmLoadingRpt.setwherefrom frmUBPrintBillsMenu
    ARptLineRptBars.GetName RptName$
    ARptLineRptBars.startrpt

   

End Sub

