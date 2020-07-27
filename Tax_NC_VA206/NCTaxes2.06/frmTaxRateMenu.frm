VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxRateMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Optional Revenue Rate Table Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxRateMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
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
      ButtonDesigner  =   "frmTaxRateMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
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
      ButtonDesigner  =   "frmTaxRateMenu.frx":0AB5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEdit 
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
      ButtonDesigner  =   "frmTaxRateMenu.frx":0CA5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAddNew 
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
      ButtonDesigner  =   "frmTaxRateMenu.frx":0E90
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
      ButtonDesigner  =   "frmTaxRateMenu.frx":1075
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
      Caption         =   "TAX OPTIONAL REVENUE RATE TABLE MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2513
      TabIndex        =   5
      Top             =   1170
      Width           =   6615
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
Attribute VB_Name = "frmTaxRateMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAddNew_Click()
  Dim One As Integer
  Dim AHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  'on error goto ERRORSTUFF
  Me.HelpContextID = hlpOptionalRevenue
  Opt1 = True
  Opt2 = True
  Opt3 = True
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  If QPTrim$(TaxMasterRec.OptRev1) = "" Then
    Opt1 = False
  End If
  If QPTrim$(TaxMasterRec.OptRev2) = "" Then
    Opt2 = False
  End If
  If QPTrim$(TaxMasterRec.OptRev3) = "" Then
    Opt3 = False
  End If
  
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Call TaxMsg(800, "There are no optional revenues saved at this time. Optional revenues are set up on the Tax System Setup screen.")
    Close
    Exit Sub
  End If
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
      If TblRec.Deleted = True Then GoTo SkipIt
      If QPTrim$(TaxMasterRec.OptRev1) = QPTrim$(TblRec.Desc) Then
        Opt1 = False
      ElseIf QPTrim$(TaxMasterRec.OptRev2) = QPTrim$(TblRec.Desc) Then
        Opt2 = False
      ElseIf QPTrim$(TaxMasterRec.OptRev3) = QPTrim$(TblRec.Desc) Then
        Opt3 = False
      End If
SkipIt:
  Next x
  Close
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Call TaxMsg(800, "All optional revenues have rate tables already set up. Please use the edit screen to make any changes.")
    Close
    Exit Sub
  End If
  
  
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\addtbl.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  frmTaxRateTables.Show
  DoEvents
  Unload Me
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxRateMenu", "cmdAddNew_Click", Erl)
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
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub cmdDelete_Click()
  frmTaxDeleteRateCode.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdEdit_Click()
  If Not Exist(TaxRateTableFile) Then
    Call TaxMsg(900, "There are no rate tables on file. Load attempt aborted.")
    Exit Sub
  End If
  
  frmTaxRateListPop.Show
  DoEvents
End Sub

Private Sub cmdExit_Click()
  KillFile "C:\CPWork\addtbl.dat"
  frmTaxBillSetUpMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrint_Click()
  frmTaxReportOpt.Show vbModal
  If frmTaxReportOpt.fptxtPrintType.Text = "Graphical" Then
    Unload frmTaxReportOpt
    Call PrintGraphics
  ElseIf frmTaxReportOpt.fptxtPrintType.Text = "Text" Then
    Call TaxMsg(900, "Pitch 10 is recommended for this report.")
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxRateMenu.")
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
  Dim TaxSURec As TaxMasterType
  Dim TMHandle As Integer
  Dim dlm$, x As Integer, y As Integer
  Dim Town$
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim RateNum As Integer
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxSURec
  Close TMHandle
  
  Town = QPTrim$(TaxSURec.Name)
  dlm = "~"
  RptFile$ = "TAXRPTS\TXOPRATE.RPT"     'Report File Name
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  Do
    RateNum = RateNum + 1
    For x = 1 To NumOfTRRecs
      Get TRHandle, x, TblRec
      If TblRec.Deleted = True Then GoTo SkipIt
      If TblRec.OptRevNum <> RateNum Then GoTo SkipIt
      '                   0                  1                      2
      Print #RptHandle, Town; dlm; QPTrim$(TblRec.Desc); dlm; TblRec.FlatAmt; dlm;
      '                         3                   4
      Print #RptHandle, TblRec.OptRevNum; dlm; TblRec.Type; dlm;
      '
      For y = 1 To 10
        '                    5 - 14                  6 - 15
        Print #RptHandle, TblRec.FromAmt(y); dlm; TblRec.ToAmt(y); dlm;
        If y < 10 Then
          '                    7 - 16                  8 - 17
          Print #RptHandle, TblRec.TaxFAmt(y); dlm; TblRec.TaxPAmt(y); dlm;
        Else
          Print #RptHandle, TblRec.TaxFAmt(y); dlm; TblRec.TaxPAmt(y)
        End If
      Next y
SkipIt:
    Next x
    If RateNum = 3 Then Exit Do
  Loop
  
  Close RptHandle
  Close TRHandle

  arTaxOptRateRpt.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxRateMenu", "PrintGraphics", Erl)
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
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub PrintText()
  Dim TaxSURec As TaxMasterType
  Dim TMHandle As Integer
  Dim dlm$, x As Integer, y As Integer
  Dim Town$
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim RateNum As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$, Page As Integer
  Dim Line1$, Method$
  Dim Use$
  Dim UseP$
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxSURec
  Close TMHandle
  
  Town = QPTrim$(TaxSURec.Name)
  Line1$ = String(80, "-")
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  Use$ = "$###,###,##0.00"
  UseP$ = "##0.0000"
  
  RptFile$ = "TAXRPTS\TXOPRATE.PRN"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  ReDim WhichOpt(1 To NumOfTRRecs) As Integer
  ReDim WhichType(1 To NumOfTRRecs) As String * 1
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
    If TblRec.Deleted = False Then
      WhichOpt(x) = TblRec.OptRevNum
      WhichType(x) = TblRec.Type
    End If
  Next x
  
  Do
    RateNum = RateNum + 1
    Print #RptHandle, "Optional Revenue #" + CStr(RateNum)
    LineCnt = LineCnt + 1
    For x = 1 To NumOfTRRecs
      If WhichOpt(x) = RateNum And WhichType(x) = "F" And LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        Print #RptHandle, "Optional Revenue #" + CStr(RateNum)
        Print #RptHandle, Line1
        LineCnt = LineCnt + 2
        Exit For
      ElseIf WhichOpt(x) = RateNum And WhichType(x) <> "F" And LineCnt >= MaxLines - 11 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        Print #RptHandle, "Optional Revenue #" + CStr(RateNum)
        Print #RptHandle, Line1
        LineCnt = LineCnt + 2
        Exit For
      End If
    Next x
    
    For x = 1 To NumOfTRRecs
      Get TRHandle, x, TblRec
      If TblRec.Deleted = True Then GoTo MoveOn
      If TblRec.OptRevNum = RateNum Then
        If TblRec.Type = "F" Then
          Method = "Flat Rate"
        ElseIf TblRec.Type = "S" Then
          Method = "Step Flat"
        ElseIf TblRec.Type = "P" Then
          Method = "Step Pct"
        Else
          Method = ""
        End If
        Print #RptHandle, Line1
        If Method = "Flat Rate" Then
          Print #RptHandle, Tab(3); QPTrim$(TblRec.Desc); Tab(35); Method; Tab(50); Using("$##,##0.00", TblRec.FlatAmt)
          Print #RptHandle, Line1
          LineCnt = LineCnt + 2
        Else
          Print #RptHandle, Tab(3); QPTrim$(TblRec.Desc); Tab(35); Method
          Print #RptHandle, Line1
          Print #RptHandle, Tab(17); "From Amt"; Tab(39); "To Amt"; Tab(52); "Flat Amt"; Tab(68); "Pct Amt"
          LineCnt = LineCnt + 3
          For y = 1 To 10
            Print #RptHandle, Tab(10); Using$(Use, TblRec.FromAmt(y)); Tab(30); Using$(Use, TblRec.ToAmt(y)); Tab(45); Using(Use, TblRec.TaxFAmt(y)); Tab(66); Using(UseP, TblRec.TaxPAmt(y)) + " %"
            LineCnt = LineCnt + 1
          Next y
        End If
        Print #RptHandle,
        LineCnt = LineCnt + 1
        If x < NumOfTRRecs Then
          If WhichType(x + 1) = "F" And LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, "Optional Revenue #" + CStr(RateNum)
            LineCnt = LineCnt + 1
          ElseIf WhichType(x + 1) <> "F" And LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, "Optional Revenue #" + CStr(RateNum)
            LineCnt = LineCnt + 1
          End If
        End If
      End If
MoveOn:
    Next x
    If RateNum = 3 Then Exit Do
  Loop
  
  Print #RptHandle, FF$
  Close

  ViewPrint RptFile$, "Optional Revenue Deleted Rate Instances", True
  
  KillFile RptFile$
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Optional Revenue Rates Report"
  Print #RptHandle, Town; Tab(65); "Page # " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, Line1
  LineCnt = 4
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxRateMenu", "PrintText", Erl)
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
    ClearInUse PWcnt
    Terminate

End Sub
