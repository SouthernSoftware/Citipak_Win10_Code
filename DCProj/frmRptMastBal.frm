VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmRptMastBal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Balance Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptMastBal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5490
      TabIndex        =   1
      Top             =   3528
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptMastBal.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5490
      TabIndex        =   0
      Top             =   2976
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptMastBal.frx":0BF8
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10080
      TabIndex        =   3
      Top             =   7248
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F10 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      TabIndex        =   2
      Top             =   7248
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "4:01 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "8/30/2005"
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2340
      Left            =   2454
      Top             =   2184
      Width           =   7284
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2922
      TabIndex        =   7
      Top             =   3552
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Order:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   7
      Left            =   3498
      TabIndex        =   6
      Top             =   3048
      Width           =   1716
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   312
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customer Balance Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   5
      Top             =   552
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   192
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmRptMastBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmDCReportsMenu.Show
  Unload frmRptMastBal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via RptMastBal by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub



Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPrintOrder.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub cmdPrint_Click()
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
      CustomerBalance
    ElseIf fpcboRptType.ListIndex = 1 Then
      CustomerBalance
      ActivateControls Me, True
    Else
      ActivateControls Me, True
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub MasterBalanceListing()
  Dim RCnt As Integer, UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim MaxRevenue As Integer, TRevName As String, AndPos As String
  Dim UsingBook As Boolean, RStatus As String, UsingName As Boolean
  Dim PageNo As Integer, UseStatus As Boolean, AcctNo As Long
  Dim Dash80 As String, IndexName As String, RealBalance As Double
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCust As Integer, UBRpt As Integer, UBSetUp1 As Integer
  Dim RevChk As Integer, CStatus As String, Book As String
  Dim SEQNUMB As String, BalType As String, ChkBalance As Double
  Dim MinBal As Double, RevSource As Integer, TCurrBalance As Double
  Dim CustCnt As Long, TPrevBalance As Double, Detail As String
  Dim GTotal As Double, CoFlag As Boolean, Stat As String, UsingAcct As Boolean
  Dim POrder As String, Bal As String, DLineCnt As Integer, bk As Integer
  Dim TCnt As Integer, First As Integer, Last As Integer, Rev As String
  Dim TabStop As Integer, Det As Boolean, order As String
  Dim ReportFile As String
  RCnt = RCnt + 1
  UsingAcct = False
  UseStatus = False
  UsingName = False
  UsingBook = False
  ReDim fmt$(1 To 3)
  fmt$(1) = "####,#.##"
  fmt$(2) = "#####"
  fmt$(3) = "######,#.##"
  'Main Body Start
  FrmShowPctComp.Label1 = "Creating Master Balance Listing"
  FrmShowPctComp.Show , Me

End Sub
'
'   'ViewPrint "UBBALIST.RPT", "Balance Listing Report."
'  If CustCnt > 0 Then
'  Load frmLoadingRpt
'  frmLoadingRpt.setwherefrom frmRptMastBal
'  ARptMastBalList.txtDate = Now
'  ARptMastBalList.txtTown = TOWNNAME$
'  ARptMastBalList.Title = "Master Customer Balance Report"
'  ARptMastBalList.txtRptParm1.Caption = ToPrintH1$
'  ARptMastBalList.txtRptParm2.Caption = ToPrintH2$
'  ARptMastBalList.txtTotCust = CustCnt
'  ARptMastBalList.txtTotCur.DataValue = TCurrBalance#
'  ARptMastBalList.txtTotPast.DataValue = TPrevBalance#
'  ARptMastBalList.txtHead = fpcboRevenues.Text
'  ARptMastBalList.txtTotAcctBal.DataValue = Round#(TCurrBalance# + TPrevBalance#)
'  ARptMastBalList.GetName ReportFile$, Report2$, DetFlag, RevSource
'  ARptMastBalList.startrpt
'  Else
'    MsgBox "No Information to print.", vbOKOnly, "No Information"
'    ActivateControls Me, True
'  End If
Private Sub CustomerBalance()
  Dim DCCustRecLen As Integer, Page As Integer
  Dim UsingName As Boolean, TotalBal As Double, TCnt As Long
  Dim CustomerCnt As Long, UsingAcct As Boolean
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, DCRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfDCRecs As Long, AcctNo As Long
  Dim DCFile As Integer, ToPrint As String, Header As String
 
  Dim Dash80 As String, CustCnt As Long
  Dim cnt As Long, RptHandle As Integer
  Dim ReportFile As String
  ReDim DCCustRec(1) As DCCustRecType
  DCCustRecLen = Len(DCCustRec(1))
  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    IndexName$ = DCPath$ + "DCCUST.IDX"
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case Else
  End Select
  Header$ = "Customer Balance Report"
  ReportFile$ = "DCCusBal.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  Linecnt = 0
  CustCnt = 0
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  GoSub PrintCustBalRptHeader
  If UsingName = True Then
    NumOfDCRecs = FileSize(IndexName$) \ 4
    ReDim IndexArray(1 To NumOfDCRecs) As DCTempIDXRecType
    'FGetAH IndexName$, IndexArray(1), , NumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = 4
    For cnt& = 1 To NumOfDCRecs
      Get #Handle, cnt&, IndexArray(cnt&)
    Next
    Close Handle

  Else
    NumOfDCRecs = FileSize(DCPath$ + "DCCUST.DAT") \ DCCustRecLen
  End If

'  ' Print Main Body
'  OpenDCCustFile NumOfDCRecs, DCFile
'  ReDim DCIdxRec(1) As DCCustIDXRecType
'  IdxFile = FreeFile
'  Open "DCCust.IDX" For Random Access Read Write Shared As #IdxFile Len = Len(DCIdxRec(1))


  For cnt& = 1 To NumOfDCRecs
    If UsingAcct Then
      AcctNo& = cnt
    ElseIf UsingName Then
      AcctNo& = IndexArray(cnt).IDXRECORD
    End If
    Get DCFile, AcctNo&, DCCustRec(1)
    If DCCustRec(1).Deleted <> "Y" Then
      If DCCustRec(1).AcctBal <> 0 Then
        If Linecnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintCustBalRptHeader
        End If
        Print #RptHandle, Using("#####", AcctNo&);
        Print #RptHandle, Tab(10); DCCustRec(1).BILLNAME;
        Print #RptHandle, Tab(50); Using("$##,###.##", DCCustRec(1).AcctBal)
        CustCnt = CustCnt + 1
        TotalBal# = Round#(TotalBal# + DCCustRec(1).AcctBal)
        Linecnt = Linecnt + 1
      End If
    End If
  Next cnt&
  GoSub PrintCustBalRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now


  ViewPrint ReportFile$, Header$
  Kill ReportFile$
  Exit Sub


PrintCustBalRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(21); "Va. Decal System : Detailed Customer Listing"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "Cust #"; Tab(10); "Customer Name"; Tab(50); "Account Balance"
  Print #RptHandle, String$(80, "=")
  Linecnt = 5
  Return

PrintCustBalRptEnding:
  Print #RptHandle, String$(79, "-")
  Print #RptHandle, "Total Customers Printed: "; Using("#####", CustCnt)
  Print #RptHandle, "  Total Balance on File: "; Using("$###,###.##", TotalBal#)
  Print #RptHandle, FF$
  Return

SelectCustBalOutput:
End Sub
