VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmDraftCustList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draft Customer Listing"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12216
   Icon            =   "frmDraftCustList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5424
      TabIndex        =   0
      Top             =   3576
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
      ColDesigner     =   "frmDraftCustList.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5424
      TabIndex        =   1
      Top             =   4128
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
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
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
      ColDesigner     =   "frmDraftCustList.frx":0BED
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F10 &Ok"
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
      Left            =   7776
      TabIndex        =   3
      Top             =   7440
      Width           =   1332
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
      Left            =   9456
      TabIndex        =   2
      Top             =   7440
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
            TextSave        =   "9:43 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "5/6/2005"
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   1032
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Draft Customer Listing"
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
      Left            =   3624
      TabIndex        =   7
      Top             =   1272
      Width           =   5004
   End
   Begin VB.Label Label4 
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
      Left            =   2856
      TabIndex        =   6
      Top             =   4152
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1524
      Left            =   2424
      Top             =   3240
      Width           =   7284
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
      Height          =   372
      Index           =   7
      Left            =   3456
      TabIndex        =   5
      Top             =   3624
      Width           =   1716
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   912
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmDraftCustList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Dim Grpt As Boolean, CycleCnt As Integer
Dim Cycle(1 To 16) As Integer

Private Sub cmdExit_Click()
  frmUBDraftMenu.Show
  Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via DrftCustList by " + PWUser$
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
        cmdOk.SetFocus
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
      cmdOk.SetFocus
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
Private Sub cmdOk_Click()
  Dim Grpt As Boolean
    
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
     'do graphic report
      Grpt = True
    ElseIf fpcboRptType.ListIndex = 1 Then
      Grpt = False
    End If
    UBDraftListing Grpt
    ActivateControls Me, True
 
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
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Grpt = False
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  fpcboPrintOrder.AddItem "Bank Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.ListIndex = 0
  CycleCnt = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub


Private Sub UBDraftListing(Grpt As Boolean)
  Dim Dash80 As String, UBSetupLen As Integer, IndexName As String
  Dim OKFlag As Boolean, UBCustRecLen As Integer, UBCust As Integer
  Dim NumOfRecs As Long, UBRpt As Integer, cnt As Long
  Dim CustCycle As Integer, CustOk As Boolean, CCnt As Integer
  Dim CstCnt As Long, llow As Long, hhigh As Long, BankCnt As Integer
  Dim PrevBank As String, GTotal As Double, TabOffSet As Integer
  Dim ReportFile As String, bnameorder As Boolean, GATot As Double
  Dim TotalDraftCustomers As Long, Temp1 As String, Temp2 As String
  Dim UsingAcct As Boolean, UsingName As Boolean, num As Long
  Dim IdxRecLen As Integer, Handle As Integer, lcnt As Long
  Dim AcctRecord As Long, RptText As String, ToPrint As String
  Dim Dosome As Integer, ReportSum As String, SumRpt As Integer
  TotalDraftCustomers = 0
  Dash80$ = String$(80, "-")
  Temp1$ = Space$(10)
  Temp2$ = Space$(12)
  MaxLines = 60
  FF$ = Chr$(12)
  ToPrint$ = ""
  'load setup file
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  ReDim BankTotals(1 To 1) As BankTotalsType

  MaxLines = 50
  FrmShowPctComp.Label1 = "Creating Draft Customer Listing"
  FrmShowPctComp.Show , Me
  UsingAcct = False
  UsingName = False
  PageNo = 0
  If fpcboPrintOrder.ListIndex = 0 Then
    bnameorder = True
  End If
  If fpcboPrintOrder.ListIndex = 1 Then
    RptText$ = "(Account Number Order)"
    UsingAcct = True
    IndexName$ = ""
  ElseIf fpcboPrintOrder.ListIndex = 2 Then
    UsingName = True
    RptText$ = "(Customer Name Order)"
    IndexName$ = NameIndexFile
  End If

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs& = LOF(UBCust) \ UBCustRecLen
  If UsingName Then
    UBLog "Loading index file: " + IndexName$
    IdxRecLen = 4
    NumOfRecs = FileSize(IndexName$) \ 4
    ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For lcnt& = 1 To NumOfRecs
      Get #Handle, lcnt&, IndexArray(lcnt&)
    Next
    Close Handle
  End If
  ReDim BDCust(1 To 16) As BDRptType

  UBRpt = FreeFile
  ReportFile$ = UBPath$ + "UBDFTCUS.RPT"
  Open ReportFile$ For Output As UBRpt

  If Not Grpt Then GoSub PrintPreHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitPreReport
    End If

    If UsingName Then
      num& = IndexArray(cnt&).RecNum
    Else
      num& = cnt&
    End If
    Get UBCust, num&, UBCustRec(1)
    '  Process Customer Here
      If UBCustRec(1).Status = "A" Or UBCustRec(1).Status = "B" Then
        If (UBCustRec(1).USEDRAFT = "Y") Then  'And (Len(QPTrim$(UBCustRec(1).BankName)) > 0)
            CstCnt = CstCnt + 1
            ReDim Preserve BDCust(1 To CstCnt) As BDRptType
            BDCust(CstCnt).BankName = QPTrim$(UBCustRec(1).BankName)
            BDCust(CstCnt).CustRec = num&
            BDCust(CstCnt).TransRec = num&
          
        End If
      End If
    
'    If AskAbandonPrint% Then
'      ABExit = True
'      GoTo NON2PrintExit:
'    End If
'    ShowPctCompL cnt&, NumOfRecs&
DFTskipem:
  Next
    
  If CstCnt <= 0 Then
    If Not Grpt Then
      Print #UBRpt, "No Info found to Print"
      Print #UBRpt, Dash80$
    Else
      MsgBox "No Info to Print", vbOKOnly, "No Drafts"
    End If
    GoTo NON2PrintExit
  End If
  llow = LBound(BDCust)
  hhigh = UBound(BDCust)
  If bnameorder Then
    BDSort BDCust(), llow, hhigh
  End If
  'SortT BDCust(1), CstCnt, 0, 20, 0, 14
  BankCnt = 1
  Get UBCust, BDCust(1).CustRec, UBCustRec(1)

  PrevBank$ = QPTrim$(BDCust(1).BankName)
  BankTotals(BankCnt).BankName = QPTrim$(BDCust(1).BankName)

  For cnt = 1 To CstCnt
    Get UBCust, BDCust(cnt).CustRec, UBCustRec(1)
    If bnameorder Then
      Dosome = 1
      If PrevBank$ <> QPTrim$(BDCust(cnt).BankName) Then
        BankCnt = BankCnt + 1
        ReDim Preserve BankTotals(1 To BankCnt) As BankTotalsType
        BankTotals(BankCnt).BankName = QPTrim$(BDCust(cnt).BankName)
        PrevBank$ = QPTrim$(BDCust(cnt).BankName)
      End If
      BankTotals(BankCnt).Amount = Round#(BankTotals(BankCnt).Amount + 1)
    Else
      Dosome = 2
    End If
     If Not Grpt Then
        If LineCnt >= MaxLines Then
          Print #UBRpt, FF$;
          GoSub PrintPreHeader
        End If
      End If
      If UBCustRec(1).USEDRAFT = "Y" Then
        TotalDraftCustomers = TotalDraftCustomers + 1
        If Not Grpt Then
          Print #UBRpt, Using("#####", BDCust(cnt).CustRec);
          Print #UBRpt, Tab(10); UBCustRec(1).CustName;
          Print #UBRpt, Tab(48); Left$(UBCustRec(1).TRANSIT, 9);
          Print #UBRpt, Tab(59); Left$(UBCustRec(1).BankAcct, 20)
          If UBCustRec(1).PreNoteFlag = 0 Then
            Print #UBRpt, Tab(2); "N";
          Else
            Print #UBRpt, Tab(2); "Y";
          End If
          Print #UBRpt, Tab(15); QPTrim$(UBCustRec(1).BankName);
          Print #UBRpt, Tab(50); Left$(UBCustRec(1).BANKLOC, 20); "  "; UBCustRec(1).BILLCYCL
          Print #UBRpt, String$(79, "-")
          LineCnt = LineCnt + 3
        Else
          ToPrint$ = Using("#####", BDCust(cnt).CustRec)
          ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).CustName)
          ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).TRANSIT, 9)
          ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).BankAcct, 20)
          If UBCustRec(1).PreNoteFlag = 0 Then
            ToPrint$ = ToPrint$ + "~" + "N"
          Else
            ToPrint$ = ToPrint$ + "~" + "Y"
          End If
          ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).BankName)
          ToPrint$ = ToPrint$ + Left$(UBCustRec(1).BANKLOC, 20)
          ToPrint$ = ToPrint$ + "~" + Str(UBCustRec(1).BILLCYCL)
          Print #UBRpt, ToPrint$
          ToPrint$ = ""
        End If
      End If
    
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If
'    ShowPctComp cnt, NumOfRecs
  Next cnt
  If Not Grpt Then
    Print #UBRpt, "Total Draft Customers on File: "; Using("####,#", TotalDraftCustomers)
    Print #UBRpt, FF$;
  End If
  If bnameorder Then
    If Not Grpt Then
      For cnt = 1 To BankCnt
        Print #UBRpt, BankTotals(cnt).BankName; Tab(30); Using$("#####.##", Str$(BankTotals(cnt).Amount))
        GTotal# = Round#(GTotal# + BankTotals(cnt).Amount)
      Next
    Else
      ReportSum$ = UBPath$ + "UBSum.RPT"
      SumRpt = FreeFile
      Open ReportSum$ For Output As SumRpt
      For cnt = 1 To BankCnt
        ToPrint$ = BankTotals(cnt).BankName + "~" + Using$("#####.##", Str$(BankTotals(cnt).Amount))
        Print #SumRpt, ToPrint$
        ToPrint$ = ""
        GTotal# = Round#(GTotal# + BankTotals(cnt).Amount)
      Next
    End If
  End If

  
  Close

  If Grpt Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmDraftCustList
    ARptDraftCustList.Title = "Draft Customer Listing " + RptText$
    ARptDraftCustList.txtDate = Now
    ARptDraftCustList.txtTown = TOWNNAME$
    ARptDraftCustList.totCust = Using("####,#", TotalDraftCustomers)
    ARptDraftCustList.GetName ReportFile$, ReportSum$, Dosome
    ARptDraftCustList.startrpt
  Else

    ViewPrint ReportFile$, "Draft Customer Report " + RptText$ ', "UBDFTCUS.RPT", LPTPort
  End If
  
  Close

  '  IF NOT ExitFlag THEN GOTO ReStart
  Exit Sub

PrintPreHeader:
  PageNo = PageNo + 1
  Print #UBRpt, " "
  Print #UBRpt, " "
  Print #UBRpt, "Utility Draft Customer Listing"; Tab(70); "Page: "; PageNo
  Print #UBRpt, "Date: "; Date$
  Print #UBRpt,
  Print #UBRpt, "Acct #"; Tab(10); "Customer Name"; Tab(48); "Transit #"; Tab(60); "Bank Acct #"
  Print #UBRpt, "Prenoted??"; Tab(15); "Bank Name & Location"
  Print #UBRpt, Dash80$
  LineCnt = 8

Return
NON2PrintExit:
 Close
ExitPreReport:

End Sub

