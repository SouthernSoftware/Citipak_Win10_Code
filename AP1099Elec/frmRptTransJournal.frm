VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmRptTransJournal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Transaction Detail Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptTransJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5520
      TabIndex        =   0
      Top             =   3024
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
      ColDesigner     =   "frmRptTransJournal.frx":08CA
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
      Left            =   8088
      TabIndex        =   2
      Top             =   4656
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
      Left            =   6384
      TabIndex        =   1
      Top             =   4656
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
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
            TextSave        =   "3:46 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "2/3/2006"
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
      Height          =   1380
      Left            =   2496
      Top             =   2472
      Width           =   7284
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
      Left            =   3072
      TabIndex        =   5
      Top             =   3048
      Width           =   2388
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3192
      Top             =   816
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Utility Tranaction Detail"
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
      TabIndex        =   4
      Top             =   984
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3192
      Top             =   696
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
Attribute VB_Name = "frmRptTransJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Sub cmdExit_Click()
  frmGetDistMenu.Show
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
        MainLog "Closed via RptTransJournal by " + PWUser$
      End If
    End If
  End If
End Sub

'Private Function ValidDate()
'  Dim TempDate1 As Integer, TempDate2 As Integer
'  If CheckValDate(txtDate1) = False And CheckValDate(txtDate2) = False Then
'    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
'    ValidDate = False
'  Else
'    TempDate1 = DateDiff("d", "12/31/1979", txtDate1)
'    TempDate2 = DateDiff("d", "12/31/1979", txtDate2)
'    If TempDate1 > TempDate2 Then
'      ValidDate = False
'      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
'    Else
'      ValidDate = True
'    End If
'  End If
'End Function

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

'Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    txtDate2.SetFocus
'  End If
'End Sub
'
'Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    fptxtRoute1.SetFocus
'  End If
'End Sub

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
        'fpcboPrintOrder.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub



Private Sub cmdPrint_Click()
 ' If ValidDate = True Then
 '  If ValidRoutes Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 2 Then
      DetailedTransJournal
      DoEvents
      TransSummary
    ElseIf fpcboRptType.ListIndex = 1 Or fpcboRptType.ListIndex = 0 Then
      DetailedTransJournal2
     DoEvents
     'dothesummary
    Else
     ActivateControls Me, True
    End If
 '  End If
 ' End If
End Sub
'Private Sub dothesummary()
'MsgBox "ready to go", vbOKOnly, "ready?"
'TransSummary
'ActivateControls Me, True
'
'End Sub
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
'  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  fpcboRptType.InsertRow = "Graphics - Landscape"
  fpcboRptType.InsertRow = "Graphics - Portrait"
  fpcboRptType.InsertRow = "Text - Condensed Print"
  fpcboRptType.ListIndex = 0
  
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub DetailedTransJournal()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim NumOfTRecs As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long, PageNo As Integer
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean, FF As String
  Dim ReportFile As String, MoFlag As Boolean, TCnt As Long, MaxLines As Integer
  Dim UBSetUpFileNum As Integer, UBSetUpLen As Integer, Linecnt As Integer
  Dim TownName As String
  MaxLines = 55
  PageNo = 0
  FF$ = Chr$(12)
  Dash120$ = String$(121, "-")
  FrmShowPctComp.Label1 = "Creating Utility Transaction Journal"
  FrmShowPctComp.Show , Me
  DoEvents
  ReDim RevTotals(1 To 15) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUpRec(1))
  ReDim UBTransRec(1) As GLUBTempRecType
  UBTransRecLen = Len(UBTransRec(1))
  UBTrans = FreeFile
  Open "GLUBTran.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfTRecs& = LOF(UBTrans) \ UBTransRecLen
   If NumOfTRecs& <= 0 Then
    Unload FrmShowPctComp
    MsgBox "No Transactions", vbOKOnly, "No Trans"
    ActivateControls Me, True
    Exit Sub
  End If
  ReportFile$ = "GLUBtmp.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  LoadUBSetUpFile UBSetUpFileNum, UBSetUpLen
  Get UBSetUpFileNum, 1, UBSetUpRec(1)
  If Len(GLUserName) = 0 Then
    GLUserName = "Undefined"
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    TownName$ = GLUserName
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUpRec(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
  End If
  GoSub DoDetailedRptHeader

  For cnt = 1 To NumOfTRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfTRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitDetailedListing
    End If


    If Linecnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoDetailedRptHeader
    End If
'*************************************
'   Main Body of Printing goes here
    BadCount = 0
      Get UBTrans, cnt, UBTransRec(1)
      GoSub DefineType
      Print #UBRpt, Num2Date$(UBTransRec(1).TransDate); Tab(11); Using("#####", UBTransRec(1).CustAcctNo);
      Print #UBRpt, Tab(20); Left$(UBTransRec(1).CustName, 33);
      Print #UBRpt, Tab(55); TransDesc$;
      Print #UBRpt, Tab(80); UBTransRec(1).OperatorNumber;
      Print #UBRpt, Tab(90); Left$(UBTransRec(1).TransDesc, 20);
      Print #UBRpt, Tab(110); Using("$###,###.##", Amount#)
      Linecnt = Linecnt + 1
      TotalTrans# = Round#(TotalTrans# + Amount#)
      TransCnt& = TransCnt& + 1
      If Detail$ = "Y" Then
        Print #UBRpt, "Revenue Source Breakdown ........................"
        Linecnt = Linecnt + 1
        For RCnt = 1 To MaxRevenue Step 3
          Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
          Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
          Print #UBRpt, Tab(60); RevenueName$(RCnt + 2); Tab(76); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
          Linecnt = Linecnt + 1
        Next RCnt
        Print #UBRpt, Dash120$
        Linecnt = Linecnt + 1
      End If
      For RCnt = 1 To MaxRevenue
        RevTotals(RCnt) = Round#(RevTotals(RCnt) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
      Next
      If Linecnt > MaxLines Then
        Print #UBRpt, FF$
        GoSub DoDetailedRptHeader
      End If
SkipThisOne:
  Next

  GoSub DoDetailedRptFooter
  Print #UBRpt, FF$;

  Close
  ViewPrint ReportFile$, "Detailed Journal Report", True
  
ExitDetailedListing:

  Exit Sub

DoDetailedRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TownName$
  Print #UBRpt, Tab(28); "Detailed Transaction Report"; Tab(90); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Beginning Transaction Date: "; Date1$;
  If Val(Operator$) = 0 Then
    Print #UBRpt, Tab(90); " Operator #: ALL"
  Else
    Print #UBRpt, Tab(90); " Operator #: "; Operator$
  End If
  Print #UBRpt, "   Ending Transaction Date: "; Date2$;
  Print #UBRpt, Tab(90); "Show Detail: "; Detail$
  Print #UBRpt, "          Transaction Type: "; "All Transactions"
  Print #UBRpt, "             Customer Type: ";

  If UseType Then
    Print #UBRpt, CUSTTYPE$
  Else
    Print #UBRpt, "N/A"
  End If

  Print #UBRpt,
  Print #UBRpt, "  Date"; Tab(11); "Acct #"; Tab(20); "Customer Name"; Tab(55); "Description"; Tab(80); "Oper#"; Tab(90); "Trans Type"; Tab(113); "   Amount"
  Print #UBRpt, Dash120$
  Linecnt = 10
  Return

DoDetailedRptFooter:
  Print #UBRpt, Dash120$
  Print #UBRpt, "Transactions: "; TransCnt&; "                                                       Total of Transactions: "; Using("$##,###,###.##", TotalTrans#)
  Print #UBRpt, FF$
  PageNo = PageNo + 1
  Print #UBRpt, TownName$
  Print #UBRpt, Tab(28); "Detailed Transaction Report"; Tab(90); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Beginning Transaction Date: "; Date1$;
  If Val(Operator$) = 0 Then
    Print #UBRpt, Tab(90); " Operator #: ALL"
  Else
    Print #UBRpt, Tab(90); " Operator #: "; Operator$
  End If
  Print #UBRpt, "   Ending Transaction Date: "; Date2$;
  Print #UBRpt, Tab(90); "Show Detail: "; Detail$
  Print #UBRpt, ""
  Print #UBRpt, "Revenue Summary"; Tab(38); "Amount"
  Print #UBRpt, Dash120$
  TotalRevsAmt# = 0
  For RCnt = 1 To MaxRevenue
    TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
    Print #UBRpt, RevenueName$(RCnt), Tab(35); Using("########.##", RevTotals(RCnt))
  Next
  Print #UBRpt,
  Print #UBRpt, "Total Amount"; Tab(35); Using("########.##", TotalRevsAmt#)
  Return
DefineType:
  Select Case UBTransRec(1).TransType
  Case 1, 101
    TransDesc$ = "Util Bill"
    Amount# = UBTransRec(1).Transamt
  Case 2, 102
    TransDesc$ = "Late Charge"
    Amount# = UBTransRec(1).Transamt
  Case 3
    TransDesc$ = "Reconnect"
    Amount# = UBTransRec(1).Transamt
  Case 4, 104
    TransDesc$ = "Reg Payment"
        Amount# = UBTransRec(1).Transamt
  Case 5, 105
    TransDesc$ = "Applied Dep"
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 6
    TransDesc$ = "Penalty Chg"
    Amount# = UBTransRec(1).Transamt
  Case 7, 107
    TransDesc$ = "Dep. Payment"
    Amount# = UBTransRec(1).Transamt
  Case 8
    TransDesc$ = "Draft Paymt"
    Amount# = UBTransRec(1).Transamt * -1
  Case 9, 109
    TransDesc$ = "Refunded Dep"
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 10, 110
    TransDesc$ = "Beg Balance"
    Amount# = UBTransRec(1).Transamt
  Case 11, 111
    TransDesc$ = "Up Adj"
    Amount# = UBTransRec(1).Transamt
  Case 12, 112
    TransDesc$ = "Down Adj"
    Amount# = UBTransRec(1).Transamt
  Case 33
    TransDesc$ = "Pay Adj"
    Amount# = UBTransRec(1).Transamt
  Case 37
    TransDesc$ = "Dep Cred Rem"
    Amount# = UBTransRec(1).Transamt
  Case 39
    TransDesc$ = "Dep Pay Void"
    Amount# = UBTransRec(1).Transamt
  Case 99
    TransDesc$ = "Misc Payment"
    Amount# = UBTransRec(1).Transamt
  Case Else
    TransDesc$ = "UNKNOWN"
    Amount# = UBTransRec(1).Transamt
  End Select
  Return
End Sub

Private Sub DetailedTransJournal2()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim NumOfTRecs As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long, TCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean
  Dim ToPrint As String, PrnH1 As String, PrnH2 As String, PrnH3 As String
  Dim SumRpt As Integer, ToPrintD As String, DetFlag As Boolean
  Dim ReportFile As String, ReportSum As String, MoFlag As Boolean
  Dim UBSetUpFileNum As Integer, UBSetUpLen As Integer, Linecnt As Integer
  Dim TownName As String, numtochk As Long, bdate As String
  TownName$ = GLUserName
  FrmShowPctComp.Label1 = "Creating Utility Transaction Journal"
  FrmShowPctComp.Show , Me
  DoEvents
  ReDim RevTotals(1 To 15) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBSetUpRec(1) As UBSetupRecType

  UBSetupreclen = Len(UBSetUpRec(1))
  ReDim UBTransRec(1) As GLUBTempRecType
  UBTransRecLen = Len(UBTransRec(1))
  UBTrans = FreeFile
  Open "GLUBTran.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfTRecs& = LOF(UBTrans) \ UBTransRecLen
  If NumOfTRecs& <= 0 Then
    Unload FrmShowPctComp
    MsgBox "No Transactions", vbOKOnly, "No Trans"
    ActivateControls Me, True
    Exit Sub
  End If
  ReportFile$ = "GLUBtmp.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  LoadUBSetUpFile UBSetUpFileNum, UBSetUpLen
  Get UBSetUpFileNum, 1, UBSetUpRec(1)
  ReportSum$ = "UBDJSUM.RPT"
  SumRpt = FreeFile
  Open ReportSum$ For Output As SumRpt

  If Len(GLUserName) = 0 Then
    GLUserName = "Undefined"
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUpRec(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
  End If
'*************************************
'   Main Body of Printing goes here
    BadCount = 0
    For TCnt = 1 To NumOfTRecs&
      FrmShowPctComp.ShowPctComp TCnt, NumOfTRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        ActivateControls Me, True
        GoTo ExitDetailedListing
      End If

      Get UBTrans, TCnt, UBTransRec(1)
      GoSub DefineType
      bdate$ = UBTransRec(1).Grabbatch
      ToPrint$ = Str$(TCnt) + "~" + Num2Date$(UBTransRec(1).TransDate) + "~" + Using("#####", UBTransRec(1).CustAcctNo)
      ToPrint$ = ToPrint$ + "~ " + Left$(UBTransRec(1).CustName, 33)
      ToPrint$ = ToPrint$ + "~" + TransDesc$
      ToPrint$ = ToPrint$ + "~" + Str$(UBTransRec(1).OperatorNumber)
      ToPrint$ = ToPrint$ + "~" + Left$(UBTransRec(1).TransDesc, 20)
      ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", Amount#)
      TotalTrans# = Round#(TotalTrans# + Amount#)
      TransCnt& = TransCnt& + 1
      If Detail$ = "Y" Then
        For RCnt = 1 To 15
          If UBTransRec(1).RevAmt(RCnt) <> 0 Then
            ToPrintD$ = ToPrintD$ + RevenueName$(RCnt) + "~" + Str$(Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt))) + "~"
          Else
            If Len(RevenueName$(RCnt)) > 0 Then
              ToPrintD$ = ToPrintD$ + RevenueName$(RCnt) + "~0.00~"
            Else
              ToPrintD$ = ToPrintD$ + " ~  ~"
            End If
          End If
        Next RCnt
      Else
        ToPrintD$ = "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~"
      End If
      For RCnt = 1 To MaxRevenue
        RevTotals(RCnt) = Round#(RevTotals(RCnt) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
      Next
      Print #UBRpt, ToPrint$ + "~" + ToPrintD$
      ToPrint$ = ""
      ToPrintD$ = ""
    Next
SkipThisOne:
  GoSub DoDetailedRptHeader
  GoSub DoDetailedRptFooter

  Close
  If TransCnt& > 0 Then

    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptTransJournal
    If fpcboRptType.ListIndex = 0 Then
      ARptTransJournal.txtDate = Now
      ARptTransJournal.txtTown = GLUserName
      ARptTransJournal.Title = "Utility Transaction Interface Journal Report"
      ARptTransJournal.txtRptParm1.Caption = PrnH1$
     ' ARptTransJournal.txtRptParm2.Caption = PrnH2$
     ' ARptTransJournal.txtPrnOrd = "In " + fpcboPrintOrder.Text
      ARptTransJournal.totCust = TransCnt&
      ARptTransJournal.GetName ReportFile$, ReportSum$, DetFlag, MaxRevenue
      ARptTransJournal.startrpt
    ElseIf fpcboRptType.ListIndex = 1 Then
      ARptTransJPortrait.txtDate = Now
      ARptTransJPortrait.txtTown = GLUserName
      ARptTransJPortrait.Title = "Utility Transaction Interface Journal Report"
      ARptTransJPortrait.txtRptParm1.Caption = PrnH1$
    '  ARptTransJPortrait.txtRptParm2.Caption = PrnH2$
     ' ARptTransJPortrait.txtPrnOrd = "In " + fpcboPrintOrder.Text
      ARptTransJPortrait.totCust = TransCnt&
      ARptTransJPortrait.GetName ReportFile$, ReportSum$, DetFlag, MaxRevenue
       
      ARptTransJPortrait.startrpt
    End If
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If
  
ExitDetailedListing:
  
  Exit Sub

DoDetailedRptHeader:
  PrnH1$ = "Detail list of Utility Transactions batch: " + bdate$
  PrnH2$ = ""
  Return

DoDetailedRptFooter:
  TotalRevsAmt# = 0
  For RCnt = 1 To MaxRevenue
    TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
    Print #SumRpt, RevenueName$(RCnt) + "~" + Using("########.##", RevTotals(RCnt))
  Next
  Return
DefineType:
  Select Case UBTransRec(1).TransType
  Case 1, 101
    TransDesc$ = "Util Bill"
    Amount# = UBTransRec(1).Transamt
  Case 2, 102
    TransDesc$ = "Late Charge"
    Amount# = UBTransRec(1).Transamt
  Case 3
    TransDesc$ = "Reconnect"
    Amount# = UBTransRec(1).Transamt
  Case 4, 104
    TransDesc$ = "Reg Payment"
        Amount# = UBTransRec(1).Transamt
  Case 5, 105
    TransDesc$ = "Applied Dep"
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 6
    TransDesc$ = "Penalty Chg"
    Amount# = UBTransRec(1).Transamt
  Case 7, 107
    TransDesc$ = "Dep. Payment"
    Amount# = UBTransRec(1).Transamt
  Case 8
    TransDesc$ = "Draft Paymt"
    Amount# = UBTransRec(1).Transamt * -1
  Case 9, 109
    TransDesc$ = "Refunded Dep"
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 10, 110
    TransDesc$ = "Beg Balance"
    Amount# = UBTransRec(1).Transamt
  Case 11, 111
    TransDesc$ = "Up Adj"
    Amount# = UBTransRec(1).Transamt
  Case 12, 112
    TransDesc$ = "Down Adj"
    Amount# = UBTransRec(1).Transamt
  Case 33
    TransDesc$ = "Pay Adj"
    Amount# = UBTransRec(1).Transamt
  Case 37
    TransDesc$ = "Dep Cred Rem"
    Amount# = UBTransRec(1).Transamt
  Case 39
    TransDesc$ = "Dep Pay Void"
    Amount# = UBTransRec(1).Transamt
  Case 99
    TransDesc$ = "Misc Payment"
    Amount# = UBTransRec(1).Transamt
  Case Else
    TransDesc$ = "UNKNOWN"
    Amount# = UBTransRec(1).Transamt
  End Select
  Return

End Sub
Public Sub TransSummary()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean
  Dim ToPrint As String, PrnH1 As String, PrnH2 As String, PrnH3 As String
  Dim SumRpt As Integer, ToPrintD As String, cntp As Integer, bdate As String
  Dim ReportFile As String, ReportSum As String, cnttype As Integer
  Dim fmt As String, TotBills As Double, TotPen As Double, TotUA As Double
  Dim TotDA As Double, TotPay As Double, TotDft As Double, TotOvA As Double
  Dim TotDP As Double, TotAP As Double, TotRD As Double, TotCr As Double
  Dim MaxLines As Integer, PageNo As Integer, FF As String, TownName As String
  Dim UBSetUpFileNum As Integer, UBSetUpLen As Integer, Linecnt As Integer

  'FrmShowPctComp.Label1 = "Creating Utility Transaction Summary"
 ' FrmShowPctComp.Show , Me
  MaxLines = 55
  PageNo = 0
  Dash120$ = String$(130, "-")
  fmt$ = "#######.##"
  ReDim RevTotals(1 To 15, 1 To 12) As Double
  ReDim RevenueName(1 To 15) As String
  TownName$ = GLUserName
  ReDim ubsetup(1) As UBSetupRecType
  UBSetupreclen = Len(ubsetup(1))

  ReDim UBTransRec(1) As GLUBTempRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUpRec(1) As UBSetupRecType

  UBSetupreclen = Len(UBSetUpRec(1))
  MaxRevenue = 15
  UBTrans = FreeFile
  Open "GLUBTran.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfRecs = FileSize("GLUBtran.DAT") \ UBTransRecLen

  ReportFile$ = "UBDJLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  ReportSum$ = "UBDJSUM.RPT"
  SumRpt = FreeFile
  Open ReportSum$ For Output As SumRpt
  LoadUBSetUpFile UBSetUpFileNum, UBSetUpLen
  Get UBSetUpFileNum, 1, UBSetUpRec(1)
'''  AcctMeth$ = QPTrim$(UBSetUpRec(1).MethAcct)
'''  If (Len(AcctMeth$) = 0) Then
'''    Unload FrmShowPctComp
'''    MsgBox "The Utility Account Method Is Not Setup", vbOKOnly, "Invalid Setup Info"
'''    GoTo SendExitUB
'''  End If
'''
'''  Select Case AcctMeth$
'''  Case "C"
'''    InterfaceMethod = 1
'''  Case "A"
'''    InterfaceMethod = 2
'''  Case Else
'''    Unload FrmShowPctComp
'''    GoTo SendExitUB
'''  End Select
  If Len(TownName$) = 0 Then
    TownName$ = "Undefined"
    ' Set Revenue Names to Nothing
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
   For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUpRec(1).Revenues(RCnt).RevName)
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = MaxRevenue - 1
        'Exit For
      Else
'''      ReDim Preserve BADebit(1 To RCnt) As String
'''      ReDim Preserve BACredit(1 To RCnt) As String
'''      ReDim Preserve PADebit(1 To RCnt) As String
'''      ReDim Preserve PACredit(1 To RCnt) As String
'''      ReDim Preserve DADebit(1 To RCnt) As String
'''      ReDim Preserve DACredit(1 To RCnt) As String
'''
'''      BADebit(RCnt) = UBSetUpRec(1).BillAcct(RCnt).DebitAcct
'''      BACredit(RCnt) = UBSetUpRec(1).BillAcct(RCnt).CreditAcct
'''      PADebit(RCnt) = UBSetUpRec(1).PayAcct(RCnt).DebitAcct
'''      PACredit(RCnt) = UBSetUpRec(1).PayAcct(RCnt).CreditAcct
'''      If UBSetUpRec(1).Revenues(RCnt).UseDep = "Y" Then
'''        DADebit(RCnt) = UBSetUpRec(1).DepAcct(RCnt).DebitAcct
'''        DACredit(RCnt) = UBSetUpRec(1).DepAcct(RCnt).CreditAcct
'''      End If
      End If
   Next
  End If
  If NumOfRecs& = 0 Then
    FrmShowPctComp.ShowPctComp 100, 100
  Else
  For Trans& = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitDetailedListing
    End If
      Get UBTrans, Trans&, UBTransRec(1)
      bdate$ = UBTransRec(1).Grabbatch
            GoSub DefineType
            TransCnt& = TransCnt& + 1
    Next
SkipThisOne:
  End If
  If fpcboRptType.ListIndex <> 2 Then
    For cntp = 1 To MaxRevenue
      For cnttype = 1 To 12
        ToPrintD$ = ToPrintD$ + Str(RevTotals(cntp, cnttype)) + "~"
      Next
      Print #UBRpt, RevenueName$(cntp) + "~" + ToPrintD$
      ToPrintD$ = ""
    Next
  Else
    GoSub DoDetailedRptHeader

    For cntp = 1 To MaxRevenue
      Print #UBRpt, Tab(1); RevenueName$(cntp)
      Print #UBRpt, Tab(3); Using(fmt$, (RevTotals(cntp, 1)));
      Print #UBRpt, Tab(14); Using(fmt$, (RevTotals(cntp, 4)));
      Print #UBRpt, Tab(25); Using(fmt$, (RevTotals(cntp, 8)));
      Print #UBRpt, Tab(36); Using(fmt$, (RevTotals(cntp, 9)));
      Print #UBRpt, Tab(47); Using(fmt$, (RevTotals(cntp, 2)));
      Print #UBRpt, Tab(58); Using(fmt$, (RevTotals(cntp, 6)));
      Print #UBRpt, Tab(69); Using(fmt$, (RevTotals(cntp, 10)));
      Print #UBRpt, Tab(80); Using(fmt$, (RevTotals(cntp, 5)));
      Print #UBRpt, Tab(91); Using(fmt$, (RevTotals(cntp, 3)));
      Print #UBRpt, Tab(102); Using(fmt$, (RevTotals(cntp, 7)));
      Print #UBRpt, Tab(113); Using(fmt$, (RevTotals(cntp, 11)))
      'Print #UBRpt, Tab(124); Using(fmt$, (RevTotals(cntp, 1)));
    Next

    GoSub DoDetailedRptFooter
    Print #UBRpt, FF$;

  End If

 Close

 ' Erase Idxbuff, UBCustRec

  'END

  'If Not AbortFlag Then
  '  PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", LptPort, RetCode, EntryPoint
 ' End If
 ' ViewPrint "UBDJLIST.RPT", "Detailed Journal Report", True
  'KillFile "UBDJLIST.RPT"


  If TransCnt& > 0 Then


    If fpcboRptType.ListIndex = 1 Then
  '    Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmRptTransJournal
      ARptTransSummary.txtDate = Now
      ARptTransSummary.txtTown = TownName$
      ARptTransSummary.LblRange.Caption = "Detail of Utility Transactions batch: " + bdate$
      ARptTransSummary.totTrans = TransCnt&
      ARptTransSummary.GetName ReportFile$ ', ReportSum$, DetFlag, MaxRevenue
      ARptTransSummary.startrpt
    ElseIf fpcboRptType.ListIndex = 0 Then
    '  Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmRptTransJournal
      ARptTransSumLand.txtDate = Now
      ARptTransSumLand.txtTown = TownName$
      ARptTransSumLand.LblRange.Caption = "Detail of Utility Transactions batch: " + bdate$
      ARptTransSumLand.totTrans = TransCnt&
      ARptTransSumLand.GetName ReportFile$ ', ReportSum$, DetFlag, MaxRevenue
      ARptTransSumLand.startrpt
    ElseIf fpcboRptType.ListIndex = 2 Then
      ViewPrint ReportFile$, "Utility Transaction Interface Summary Report", True
      KillFile ReportFile$
      ActivateControls Me, True
      'TransSummaryPerCust
    Else
      ActivateControls Me, True
    End If
  Else
    'MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If
GLUBKill = 1
ExitDetailedListing:

  Exit Sub

DoDetailedRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TownName$
  Print #UBRpt, Tab(48); "Utility Transaction Interface Summary Report"; Tab(113); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Detail of Utility Transactions batch: " + bdate$
  Print #UBRpt, " "
  Print #UBRpt, Tab(3); "    Bills"; Tab(14); " Penalties"; Tab(25); "    Up Adj"; Tab(36); "   Dwn Adj"; Tab(47); "  Payments";
  Print #UBRpt, Tab(58); "  Drft Pay"; Tab(69); "    OV Adj"; Tab(80); "   Dep Pay"; Tab(91); "   App Dep"; Tab(102); "   Ref Dep"; Tab(113); "  Dep CrRem"
  Print #UBRpt, Dash120$
'  Linecnt = 10
  Return

DoDetailedRptFooter:
  Print #UBRpt, Dash120$
  TotalRevsAmt# = 0
  For RCnt = 1 To MaxRevenue
    TotBills# = Round#(TotBills# + RevTotals(RCnt, 1))
    TotPen# = Round#(TotPen# + RevTotals(RCnt, 4))
    TotUA# = Round#(TotUA# + RevTotals(RCnt, 8))
    TotDA# = Round#(TotDA# + RevTotals(RCnt, 9))
    TotPay# = Round#(TotPay# + RevTotals(RCnt, 2))
    TotDft# = Round#(TotDft# + RevTotals(RCnt, 6))
    TotOvA# = Round#(TotOvA# + RevTotals(RCnt, 10))
    TotDP# = Round#(TotDP# + RevTotals(RCnt, 5))
    TotAP# = Round#(TotAP# + RevTotals(RCnt, 3))
    TotRD# = Round#(TotRD# + RevTotals(RCnt, 7))
    TotCr# = Round#(TotCr# + RevTotals(RCnt, 11))
  Next
  Print #UBRpt, "Totals"
  Print #UBRpt, Tab(2); Using("########.##", TotBills#);
  Print #UBRpt, Tab(13); Using("########.##", TotPen#);
  Print #UBRpt, Tab(24); Using("########.##", TotUA#);
  Print #UBRpt, Tab(35); Using("########.##", TotDA#);
  Print #UBRpt, Tab(46); Using("########.##", TotPay#);
  Print #UBRpt, Tab(57); Using("########.##", TotDft#);
  Print #UBRpt, Tab(68); Using("########.##", TotOvA#);
  Print #UBRpt, Tab(79); Using("########.##", TotDP#);
  Print #UBRpt, Tab(90); Using("########.##", TotAP#);
  Print #UBRpt, Tab(101); Using("########.##", TotRD#);
  Print #UBRpt, Tab(112); Using("########.##", TotCr#)
  Print #UBRpt, Dash120$

  Print #UBRpt, "Transactions: "; TransCnt&

  Print #UBRpt, Dash120$

'  TotalRevsAmt# = 0
'  For RCnt = 1 To MaxRevenue
'    TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
'    Print #SumRpt, RevenueName$(RCnt) + "~" + Using("########.##", RevTotals(RCnt))
'  Next
'  Print #UBRpt,
'  Print #UBRpt, "Total Amount"; Tab(35); Using("########.##", TotalRevsAmt#)
  Return
DefineType:
  Select Case UBTransRec(1).TransType
  Case 1, 101  'Bills col 1 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 1) = Round#(RevTotals(RCnt, 1) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 4, 104 'Payments col 2 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 2) = Round#(RevTotals(RCnt, 2) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 5, 105 'Applied Dep col 3 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 3) = Round#(RevTotals(RCnt, 3) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 6 'Penalties col 4 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 4) = Round#(RevTotals(RCnt, 4) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 7, 107 'Dep Payment col 5 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 5) = Round#(RevTotals(RCnt, 5) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 8  'Draft Payment col 6 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 6) = Round#(RevTotals(RCnt, 6) + ((UBTransRec(1).RevAmt(RCnt) * -1) + (UBTransRec(1).TaxAmt(RCnt) * -1)))
    Next
'''''??????'    Amount# = UBTransRec(1).Transamt * -1
  Case 9, 109  'Refund Dep col 7 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 7) = Round#(RevTotals(RCnt, 7) + (Abs(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt))))
    Next
'''''???????'    Amount# = Abs(UBTransRec(1).Transamt)
  Case 11, 111 'Up Adj Bill col 8 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 8) = Round#(RevTotals(RCnt, 8) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 12, 112  'Down Adj Bill col 9 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 9) = Round#(RevTotals(RCnt, 9) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 33   'Over Pay Adj col 10 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 10) = Round#(RevTotals(RCnt, 10) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 37   'Dep Credit Rem col 11 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 11) = Round#(RevTotals(RCnt, 11) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case 39
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 11) = Round#(RevTotals(RCnt, 11) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  Case Else '99 or any other type col 12 in matrix
    For RCnt = 1 To MaxRevenue
      RevTotals(RCnt, 12) = Round#(RevTotals(RCnt, 12) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
    Next
  End Select
  Return

End Sub

'''  Print #UBRpt, Dash120$
'''  TotalRevsAmt# = 0
'''  For RCnt = 1 To MaxRevenue
'''     RevTotals(RCnt, 1))
'''    TotPen# = Round#(TotPen# + RevTotals(RCnt, 4))
'''    TotUA# = Round#(TotUA# + RevTotals(RCnt, 8))
'''    TotDA# = Round#(TotDA# + RevTotals(RCnt, 9))
'''    TotPay# = Round#(TotPay# + RevTotals(RCnt, 2))
'''    TotDft# = Round#(TotDft# + RevTotals(RCnt, 6))
'''    TotOvA# = Round#(TotOvA# + RevTotals(RCnt, 10))
'''    TotDP# = Round#(TotDP# + RevTotals(RCnt, 5))
'''    TotAP# = Round#(TotAP# + RevTotals(RCnt, 3))
'''    TotRD# = Round#(TotRD# + RevTotals(RCnt, 7))
'''    TotCr# = Round#(TotCr# + RevTotals(RCnt, 11))
'''  Next
'''
'''    Select Case InterfaceMethod
'''    Case 1      'Cash Central
'''      Case TranBillPayment      ' 4=Bill Payment
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''
'''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranAppliedDeposit   ' 5=Applied Deposit
'''        'no action
'''      Case TranPenaltyCharge    ' 6=Penalty Charge
'''        'no action
'''      Case TranDepositPayment   ' 7=Deposit Payment
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''
'''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranDraftPayment     ' 8=Draft Payment
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''
'''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranRefundDeposit    ' 9=Refund Deposit
'''        'no action
'''      Case TranBeginBalance     '10=Beginning Balance
'''        'no action
'''      Case TranUpwardAdjustment '11=Upward Adjustments
'''        'no action
'''      Case TranDownwardAdjustment  '12=Downward Adjustments
'''        'no action
'''      Case TranOverPayAdjustment   '33=OverPayment Adjustments
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranDepCreditRemoval    '37=Deposit Credit Removal Not to Interface w/GL
'''        'No Action !!!
'''      Case TranDepPaymentVoid         ' 39=Deposit Void
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      End Select
'''
'''    Case 2      'Accrual
'''      Select Case UBTransRec(1).TransType
'''      Case TranUtilityBill      ' 1=Utility bill
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranLateCharge       ' 2=late charge
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranReconnectFee     ' 3=reconnect fee
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranBillPayment      ' 4=Bill Payment
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranAppliedDeposit   ' 5=Applied Deposit
'''        'no action
'''        'FOR RevCnt = 1 TO NumOfRevs
'''        '  GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
'''        '  GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''        'NEXT
'''
'''      Case TranPenaltyCharge    ' 6=Penalty Charge
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranDepositPayment   ' 7=Deposit Payment
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranDraftPayment     ' 8=Draft Payment
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranRefundDeposit    ' 9=Refund Deposit
'''        'no action
'''        '  FOR RevCnt = 1 TO NumOfRevs
'''        '    GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
'''        '    GJInfo(RevCnt).dacctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
'''        '  NEXT
'''      Case TranBeginBalance     '10=Beginning Balance
'''        'no action
'''      Case TranUpwardAdjustment '11=Upward Adjustments
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranDownwardAdjustment               '12=Downward Adjustments
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranOverPayAdjustment   '33=OverPayment Adjustments
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      Case TranDepCreditRemoval    '37=Deposit Credit Removal Not to Interface w/GL
'''        'No Action !!!
'''      Case TranDepPaymentVoid         ' 39=Deposit Void
'''        For RevCnt = 1 To NumOfRevs
'''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
'''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
'''        Next
'''        GoSub UpdateTempFile
'''      End Select
'''    End Select
'''    'SmallPause
'''
'''  Next
'''
