VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmRptFlatRate 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Flat Rate Listing"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptFlatRate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRevenues 
      Height          =   348
      Left            =   5472
      TabIndex        =   0
      Top             =   3096
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptFlatRate.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5472
      TabIndex        =   2
      Top             =   4200
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptFlatRate.frx":0C5D
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5472
      TabIndex        =   1
      Top             =   3648
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptFlatRate.frx":0FFB
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
      TabIndex        =   4
      Top             =   7272
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
      TabIndex        =   3
      Top             =   7272
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
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
            TextSave        =   "3:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "6/17/2003"
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
      Height          =   2484
      Left            =   2460
      Top             =   2568
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
      Height          =   324
      Index           =   7
      Left            =   3564
      TabIndex        =   9
      Top             =   3684
      Width           =   1716
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
      Left            =   2964
      TabIndex        =   8
      Top             =   4200
      Width           =   2388
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Revenue:"
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
      Height          =   348
      Left            =   3264
      TabIndex        =   7
      Top             =   3120
      Width           =   2004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1080
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customer Flat Rate List"
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
      Left            =   3618
      TabIndex        =   6
      Top             =   1320
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   960
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
Attribute VB_Name = "frmRptFlatRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptFlatRate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
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
        fpcboRevenues.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboRevenues_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRevenues.ListDown = True
  End If
  If fpcboRevenues.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
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
    CustomerFlatRates2
  ElseIf fpcboRptType.ListIndex = 1 Then
    CustomerFlatRates
  End If
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
  StatusBar1.Panels.Item(1).Text = TownName$
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.ListIndex = 0
  FillRevList fpcboRevenues
  fpcboRevenues.RemoveItem (0)
  fpcboRevenues.ListIndex = 0
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
Private Sub CustomerFlatRates()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAcct As Boolean
  Dim PageNo As Integer, AcctNumber As Long, UBCust As Integer
  Dim Dash80 As String, IndexName As String, UBRpt As Integer
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, RTotal As Double, NTotal As Double, AcctNo As Long
  Dim Book As String, CustCnt As Long, Order As String, TotalFlat As Double
  Dim SEQNUMB As String, RevSource As Integer, POrder As String
  Dim DidOne As Boolean, FRCnt As Integer, CustFlat As Double
  Dim ReportFile As String
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  TownName$ = QPTrim$(UBSetUp(1).UTILNAME)

  FrmShowPctComp.Label1 = "Creating Customer Flat Rate Listing"
  FrmShowPctComp.Show , Me

  UsingBook = False
  UsingAcct = False
  UsingName = False
  Order$ = Mid$(fpcboPrintOrder.Text, 1, 1)
  RevSource = fpcboRevenues.ListIndex + 1
  PageNo = 0
  Select Case Order$
  Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
    POrder$ = " CUSTOMER NAME"
  Case "A"
    POrder$ = " ACCOUNT NUMBER"
        IndexName$ = ""
    UsingAcct = True
  Case "L"
    POrder$ = " LOCATION NUMBER"
    IndexName$ = BookIndexFile
    UsingBook = True
  Case Else
    MsgBox "Invalid Printing Order", vbOKOnly, "Invalid Parameter"
    fpcboPrintOrder.SetFocus
    GoTo ExitCustomerFlat
  End Select
    '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = String$(80, "-")

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
      NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  End If


  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  ReportFile$ = UBPath$ + "UBFRLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  GoSub DoMastFlatHeader

  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Then
      AcctNo = IdxBuff(cnt).RecNum
    Else
      AcctNo = cnt
    End If
     Get UBCust, AcctNo, UBCustRec(1)
     FrmShowPctComp.ShowPctComp cnt, NumOfRecs
     If FrmShowPctComp.Out = True Then
       Close
       FrmShowPctComp.Out = False
       GoTo ExitCustomerFlat
     End If

    If Linecnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoMastFlatHeader
    End If

    If UBCustRec(1).DelFlag <> -1 And UBCustRec(1).Status = "A" Then
      DidOne = False
      For FRCnt = 1 To 4
        If UBCustRec(1).FlatRates(FRCnt).REVSRC = RevSource Then
          DidOne = True
          Print #UBRpt, "Acct:" + Using("#####", AcctNo);
          Print #UBRpt, Tab(15); Left$(UBCustRec(1).CustName, 30);
          If UBCustRec(1).FlatRates(FRCnt).NumMin < 1 Then
            UBCustRec(1).FlatRates(FRCnt).NumMin = 1
          End If
          CustFlat# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * UBCustRec(1).FlatRates(FRCnt).NumMin)
          Print #UBRpt, Tab(45); QPTrim$(UBCustRec(1).FlatRates(FRCnt).FRDESC); Tab(65); Using("#####.##", CustFlat#)
          TotalFlat# = Round#(TotalFlat# + CustFlat#)
          Exit For
        End If
SkipEmFlat:
      Next
      If DidOne Then
        Print #UBRpt, Dash80$
        Linecnt = Linecnt + 1
        CustCnt = CustCnt + 1
      End If
      '*************************************
    End If
  Next

  GoSub DoMastFlatFooter:

  Close UBCust, UBRpt
  Erase IdxBuff, UBCustRec

  If CustCnt > 0 Then
    ViewPrint ReportFile$, "Customer Flat Rate Report."
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
  End If

  KillFile ReportFile$

ExitCustomerFlat:

  Exit Sub

DoMastFlatHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TownName$
  Print #UBRpt, Tab(28); "Customer Flat Rate Listing"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "   Account"; Tab(15); "Customer Name"; Tab(45); "Description"; Tab(67); "Amount"
  Print #UBRpt, Dash80$
  Linecnt = 5
  Return

DoMastFlatFooter:
  Print #UBRpt, "Totals:"; Tab(10); "Customers: "; Using("#####,#", CustCnt)
  Print #UBRpt, Tab(10); "   Amount:"; Using("######,#.##", TotalFlat#)
  Return

End Sub
Private Sub CustomerFlatRates2()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean
  Dim AcctNumber As Long, UBCust As Integer, UsingAcct As Boolean
  Dim IndexName As String, UBRpt As Integer, ToPrint As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, RTotal As Double, NTotal As Double, AcctNo As Long
  Dim Book As String, CustCnt As Long, Order As String, TotalFlat As Double
  Dim SEQNUMB As String, RevSource As Integer, POrder As String
  Dim DidOne As Boolean, FRCnt As Integer, CustFlat As Double
  Dim ReportFile As String
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  TownName$ = QPTrim$(UBSetUp(1).UTILNAME)

  FrmShowPctComp.Label1 = "Creating Customer Flat Rate Listing"
  FrmShowPctComp.Show , Me

  UsingBook = False
  UsingAcct = False
  UsingName = False
  Order$ = Mid$(fpcboPrintOrder.Text, 1, 1)
  RevSource = fpcboRevenues.ListIndex + 1
  Select Case Order$
  Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
    POrder$ = " CUSTOMER NAME"
  Case "A"
    POrder$ = " ACCOUNT NUMBER"
        IndexName$ = ""
    UsingAcct = True
  Case "L"
    POrder$ = " LOCATION NUMBER"
    IndexName$ = BookIndexFile
    UsingBook = True
  Case Else
    MsgBox "Invalid Printing Order", vbOKOnly, "Invalid Parameter"
    fpcboPrintOrder.SetFocus
    GoTo ExitCustomerFlat
  End Select
    '***************

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
      NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  End If


  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  ReportFile$ = UBPath$ + "UBFRLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  
  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Then
      AcctNo = IdxBuff(cnt).RecNum
    Else
      AcctNo = cnt
    End If
     Get UBCust, AcctNo, UBCustRec(1)
     FrmShowPctComp.ShowPctComp cnt, NumOfRecs
     If FrmShowPctComp.Out = True Then
       Close
       FrmShowPctComp.Out = False
       GoTo ExitCustomerFlat
     End If


    If UBCustRec(1).DelFlag <> -1 And UBCustRec(1).Status = "A" Then
      DidOne = False
      For FRCnt = 1 To 4
        If UBCustRec(1).FlatRates(FRCnt).REVSRC = RevSource Then
          DidOne = True
          ToPrint$ = Str$(AcctNo) + "~"
          ToPrint$ = ToPrint$ + Left$(UBCustRec(1).CustName, 30)
          If UBCustRec(1).FlatRates(FRCnt).NumMin < 1 Then
            UBCustRec(1).FlatRates(FRCnt).NumMin = 1
          End If
          CustFlat# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * UBCustRec(1).FlatRates(FRCnt).NumMin)
          ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).FlatRates(FRCnt).FRDESC) + "~" + Str$(CustFlat#)
          TotalFlat# = Round#(TotalFlat# + CustFlat#)
          Exit For
        End If
SkipEmFlat:
      Next
      If DidOne Then
        Print #UBRpt, ToPrint$
        ToPrint$ = ""
        CustCnt = CustCnt + 1
      End If
      '*************************************
    End If
  Next

  Close UBCust, UBRpt
  Erase IdxBuff, UBCustRec

  If CustCnt > 0 Then
    Load frmLoadingRpt
    ARptFlatRate.txtDate = Now
    ARptFlatRate.txtTown = TownName$
    ARptFlatRate.Title = "Customer Flat Rate Report."
    ARptFlatRate.totCust = CustCnt
    ARptFlatRate.GetName ReportFile$
    ARptFlatRate.startrpt
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
  End If


ExitCustomerFlat:

  Exit Sub

'DoMastFlatHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(28); "Customer Flat Rate Listing"; Tab(70); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; Date$
'  Print #UBRpt, Dash80$
'  Linecnt = 4
'  Return

'DoMastFlatFooter:
'  Print #UBRpt, "Totals:"; Tab(10); "Customers: "; Using("#####,#", CustCnt)
'  Print #UBRpt, Tab(10); "   Amount:"; Using("######,#.##", TotalFlat#)
'  Return

End Sub


