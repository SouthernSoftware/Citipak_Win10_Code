VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptCustListings 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Listings"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptCustListings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.OptionButton OptNonOwner 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NonOwner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   7320
      TabIndex        =   12
      Top             =   2736
      Width           =   1932
   End
   Begin VB.OptionButton OptOwner 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Owner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   5208
      TabIndex        =   11
      Top             =   2736
      Width           =   1932
   End
   Begin VB.OptionButton OptResident 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Residential"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3096
      TabIndex        =   10
      Top             =   2736
      Width           =   1932
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5256
      TabIndex        =   2
      Top             =   4656
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
      ColDesigner     =   "frmRptCustListings.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5256
      TabIndex        =   1
      Top             =   4116
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
      ColDesigner     =   "frmRptCustListings.frx":0BF8
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
      Top             =   7560
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
      Top             =   7560
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
            TextSave        =   "3:08 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "9/2/2005"
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
   Begin EditLib.fpDateTime txtOpnDate 
      Height          =   348
      Left            =   5256
      TabIndex        =   0
      Top             =   3552
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acct Open Date:"
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
      Index           =   1
      Left            =   3048
      TabIndex        =   9
      Top             =   3624
      Width           =   2076
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
      Left            =   3408
      TabIndex        =   8
      Top             =   4140
      Width           =   1716
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3396
      Left            =   2460
      Top             =   2232
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
      Left            =   2808
      TabIndex        =   7
      Top             =   4680
      Width           =   2388
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3192
      Top             =   312
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customer Listings"
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
      TabIndex        =   6
      Top             =   480
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
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
Attribute VB_Name = "frmRptCustListings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmDCReportsMenu.Show
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
        DCLog "Closed via RptTransJournal by " + PWUser$
        CitiTerminate
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


Private Sub fpDecalCat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpDecalCat.ListDown = True
  End If
  If fpDecalCat.ListDown <> True Then
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
        fpDecalCat.SetFocus
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


Private Sub cmdPrint_Click()
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 1 Then
      CustomerListing
      ActivateControls Me, True
    ElseIf fpcboRptType.ListIndex = 0 Then
      CustomerListing
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
  FillCatCMBOAll fpDecalCat
  fpDecalCat.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub CustomerListing()
  Dim DCCustRecLen As Integer, Page As Integer
  Dim UsingName As Boolean, TotalBal As Double, TCnt As Long
  Dim CustomerCnt As Long, UsingAcct As Boolean
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, DCRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfDCRecs As Long, AcctNo As Long
  Dim DCFile As Integer, ToPrint As String, Header As String
  Dim CatSel As String, Catdo As Boolean, DCvFile As Integer
  Dim RptType As Integer, DCCFile As Integer, NumOfVRecs As Long
  Dim Dash80 As String, CustCnt As Long, VehRecord As Long
  Dim cnt As Long, RptHandle As Integer, Category As String
  Dim ReportFile As String, DCVehReclen As Integer
  Dim DCCodeRecLen As Integer, NumCodeRecs As Integer
  ReDim DCCustRec(1) As DCCustRecType
  ReDim DCCodeRec(1) As DCCatCodeRecType
  RptType = fpcboRptType.ListIndex
  DCCodeRecLen = Len(DCCodeRec(1))
  DCCustRecLen = Len(DCCustRec(1))
  NumCodeRecs = FileSize(DCPath + "DCCODE.DAT") \ DCCodeRecLen
  Dash120$ = String$(121, "-")
  
  If fpDecalCat.ListIndex <> 0 Then
    fpDecalCat.col = 1
    CatSel = QPTrim(fpDecalCat.ColText)
    Catdo = True
    fpDecalCat.col = 2
    Category$ = CatSel + " " + QPTrim(fpDecalCat.ColText)
  Else
    Category$ = "All"
    CatSel = "All"
    Catdo = False
  End If

  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    IndexName$ = DCPath$ + "DCCUST.IDX"
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case Else
  End Select
  ReportFile$ = "DCDetCus.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  Linecnt = 0
  CustCnt = 0
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  GoSub PrintDetailCustomerRptHeader
  ' Print Main Body
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
  'Open Vehicle File
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  
  DCFile = FreeFile
  Open DCPath + "DCCUST.DAT" For Random Shared As DCFile Len = DCCustRecLen
  For cnt = 1 To NumOfDCRecs
  'If cnt = NumOfRecs Then Stop
    FrmShowPctComp.ShowPctComp cnt, NumOfDCRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitPrint
    End If

    If UsingAcct Then
      AcctNo& = cnt
    ElseIf UsingName Then
      AcctNo& = IndexArray(cnt).IDXRECORD
    End If
    Get DCFile, AcctNo&, DCCustRec(1)
    If DCCustRec(1).Deleted <> "Y" Then
      If Linecnt >= MaxLines Then
        If RptType = 1 Then Print #RptHandle, FF$
        GoSub PrintDetailCustomerRptHeader
      End If
      If Catdo Then
        VehRecord = DCCustRec(1).FirstCar
        While VehRecord <> 0
          Get DCvFile, VehRecord, DCVRec(1)
          If RTrim$(DCVRec(1).DecalCat) = CatSel Then GoTo ProcessThisOne
          'HERE DALE
          If VehRecord = DCVRec(1).NextRec Then
            GoTo ExitWhile
          End If
          VehRecord = DCVRec(1).NextRec
        Wend
ExitWhile:
        GoTo SkipRecord
      End If
ProcessThisOne:
    If RptType = 1 Then
      Print #RptHandle, "Cust #"; AcctNo&
      Print #RptHandle, QPTrim$(DCCustRec(1).BILLNAME)
      Print #RptHandle, RTrim$(DCCustRec(1).ADDRESS1); Tab(50); " H Phone: "; DCCustRec(1).HPHONE
      Print #RptHandle, RTrim$(DCCustRec(1).ADDRESS2); Tab(50); " W Phone: "; DCCustRec(1).WPHONE
      Print #RptHandle, RTrim$(DCCustRec(1).City); ", "; DCCustRec(1).State; ""; DCCustRec(1).ZIPCODE
      Print #RptHandle, String$(80, "-")
      CustCnt = CustCnt + 1
      Linecnt = Linecnt + 6
    Else
      ToPrint$ = Str(AcctNo&) + "~" + QPTrim$(DCCustRec(1).BILLNAME)
      ToPrint$ = ToPrint$ + "~" + RTrim$(DCCustRec(1).ADDRESS1) + "~" + RTrim$(DCCustRec(1).ADDRESS2)
      ToPrint$ = ToPrint$ + "~" + RTrim$(DCCustRec(1).City) + "~" + DCCustRec(1).State + "~" + DCCustRec(1).ZIPCODE
      ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).HPHONE) + "~" + QPTrim$(DCCustRec(1).WPHONE)
      CustCnt = CustCnt + 1
      Print #RptHandle, ToPrint$
    End If
   End If
SkipRecord:
  Next cnt
    GoSub PrintDetailCustomerRptEnding
    If RptType <> 0 Then Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
    Close         'Close all open files now
    If RptType = 1 Then
      ViewPrint ReportFile$, Header$
      Kill ReportFile$
    Else
      Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmRptCustDetList
      ARptMastCustList.txtDate = Now
      ARptMastCustList.txtTown = TOWNNAME$
      ARptMastCustList.Catfield = "Category -" + Category$
      ARptMastCustList.Title = "Customer Detail Report"
      ARptMastCustList.totCust = CustCnt
      ARptMastCustList.GetName ReportFile$
      ARptMastCustList.startrpt
    End If
ExitPrint:
  Exit Sub
PrintDetailCustomerRptHeader:
If RptType <> 0 Then
  Page = Page + 1
  Print #RptHandle, Tab(22); "Va. Decal System : Detailed Customer Listing"
  Print #RptHandle, "Category: "; Category$
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, String$(80, "=")
  Linecnt = 4
End If
  Return

PrintDetailCustomerRptEnding:
If RptType <> 0 Then
  Print #RptHandle, "Total Customers Printed: "; Using("#####", CustCnt)
  Print #RptHandle,
  Print #RptHandle, FF$
End If
  Return

End Sub
'Private Sub TransactionJournal()
'  Dim DCCustRecLen As Integer, Page As Integer, TrNumRecs As Long
'  Dim UsingName As Boolean, Totaltot As Double, TCnt As Long
'  Dim CustomerCnt As Long, DCTrFile As Integer, UsingAcct As Boolean
'  Dim IndexName As String, Handle As Integer, Dash120 As String
'  Dim IdxRecLen As Integer, IdxFileSize As Long, DCRpt As Integer
'  Dim IdxNumOfRecs As Long, NumOfDCRecs As Long, AcctNo As Long
'  Dim cnt As Long, DCFile As Integer, UseType As Boolean, ToPrint As String
'  Dim ThisType As String, CatCnt As Integer, CatLoop As Integer
'  Dim BadCount As Long, CatCnt1 As Integer, Lp As Integer, PrnH1 As String
'  Dim DCTransLen As Integer, BegDate As Integer, Trans As Long, PrnH2 As String
'  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
'  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
'  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
'  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
'  Dim TrType As String, RptHandle As Integer, CatSel As String
'  Dim TrTyp As Integer, OperatorNo As String, Catdo As Boolean
'  Dim ReportFile As String, CatCntV As Integer, CatCntV1 As Integer
'  Dim CatCnt2 As Integer, CatCnt3 As Integer, DCCFile As Integer
'  Dim DCCodeRecLen As Integer, NumCodeRecs As Integer, RptType As Integer
'  Dim SumRpt As Integer, ReportSum As String
'  ReDim DCCodeRec(1) As DCCatCodeRecType
'  RptType = fpcboRptType.ListIndex
'  DCCodeRecLen = Len(DCCodeRec(1))
'  NumCodeRecs = FileSize(DCPath + "DCCODE.DAT") \ DCCodeRecLen
'  CatCnt = NumCodeRecs
'  CatCnt1 = NumCodeRecs
'  CatCnt2 = NumCodeRecs
'  CatCnt3 = NumCodeRecs
'  Dash120$ = String$(121, "-")
'  Dim Cat$(250), Cat1$(250), CatAmt#(250), CatAmt1#(250)
'  Dim Cat2$(250), Cat3$(250), CatAmt2#(250), CatAmt3#(250)
'  DCCFile = FreeFile
'  Open DCPath + "DCCODE.DAT" For Random Shared As DCCFile Len = DCCodeRecLen
'    For cnt = 1 To NumCodeRecs
'      Get DCCFile, cnt, DCCodeRec(1)
'      Cat$(cnt) = QPTrim$(DCCodeRec(1).CATCODE)
'      Cat1$(cnt) = QPTrim$(DCCodeRec(1).CATCODE)
'      Cat2$(cnt) = QPTrim$(DCCodeRec(1).CATCODE)
'      Cat3$(cnt) = QPTrim$(DCCodeRec(1).CATCODE)
'    Next
'  Close DCCFile
'
'  ReportFile$ = DCPath$ + "DCTrans.PRN"   'Report File Name
'  ReportSum$ = DCPath$ + "DCTrSum.prn"
'  FF$ = Chr$(12)
'  MaxLines = 53
'  Linecnt = 0
'  FrmShowPctComp.Label1 = "Creating Transaction Journal"
'  FrmShowPctComp.Show , Me
'  ReDim DCCustRec(1) As DCCustRecType
'  DCCustRecLen = Len(DCCustRec(1))
'
'  ReDim DCTransRec(1) As DCTransRecType
'  ReDim Totalamt(1 To 99) As Double
'  GoSub GetReportInfo
'
'
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
'  SumRpt = FreeFile
'  Open ReportSum$ For Output As SumRpt
'
'  If UsingName = True Then
'    NumOfDCRecs = FileSize(IndexName$) \ 4
'    ReDim IndexArray(1 To NumOfDCRecs) As DCTempIDXRecType
'    'FGetAH IndexName$, IndexArray(1), , NumOfRecs
'    Handle = FreeFile
'    Open IndexName$ For Random Shared As Handle Len = 4
'    For cnt& = 1 To NumOfDCRecs
'      Get #Handle, cnt&, IndexArray(cnt&)
'    Next
'    Close Handle
'
'  Else
'    NumOfDCRecs = FileSize(DCPath$ + "DCCUST.DAT") \ DCCustRecLen
'  End If
'  GoSub PrintRptHeader2
'  DCFile = FreeFile
'  Open DCPath + "DCCUST.DAT" For Random Shared As DCFile Len = DCCustRecLen
'  DCTransLen = Len(DCTransRec(1))
'  DCTrFile = FreeFile
'  Open DCPath + "DCTRANS.DAT" For Random Shared As DCTrFile Len = DCTransLen
'
'  For cnt = 1 To NumOfDCRecs
'  'If cnt = NumOfRecs Then Stop
'    FrmShowPctComp.ShowPctComp cnt, NumOfDCRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      GoTo ExitPrint
'    End If
'
'    If UsingAcct Then
'      AcctNo& = cnt
'    ElseIf UsingName Then
'      AcctNo& = IndexArray(cnt).IDXRECORD
'    End If
'    Get DCFile, AcctNo&, DCCustRec(1)
'    If DCCustRec(1).Deleted <> "Y" Then
'
'     BadCount = 0
'      Trans& = DCCustRec(1).FirstTrans
'      Do While Trans& <> 0
'        Get DCTrFile, Trans&, DCTransRec(1)
'
'        If Linecnt >= MaxLines And RptType > 1 Then
'          Print #RptHandle, FF$
'          GoSub PrintRptHeader2
'        End If
'
''      'Check Date, Operator and Trans Type
''      If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
''        If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
''          If (UBTransRec(1).TransType >= BegTrans And UBTransRec(1).TransType <= EndTrans Or (UBTransRec(1).TransType >= BegTrans + 100 And UBTransRec(1).TransType <= EndTrans + 100)) Then
'
'      If DCTransRec(1).TransDate >= BegDate And DCTransRec(1).TransDate <= EndDate Then
'        If (DCTransRec(1).OperNum >= BegOperator And DCTransRec(1).OperNum <= EndOperator) Then
'          If TrType > 0 Then
'            If DCTransRec(1).TransType <> TrType Then
'               GoTo NextTrans
'            End If
'          End If
'          If Catdo = True Then
'            If CatSel <> QPTrim$(DCTransRec(1).DecalCat) Then
'               GoTo NextTrans
'            End If
'          End If
'          If RptType = 2 Then
'            Print #RptHandle, Num2Date$(DCTransRec(1).TransDate);
'            Print #RptHandle, Tab(11); Str$(AcctNo&);
'            Print #RptHandle, Tab(16); Left$(DCCustRec(1).BILLNAME, 20);
'            Print #RptHandle, Tab(37); "";
'            If DCTransRec(1).TransType = 1 Then
'                Print #RptHandle, "Charge";
'                GoSub ChargeSub
'            End If
'            If DCTransRec(1).TransType = 2 Then
'                Print #RptHandle, "Payment";
'                GoSub PaymentSub
'            End If
'            If DCTransRec(1).TransType = 3 Then
'                Print #RptHandle, "V Charge";
'                GoSub VChargeSub
'            End If
'            If DCTransRec(1).TransType = 4 Then
'                Print #RptHandle, "V Payment";
'                GoSub VPaymentSub
'            End If
'            Print #RptHandle, Tab(46); Left$(DCTransRec(1).TRVinDesc, 18);
'            Print #RptHandle, Tab(65); Str$(DCTransRec(1).OperNum);
'            Print #RptHandle, Tab(69); Using("$###,###.##", DCTransRec(1).TransAmount)
'          Else
'            ToPrint$ = Str$(Trans&) + "~" + Num2Date$(DCTransRec(1).TransDate)
'            ToPrint$ = ToPrint$ + "~" + Str$(AcctNo&)
'            ToPrint$ = ToPrint$ + "~" + Left$(DCCustRec(1).BILLNAME, 20)
'            If DCTransRec(1).TransType = 1 Then
'                ToPrint$ = ToPrint$ + "~" + "Charge"
'                GoSub ChargeSub
'            End If
'            If DCTransRec(1).TransType = 2 Then
'                ToPrint$ = ToPrint$ + "~" + "Payment"
'                GoSub PaymentSub
'            End If
'            If DCTransRec(1).TransType = 3 Then
'                ToPrint$ = ToPrint$ + "~" + "V Charge"
'                GoSub VChargeSub
'            End If
'            If DCTransRec(1).TransType = 4 Then
'                ToPrint$ = ToPrint$ + "~" + "V Payment"
'                GoSub VPaymentSub
'            End If
'            ToPrint$ = ToPrint$ + "~" + Left$(DCTransRec(1).TRVinDesc, 18)
'            ToPrint$ = ToPrint$ + "~" + Str$(DCTransRec(1).OperNum)
'            ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", DCTransRec(1).TransAmount)
'            Print #RptHandle, ToPrint$
'          End If
'          ToPrint$ = ""
'          TotalTrans = TotalTrans + 1
'          Linecnt = Linecnt + 1
'          Totaltot# = Totaltot# + DCTransRec(1).TransAmount
'          Totalamt#(DCTransRec(1).TransType) = Totalamt#(DCTransRec(1).TransType) + DCTransRec(1).TransAmount
'        End If
'      End If
'NextTrans:
'  Trans& = DCTransRec(1).NextTrans
'  Loop
'  End If
'Next cnt
'Close DCFile
'    'Now Subtotal by Decal Type
'
'  GoSub PrintRptEnding2
'  If RptType = 2 Then Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
'  Close         'Close all open files now
'
'If TotalTrans > 0 Then
'  If RptType = 2 Then
'    ViewPrint ReportFile$, "Transaction Journal"
'    'Kill ReportFile$
'  Else
'    Load frmLoadingRpt
'    frmLoadingRpt.setwherefrom frmRptTransJournal
'    If RptType = 0 Then
'      ARptTransJournal.txtDate = Now
'      ARptTransJournal.txtTown = TOWNNAME$
'      ARptTransJournal.Title = "Transaction Journal Report"
'      ARptTransJournal.txtRptParm1.Caption = PrnH1$
'      ARptTransJournal.txtRptParm2.Caption = PrnH2$
'      ARptTransJournal.txtPrnOrd = "In " + fpcboPrintOrder.Text
'      ARptTransJournal.totTrans = TotalTrans
'      ARptTransJournal.GetName ReportFile$, ReportSum$
'      ARptTransJournal.startrpt
'    ElseIf RptType = 1 Then
'      ARptTransJPortrait.txtDate = Now
'      ARptTransJPortrait.txtTown = TOWNNAME$
'      ARptTransJPortrait.Title = "Transaction Journal Report"
'      ARptTransJPortrait.txtRptParm1.Caption = PrnH1$
'      ARptTransJPortrait.txtRptParm2.Caption = PrnH2$
'      ARptTransJPortrait.txtPrnOrd = "In " + fpcboPrintOrder.Text
'      ARptTransJPortrait.totTrans = TotalTrans
'      ARptTransJPortrait.GetName ReportFile$, ReportSum$
'      ARptTransJPortrait.startrpt
'   End If
'  End If
'Else
'  MsgBox "No Information to print.", vbOKOnly, "No Information"
'  ActivateControls Me, True
'End If
'ExitPrint:
'  Exit Sub
'
'ChargeSub:
'
'  If DCTransRec(1).TransType = 1 Then
'    For CatLoop = 1 To NumCodeRecs
'      If Cat$(CatLoop) = QPTrim$(DCTransRec(1).DecalCat) Then
'        CatAmt#(CatLoop) = CatAmt#(CatLoop) + DCTransRec(1).TransAmount
'        Return
'      End If
'    Next CatLoop
'  End If
'  Return
'
'PaymentSub:
'  If DCTransRec(1).TransType = 2 Then
'    For CatLoop = 1 To NumCodeRecs
'      If Cat1$(CatLoop) = QPTrim$(DCTransRec(1).DecalCat) Then
'        CatAmt1#(CatLoop) = CatAmt1#(CatLoop) + DCTransRec(1).TransAmount
'        Return
'      End If
'    Next CatLoop
'  End If
'  Return
'VChargeSub:
'  If DCTransRec(1).TransType = 3 Then
'    For CatLoop = 1 To NumCodeRecs
'      If Cat2$(CatLoop) = QPTrim$(DCTransRec(1).DecalCat) Then
'        CatAmt2#(CatLoop) = CatAmt2#(CatLoop) + DCTransRec(1).TransAmount
'        Return
'      End If
'    Next CatLoop
'  End If
'  Return
'
'VPaymentSub:
'  If DCTransRec(1).TransType = 4 Then
'    For CatLoop = 1 To NumCodeRecs
'      If Cat3$(CatLoop) = QPTrim$(DCTransRec(1).DecalCat) Then
'        CatAmt3#(CatLoop) = CatAmt3#(CatLoop) + DCTransRec(1).TransAmount
'        Return
'      End If
'    Next CatLoop
'  End If
'Return
'
'PrintRptHeader2:
'  PrnH1$ = " Beginning Transaction Date: " + Date1$
'  If Val(Operator$) = 0 Then
'    PrnH1$ = PrnH1$ + "     Operator #: ALL" + "      Category: " + CatSel$
'  Else
'    PrnH1$ = PrnH1$ + "     Operator #: " + Mid$(Operator$, 1, 3) + "      Category: " + CatSel$
'  End If
'  PrnH2$ = "    Ending Transaction Date: " + Date2$ + "     Transaction Type: " + fpcboTransType.Text
'
'If RptType = 2 Then
'  Page = Page + 1
'  Print #RptHandle, Tab(20); "Va. Vehicle Decals : Transactions Journal "
'  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
'  Print #RptHandle, "Printing Order: " + QPTrim$(fpcboPrintOrder.Text)
'  Print #RptHandle, PrnH1$
'  Print #RptHandle, PrnH2$
'  Print #RptHandle, ""
'  Print #RptHandle, "  Date"; Tab(13); "Customer Acct/Name"; Tab(37); "Type"; Tab(46); "Description"; Tab(64); "Oper"; Tab(73); "Amount"
'  Print #RptHandle, String$(80, "=")
'  Linecnt = 8
'End If
'  Return
'
'PrintRptEnding2:
'If RptType = 2 Then
'  Print #RptHandle, String$(80, "-")
'  Print #RptHandle, Tab(69); Using("$###,###.##", Totaltot#)
'  For cnt& = 1 To 4
'    If Totalamt#(cnt&) <> 0 Then
'      Print #RptHandle, "Trans Type : ";
'      If cnt& = 1 Then Print #RptHandle, "     Charges  ";
'      If cnt& = 2 Then Print #RptHandle, "     Payments ";
'      If cnt& = 3 Then Print #RptHandle, " Void Charges ";
'      If cnt& = 4 Then Print #RptHandle, "Void Payments ";
'      Print #RptHandle, "     Total Amount: "; Using("$#,###,###.##", Totalamt#(cnt&))
'    End If
'  Next cnt&
'  Print #RptHandle, String$(80, "-")
'  If Totalamt#(1) <> 0 Then
'    Print #RptHandle, "Catagory Totals : Charges"
'    Print #RptHandle, " "
'    Print #RptHandle, "Catagory"; Tab(20); "       Amount"
'    For Lp = 1 To CatCnt
'      Print #RptHandle, Cat$(Lp); Tab(20); Using("$#,###,###.##", CatAmt#(Lp))
'    Next
'  End If
'  Print #RptHandle, " "
'  If Totalamt#(2) <> 0 Then
'    Print #RptHandle, " "
'    Print #RptHandle, "Catagory Totals : Payments"
'    Print #RptHandle, " "
'    Print #RptHandle, "Catagory"; Tab(20); "       Amount"
'    For Lp = 1 To CatCnt1
'      Print #RptHandle, Cat1$(Lp); Tab(20); Using("$#,###,###.##", CatAmt1#(Lp))
'    Next
'    Print #RptHandle, " "
'  End If
'  If Totalamt#(3) <> 0 Then
'    Print #RptHandle, " "
'    Print #RptHandle, "Catagory Totals : Void Charges"
'    Print #RptHandle, " "
'    Print #RptHandle, "Catagory"; Tab(20); "       Amount"
'    For Lp = 1 To CatCnt2
'      Print #RptHandle, Cat2$(Lp); Tab(20); Using("$#,###,###.##", CatAmt2#(Lp))
'    Next
'    Print #RptHandle, " "
'  End If
'  If Totalamt#(4) <> 0 Then
'    Print #RptHandle, " "
'    Print #RptHandle, "Catagory Totals : Void Payments"
'    Print #RptHandle, " "
'    Print #RptHandle, "Catagory"; Tab(20); "       Amount"
'    For Lp = 1 To CatCnt3
'      Print #RptHandle, Cat3$(Lp); Tab(20); Using("$#,###,###.##", CatAmt3#(Lp))
'    Next
'  End If
'  Print #RptHandle, FF$
' Else
'  For cnt& = 1 To 4
'    If Totalamt#(cnt&) <> 0 Then
'      If cnt& = 1 Then Print #SumRpt, "Charges" + "~" + Using("$#,###,###.##", Totalamt#(cnt&))
'      If cnt& = 2 Then Print #SumRpt, "Payments" + "~" + Using("$#,###,###.##", Totalamt#(cnt&))
'      If cnt& = 3 Then Print #SumRpt, "Void Charges" + "~" + Using("$#,###,###.##", Totalamt#(cnt&))
'      If cnt& = 4 Then Print #SumRpt, "Void Payments" + "~" + Using("$#,###,###.##", Totalamt#(cnt&))
'    End If
'  Next cnt&
'  If Totalamt#(1) <> 0 Then
'    Print #SumRpt, " ~ "
'    Print #SumRpt, "Charges~ "
'    Print #SumRpt, "Catagory~Amount"
'    For Lp = 1 To CatCnt
'      Print #SumRpt, Cat$(Lp) + "~" + Using("$#,###,###.##", CatAmt#(Lp))
'    Next
'  End If
'  If Totalamt#(2) <> 0 Then
'    Print #SumRpt, " ~ "
'    Print #SumRpt, " ~ "
'    Print #SumRpt, "Payments~ "
'    Print #SumRpt, "Catagory~Amount"
'    For Lp = 1 To CatCnt1
'      Print #SumRpt, Cat1$(Lp) + "~" + Using("$#,###,###.##", CatAmt1#(Lp))
'    Next
'  End If
'  If Totalamt#(3) <> 0 Then
'    Print #SumRpt, " ~ "
'    Print #SumRpt, " ~ "
'    Print #SumRpt, "Void Charges~ "
'    Print #SumRpt, "Catagory~Amount"
'    For Lp = 1 To CatCnt2
'      Print #SumRpt, Cat2$(Lp) + "~" + Using("$#,###,###.##", CatAmt2#(Lp))
'    Next
'  End If
'  If Totalamt#(4) <> 0 Then
'    Print #SumRpt, " ~ "
'    Print #SumRpt, " ~ "
'    Print #SumRpt, "Void Payments~ "
'    Print #SumRpt, "Catagory~Amount"
'    For Lp = 1 To CatCnt3
'      Print #SumRpt, Cat3$(Lp) + "~" + Using("$#,###,###.##", CatAmt3#(Lp))
'    Next
'  End If
' End If
'  Return
'GetReportInfo:
'  Date1$ = txtDate1
'  Date2$ = txtDate2
'
'  BegDate = Date2Num%(Date1$)
'  EndDate = Date2Num%(Date2$)
'
'  If fpcboTransType.ListIndex <> -1 Then
'    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
'    TrTyp = Val(TrType$)
'  Else
'    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
'    fpcboTransType.SetFocus
'    GoTo ExitPrint
'  End If
'  If TrTyp = 0 Then
'    BegTrans = 1
'    EndTrans = 999
'  Else
'    BegTrans = TrTyp
'    EndTrans = TrTyp
'  End If
'
'  OperatorNo$ = fptxtOperator
'  Operator = Val(OperatorNo$)
'  If Operator = 0 Then
'    BegOperator = 0
'    EndOperator = 9999
'  Else
'    BegOperator = Operator
'    EndOperator = Operator
'  End If
'
'  If fpDecalCat.ListIndex <> 0 Then
'    fpDecalCat.col = 1
'    CatSel = QPTrim(fpDecalCat.ColText)
'    Catdo = True
'  Else
'    CatSel = "All"
'    Catdo = False
'  End If
'
'  Select Case Left$(fpcboPrintOrder.Text, 1)
'    Case "C"
'    IndexName$ = DCPath$ + "DCCUST.IDX"
'    UsingName = True
'  Case "A"
'    IndexName$ = ""
'    UsingAcct = True
'  Case Else
'  End Select
'Return
'
'End Sub
