VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptTransSummary 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Summary"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmRptTransSummary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   5430
      TabIndex        =   2
      Top             =   4905
      Width           =   3600
      _Version        =   196608
      _ExtentX        =   6350
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ColDesigner     =   "frmRptTransSummary.frx":08CA
   End
   Begin VB.CheckBox chkOther 
      Caption         =   "Include Non-Deleted Customers Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3810
      TabIndex        =   12
      Top             =   3030
      Width           =   5148
   End
   Begin VB.CheckBox DelOnly 
      Caption         =   "Include Deleted Customers Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3810
      TabIndex        =   11
      Top             =   2400
      Width           =   5148
   End
   Begin VB.CheckBox QckSrch 
      BackColor       =   &H008F8265&
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   10
      Top             =   5550
      Width           =   2052
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
      Left            =   7344
      TabIndex        =   3
      Top             =   6435
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
      Left            =   9024
      TabIndex        =   4
      Top             =   6435
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "4:00 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7/20/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtDate2 
      Height          =   345
      Left            =   5430
      TabIndex        =   1
      Top             =   4365
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   345
      Left            =   5430
      TabIndex        =   0
      Top             =   3855
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      Height          =   375
      Left            =   3195
      TabIndex        =   9
      Top             =   4950
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Transaction Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3654
      TabIndex        =   8
      Top             =   984
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3222
      Top             =   816
      Width           =   5772
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
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
      Height          =   375
      Index           =   0
      Left            =   3765
      TabIndex        =   7
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date:"
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
      Height          =   375
      Left            =   3675
      TabIndex        =   6
      Top             =   3900
      Width           =   1665
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4050
      Left            =   2760
      Top             =   2040
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3222
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
Attribute VB_Name = "frmRptTransSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
'Dim BegRoute As String, EndRoute As String
'Dim UseCycle As Boolean
Private Sub cmdExit_Click()
  frmUBEditMenu.Show
  Unload frmRptTransSummary
End Sub

Private Sub cmdPrint_Click()
  If ValidDate = True Then
    DeActivateControls Me, True
    TransSummary
    
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      End If
    End If
  End If
End Sub

Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtDate1) = False And CheckValDate(txtDate2) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    TempDate1 = DateDiff("d", "12/31/1979", txtDate1)
    TempDate2 = DateDiff("d", "12/31/1979", txtDate2)
    If TempDate1 > TempDate2 Then
      ValidDate = False
      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
    Else
      ValidDate = True
    End If
  End If
End Function

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
        txtDate2.SetFocus
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
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics - Portrait"
  fpcboRptType.InsertRow = "Graphics - Landscape"
  fpcboRptType.InsertRow = "Text - Condensed Print"
  fpcboRptType.ListIndex = 0
  QckSrch.Value = 1
End Sub
Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' 'Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub
Private Sub TransSummary()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfTRecs As Long, MaxRevenue As Integer
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
  Dim SumRpt As Integer, ToPrintD As String, cntp As Integer
  Dim ReportFile As String, ReportSum As String, cnttype As Integer
  Dim fmt As String, TotBills As Double, TotPen As Double, TotUA As Double
  Dim TotDA As Double, TotPay As Double, TotDft As Double, TotOvA As Double
  Dim TotDP As Double, TotAP As Double, TotRD As Double, TotCR As Double
  Dim showall As Boolean, showdelonly As Boolean, showregonly As Boolean

 'get report parameters
  GoSub CheckDetailParms
  FrmShowPctComp.Label1 = "Creating Transaction Summary"
  FrmShowPctComp.Show , Me
  MaxLines = 55
  PageNo = 0
  Dash120$ = String$(130, "-")
  fmt$ = "#######.##"
  ReDim RevTotals(1 To 15, 1 To 12) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

  UBTrans = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfTRecs = LOF(UBTrans) \ UBTransRecLen
  
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  ReportFile$ = UBPath$ + "UBDJLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  ReportSum$ = UBPath$ + "UBDJSUM.RPT"
  SumRpt = FreeFile
  Open ReportSum$ For Output As SumRpt
'  ubsetup = FreeFile
'  Open "UBSETUP.DAT" For Random Shared As ubsetup Len = UBSetupreclen
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If Len(TOWNNAME$) = 0 Then
    TOWNNAME$ = "Undefined"
    ' Set Revenue Names to Nothing
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    'Get ubsetup, 1, ubsetup(1)
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
'    TownName$ = ubsetup(1).UTILNAME
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
  End If

  showall = False
  showdelonly = False
  showregonly = False
    If DelOnly.Value = 1 And chkOther.Value = 1 Then
      showall = True
    ElseIf DelOnly.Value = 1 And chkOther.Value = 0 Then
      showdelonly = True
    ElseIf DelOnly.Value = 0 And chkOther.Value = 1 Then
      showregonly = True
    End If

  If NumOfTRecs& = 0 Then
    FrmShowPctComp.ShowPctComp 100, 100
  Else
  For cnt = 1 To NumOfTRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfTRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitDetailedListing
    End If

    Get UBTrans, cnt, UBTransRec(1)
    Get UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)
   If showall = True Then
      'dont skip any
    ElseIf showdelonly = True Then
      If UBCustRec(1).DelFlag = 0 Then
        GoTo SkipThisOne
      End If
    ElseIf showregonly = True Then
      If UBCustRec(1).DelFlag <> 0 Then
        GoTo SkipThisOne
      End If
    End If


    
       If QckSrch.Value = 1 Then
        If UBTransRec(1).TransDate < BegDate Then
          BadCount = BadCount + 1
          If BadCount > 3 Then
            Exit For
          End If
        End If
      End If
      'Check Date, Operator and Trans Type
      If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
        If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
          If (UBTransRec(1).TransType >= BegTrans And UBTransRec(1).TransType <= EndTrans Or (UBTransRec(1).TransType >= BegTrans + 100 And UBTransRec(1).TransType <= EndTrans + 100)) Then
            GoSub DefineType
            TransCnt& = TransCnt& + 1
          End If
        End If
      End If
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
'      Trans& = UBTransRec(1).PrevTrans
   
SkipThisOne:
'    ShowPctComp cnt, NumOfRecs
  Next
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

  'Erase IdxBuff, UBCustRec

  'END

  'If Not AbortFlag Then
  '  PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", LptPort, RetCode, EntryPoint
 ' End If
 ' ViewPrint "UBDJLIST.RPT", "Detailed Journal Report", True
  'KillFile "UBDJLIST.RPT"
 
  
  If TransCnt& > 0 Then
    
    
    If fpcboRptType.ListIndex = 0 Then
      Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmRptTransSummary
      ARptTransSummary.txtDate = Now
      ARptTransSummary.txtTown = TOWNNAME$
      ARptTransSummary.LblRange.Caption = "Date Range: " + Date1$ + " to " + Date2$
      ARptTransSummary.totTrans = TransCnt&
      ARptTransSummary.GetName ReportFile$ ', ReportSum$, DetFlag, MaxRevenue
      ARptTransSummary.startrpt
    ElseIf fpcboRptType.ListIndex = 1 Then
      Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmRptTransSummary
      ARptTransSumLand.txtDate = Now
      ARptTransSumLand.txtTown = TOWNNAME$
      ARptTransSumLand.LblRange.Caption = "Date Range: " + Date1$ + " to " + Date2$
      ARptTransSumLand.totTrans = TransCnt&
      ARptTransSumLand.GetName ReportFile$ ', ReportSum$, DetFlag, MaxRevenue
      ARptTransSumLand.startrpt
    ElseIf fpcboRptType.ListIndex = 2 Then
      ViewPrint ReportFile$, "Transaction Summary Report", True
      KillFile ReportFile$
      ActivateControls Me, True
      TransSummaryPerCust
    Else
      ActivateControls Me, True
    End If
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If

ExitDetailedListing:
  
  Exit Sub

DoDetailedRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(48); "Transaction Summary Report"; Tab(113); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Date Range: " + Date1$ + " To " + Date2$
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
    TotCR# = Round#(TotCR# + RevTotals(RCnt, 11))
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
  Print #UBRpt, Tab(112); Using("########.##", TotCR#)
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

CheckDetailParms:

  Date1$ = txtDate1
  Date2$ = txtDate2

  BegDate = Date2Num%(Date1$)
  EndDate = Date2Num%(Date2$)

  FromBook = 0 'Val(BegRoute)
  ThruBook = 99 'Val(EndRoute)
'  If fpcboTransType.ListIndex <> -1 Then
'    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
'    TrTyp = Val(TrType$)
'  Else
'    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
'    fpcboTransType.SetFocus
'    GoSub ExitDetailedListing
'  End If
TrTyp = 0
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
  If TrTyp = 0 Then
    BegTrans = 1
    EndTrans = 999
  Else
    BegTrans = TrTyp
    EndTrans = TrTyp
  End If

'  OperatorNo$ = fptxtOperator
'  Operator = Val(Operator$)
Operator = 0
  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 99
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If
  
  'Detail$ = QPTrim$(Left$(fpcboDetail.Text, 1))

  'CUSTTYPE$ = QPTrim$(fptxtCustType)
'  If Len(CUSTTYPE$) > 0 Then
'    UseType = True
'  End If

'  Select Case Left$(fpcboPrintOrder.Text, 1)
'    Case "C"
'    IndexName$ = NameIndexFile
'    UsingName = True
'  Case "A"
'    IndexName$ = ""
'    UsingAcct = True
'  Case "L"
'    IndexName$ = BookIndexFile
'    UsingBook = True
'  Case "S"
'    IndexName$ = TempIndexName
'    UsingAddr = True
'  Case Else
'  End Select
Return
End Sub
Private Sub TransSummaryPerCust()
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
  Dim SumRpt As Integer, ToPrintD As String, cntp As Integer
  Dim ReportFile As String, ReportSum As String, cnttype As Integer
  Dim fmt As String, TotBills As Double, TotPen As Double, TotUA As Double
  Dim TotDA As Double, TotPay As Double, TotDft As Double, TotOvA As Double
  Dim TotDP As Double, TotAP As Double, TotRD As Double, TotCR As Double
  Dim BAltouse As Double, begtouse As Double
  Dim showall As Boolean, showdelonly As Boolean, showregonly As Boolean

 'get report parameters
  GoSub CheckDetailParms
  FrmShowPctComp.Label1 = "Creating Transaction Summary"
  FrmShowPctComp.Show , Me
  MaxLines = 55
  PageNo = 0
  Dash120$ = String$(130, "-")
  fmt$ = "#######.##"
 ' ReDim RevTotals(1 To 15, 1 To 12) As Double
 ' ReDim RevenueName(1 To 15) As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

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
'  ElseIf UsingAddr Then
''unrem
'    SortServiceAddrs frmRptMastCust
'    IdxRecLen = 4               'we are using a long integer
'    IdxFileSize& = FileSize&(IndexName$)
'    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
'    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
'    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
'    NumOfRecs = IdxNumOfRecs
'    Handle = FreeFile
'    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
'    For cnt& = 1 To IdxNumOfRecs
'      Get #Handle, cnt&, IdxBuff(cnt&)
'    Next
'    Close Handle
'
  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
  
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBDJLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  ReportSum$ = UBPath$ + "UBDJSUM.RPT"
  SumRpt = FreeFile
  Open ReportSum$ For Output As SumRpt
'  ubsetup = FreeFile
'  Open "UBSETUP.DAT" For Random Shared As ubsetup Len = UBSetupreclen
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If Len(TOWNNAME$) = 0 Then
    TOWNNAME$ = "Undefined"
    ' Set Revenue Names to Nothing
'    For RCnt = 1 To 15
'      RevenueName$(RCnt) = "Not Set"
'    Next RCnt
  Else
    'Get ubsetup, 1, ubsetup(1)
'    For RCnt = 1 To 15
'      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
'    Next RCnt
'    RCnt = 1
'    Do While RCnt <= 15
'      If RevenueName$(RCnt) = "" Then
'        MaxRevenue = RCnt - 1
'        Exit Do
'      End If
'      RCnt = RCnt + 1
'    Loop
    TOWNNAME$ = UBSetUp(1).UTILNAME
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
  End If
  'Close ubsetup

  'Special Code just for ellenboro!!
'  If InStr(TownName$, "ELLENBO") > 0 Then
'    EllenFlag = True
'  End If

'  BlockClear
'  ShowProcessingScrn "Detailed Journal Report."
  showall = False
  showdelonly = False
  showregonly = False
    If DelOnly.Value = 1 And chkOther.Value = 1 Then
      showall = True
    ElseIf DelOnly.Value = 1 And chkOther.Value = 0 Then
      showdelonly = True
    ElseIf DelOnly.Value = 0 And chkOther.Value = 1 Then
      showregonly = True
    End If

'  GoSub DoDetailedRptHeader
  If NumOfRecs& = 0 Then
    FrmShowPctComp.ShowPctComp 100, 100
  Else
  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitDetailedListing
    End If

' '   If UsingName Or UsingBook Or UsingAddr Then
'      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
'    Else
      Get UBCust, cnt, UBCustRec(1)
'    End If

    If showall = True Then
      'dont skip any
    ElseIf showdelonly = True Then
      If UBCustRec(1).DelFlag = 0 Then
        GoTo SkipThisOne
      End If
    ElseIf showregonly = True Then
      If UBCustRec(1).DelFlag <> 0 Then
        GoTo SkipThisOne
      End If
    End If

'    If Linecnt > MaxLines Then
'      Print #UBRpt, FF$
'      GoSub DoDetailedRptHeader
'    End If
'*************************************
'   Main Body of Printing goes here
    BadCount = 0
    Trans& = UBCustRec(1).LastTrans
    TransCnt = 0
    Do While Trans& <> 0
      Get UBTrans, Trans&, UBTransRec(1)
      'If Not EllenFlag Then
'        If UBTransRec(1).TransDate < BegDate Then
'          BadCount = BadCount + 1
'          If BadCount > 3 Then
'            Exit Do
'          End If
'        End If
      'End If
      'Check Date, Operator and Trans Type
      If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
        'If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
         'If (UBTransRec(1).TransType >= BegTrans And UBTransRec(1).TransType <= EndTrans Or (UBTransRec(1).TransType >= BegTrans + 100 And UBTransRec(1).TransType <= EndTrans + 100)) Then
            GoSub DefineType
            TransCnt& = TransCnt& + 1
            If TransCnt& = 1 Then
              BAltouse = UBTransRec(1).RunBalance
            End If
         ' End If
        'End If
      End If
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
      Trans& = UBTransRec(1).PrevTrans
    Loop
    GoSub DoDetailed
   
SkipThisOne:
'    ShowPctComp cnt, NumOfRecs
  Next
  End If
  If fpcboRptType.ListIndex <> 2 Then
'    For cntp = 1 To MaxRevenue
'      For cnttype = 1 To 12
'        ToPrintD$ = ToPrintD$ + Str(RevTotals(cntp, cnttype)) + "~"
'      Next
'      Print #UBRpt, RevenueName$(cntp) + "~" + ToPrintD$
'      ToPrintD$ = ""
'    Next
  Else

'    For cntp = 1 To MaxRevenue
'      Print #UBRpt, Tab(1); RevenueName$(cntp)
'      Print #UBRpt, Tab(3); Using(fmt$, (RevTotals(cntp, 1)));
'      Print #UBRpt, Tab(14); Using(fmt$, (RevTotals(cntp, 4)));
'      Print #UBRpt, Tab(25); Using(fmt$, (RevTotals(cntp, 8)));
'      Print #UBRpt, Tab(36); Using(fmt$, (RevTotals(cntp, 9)));
'      Print #UBRpt, Tab(47); Using(fmt$, (RevTotals(cntp, 2)));
'      Print #UBRpt, Tab(58); Using(fmt$, (RevTotals(cntp, 6)));
'      Print #UBRpt, Tab(69); Using(fmt$, (RevTotals(cntp, 10)));
'      Print #UBRpt, Tab(80); Using(fmt$, (RevTotals(cntp, 5)));
'      Print #UBRpt, Tab(91); Using(fmt$, (RevTotals(cntp, 3)));
'      Print #UBRpt, Tab(102); Using(fmt$, (RevTotals(cntp, 7)));
'      Print #UBRpt, Tab(113); Using(fmt$, (RevTotals(cntp, 11)))
'      'Print #UBRpt, Tab(124); Using(fmt$, (RevTotals(cntp, 1)));
'    Next
      
'    Print #UBRpt, FF$;

  End If
 
 Close

  Erase IdxBuff, UBCustRec

  'END

  'If Not AbortFlag Then
  '  PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", LptPort, RetCode, EntryPoint
 ' End If
 ' ViewPrint "UBDJLIST.RPT", "Detailed Journal Report", True
  'KillFile "UBDJLIST.RPT"
 
  
 ' If TransCnt& > 0 Then
    
    
'    If fpcboRptType.ListIndex = 0 Then
'      Load frmLoadingRpt
'      frmLoadingRpt.setwherefrom frmRptTransSummary
'      ARptTransSummary.txtDate = Now
'      ARptTransSummary.txtTown = TOWNNAME$
'      ARptTransSummary.LblRange.Caption = "Date Range: " + Date1$ + " to " + Date2$
'      ARptTransSummary.totTrans = TransCnt&
'      ARptTransSummary.GetName ReportFile$ ', ReportSum$, DetFlag, MaxRevenue
'      ARptTransSummary.startrpt
'    ElseIf fpcboRptType.ListIndex = 1 Then
'      Load frmLoadingRpt
'      frmLoadingRpt.setwherefrom frmRptTransSummary
'      ARptTransSumLand.txtDate = Now
'      ARptTransSumLand.txtTown = TOWNNAME$
'      ARptTransSumLand.LblRange.Caption = "Date Range: " + Date1$ + " to " + Date2$
'      ARptTransSumLand.totTrans = TransCnt&
'      ARptTransSumLand.GetName ReportFile$ ', ReportSum$, DetFlag, MaxRevenue
'      ARptTransSumLand.startrpt
'    ElseIf fpcboRptType.ListIndex = 2 Then
      ViewPrint ReportFile$, "Transaction Summary Report", True
      KillFile ReportFile$
      ActivateControls Me, True
'    Else
'      ActivateControls Me, True
'    End If
'  Else
'    MsgBox "No Information to print.", vbOKOnly, "No Information"
'    ActivateControls Me, True
'  End If

ExitDetailedListing:
  
  Exit Sub

DoDetailedRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(48); "Transaction Summary Report"; Tab(113); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Date Range: " + Date1$ + " To " + Date2$
  Print #UBRpt, " "
  Print #UBRpt, Tab(3); "    Bills"; Tab(14); " Penalties"; Tab(25); "    Up Adj"; Tab(36); "   Dwn Adj"; Tab(47); "  Payments";
  Print #UBRpt, Tab(58); "  Drft Pay"; Tab(69); "    OV Adj"; Tab(80); "   Dep Pay"; Tab(91); "   App Dep"; Tab(102); "   Ref Dep"; Tab(113); "  Dep CrRem"
  Print #UBRpt, Dash120$
'  Linecnt = 10
  Return
DoDetailed:
  begtouse = Round#(BAltouse - TotBills#)
  begtouse = Round#(begtouse - TotPen#)
  begtouse = Round#(begtouse - TotUA#)
  begtouse = Round#(begtouse + TotDA#)
  begtouse = Round#(begtouse + TotPay#)
  begtouse = Round#(begtouse + TotDft#)
  begtouse = Round#(begtouse - TotOvA#)
  begtouse = Round#(begtouse + TotAP#)
  Print #UBRpt, "Customer - "; Str(cnt); ","; QPTrim$(UBCustRec(1).CustName);
  Print #UBRpt, "Beg. Bal "; Using("########.##", begtouse#)
  Print #UBRpt, "End  Bal "; Using("########.##", BAltouse#)
  
  Print #UBRpt, "Totals"
  Print #UBRpt, Tab(2); Using("########.##", TotBills#);
  Print #UBRpt, Tab(13); Using("########.##", TotPen#);
  Print #UBRpt, Tab(24); Using("########.##", TotUA#);
  Print #UBRpt, Tab(35); Using("########.##", TotDA#);
  Print #UBRpt, Tab(46); Using("########.##", TotPay#);
  Print #UBRpt, Tab(57); Using("########.##", TotDft#);
  Print #UBRpt, Tab(68); Using("########.##", TotOvA#);
  'Print #UBRpt, Tab(79); Using("########.##", TotDP#);
  Print #UBRpt, Tab(90); Using("########.##", TotAP#);
  'Print #UBRpt, Tab(101); Using("########.##", TotRD#);
 ' Print #UBRpt, Tab(112); Using("########.##", TotCR#)
  Print #UBRpt, Dash120$

  Print #UBRpt, "Transactions: "; TransCnt&
  
  Print #UBRpt, Dash120$
  begtouse# = 0
  TotBills# = 0
  TotPen# = 0
  TotUA# = 0
  TotDA# = 0
  TotPay# = 0
  TotDft# = 0
  TotOvA# = 0
 ' TotDP# = 0
  TotAP# = 0
  'TotRD# = 0
  'TotCR# = 0
  
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
    TotBills# = Round#(TotBills# + UBTransRec(1).Transamt)
  Case 4, 104 'Payments col 2 in matrix
    TotPay# = Round#(TotPay# + UBTransRec(1).Transamt)
  Case 5, 105 'Applied Dep col 3 in matrix
    TotAP# = Round#(TotAP# + UBTransRec(1).Transamt)
  Case 6 'Penalties col 4 in matrix
    TotPen# = Round#(TotPen# + UBTransRec(1).Transamt)
  Case 7, 107 'Dep Payment col 5 in matrix
    TotDP# = Round#(TotDP# + UBTransRec(1).Transamt)
  Case 8  'Draft Payment col 6 in matrix
    TotDft# = Round#(TotDft# + UBTransRec(1).Transamt)
'''''??????'    Amount# = UBTransRec(1).Transamt * -1
  Case 9, 109  'Refund Dep col 7 in matrix
    TotRD# = Round#(TotRD# + UBTransRec(1).Transamt)
'''''???????'    Amount# = Abs(UBTransRec(1).Transamt)
  Case 11, 111 'Up Adj Bill col 8 in matrix
    TotUA# = Round#(TotUA# + UBTransRec(1).Transamt)
  Case 12, 112  'Down Adj Bill col 9 in matrix
    TotDA# = Round#(TotDA# + UBTransRec(1).Transamt)
  Case 33   'Over Pay Adj col 10 in matrix
    TotOvA# = Round#(TotOvA# + UBTransRec(1).Transamt)
  Case 37   'Dep Credit Rem col 11 in matrix
    TotCR# = Round#(TotCR# + UBTransRec(1).Transamt)
  Case 39
    TotCR# = Round#(TotCR# + UBTransRec(1).Transamt)
  Case Else '99 or any other type col 12 in matrix
  End Select
  Return

CheckDetailParms:

  Date1$ = txtDate1
  Date2$ = txtDate2

  BegDate = Date2Num%(Date1$)
  EndDate = Date2Num%(Date2$)

  FromBook = 0 'Val(BegRoute)
  ThruBook = 99 'Val(EndRoute)
'  If fpcboTransType.ListIndex <> -1 Then
'    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
'    TrTyp = Val(TrType$)
'  Else
'    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
'    fpcboTransType.SetFocus
'    GoSub ExitDetailedListing
'  End If
TrTyp = 0
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
  If TrTyp = 0 Then
    BegTrans = 1
    EndTrans = 999
  Else
    BegTrans = TrTyp
    EndTrans = TrTyp
  End If

'  OperatorNo$ = fptxtOperator
'  Operator = Val(Operator$)
'Operator = 0
'  If Operator = 0 Then
'    BegOperator = 0
'    EndOperator = 99
'  Else
'    BegOperator = Operator
'    EndOperator = Operator
'  End If
  
  'Detail$ = QPTrim$(Left$(fpcboDetail.Text, 1))

  'CUSTTYPE$ = QPTrim$(fptxtCustType)
'  If Len(CUSTTYPE$) > 0 Then
'    UseType = True
'  End If

'  Select Case Left$(fpcboPrintOrder.Text, 1)
'    Case "C"
'    IndexName$ = NameIndexFile
'    UsingName = True
'  Case "A"
'    IndexName$ = ""
'    UsingAcct = True
'  Case "L"
'    IndexName$ = BookIndexFile
'    UsingBook = True
'  Case "S"
'    IndexName$ = TempIndexName
'    UsingAddr = True
'  Case Else
'  End Select
Return
End Sub




