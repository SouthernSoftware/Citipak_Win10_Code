VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnAPChecks 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Printing"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnAPChecks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrinters 
      Height          =   405
      Left            =   5820
      TabIndex        =   3
      Top             =   5310
      Width           =   3570
      _Version        =   196608
      _ExtentX        =   6297
      _ExtentY        =   714
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   2
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnAPChecks.frx":08CA
   End
   Begin LpLib.fpCombo fpcboBanks 
      Height          =   405
      Left            =   5835
      TabIndex        =   0
      Top             =   3405
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1931
      _ExtentY        =   714
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
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
      ListWidth       =   3495
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
      ColDesigner     =   "frmPrnAPChecks.frx":0C69
   End
   Begin VB.CommandButton cmdAlign 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F5 &Align"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3657
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1236
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D0D0D0&
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
      Height          =   468
      Left            =   5478
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1236
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
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
      Height          =   468
      Left            =   7299
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1236
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8484
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "1:00 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/28/2008"
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
   Begin EditLib.fpText fptxtChkNum 
      Height          =   372
      Left            =   5832
      TabIndex        =   1
      Top             =   4032
      Width           =   1308
      _Version        =   196608
      _ExtentX        =   2307
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fpChkDate 
      Height          =   372
      Left            =   5844
      TabIndex        =   2
      Top             =   4656
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Text            =   "10/01/2001"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Printer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   348
      Left            =   3840
      TabIndex        =   12
      Top             =   5376
      Width           =   1836
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Checks:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   348
      Index           =   1
      Left            =   3456
      TabIndex        =   11
      Top             =   4728
      Width           =   2100
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Check Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Left            =   2928
      TabIndex        =   10
      Top             =   4104
      Width           =   2628
   End
   Begin VB.Label Label4b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   3144
      TabIndex        =   9
      Top             =   3432
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4572
      Left            =   2304
      Top             =   2664
      Width           =   7548
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   960
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print A/P Checks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4332
      TabIndex        =   8
      Top             =   1200
      Width           =   3540
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   840
      Width           =   7020
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
Attribute VB_Name = "frmPrnAPChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim Vendor As VendorRecType
Dim APChkInf As CheckInfoType3
Dim TPayList2 As TPayListType
Dim APCheck As Integer
Dim DePrn As String, DefBnk As Integer
Dim LPDate As Integer, HPDate As Integer, StartCnt As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim CHKinfo() As CheckInfoType3, InvList() As TPayListType
Dim TCheckNum As Long, CheckDate As Integer, Bankcode As Integer
Dim ChkinfoFile As Integer, VCnt As Integer, FirstBadNum As Long

Private Sub cmdAlign_Click()
  If oktoprint = True Then
    PrintAln
  End If
End Sub

Private Sub cmdExit_Click()
  frmAPChkProcessMenu.Show
  frmAPChkProcessMenu.cmdPrintChkRegister.SetFocus
  Unload frmPrnAPChecks
End Sub

Private Sub cmdOk_Click()
  Dim Bank As Integer, chk As Long, CkDate As String, DePrn As String
  Dim DPName As String
  If oktoprint = True Then
    fpcboBanks.col = 0
    Bank = QPTrim(fpcboBanks.ColText)
    chk = QPTrim(fptxtChkNum)
    CkDate = fpChkDate
    fpcboPrinters.col = 0
    DPName = QPTrim(fpcboPrinters.ColText)
    fpcboPrinters.col = 1
    'DefPrinter = fpcboPrinters.ColText
    'for allowance of winxp network printer port situation did following
    'look at frmprint for further details.
    If InStr(1, DPName, "\\", vbTextCompare) Then
      DePrn = DPName
    Else
      DePrn = QPTrim(fpcboPrinters.ColText)
    End If
    Call MainLog("Begin Print APChk: " + Str$(chk) + "," + CkDate)
    ChkCodeDir Bank, chk, CkDate, DePrn
    cmdExit_Click
  End If
End Sub
Private Sub PrintAln()
  Dim ToPrintA As String, LPTA As Integer, RptA As Integer, DPName As String
  Dim DefPrinter As String, Copies As Integer, AlignRpt As String
  On Error GoTo Cancel
  Select Case APCheck
    Case 1
      AlignRpt$ = "Chk1Aln.MSK"
    Case 2
      AlignRpt$ = "Chk2Aln.MSK"
    Case 3
      AlignRpt$ = "Chk3Aln.MSK"
    Case 6
      AlignRpt$ = "Chk6Aln.MSK"
    Case Else
    'laser no align
      MsgBox "Laser Checks Have No Alignment Mask."
      Exit Sub
  End Select
10:
  If fpcboPrinters.ListIndex <> -1 Then
    fpcboPrinters.col = 0
    DPName = QPTrim(fpcboPrinters.ColText)
20:
    fpcboPrinters.col = 1
    'DefPrinter = fpcboPrinters.ColText
    'for allowance of winxp network printer port situation did following
    'look at frmprint for further details.
    If InStr(1, DPName, "\\", vbTextCompare) Then
      DefPrinter = DPName
    Else
      DefPrinter = QPTrim(fpcboPrinters.ColText)
    End If
    Copies = 1
    GoSub PrintAlignMask
  End If
PrintAlignMask:
30:
    LPTA = FreeFile
    Open DefPrinter For Output As LPTA
    RptA = FreeFile
40:
    Open AlignRpt For Input As RptA
    Do
        Line Input #RptA, ToPrintA$
        
        ToPrintA$ = RTrim$(ToPrintA$)
        Print #LPTA, ToPrintA$
50:
    Loop Until eof(RptA)
    Close LPTA, RptA
    fptxtChkNum = fptxtChkNum + 1
    Printer.EndDoc
    If MsgBox("Do You Wish to Print Another Mask?", vbYesNo, "Print Mask") = vbYes Then
      GoSub PrintAlignMask
    End If
Cancel:
  If Err > 0 Then
    MsgBox "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
  End If
  Close
  Exit Sub
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Call MainLog("Close via AP Chk Print.")
        KillFile ("APCHK.opn")
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
  End If
End Sub


Private Sub fpcboBanks_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboBanks.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboBanks.ListIndex = -1
    fpcboBanks.Action = ActionClearSearchBuffer
  End If
  If fpcboBanks.ListDown <> True Then
    If KeyCode = vbKeyDown Then
        SendKeys "{Tab}"
        KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcboPrinters_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrinters.ListDown = True
  End If
  If fpcboPrinters.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub
Private Sub FillPrinters(combo As fpCombo)
Dim cnt As Integer

For cnt = 0 To (Printers.Count - 1)
  fpcboPrinters.InsertRow = Printers(cnt).DeviceName & Chr(9) & Printers(cnt).Port
Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpPrintAPChks
  GetBankList fpcboBanks
  SetDefBank "C", DefBnk
  If DefBnk > 0 Then
    'txtBanks.Col = 0
    fpcboBanks.SearchText = Trim(DefBnk)
    fpcboBanks.Action = 0
    If fpcboBanks.SearchIndex <> -1 Then
      fpcboBanks.ListIndex = fpcboBanks.SearchIndex
    End If
  End If
  
  GetAPCheck APCheck
  GetPostDates LPDate, HPDate
  fpChkDate.Text = Format(Now, "mm/dd/yyyy")
  If APCheck < 4 Or APCheck = 6 Then
    FillPrinters fpcboPrinters
    fpcboPrinters.col = 1
    fpcboPrinters.SearchText = Printer.Port
    fpcboPrinters.Action = 0
    If fpcboPrinters.SearchIndex <> -1 Then
      fpcboPrinters.ListIndex = fpcboPrinters.SearchIndex
    Else
      fpcboPrinters.ListIndex = 0
    End If
  Else
    Label1.Visible = False
    fpcboPrinters.Enabled = False
    fpcboPrinters.Visible = False
    cmdAlign.Enabled = False
    cmdAlign.Visible = False
  End If
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Function oktoprint()
  Dim CheckDate As Integer
  If fpcboBanks.ListIndex = -1 Then
    MsgBox "You Must Select A Bank Code.", vbOKOnly, "Invlaid Bank"
    fpcboBanks.SetFocus
    oktoprint = False
    Exit Function
  End If
  If Val(fptxtChkNum) <= 0 Then
    fptxtChkNum.SetFocus
    oktoprint = False
    MsgBox "You Must Enter A Valid Check Number.", vbOKOnly, "Invalid Check"
    Exit Function
  End If
  CheckDate = DateDiff("d", "12/31/1979", fpChkDate)
  If CheckValDate(fpChkDate) = True Then
    If (CheckDate < LPDate) Or (CheckDate > HPDate) Then
      MsgBox "This Date Is Not Within Allowable Posting Range. Please Correct or Change Setup.", vbOKOnly, "Invalid Date"
      fpChkDate.SetFocus
      oktoprint = False
      Exit Function
    Else
      oktoprint = True
    End If
  Else
    MsgBox "This Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    oktoprint = False
    Exit Function
  End If
  If fpcboPrinters.Enabled = True Then
  If fpcboPrinters.ListIndex <> -1 Then
    oktoprint = True
  Else
    MsgBox "Make A Printer Selection Or Cancel.", vbOKOnly, "Invalid Printer Selection"
    oktoprint = False
    Exit Function
  End If
  End If
End Function
 
Public Sub ChkCodeDir(Bank As Integer, chk As Long, CkDate As String, DePrn As String, Optional RestartFlag As Boolean, Optional oldchk1 As Long, Optional oldchk2 As Long)
  Dim cnt As Integer, ccnt As Integer, BadNum2 As Long
  Dim ChkInfoRecLen As Integer, PayListRecLen As Integer
  Dim TPayListFile As Integer, TNumInvoices As Integer
  Dim LCnt As Integer, VRecNum1 As Long, Pcnt As Integer
  Dim VRecNum2 As Long, ChkCnt As Integer, Lvcnt As Integer
  Dim TCnt As Integer
  
  If RestartFlag = True Then
    FirstBadNum& = oldchk1
    BadNum2 = oldchk2
  End If
  TCheckNum& = chk
  ReDim PayListRec(1) As TPayListType
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  Bankcode = Bank
  ChkInfoRecLen = Len(CHKinfo(1))
  PayListRecLen = Len(PayListRec(1))
  TNumInvoices = (FileSize("TPAYLIST2.LST") \ PayListRecLen)
  CheckDate = DateDiff("d", "12/31/1979", CkDate)
  TPayListFile = FreeFile
'  TNumInvoices = LOF(TPayListFile) \ 6
  ReDim InvList(1 To TNumInvoices) As TPayListType
  'FGetAH "TPAYLIST.LST", InvList(1), PayListRecLen, TNumInvoices
  Open "TPAYLIST2.LST" For Random Shared As TPayListFile Len = PayListRecLen

  For Pcnt = 1 To TNumInvoices
    Get TPayListFile, Pcnt, TPayList2
    InvList(Pcnt).VendorRecNum = TPayList2.VendorRecNum
    InvList(Pcnt).LedgerRecNum = TPayList2.LedgerRecNum
  Next
  Close TPayListFile
  
  If Not RestartFlag Then
    VCnt = 1
    LCnt = 1
    VRecNum1& = InvList(LCnt).VendorRecNum
    Do Until LCnt >= TNumInvoices
      LCnt = LCnt + 1
      VRecNum2& = InvList(LCnt).VendorRecNum
      If VRecNum1& <> VRecNum2& Then
        VCnt = VCnt + 1
        VRecNum1& = VRecNum2&
      End If
    Loop
    If Exist("APCHKINF.DAT") Then
      
      If MsgBox("An Unposted Check File Exists - Continue and Delete, or Cancel", vbOKCancel, "Continue?") = vbCancel Then
        Call MainLog("AP Chk Unposted Chk File -Exit.")
        Exit Sub
      Else
        Call MainLog("AP Chk Unposted Chk File -Continue and delete file.")
        Kill ("APCHKINF.DAT")
      End If
    End If
    ReDim CHKinfo(1 To VCnt) As CheckInfoType3
    ChkCnt = 1
    VRecNum1& = InvList(ChkCnt).VendorRecNum
    CHKinfo(ChkCnt).ListFirst = 1
    CHKinfo(ChkCnt).VendorRecNum = VRecNum1&
    For cnt = 2 To TNumInvoices
      VRecNum2& = InvList(cnt).VendorRecNum
      If VRecNum1& <> VRecNum2& Then
        CHKinfo(ChkCnt).ListLast = cnt - 1
        CHKinfo(ChkCnt).VendorRecNum = InvList(cnt - 1).VendorRecNum
        CHKinfo(ChkCnt).Bankcode = Bankcode
        ChkCnt = ChkCnt + 1
        CHKinfo(ChkCnt).ListFirst = cnt
        VRecNum1& = VRecNum2&
      End If
    Next
    CHKinfo(ChkCnt).ListLast = TNumInvoices
    CHKinfo(ChkCnt).VendorRecNum = InvList(TNumInvoices).VendorRecNum
    CHKinfo(ChkCnt).Bankcode = Bankcode
    KillFile "APCHKINF.DAT"
    ChkinfoFile = FreeFile
    Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
    For ccnt = 1 To VCnt
      Put ChkinfoFile, ccnt, CHKinfo(ccnt)
    Next
    Close ChkinfoFile
    'FPutAH "APCHKINF.DAT", Chkinfo(1), ChkInfoRecLen, ChkCnt
    StartCnt = 1

'***************************
  Else          'We are restarting printing
    VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
    ChkinfoFile = FreeFile
    Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen

    ReDim CHKinfo(1 To VCnt) As CheckInfoType3
    For TCnt = 1 To VCnt
    Get ChkinfoFile, TCnt, CHKinfo(TCnt)
      If CHKinfo(TCnt).LastChk = FirstBadNum& Then
        StartCnt = TCnt
        Exit For
      End If
    Next
    For TCnt = 1 To VCnt
    Get ChkinfoFile, TCnt, CHKinfo(TCnt)
     If CHKinfo(TCnt).LastChk = BadNum2 Then
        Lvcnt = TCnt
        Exit For
      End If
    Next
  VCnt = Lvcnt
  End If
  FrmShowPctComp.Label1 = "Creating Check Print File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnAPChecks, True
  Select Case APCheck
    Case 1
      Call MainLog("Begin APChkPrint 1.")
      Check1prn DePrn
    Case 2
      Call MainLog("Begin APChkPrint 2.")
      Check2prn DePrn
    Case 3
      Call MainLog("Begin APChkPrint 3.")
      Check3prn DePrn
    Case 4
      Call MainLog("Begin APChkPrint 4.")
      Check4prn
    Case 5
      Call MainLog("Begin APChkPrint 5.")
      Check5prn
    Case 6
      Call MainLog("Begin APChkPrint 6.")
      Check6prn DePrn
  End Select
  
ExitCheckPrinting:

End Sub
'************************************* Check 1
Private Sub Check1prn(DePrn As String)
 '---New Std 10 cpi   (9013), Blank Top Stub- 39 Line
  Dim PrintFile As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, DistRecord As Long
  Dim APDistFile As Integer, NumDistRecs As Long, TabStop As Integer
  Dim DoStubHeader As Boolean, TChkAmt As Double, ShowDist As String
  Dim TopStubCnt As Integer, Cnt2 As Integer, RecLen As Integer, Here As Integer
  Dim CntZZ As Integer, VLCnt As Integer, cnt As Integer, ccnt As Integer
  Dim Void As String, MaxTopStub As Integer, ChkPrn As String, Toolong As String
  Dim ToPrint As String, APDistRecLen As Integer, ChkInfoRecLen As Integer
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yyyy")
  ToPrint$ = Space$(80)
  Void$ = "* VOID * VOID * VOID * VOID * VOID * VOID * VOID * VOID * VOID * VO"
  MaxTopStub = 15  'detail lines on stub, 18 total lines
  PrintFile = FreeFile
  Open "APCHECK.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  DoStubHeader = True

  '--Don't change this loop
  For cnt = StartCnt To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    TChkAmt# = 0
    TopStubCnt = 0
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2 = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast
      Get APLedgerFile, InvList(Cnt2).LedgerRecNum, APLedgerRec(1)
      If Cnt2 = CHKinfo(cnt).ListFirst Then
        CHKinfo(cnt).StartChk = TCheckNum&
      End If
      GoSub PRINTChkInfo1        'go print some stuff
    Next
    CHKinfo(cnt).LastChk = TCheckNum&
    CHKinfo(cnt).ChkAmt = TChkAmt#
    CHKinfo(cnt).chkdate = CheckDate
    GoSub FinishChk1
  
  Next
  Print #PrintFile, Chr$(12) '--For Novell
  Close
  'KillFile "APCHKINF.DAT"
  ChkinfoFile = FreeFile
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For ccnt = StartCnt To VCnt
    
    Put ChkinfoFile, ccnt, CHKinfo(ccnt)
  Next
  Close ChkinfoFile
  
'  FPutAH "APCHKINF.DAT", Chkinfo(1), ChkInfoRecLen, VCnt
  ToPrint$ = ""
  Erase APLedgerRec, InvList, CHKinfo
  ActivateControls frmPrnAPChecks, True
  ViewPrnChks "APCHECK.PRN", DePrn, True
  Call MainLog("APChkPrint 1 Complete.")
  GoTo ExitCheckPrinting

PRINTChkInfo1:
  '--printing the stub detail lines here.
  If TopStubCnt >= MaxTopStub Then  '--if listing more invoices that will
    GoSub PrintVoidChk             '--fit on a stub void the check and
  End If                           '--contine on next check
  If DoStubHeader Then             '--check if we need to do a header
    GoSub StubHeader
  End If
  LSet ToPrint$ = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")   '--Invoice Date
  Mid$(ToPrint$, 13) = Left$(APLedgerRec(1).DOCNum, 14)       '--Invoice Numbe
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
   Mid$(ToPrint$, 28) = Left$(APLedgerRec(1).PONum, 10)        '--PO Number
  Else
   Mid$(ToPrint$, 28) = Left$(APLedgerRec(1).MPONum, 10)
  End If
  Mid$(ToPrint$, 40) = Left$(APLedgerRec(1).Comment, 23)        '--PO Number
  Mid$(ToPrint$, 65) = Using("###,###,###.##", Str$(APLedgerRec(1).Amt)) '-Amt
  Print #PrintFile, ToPrint$
  TChkAmt# = Round(TChkAmt# + APLedgerRec(1).Amt)
  TopStubCnt = TopStubCnt + 1
  DistRecord& = APLedgerRec(1).FrstDist
  TabStop = 1
  ShowDist$ = "N"
  If ShowDist$ = "Y" Then
   While DistRecord& > 0
   Get #APDistFile, DistRecord&, APDistRec(1)
   Print #PrintFile, Tab(TabStop); QPTrim$(APDistRec(1).DistAcctNum);
   Print #PrintFile, " "; Using("########.##", APDistRec(1).DistAmt);
   If TabStop = 1 Then
    TabStop = 40
    TopStubCnt = TopStubCnt + 1
    Else
    TabStop = 1
   End If
   DistRecord& = APDistRec(1).NextDist
   Wend
   Print #PrintFile, ""
  End If
  Return
FinishChk1:
 ' New Code Added to Page Up Check Upon Adding the Distribution
 If ShowDist$ = "Y" Then
  For CntZZ = TopStubCnt To (MaxTopStub + 3)
    Print #PrintFile, " "
  Next
 Else
  '--area from last detail line on stub to summary line
  For CntZZ = TopStubCnt To MaxTopStub
    Print #PrintFile, ""
  Next
  '--Stub summary line
  Print #PrintFile,
  LSet ToPrint$ = ""
  Mid$(ToPrint$, 10) = "Check - " + Using("########", TCheckNum&) + ", " + ChkPrn
  Mid$(ToPrint$, 40) = "Total Invoices:"
  Mid$(ToPrint$, 66) = Using("###,###,###.##", Str$(TChkAmt#))
  Print #PrintFile, ToPrint$
  Print #PrintFile, "~"   '"End of Stub Line"
 End If                 'Finish Distribution Check
  '-------body of check
  Print #PrintFile, "~"   '"Check Line1"
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, Tab(70); Using("########", TCheckNum&)
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  'Print #PrintFile,
  'Print #PrintFile, Tab(11); SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  'Print #PrintFile,
  Toolong$ = ""
  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  If Len(Toolong$) > 72 Then
    Here = InStr(Toolong$, " and ")
    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
  Else
    Print #PrintFile, Tab(11); QPTrim$(Toolong$)
    Print #PrintFile,
  End If

  Print #PrintFile,
  Print #PrintFile, Tab(52); ChkPrn; Tab(62); Using("$###,###,###.##", Str$(TChkAmt#))
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoName)
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoAddr)
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoAddr2)
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim$(Vendor.PaytoZip)
  Print #PrintFile,
  Print #PrintFile, Tab(5); QPTrim(Vendor.Memo)
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, "~" '"End of Form"
  LSet ToPrint$ = ""
  DoStubHeader = True
  TCheckNum& = TCheckNum& + 1
Return
PrintVoidChk:
  '--finish stub
  For CntZZ = TopStubCnt To 29
    Print #PrintFile, ""
  Next
  For VLCnt = 1 To 6 '--24 lines on check
    Print #PrintFile, Void$
  Next
  For VLCnt = 36 To 39
    Print #PrintFile, '"Finish Void Chk"; VCnt
  Next
  TopStubCnt = 0
  TCheckNum& = TCheckNum& + 1
  DoStubHeader = True
Return
StubHeader:
  '--number of lines from top of form to first invoice item
'HERE
  Print #PrintFile, "~" '"Top of Form"
  Print #PrintFile, "Date        Invoice        PO          Desc                             Amount"
  Print #PrintFile, String$(78, "-")
  TopStubCnt = 4
  DoStubHeader = False
Return
'*******************************************************
ExitCheckPrinting:

End Sub
'************************************** Check 2
Private Sub Check2prn(DePrn As String)
 '---New Std 10 cpi   (9013), Blank Top Stub- 42 Line
  Dim PrintFile As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, DistRecord As Long
  Dim APDistFile As Integer, NumDistRecs As Long, TabStop As Integer
  Dim DoStubHeader As Boolean, TChkAmt As Double, ShowDist As String
  Dim TopStubCnt As Integer, Cnt2 As Integer, RecLen As Integer, Here As Integer
  Dim CntZZ As Integer, VLCnt As Integer, cnt As Integer, ccnt As Integer
  Dim Void As String, MaxTopStub As Integer, ChkPrn As String, Toolong As String
  Dim ToPrint As String, APDistRecLen As Integer, ChkInfoRecLen As Integer
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yyyy")

  ToPrint$ = Space$(80)
  Void$ = "* VOID * VOID * VOID * VOID * VOID * VOID * VOID * VOID * VOID * VOid *"
  MaxTopStub = 18  'detail lines on stub, 18 total lines

  PrintFile = FreeFile
  Open "APCHECK.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen

  DoStubHeader = True

  '--Don't change this loop
  For cnt = StartCnt To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    TChkAmt# = 0
    TopStubCnt = 0
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2 = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast
      Get APLedgerFile, InvList(Cnt2).LedgerRecNum, APLedgerRec(1)
      If Cnt2 = CHKinfo(cnt).ListFirst Then
        CHKinfo(cnt).StartChk = TCheckNum&
      End If
      GoSub PRINTChkInfo        'go print some stuff
    Next
    CHKinfo(cnt).LastChk = TCheckNum&
    CHKinfo(cnt).ChkAmt = TChkAmt#
    CHKinfo(cnt).chkdate = CheckDate
    GoSub FinishChk
  Next

  Print #PrintFile, Chr$(12) '--For Novell

  Close
  'KillFile "APCHKINF.DAT"
  ChkinfoFile = FreeFile
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For ccnt = StartCnt To VCnt
    
    Put ChkinfoFile, ccnt, CHKinfo(ccnt)
  Next
  Close ChkinfoFile
  

  'FPutAH "APCHKINF.DAT", CHKinfo(1), ChkInfoRecLen, VCnt

  ToPrint$ = ""
  Erase APLedgerRec, CHKinfo, InvList
  ActivateControls frmPrnAPChecks, True
  ViewPrnChks "APCHECK.PRN", DePrn, True
  Call MainLog("APChkPrint 2 Complete.")
  GoTo ExitCheckPrinting

PRINTChkInfo:
  '--printing the stub detail lines here.
  If TopStubCnt = MaxTopStub Then  '--if listing more invoices that will
    GoSub PrintVoidChk             '--fit on a stub void the check and
  End If                           '--contine on next check
  If DoStubHeader Then             '--check if we need to do a header
    GoSub StubHeader
  End If
  LSet ToPrint$ = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")   '--Invoice Date
  Mid$(ToPrint$, 13) = Left$(QPTrim$(APLedgerRec(1).DOCNum), 14)       '--Invoice Numbe
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    Mid$(ToPrint$, 28) = Left$(QPTrim$(APLedgerRec(1).PONum), 10)        '--PO Number
  Else
    Mid$(ToPrint$, 28) = Left$(QPTrim$(APLedgerRec(1).MPONum), 10)
  End If
  Mid$(ToPrint$, 40) = Left$(QPTrim$(APLedgerRec(1).Comment), 23)        '--Inv Desc
  Mid$(ToPrint$, 66) = Using("###,###,###.##", Str$(APLedgerRec(1).Amt)) '-Amt
  Print #PrintFile, ToPrint$
  TChkAmt# = Round(TChkAmt# + APLedgerRec(1).Amt)
  TopStubCnt = TopStubCnt + 1
Return

FinishChk:
  '--area from last detail line on stub to summary line
  For CntZZ = TopStubCnt To MaxTopStub - 1
    Print #PrintFile, '"CntZZ:"; CntZZ
  Next

  '--Stub summary line
  Print #PrintFile,
  LSet ToPrint$ = ""
  Mid$(ToPrint$, 10) = "Check - " + Using("########", TCheckNum&) + ", " + ChkPrn
  Mid$(ToPrint$, 40) = "Total Invoices:"
  Mid$(ToPrint$, 66) = Using("###,###,###.##", Str$(TChkAmt#))
  Print #PrintFile, ToPrint$
  Print #PrintFile, '"End of Stub Line"

  '-------body of check
  Print #PrintFile, '"Check Line1"
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, Tab(70); Using("########", TCheckNum&)
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, Tab(52); ChkPrn; Tab(62); Using("$###,###,###.##", Str$(TChkAmt#))
  'Print #PrintFile, Tab(11); SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  'Print #PrintFile,
  Toolong$ = ""
  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  If Len(Toolong$) > 72 Then
    Here = InStr(Toolong$, " and ")
    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
  Else
    Print #PrintFile, Tab(11); QPTrim$(Toolong$)
    Print #PrintFile,
  End If

  Print #PrintFile,
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoName)
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoAddr)
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoAddr2)
  Print #PrintFile, Tab(11); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim(Vendor.PaytoZip)
  Print #PrintFile,
  Print #PrintFile, Tab(5); QPTrim(Vendor.Memo)
  Print #PrintFile,
  'Print #PrintFile,
  Print #PrintFile, "~" '"End of Form"

  LSet ToPrint$ = ""
  DoStubHeader = True
  TCheckNum& = TCheckNum& + 1
Return
PrintVoidChk:
  '--finish stub
  For VLCnt = 1 To 3
    Print #PrintFile, '"Finish Stub"; VCnt
  Next

  For VLCnt = 1 To 6 '--24 lines on check
    Print #PrintFile,
    Print #PrintFile, Void$
    Print #PrintFile,
  Next

  For VLCnt = 1 To 3
    Print #PrintFile, '"Finish Void Chk"; VCnt
  Next

  TopStubCnt = 0
  TCheckNum& = TCheckNum& + 1
  DoStubHeader = True
Return

StubHeader:
  '--number of lines from top of form to first invoice item
  Print #PrintFile, "~" '; TAB(30); QPTrim$(VENDOR.PayToName)
  Print #PrintFile, "Date        Invoice        PO          Desc                             Amount"
                   ' 1234567890123456789012345678901234567890123456789012345674567890123456789012345678901234567890
                   '          1         2         3         4         5      5         6         7         8

  Print #PrintFile, String$(78, "-")
  TopStubCnt = 3
  DoStubHeader = False
Return
ExitCheckPrinting:
End Sub
'*********************************** Check 3
Private Sub Check3prn(DePrn As String)
 '---Old Std   (9028), Old Citipak standard
  Dim PrintFile As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, DistRecord As Long
  Dim APDistFile As Integer, NumDistRecs As Long, TabStop As Integer
  Dim DoStubHeader As Boolean, TChkAmt As Double, ShowDist As String
  Dim TopStubCnt As Integer, Cnt2 As Integer, RecLen As Integer, Here As Integer
  Dim CntZZ As Integer, VLCnt As Integer, cnt As Integer, ccnt As Integer
  Dim Void As String, MaxTopStub As Integer, ChkPrn As String, Toolong As String
  Dim ToPrint As String, APDistRecLen As Integer, ChkInfoRecLen As Integer
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yyyy")
  ToPrint$ = Space$(80)

  MaxTopStub = 19        'actually 21

  PrintFile = FreeFile
  Open "APCHECK.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen

  DoStubHeader = True

  For cnt = StartCnt To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    TChkAmt# = 0
    TopStubCnt = 0
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2 = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast
      Get APLedgerFile, InvList(Cnt2).LedgerRecNum, APLedgerRec(1)
      If Cnt2 = CHKinfo(cnt).ListFirst Then
        CHKinfo(cnt).StartChk = TCheckNum&
      End If
      GoSub PRINTChkInfo        'go print some stuff
    Next
    CHKinfo(cnt).LastChk = TCheckNum&
    CHKinfo(cnt).ChkAmt = TChkAmt#
    CHKinfo(cnt).chkdate = CheckDate
    GoSub FinishChk
  Next
  Close

  'KillFile "APCHKINF.DAT"
  ChkinfoFile = FreeFile
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For ccnt = StartCnt To VCnt
    Put ChkinfoFile, ccnt, CHKinfo(ccnt)
  Next
  Close ChkinfoFile

  ToPrint$ = ""
  Erase APLedgerRec, CHKinfo, InvList
  ActivateControls frmPrnAPChecks, True
  ViewPrnChks "APCHECK.PRN", DePrn, True
  Call MainLog("APChkPrint 3 Complete.")
  GoTo ExitCheckPrinting

PRINTChkInfo:

  If TopStubCnt = MaxTopStub Then
    GoSub PrintVoidChk
  End If
  If DoStubHeader Then
    GoSub StubHeader
  End If

  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 6) = Left$(APLedgerRec(1).DOCNum, 11)
  Mid$(ToPrint$, 18) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 39) = Using("###,###.##", Str$(APLedgerRec(1).Amt))
  Mid$(ToPrint$, 71) = Using("###,###.##", Str$(APLedgerRec(1).Amt))
  Print #PrintFile, ToPrint$
  TChkAmt# = Round(TChkAmt# + APLedgerRec(1).Amt)
  TopStubCnt = TopStubCnt + 1

  Return

FinishChk:
  For CntZZ = TopStubCnt To MaxTopStub - 1
    Print #PrintFile,
  Next
  LSet ToPrint$ = ""
  'MID$(ToPrint$, 44) = "Total Amt: "
  Mid$(ToPrint$, 30) = "Check - " + Using("#######", TCheckNum&) + ", " + ChkPrn
  Mid$(ToPrint$, 71) = Using("###,###.##", Str$(TChkAmt#))
  Print #PrintFile, ToPrint$
  '-------body of check
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, Tab(50); Using("#######", TCheckNum&);
  Print #PrintFile, Tab(58); ChkPrn;
  Print #PrintFile, Tab(69); Vendor.vnum
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, Tab(62); Using("$###,###,###.##", Str$(TChkAmt#))
  Toolong$ = ""
  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  If Len(Toolong$) > 72 Then
    Here = InStr(Toolong$, " and ")
    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
  Else
    Print #PrintFile, Tab(8); QPTrim$(Toolong$)
    Print #PrintFile,
  End If
  Print #PrintFile, Tab(8); QPTrim$(Vendor.PaytoName)
  Print #PrintFile, Tab(8); QPTrim$(Vendor.PaytoAddr)
  Print #PrintFile, Tab(8); QPTrim$(Vendor.PaytoAddr2)
  Print #PrintFile, Tab(8); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim$(Vendor.PaytoZip)
  LSet ToPrint$ = ""
  Print #PrintFile,
  Print #PrintFile, Tab(5); QPTrim(Vendor.Memo)
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, "~"
  DoStubHeader = True
  TCheckNum& = TCheckNum& + 1

  Return

PrintVoidChk:
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""

  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  For CntZZ = 11 To MaxTopStub
    Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
  Next
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""

  TopStubCnt = 0
  TCheckNum& = TCheckNum& + 1
  DoStubHeader = True
  Return

StubHeader:
  Print #PrintFile, "~"
  Print #PrintFile, Tab(6); Vendor.VNAME
  Print #PrintFile,
  Print #PrintFile,
  TopStubCnt = 5
  DoStubHeader = False
  Return
ExitCheckPrinting:
End Sub
'*********************************Check 4
Private Sub Check4prn()
 '---Laser  Middle Check
  Dim PrintFile As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, DistRecord As Long
  Dim APDistFile As Integer, NumDistRecs As Long, TabStop As Integer
  Dim DoStubHeader As Boolean, TChkAmt As Double, ShowDist As String
  Dim TopStubCnt As Integer, Cnt2 As Integer, RecLen As Integer, Here As Integer
  Dim CntZZ As Integer, VLCnt As Integer, cnt As Integer, ccnt As Integer
  Dim Void As String, MaxTopStub As Integer, ChkPrn As String, Toolong As String
  Dim ToPrint As String, APDistRecLen As Integer, ChkInfoRecLen As Integer
  Dim BtmStubCnt As Integer, ChkLineCnt As Integer, CntBot As Integer
  Dim TPTST As String, TPCk As String, TPBST As String, TPTSH As String
  Dim TPTSG As String, TPBSH As String, TPBSG As String
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yyyy")
  ToPrint$ = Space$(78)

  MaxTopStub = 18               'actually 21

  ReDim BotStub$(1 To MaxTopStub)
  PrintFile = FreeFile
  Open "APCHECK.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen

  DoStubHeader = True


  For cnt = StartCnt To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    TChkAmt# = 0
    TopStubCnt = 0
    BtmStubCnt = 0
    ChkLineCnt = 0
    TPTSH$ = ""
    TPTSG$ = ""
    TPTST$ = ""
    TPCk$ = ""
    TPBSH$ = ""
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2 = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast
      Get APLedgerFile, InvList(Cnt2).LedgerRecNum, APLedgerRec(1)
      If Cnt2 = CHKinfo(cnt).ListFirst Then
        CHKinfo(cnt).StartChk = TCheckNum&
      End If
      GoSub PRINTChkInfo        'go print some stuff
    Next
    CHKinfo(cnt).LastChk = TCheckNum&
    CHKinfo(cnt).ChkAmt = TChkAmt#
    CHKinfo(cnt).chkdate = CheckDate
    GoSub FinishChk
    Print #PrintFile, TPTSH$ + "~" + TPTSG$ + TPTST$ + "~" + TPCk$ + "~" + TPBSH$ '+ "~" ' + TPBSG$ + "~" + TPBST$
  Next
  Close

  'KillFile "APCHKINF.DAT"
  ChkinfoFile = FreeFile
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For ccnt = StartCnt To VCnt
    Put ChkinfoFile, ccnt, CHKinfo(ccnt)
  Next
  Close ChkinfoFile

  ToPrint$ = ""
  Erase APLedgerRec, CHKinfo, InvList
  ActivateControls frmPrnAPChecks, True
  Load frmLoadingRpt
  'ViewPrnChks "APCHECK.PRN", DePrn, True
  ARptCheck4.GetName "APCHECK.PRN"
  ARptCheck4.startrpt
  Call MainLog("APChkPrint 4 Complete.")
  
  GoTo ExitCheckPrinting
Exit Sub
PRINTChkInfo:

  If TopStubCnt = MaxTopStub - 1 Then
    GoSub PrintVoidChk
  End If
  If DoStubHeader Then
    GoSub StubHeader
  End If

'  LSet ToPrint$ = "   " + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
'  Mid$(ToPrint$, 17) = APLedgerRec(1).DOCNum
'  Mid$(ToPrint$, 44) = APLedgerRec(1).PONum
'  Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(APLedgerRec(1).Amt))
'  Print #PrintFile, ToPrint$
  TPTSG$ = TPTSG$ + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  TPTSG$ = TPTSG$ + "~" + QPTrim(APLedgerRec(1).DOCNum) + "/" + QPTrim(APLedgerRec(1).Comment)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    TPTSG$ = TPTSG$ + "~" + QPTrim(APLedgerRec(1).PONum)
  Else
    TPTSG$ = TPTSG$ + "~" + QPTrim(APLedgerRec(1).MPONum)
  End If
  TPTSG$ = TPTSG$ + "~" + Using("###,###,###.##", Str$(APLedgerRec(1).Amt)) + "~"
  TChkAmt# = Round(TChkAmt# + APLedgerRec(1).Amt)
  TopStubCnt = TopStubCnt + 1
  BotStub$(TopStubCnt) = ToPrint$

  Return

FinishChk:
  For CntZZ = TopStubCnt To MaxTopStub - 2
  'THINK ABOUT IT - TO KEEP NUM OF LINES EXACT TOPSTUBCNT ALREADY AT LINE
  'THAT NEED TO ADD TO, IF ONLY SUBTRACT 1 HAVE 1 LINE TOO MANY
    'Print #PrintFile,
    TPTSG$ = TPTSG$ + " ~ ~ ~ ~"
  Next
  'LSet ToPrint$ = ""
  'Mid$(ToPrint$, 44) = "Total Amt: "
 ' Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(TChkAmt#))
  'Print #PrintFile, ToPrint$
  TPTST$ = "~~Total Amt: ~" + Using("###,###,###.##", Str$(TChkAmt#))
  '-------body of check
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile,
'  Print #PrintFile,
'  Print #PrintFile, Tab(72); Using("#######", TCheckNum&)
'  Print #PrintFile,
'  Print #PrintFile,
'  Print #PrintFile,
'  Print #PrintFile, Tab(52); ChkPrn; Tab(64); Using("$###,###,###.##", Str$(TChkAmt#))
  TPCk$ = Using("#######", TCheckNum&) + "~" + ChkPrn + "~" + Using("$###,###,###.##", Str$(TChkAmt#))
  Toolong$ = ""
  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  If Len(Toolong$) > 72 Then
    Here = InStr(Toolong$, " and ")
'    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
'    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
    TPCk$ = TPCk$ + "~" + Mid$(Toolong$, 1, Here - 1) + "~" + Mid$(Toolong$, Here + 1)
  Else
'    Print #PrintFile, Tab(8); QPTrim$(Toolong$)
'    Print #PrintFile,
     TPCk$ = TPCk$ + "~" + QPTrim(Toolong$) + " ~"
  End If
  
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoName)
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr)
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr2)
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim$(Vendor.PaytoZip)
  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PaytoName) + "~" + QPTrim$(Vendor.PaytoAddr)
  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PaytoAddr2)
  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PayToCity) + " " + QPTrim$(Vendor.PaytoState) + " " + QPTrim$(Vendor.PaytoZip)
  GoSub PrintBotStub
  LSet ToPrint$ = ""
'  Mid$(ToPrint$, 44) = "Total Amt: "
'  Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(TChkAmt#))
'  Print #PrintFile, ToPrint$
'  Print #PrintFile, Chr$(12)
  TPBST$ = "Total Amt: " + "~" + Using("###,###,###.##", Str$(TChkAmt#))
  DoStubHeader = True
  TCheckNum& = TCheckNum& + 1

  Return

PrintVoidChk:
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  For CntZZ = 11 To MaxTopStub
'    Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
'  Next
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  TPTSG$ = TPTSG$ + "~~~~~~~~~~~~~~~~~~~~~~~~"
'  For CntZZ = 11 To MaxTopStub
'    TPTSG$ = TPTSG$ + "VOID VOID VOID~VOID VOID VOID VOID~VOID VOID VOID VOID~VOID VOID VOID VOID"
'  Next
'  TPTSG$ = TPTSG$ + "~~~~~~~~~~~~~~~~~~~~"
  TPCk$ = "VOID VOID~VOID VOID~VOID VOID VOID~VOID VOID VOID VOID VOID VOID VOID VOID VOID VOID~"
  TPCk$ = TPCk$ + "VOID VOID VOID VOID VOID VOID VOID VOID VOID~VOID VOID VOID~VOID VOID VOID~"
  TPCk$ = TPCk$ + "VOID VOID VOID~VOID VOID VOID"
             '~~~~~~
  GoSub PrintBotStub
  'Print #PrintFile, Chr$(12)
  TopStubCnt = 0
  TCheckNum& = TCheckNum& + 1
  DoStubHeader = True
  TPTST$ = "~~Total Amt: ~Continued..."
  Print #PrintFile, TPTSH$ + "~" + TPTSG$ + TPTST$ + "~" + TPCk$ + "~" + TPBSH$
  TPTSH$ = ""
  TPTSG$ = ""
  TPTST$ = ""
  TPCk$ = ""
  TPBSH$ = ""
  Return

StubHeader:
'  LSet ToPrint$ = "   Date         Inv No.                    P.O. No.             Amt"
'  Print #PrintFile, ToPrint$
  TPTSH$ = Space(78)
  TPTSH$ = "Date~Inv No.\Desc.~P.O. No.~Amt"
  TopStubCnt = 1
  DoStubHeader = False
  Return

PrintBotStub:
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, Tab(40); "Vendor: " + Vendor.VNAME
'
'  LSet ToPrint$ = "   Date         Inv No.                    P.O. No.             Amt"
'  Print #PrintFile, ToPrint$
'  For CntBot = 4 To TopStubCnt
'    LSet ToPrint$ = BotStub$(CntBot)
'    Print #PrintFile, ToPrint$
'  Next
  TPBSH$ = "Vendor: " + QPTrim(Vendor.VNAME) + "~" + QPTrim$(Vendor.Memo) '+ "~Date~Inv No.~P.O. No.~Amt"
  TPBSG$ = TPTSG$
  Return
ExitCheckPrinting:
  
End Sub
'*********************************Check 5
'use exact format of chk4 all this does is creat file to send to
'active report which is different layout for top stub check...
Private Sub Check5prn()
  Dim PrintFile As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, DistRecord As Long
  Dim APDistFile As Integer, NumDistRecs As Long, TabStop As Integer
  Dim DoStubHeader As Boolean, TChkAmt As Double, ShowDist As String
  Dim TopStubCnt As Integer, Cnt2 As Integer, RecLen As Integer, Here As Integer
  Dim CntZZ As Integer, VLCnt As Integer, cnt As Integer, ccnt As Integer
  Dim Void As String, MaxTopStub As Integer, ChkPrn As String, Toolong As String
  Dim ToPrint As String, APDistRecLen As Integer, ChkInfoRecLen As Integer
  Dim BtmStubCnt As Integer, ChkLineCnt As Integer, CntBot As Integer
  Dim TPTST As String, TPCk As String, TPBST As String, TPTSH As String
  Dim TPTSG As String, TPBSH As String, TPBSG As String
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yyyy")
  ToPrint$ = Space$(78)

  MaxTopStub = 18               'actually 21

  ReDim BotStub$(1 To MaxTopStub)
  PrintFile = FreeFile
  Open "APCHECK.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen

  DoStubHeader = True


  For cnt = StartCnt To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    TChkAmt# = 0
    TopStubCnt = 0
    BtmStubCnt = 0
    ChkLineCnt = 0
    TPTSH$ = ""
    TPTSG$ = ""
    TPTST$ = ""
    TPCk$ = ""
    TPBSH$ = ""
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2 = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast
      Get APLedgerFile, InvList(Cnt2).LedgerRecNum, APLedgerRec(1)
      If Cnt2 = CHKinfo(cnt).ListFirst Then
        CHKinfo(cnt).StartChk = TCheckNum&
      End If
      GoSub PRINTChkInfo        'go print some stuff
    Next
    CHKinfo(cnt).LastChk = TCheckNum&
    CHKinfo(cnt).ChkAmt = TChkAmt#
    CHKinfo(cnt).chkdate = CheckDate
    GoSub FinishChk
    Print #PrintFile, TPTSH$ + "~" + TPTSG$ + TPTST$ + "~" + TPCk$ + "~" + TPBSH$ '+ "~" ' + TPBSG$ + "~" + TPBST$
  Next
  Close

  'KillFile "APCHKINF.DAT"
  ChkinfoFile = FreeFile
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For ccnt = StartCnt To VCnt
    Put ChkinfoFile, ccnt, CHKinfo(ccnt)
  Next
  Close ChkinfoFile

  ToPrint$ = ""
  Erase APLedgerRec, CHKinfo, InvList
  ActivateControls frmPrnAPChecks, True
  Load frmLoadingRpt
  'ViewPrnChks "APCHECK.PRN", DePrn, True
  ARptCheck5.GetName "APCHECK.PRN"
  ARptCheck5.startrpt
  Call MainLog("APChkPrint 5 Complete.")
  
  GoTo ExitCheckPrinting
Exit Sub
PRINTChkInfo:

  If TopStubCnt = MaxTopStub - 1 Then
    GoSub PrintVoidChk
  End If
  If DoStubHeader Then
    GoSub StubHeader
  End If

'  LSet ToPrint$ = "   " + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
'  Mid$(ToPrint$, 17) = APLedgerRec(1).DOCNum
'  Mid$(ToPrint$, 44) = APLedgerRec(1).PONum
'  Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(APLedgerRec(1).Amt))
'  Print #PrintFile, ToPrint$
  TPTSG$ = TPTSG$ + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  TPTSG$ = TPTSG$ + "~" + QPTrim(APLedgerRec(1).DOCNum) + "/" + QPTrim(APLedgerRec(1).Comment)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    TPTSG$ = TPTSG$ + "~" + QPTrim(APLedgerRec(1).PONum)
  Else
    TPTSG$ = TPTSG$ + "~" + QPTrim(APLedgerRec(1).MPONum)
  End If
  TPTSG$ = TPTSG$ + "~" + Using("###,###,###.##", Str$(APLedgerRec(1).Amt)) + "~"
  TChkAmt# = Round(TChkAmt# + APLedgerRec(1).Amt)
  TopStubCnt = TopStubCnt + 1
  BotStub$(TopStubCnt) = ToPrint$

  Return

FinishChk:
  For CntZZ = TopStubCnt To MaxTopStub - 2
  'THINK ABOUT IT - TO KEEP NUM OF LINES EXACT TOPSTUBCNT ALREADY AT LINE
  'THAT NEED TO ADD TO, IF ONLY SUBTRACT 1 HAVE 1 LINE TOO MANY
    'Print #PrintFile,
    TPTSG$ = TPTSG$ + " ~ ~ ~ ~"
  Next
  'LSet ToPrint$ = ""
  'Mid$(ToPrint$, 44) = "Total Amt: "
 ' Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(TChkAmt#))
  'Print #PrintFile, ToPrint$
  TPTST$ = "~~Total Amt: ~" + Using("###,###,###.##", Str$(TChkAmt#))
  '-------body of check
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile,
'  Print #PrintFile,
'  Print #PrintFile, Tab(72); Using("#######", TCheckNum&)
'  Print #PrintFile,
'  Print #PrintFile,
'  Print #PrintFile,
'  Print #PrintFile, Tab(52); ChkPrn; Tab(64); Using("$###,###,###.##", Str$(TChkAmt#))
  TPCk$ = Using("#######", TCheckNum&) + "~" + ChkPrn + "~" + Using("$###,###,###.##", Str$(TChkAmt#))
  Toolong$ = ""
  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  If Len(Toolong$) > 72 Then
    Here = InStr(Toolong$, " and ")
'    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
'    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
    TPCk$ = TPCk$ + "~" + Mid$(Toolong$, 1, Here - 1) + "~" + Mid$(Toolong$, Here + 1)
  Else
'    Print #PrintFile, Tab(8); QPTrim$(Toolong$)
'    Print #PrintFile,
     TPCk$ = TPCk$ + "~" + QPTrim(Toolong$) + " ~"
  End If
  
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoName)
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr)
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr2)
'  Print #PrintFile, Tab(12); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim$(Vendor.PaytoZip)
  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PaytoName) + "~" + QPTrim$(Vendor.PaytoAddr)
  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PaytoAddr2)
  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PayToCity) + " " + QPTrim$(Vendor.PaytoState) + " " + QPTrim$(Vendor.PaytoZip)
  GoSub PrintBotStub
  LSet ToPrint$ = ""
'  Mid$(ToPrint$, 44) = "Total Amt: "
'  Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(TChkAmt#))
'  Print #PrintFile, ToPrint$
'  Print #PrintFile, Chr$(12)
  TPBST$ = "Total Amt: " + "~" + Using("###,###,###.##", Str$(TChkAmt#))
  DoStubHeader = True
  TCheckNum& = TCheckNum& + 1

  Return

PrintVoidChk:
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  For CntZZ = 11 To MaxTopStub
'    Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
'  Next
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  TPTSG$ = TPTSG$ + "~~~~~~~~~~~~~~~~~~~~~~~~"
'  For CntZZ = 11 To MaxTopStub
'    TPTSG$ = TPTSG$ + "VOID VOID VOID~VOID VOID VOID VOID~VOID VOID VOID VOID~VOID VOID VOID VOID"
'  Next
'  TPTSG$ = TPTSG$ + "~~~~~~~~~~~~~~~~~~~~"
  TPCk$ = "VOID VOID~VOID VOID~VOID VOID VOID~VOID VOID VOID VOID VOID VOID VOID VOID VOID VOID~"
  TPCk$ = TPCk$ + "VOID VOID VOID VOID VOID VOID VOID VOID VOID~VOID VOID VOID~VOID VOID VOID~"
  TPCk$ = TPCk$ + "VOID VOID VOID~VOID VOID VOID"
             '~~~~~~
  GoSub PrintBotStub
  'Print #PrintFile, Chr$(12)
  TopStubCnt = 0
  TCheckNum& = TCheckNum& + 1
  DoStubHeader = True
  TPTST$ = "~~Total Amt: ~Continued..."
  Print #PrintFile, TPTSH$ + "~" + TPTSG$ + TPTST$ + "~" + TPCk$ + "~" + TPBSH$
  TPTSH$ = ""
  TPTSG$ = ""
  TPTST$ = ""
  TPCk$ = ""
  TPBSH$ = ""
  Return

StubHeader:
'  LSet ToPrint$ = "   Date         Inv No.                    P.O. No.             Amt"
'  Print #PrintFile, ToPrint$
  TPTSH$ = Space(78)
  TPTSH$ = "Date~Inv No.\Desc.~P.O. No.~Amt"
  TopStubCnt = 1
  DoStubHeader = False
  Return

PrintBotStub:
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, ""
'  Print #PrintFile, Tab(40); "Vendor: " + Vendor.VNAME
'
'  LSet ToPrint$ = "   Date         Inv No.                    P.O. No.             Amt"
'  Print #PrintFile, ToPrint$
'  For CntBot = 4 To TopStubCnt
'    LSet ToPrint$ = BotStub$(CntBot)
'    Print #PrintFile, ToPrint$
'  Next
  TPBSH$ = "Vendor: " + QPTrim(Vendor.VNAME) + "~" + QPTrim$(Vendor.Memo) '+ "~Date~Inv No.~P.O. No.~Amt"
  TPBSG$ = TPTSG$
  Return
ExitCheckPrinting:
  
End Sub
Private Sub Check6prn(DePrn As String)
  'Hamlet, NC
  Dim PrintFile As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, DistRecord As Long
  Dim APDistFile As Integer, NumDistRecs As Long, TabStop As Integer
  Dim DoStubHeader As Boolean, TChkAmt As Double, ShowDist As String
  Dim TopStubCnt As Integer, Cnt2 As Integer, RecLen As Integer, Here As Integer
  Dim CntZZ As Integer, VLCnt As Integer, cnt As Integer, ccnt As Integer
  Dim Void As String, MaxTopStub As Integer, ChkPrn As String, Toolong As String
  Dim ToPrint As String, APDistRecLen As Integer, ChkInfoRecLen As Integer
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yyyy")
  ToPrint$ = Space$(80)
  MaxTopStub = 19               'actually 21
  ReDim BotStub$(1 To MaxTopStub)
  PrintFile = FreeFile
  Open "APCHECK.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  DoStubHeader = True
  If StartCnt = 0 Then StartCnt = 1
  For cnt = StartCnt To VCnt
    FrmShowPctComp.ShowPctComp cnt, VCnt
    TChkAmt# = 0
    TopStubCnt = 0
    'BtmStubCnt = 0
   ' ChkLineCnt = 0
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2 = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast
          Get APLedgerFile, InvList(Cnt2).LedgerRecNum, APLedgerRec(1)
      If Cnt2 = CHKinfo(cnt).ListFirst Then
        CHKinfo(cnt).StartChk = TCheckNum&
      End If
      GoSub PRINTChkInfo        'go print some stuff
    Next
    CHKinfo(cnt).LastChk = TCheckNum&
    CHKinfo(cnt).ChkAmt = TChkAmt#
    CHKinfo(cnt).chkdate = CheckDate
    GoSub FinishChk
  Next
  Close
  
  ChkinfoFile = FreeFile
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For ccnt = StartCnt To VCnt
    
    Put ChkinfoFile, ccnt, CHKinfo(ccnt)
  Next
  Close ChkinfoFile

  ToPrint$ = ""
  Erase APLedgerRec, CHKinfo, InvList
  ActivateControls frmPrnAPChecks, True
  ViewPrnChks "APCHECK.PRN", DePrn, True
  Call MainLog("APChkPrint 6 Complete.")
  GoTo ExitCheckPrinting
PRINTChkInfo:

  If TopStubCnt = MaxTopStub Then
    GoSub PrintVoidChk
  End If
  If DoStubHeader Then
    GoSub StubHeader
  End If

  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 1) = QPTrim$(Left$(APLedgerRec(1).DOCNum, 7))

  'Parse Out Invoice Date to xx/xx/xx format
  'IDate$ = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  'IDate$ = Left$(IDate$, 6) + Right$(IDate$, 2)

  Mid$(ToPrint$, 9) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yy")
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    Mid$(ToPrint$, 26) = QPTrim(Left$(APLedgerRec(1).PONum, 8))
  Else
    Mid$(ToPrint$, 26) = QPTrim(Left$(APLedgerRec(1).MPONum, 8))
  End If

  Mid$(ToPrint$, 35) = QPTrim$(Left$(APLedgerRec(1).Comment, 30))
  Mid$(ToPrint$, 69) = Using("#,###,###.##", Str$(APLedgerRec(1).Amt))
  Print #PrintFile, ToPrint$
  TChkAmt# = Round(TChkAmt# + APLedgerRec(1).Amt)
  TopStubCnt = TopStubCnt + 1
    Return

FinishChk:
  For CntZZ = TopStubCnt To MaxTopStub
    Print #PrintFile,
  Next
  Print #PrintFile, Tab(7); "Total for Check #: "; Using("######", TCheckNum&);
  Print #PrintFile, Tab(69); Using("$###,###.##", Str$(TChkAmt#))
  Print #PrintFile,
  '-------body of check

    'Parse Out Check Date to xx/xx/xx form here
  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yy")

  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, Tab(53); Using("######", TCheckNum&);
  Print #PrintFile, Tab(61); ChkPrn;
  Print #PrintFile, Tab(69); Using("$###,###.##", Str$(TChkAmt#))
  Print #PrintFile,
  Print #PrintFile,
  Toolong$ = ""
  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  If Len(Toolong$) > 72 Then
    Here = InStr(Toolong$, " and ")
    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
  Else
    Print #PrintFile, Tab(11); QPTrim$(Toolong$)
    Print #PrintFile,
  End If
  'Print #PrintFile, Tab(10); SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
  'Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoName)
  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr)
  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr2)
  Print #PrintFile, Tab(12); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim$(Vendor.PaytoZip)
  Print #PrintFile,
  LSet ToPrint$ = ""
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  Print #PrintFile,
  DoStubHeader = True
  TCheckNum& = TCheckNum& + 1
  Return

PrintVoidChk:
  'Starts at Line 19
  Print #PrintFile, ""
  Print #PrintFile, Tab(7); "Total for Check #: "; Using("######", TCheckNum&);
  Print #PrintFile, Tab(53); "Continued on Next Check"
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
  Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
  Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
  Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
  Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""
  Print #PrintFile, ""

  TopStubCnt = 0
  TCheckNum& = TCheckNum& + 1
  DoStubHeader = True
  Return

StubHeader:
  Print #PrintFile, "~"; Tab(75); Using("######", TCheckNum&)
  Print #PrintFile, ""
  TopStubCnt = 3
  DoStubHeader = False
  Return
  
ExitCheckPrinting:
End Sub
'save this in case no work
''' '---Laser  Top Check
'''  Dim PrintFile As Integer, VendorFile As Integer, NumVRecs As Integer
'''  Dim APLedgerFile As Integer, NumTran As Long, DistRecord As Long
'''  Dim APDistFile As Integer, NumDistRecs As Long, TabStop As Integer
'''  Dim DoStubHeader As Boolean, TChkAmt As Double, ShowDist As String
'''  Dim TopStubCnt As Integer, Cnt2 As Integer, RecLen As Integer, Here As Integer
'''  Dim CntZZ As Integer, VLCnt As Integer, cnt As Integer, ccnt As Integer
'''  Dim Void As String, MaxTopStub As Integer, ChkPrn As String, Toolong As String
'''  Dim ToPrint As String, APDistRecLen As Integer, ChkInfoRecLen As Integer
'''  Dim BtmStubCnt As Integer, ChkLineCnt As Integer, CntBot As Integer
'''  Dim SCnt As Integer, LLL As Integer
'''  Dim TPTST As String, TPCk As String, TPBST As String, TPTSH As String
'''  Dim TPTSG As String, TPBSH As String, TPBSG As String
'''  ReDim APLedgerRec(1) As APLedger81RecType
'''  ReDim APDistRec(1) As APDistRecType
'''  APDistRecLen = Len(APDistRec(1))
'''  RecLen = Len(APLedgerRec(1))
'''  ChkInfoRecLen = Len(CHKinfo(1))
'''  ChkPrn = Format(DateAdd("d", (CheckDate), "12-31-1979"), "mm/dd/yyyy")
'''  ToPrint$ = Space$(78)
'''
'''  MaxTopStub = 18               'actually 21
'''
'''  ReDim BotStub$(1 To MaxTopStub)
'''  PrintFile = FreeFile
'''  Open "APCHECK.PRN" For Output As PrintFile
'''  OpenVendorFile VendorFile, NumVRecs
'''  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
'''
'''
'''  For cnt = StartCnt To VCnt
'''    FrmShowPctComp.ShowPctComp cnt, VCnt
'''    TChkAmt# = 0
'''    TopStubCnt = 0
'''    BtmStubCnt = 0
'''    ChkLineCnt = 0
'''    TPTSH$ = ""
'''    TPTSG$ = ""
'''    TPTST$ = ""
'''    TPCk$ = ""
'''    TPBSH$ = ""
'''
'''    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
'''    For Cnt2 = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast
'''      Get APLedgerFile, InvList(Cnt2).LedgerRecNum, APLedgerRec(1)
'''      If Cnt2 = CHKinfo(cnt).ListFirst Then
'''        CHKinfo(cnt).StartChk = TCheckNum&
'''      End If
'''      GoSub PRINTChkInfo        'go print some stuff
'''
'''    Next
'''    CHKinfo(cnt).LastChk = TCheckNum&
'''    CHKinfo(cnt).ChkAmt = TChkAmt#
'''    CHKinfo(cnt).ChkDate = CheckDate
'''    GoSub FinishChk
'''   Print #PrintFile, TPTSH$ + "~" + TPTSG$ + TPTST$ + "~" + TPCk$ + "~" + TPBSH$ '+ "~" ' + TPBSG$ + "~" + TPBST$
'''  Next
'''  Close
'''
'''  'KillFile "APCHKINF.DAT"
'''  ChkinfoFile = FreeFile
'''  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
'''  For ccnt = StartCnt To VCnt
'''    Put ChkinfoFile, ccnt, CHKinfo(ccnt)
'''  Next
'''  Close ChkinfoFile
'''
'''  ToPrint$ = ""
'''  Erase APLedgerRec, CHKinfo, InvList
'''  ActivateControls frmPrnAPChecks, True
'''  Load frmLoadingRpt
'''  'ViewPrnChks "APCHECK.PRN", DePrn, True
'''  ARptCheck5.GetName "APCHECK.PRN"
'''  ARptCheck5.startrpt
'''  Call MainLog("APChkPrint 5 Complete.")
'''  GoTo ExitCheckPrinting
'''PRINTChkInfo:
''''  LSet ToPrint$ = "   " + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
''''  Mid$(ToPrint$, 17) = APLedgerRec(1).DOCNum
''''  Mid$(ToPrint$, 44) = APLedgerRec(1).PONum
''''  Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(APLedgerRec(1).Amt))
'''  TPTSG$ = TPTSG$ + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
'''  TPTSG$ = TPTSG$ + "~" + QPTrim(APLedgerRec(1).DOCNum) + "~" + QPTrim(APLedgerRec(1).PONum)
'''  TPTSG$ = TPTSG$ + "~" + Using("###,###,###.##", Str$(APLedgerRec(1).Amt)) + "~"
'''  TChkAmt# = Round(TChkAmt# + APLedgerRec(1).Amt)
'''  TopStubCnt = TopStubCnt + 1
'''  BotStub$(TopStubCnt) = ToPrint$
'''  Return
'''FinishChk:
'''  For CntZZ = TopStubCnt To MaxTopStub - 2
'''  'THINK ABOUT IT - TO KEEP NUM OF LINES EXACT TOPSTUBCNT ALREADY AT LINE
'''  'THAT NEED TO ADD TO, IF ONLY SUBTRACT 1 HAVE 1 LINE TOO MANY
'''    'Print #PrintFile,
'''    TPTSG$ = TPTSG$ + " ~ ~ ~ ~"
'''  Next
'''  'LSet ToPrint$ = ""
'''  'Mid$(ToPrint$, 44) = "Total Amt: "
''' ' Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(TChkAmt#))
'''  'Print #PrintFile, ToPrint$
'''  TPTST$ = "~~Total Amt: ~" + Using("###,###,###.##", Str$(TChkAmt#))
'''  '-------body of check
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile,
''''  Print #PrintFile,
''''  Print #PrintFile, Tab(72); Using("#######", TCheckNum&)
''''  Print #PrintFile,
''''  Print #PrintFile,
''''  Print #PrintFile,
''''  Print #PrintFile, Tab(52); ChkPrn; Tab(64); Using("$###,###,###.##", Str$(TChkAmt#))
'''  TPCk$ = Using("#######", TCheckNum&) + "~" + ChkPrn + "~" + Using("$###,###,###.##", Str$(TChkAmt#))
'''  Toolong$ = ""
'''  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
'''  If Len(Toolong$) > 72 Then
'''    Here = InStr(Toolong$, " and ")
''''    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
''''    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
'''    TPCk$ = TPCk$ + "~" + Mid$(Toolong$, 1, Here - 1) + "~" + Mid$(Toolong$, Here + 1)
'''  Else
''''    Print #PrintFile, Tab(8); QPTrim$(Toolong$)
''''    Print #PrintFile,
'''     TPCk$ = TPCk$ + "~" + QPTrim(Toolong$) + " ~"
'''  End If
'''
''''  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoName)
''''  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr)
''''  Print #PrintFile, Tab(12); QPTrim$(Vendor.PaytoAddr2)
''''  Print #PrintFile, Tab(12); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim$(Vendor.PaytoZip)
'''  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PaytoName) + "~" + QPTrim$(Vendor.PaytoAddr)
'''  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PaytoAddr2)
'''  TPCk$ = TPCk$ + "~" + QPTrim$(Vendor.PayToCity) + " " + QPTrim$(Vendor.PaytoState) + " " + QPTrim$(Vendor.PaytoZip)
'''  GoSub PrintBotStub
'''  LSet ToPrint$ = ""
''''  Mid$(ToPrint$, 44) = "Total Amt: "
''''  Mid$(ToPrint$, 56) = Using("###,###,###.##", Str$(TChkAmt#))
''''  Print #PrintFile, ToPrint$
''''  Print #PrintFile, Chr$(12)
'''  TPBST$ = "Total Amt: " + "~" + Using("###,###,###.##", Str$(TChkAmt#))
''' ' DoStubHeader = True
'''  TCheckNum& = TCheckNum& + 1
'''
'''  Return
'''
'''''''FinishChk:
'''''''  '-------body of check
'''''''  Print #PrintFile, ""
'''''''  Print #PrintFile, ""
'''''''  Print #PrintFile, ""
'''''''  Print #PrintFile, ""
'''''''  Print #PrintFile,
'''''''  Toolong$ = ""
'''''''  Toolong$ = SpellNumber$(Using("#########.##", Str$(TChkAmt#)))
'''''''  If Len(Toolong$) > 72 Then
'''''''    Here = InStr(Toolong$, " and ")
'''''''    Print #PrintFile, Tab(8); Mid$(Toolong$, 1, Here - 1)
'''''''    Print #PrintFile, Tab(8); Mid$(Toolong$, Here + 1)
'''''''  Else
'''''''    Print #PrintFile, Tab(8); QPTrim$(Toolong$)
'''''''    Print #PrintFile,
'''''''  End If
'''''''  Print #PrintFile, Tab(50); ChkPrn; Tab(65); Using("###,###,###.##", Str$(TChkAmt#))
'''''''  Print #PrintFile, ""
'''''''  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoName)
'''''''  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoAddr)
'''''''  Print #PrintFile, Tab(11); QPTrim$(Vendor.PaytoAddr2)
'''''''  Print #PrintFile, Tab(11); QPTrim$(Vendor.PayToCity); " "; QPTrim$(Vendor.PaytoState); " "; QPTrim$(Vendor.PaytoZip)
'''''''  For SCnt = 17 To 24
'''''''    Print #PrintFile, ""
'''''''  Next SCnt
'''''''
'''''''  GoSub PrintBotStub
'''''''  GoSub PrintBotStub1    'Repeat for 2nd Bottom Stub
'''''''  Print #PrintFile, Chr$(12);
'''''''
'''''''  TCheckNum& = TCheckNum& + 1
'''''''  Return
'''
'''PrintVoidChk:
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  For CntZZ = 11 To MaxTopStub
''''    Print #PrintFile, "         VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  VOID  "
''''  Next
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
''''  Print #PrintFile, ""
'''  TPCk$ = "VOID VOID~VOID VOID~VOID VOID VOID~VOID VOID VOID VOID VOID VOID VOID VOID VOID VOID~"
'''  TPCk$ = TPCk$ + "VOID VOID VOID VOID VOID VOID VOID VOID VOID~VOID VOID VOID~VOID VOID VOID~"
'''  TPCk$ = TPCk$ + "VOID VOID VOID~VOID VOID VOID"
'''
'''  GoSub PrintBotStub
''''  Print #PrintFile, Chr$(12)
'''
'''  TopStubCnt = 0
'''  TCheckNum& = TCheckNum& + 1
'''  'DoStubHeader = True
'''  Return
'''PrintBotStub:
'''  If TopStubCnt > 20 Then TopStubCnt = 20
'''  TPTSH$ = Space(78)
'''  TPTSH$ = "Date~Inv No.~P.O. No.~Amt"
'''  TopStubCnt = 1
'''
'''
''' ' LSet ToPrint$ = "   Date         Inv No.                    P.O. No.             Amt"
'''  'Print #PrintFile, ToPrint$
''''  For CntBot = 1 To TopStubCnt
''''    'LSet ToPrint$ = BotStub$(CntBot)
''''    'Print #PrintFile, ToPrint$
''''  Next CntBot
''''  If TopStubCnt < 20 Then
''''   For LLL = 1 To 21 - CntBot
''''    'Print #PrintFile, ""
''''   Next LLL
''''  End If
'''  Return
''''PrintBotStub1:
''''  If TopStubCnt > 20 Then TopStubCnt = 16
''''  TPBSH$ = "Date~Inv No.~P.O. No.~Amt"
''''  'Print #PrintFile, ToPrint$
''''  For CntBot = 1 To TopStubCnt
''''    LSet ToPrint$ = BotStub$(CntBot)
''''    Print #PrintFile, ToPrint$
''''  Next CntBot
''''  Return
'''
'''ExitCheckPrinting:
'''
'''End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
