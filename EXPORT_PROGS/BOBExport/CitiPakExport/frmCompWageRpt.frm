VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmCompWageRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Annual Workers Compensation Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmCompWageRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8840
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4908
      Left            =   2160
      TabIndex        =   3
      Top             =   1962
      Width           =   7356
      _Version        =   196609
      _ExtentX        =   12975
      _ExtentY        =   8657
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmCompWageRpt.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3390
         TabIndex        =   2
         Top             =   3075
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
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
         Columns         =   0
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
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmCompWageRpt.frx":08E6
      End
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   348
         Left            =   3792
         TabIndex        =   0
         Top             =   1728
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   614
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
         ButtonStyle     =   2
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
         CaretInsert     =   0
         CaretOverWrite  =   3
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
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fptxtEnd 
         Height          =   348
         Left            =   3792
         TabIndex        =   1
         Top             =   2400
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   614
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
         ButtonStyle     =   2
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
         CaretInsert     =   0
         CaretOverWrite  =   3
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
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4200
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to process the 'Annual Workers Comp' report."
         Top             =   3840
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmCompWageRpt.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1320
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   3840
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmCompWageRpt.frx":0DF4
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Print Option:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1635
         TabIndex        =   7
         Top             =   3165
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1476
         Top             =   528
         Width           =   4572
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D0D0D0&
         Caption         =   "End Date:"
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
         Left            =   2112
         TabIndex        =   6
         Top             =   2448
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2016
         TabIndex        =   5
         Top             =   1872
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Annual Workers Comp Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   492
         Left            =   1620
         TabIndex        =   4
         Top             =   672
         Width           =   4284
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5232
      Left            =   1980
      Top             =   1818
      Width           =   7692
   End
End
Attribute VB_Name = "frmCompWageRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmCompWageRpt
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
  End If
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
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadCompWageRptScreen
  Me.HelpContextID = hlpAnnualWorkers
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub
Private Sub LoadCompWageRptScreen()
   Dim Today As String * 10
'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   fptxtStart.Text = "01-01-" + Mid(Today, 7, 4)
   fptxtEnd.Text = Today
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
End Sub
Private Sub PrintGraphics()
  Dim cnt As Integer
  Dim LowDate As Long, HiDate As Long
  Dim Image1$, Image2$, Image3$
  Dim MaxLines As Integer, LineCnt As Integer
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, RptTitle$, RptName$
  Dim RptFile As Integer, THandle As Integer
  Dim DHandle As Integer, RecNo As Long
  Dim UsingThisOne As Boolean, PayFreq$, AnnualGross#
  Dim PayType$, TransRecNum&, Page As Integer
  Dim EGross#, EOTWage#, EmpInfo$, CompCode$
  Dim TEGross#, TAnnualGross#, TEOTWage#, FF$
  Dim EmpIdxNNameHandle As Integer, x As Long
  Dim UnitHandle As Integer, Hrs#, UTemp$
  Dim Emp2Rec As EmpData2Type
  Dim Found As Long, OutOfOrder As Boolean
  Dim TempWIdxRec As CompSortType
  Dim dlm$
  Dim ThisCnt As Integer
  
  dlm$ = "~"
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  ReDim CompSort(1 To 1) As CompSortType
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  
  If HiDate < LowDate Then
    MsgBox "ERROR: The Start Date is after the End Date"
    fptxtStart.SetFocus
    Exit Sub
  End If
  
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  Image3$ = "#,###,##0.00"
  
  Dim Dash As String * 80
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  
  ReDim GrsRpt(1 To 3) As GrossWageRptType
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
    
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  NumOfRecs = LOF(EmpIdxNNameHandle) / 2
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get #EmpIdxNNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxNNameHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  RptTitle$ = "Annual Workers Comp Report"
  RptName$ = "PRRPTS\COMPWAGEG.RPT"
  
  RptFile = FreeFile
'  On Error GoTo ErrorHandler
  Open RptName$ For Output As RptFile
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize

  OpenEmpData2File DHandle
  
  GoSub MakeCompIndex
  
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    Get DHandle, CompSort(RecNo).RecNo, Emp2Rec

    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEmComp1
    End If

    PayFreq$ = QPTrim$(Emp2Rec.EMPPFREQ)
    Select Case UCase(PayFreq$)
    Case UCase("Weekly")
      PayFreq = 52
      Hrs# = 40
    Case UCase("Bi-Weekly")
      PayFreq = 26
      Hrs# = 80
    Case UCase("Semi-Monthly")
      PayFreq = 24
      Hrs# = 86.66
    Case UCase("Monthly")
      PayFreq = 12
      Hrs# = 173.33
    Case UCase("Quarterly")
      PayFreq = 4
      Hrs# = 520
    Case UCase("Semi-Annually")
      PayFreq = 2
      Hrs# = 1040
    Case UCase("Annually")
      PayFreq = 1
      Hrs# = 2080
    End Select
    AnnualGross# = 0
    If Emp2Rec.EMPPRATE < 0 Then
      Emp2Rec.EMPPRATE = 0
    End If

    PayType$ = UCase$(QPTrim$(Emp2Rec.EMPPTYPE))
    Select Case PayType$
    
    Case "HOURLY"
       AnnualGross# = OldRound#(Hrs# * Emp2Rec.EMPPRATE) * (0# + PayFreq)
    Case "SALARIED"
       AnnualGross# = OldRound#(Emp2Rec.EMPPRATE * (0# + PayFreq))
    End Select

    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      
      Select Case TransHRec(1).CheckDate
        
      Case LowDate To HiDate
        GoSub SumEmpGross
        UsingThisOne = True
      Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
        If UsingThisOne Then    'if there were any at all
          GoSub PrintEmpGrossTotals             'then write emp totals
        End If
        Exit Do
      Else      'there are more trans to check
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEmComp1:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close DHandle
  Close THandle
  
  Close RptFile
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  arCompWageRpt.Show
  frmLoadingRpt.Show
  MainLog ("Annual Worker's Comp Report processed.")
  
  Exit Sub

SumEmpGross:
  EGross# = OldRound#(EGross# + TransHRec(1).GrossPay)
  EOTWage# = OldRound#(EOTWage# + TransHRec(1).TotOTWage)
Return
  
PrintEmpGrossTotals:

  TEGross# = OldRound#(TEGross# + EGross#)
  TAnnualGross# = OldRound#(TAnnualGross# + AnnualGross#)
  TEOTWage# = OldRound#(TEOTWage# + EOTWage#)
  ThisCnt = ThisCnt + 1
  '                        0                              1
  Print #RptFile, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(LowDate) + " through " + MakeRegDate(HiDate); dlm;
  '                         2                                  3
  Print #RptFile, QPTrim$(Emp2Rec.EmpNo); dlm; QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName); dlm;
  '                         4                                  5                              6
  Print #RptFile, Using(Image2$, EGross#); dlm; Using(Image2$, AnnualGross#); dlm; Using(Image2$, EOTWage#); dlm;
  '                          7                                 8                                9
  Print #RptFile, QPTrim$(Emp2Rec.EMPWCCLS); dlm; Using(Image3$, TEGross#); dlm; Using(Image3$, TAnnualGross#); dlm;
  '                         10
  Print #RptFile, Using(Image3$, TEOTWage#)

  EGross# = 0
  EOTWage# = 0
  AnnualGross# = 0

Return

MakeCompIndex:
  ReDim CompSort(1 To NumOfRecs) As CompSortType
  
  Found = 1
  For RecNo = 1 To NumOfRecs
    Get DHandle, CLng(RecNo), Emp2Rec
    Found = Found + 1
    RSet CompSort(RecNo).CompCode = QPTrim$(Emp2Rec.EMPWCCLS)
    CompSort(RecNo).RecNo = RecNo
  Next
  
  FrmShowPctComp.Label1 = "Annual Workers Comp Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  Do  'Sort the employee Names
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumOfRecs - 1
      If Val(CompSort(x).CompCode) > Val(CompSort(x + 1).CompCode) Then
        LSet TempWIdxRec = CompSort(x)
        LSet CompSort(x) = CompSort(x + 1)
        LSet CompSort(x + 1) = TempWIdxRec
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder

Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmCompWageRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim cnt As Integer
  Dim LowDate As Long, HiDate As Long
  Dim Image1$, Image2$, Image3$
  Dim MaxLines As Integer, LineCnt As Integer
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, RptTitle$, RptName$
  Dim RptFile As Integer, THandle As Integer
  Dim DHandle As Integer, RecNo As Long
  Dim UsingThisOne As Boolean, PayFreq$, AnnualGross#
  Dim PayType$, TransRecNum&, Page As Integer
  Dim EGross#, EOTWage#, EmpInfo$, CompCode$
  Dim TEGross#, TAnnualGross#, TEOTWage#, FF$
  Dim EmpIdxNNameHandle As Integer, x As Long
  Dim UnitHandle As Integer, Hrs#, UTemp$
  Dim Emp2Rec As EmpData2Type
  Dim Found As Long, OutOfOrder As Boolean
  Dim TempWIdxRec As CompSortType
  Dim ThisCnt As Integer
  
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  FF$ = Chr(12)
  ReDim CompSort(1 To 1) As CompSortType
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  
  If HiDate < LowDate Then
    MsgBox "ERROR: The Start Date is after the End Date"
    fptxtStart.SetFocus
    Exit Sub
  End If
  
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  Image3$ = "#,###,##0.00"
  
  Dim Dash As String * 80
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  
  ReDim GrsRpt(1 To 3) As GrossWageRptType
  MaxLines = 57
  LineCnt = 0
  Dash = String$(80, "-")
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
    
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  NumOfRecs = LOF(EmpIdxNNameHandle) / 2
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get #EmpIdxNNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxNNameHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  RptTitle$ = "Annual Workers Comp Report"
  RptName$ = "PRRPTS\COMPWAGE.RPT"
  
  RptFile = FreeFile
  Open RptName$ For Output As RptFile
  RPTSetupPRN 17, RptFile
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize

  OpenEmpData2File DHandle
  
  GoSub MakeCompIndex
  GoSub PrintGrossHeader
  
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    Get DHandle, CompSort(RecNo).RecNo, Emp2Rec

    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEmComp1
    End If

    PayFreq$ = QPTrim$(Emp2Rec.EMPPFREQ)
    Select Case UCase(PayFreq$)
    Case UCase("Weekly")
      PayFreq = 52
      Hrs# = 40
    Case UCase("Bi-Weekly")
      PayFreq = 26
      Hrs# = 80
    Case UCase("Semi-Monthly")
      PayFreq = 24
      Hrs# = 86.66
    Case UCase("Monthly")
      PayFreq = 12
      Hrs# = 173.33
    Case UCase("Quarterly")
      PayFreq = 4
      Hrs# = 520
    Case UCase("Semi-Annually")
      PayFreq = 2
      Hrs# = 1040
    Case UCase("Annually")
      PayFreq = 1
      Hrs# = 2080
    End Select
    AnnualGross# = 0
    If Emp2Rec.EMPPRATE < 0 Then
      Emp2Rec.EMPPRATE = 0
    End If

    PayType$ = UCase$(QPTrim$(Emp2Rec.EMPPTYPE))
    Select Case PayType$
    
    Case "HOURLY"
       AnnualGross# = OldRound#(Hrs# * Emp2Rec.EMPPRATE) * (0# + PayFreq)
    Case "SALARIED"
       AnnualGross# = OldRound#(Emp2Rec.EMPPRATE * (0# + PayFreq))
    End Select

    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      
      Select Case TransHRec(1).CheckDate
        
      Case LowDate To HiDate
        GoSub SumEmpGross
        UsingThisOne = True
      Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
        If UsingThisOne Then    'if there were any at all
          GoSub PrintEmpGrossTotals             'then write emp totals
          If LineCnt > MaxLines Then            'bottom of page?
            Print #RptFile, FF$
            GoSub PrintGrossHeader              'write title lines
          End If
        End If
        Exit Do
      Else      'there are more trans to check
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEmComp1:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  GoSub PrintGrossGTotals

  Close DHandle
  Close THandle
  RPTSetupPRN 123, RptFile '8/15
  
  Close RptFile
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$
  
  MainLog ("Annual Worker's Comp Report processed.")
  
  Exit Sub
  
PrintGrossHeader:
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(80)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 65) = "Page:" + Pg(1)
  Print #RptFile, UTemp$
  Print #RptFile, "Annual Workers Comp Report"
  Print #RptFile, MakeRegDate(LowDate) + " thru " + MakeRegDate(HiDate); "           Fiscal Proj Annual      Fiscal       Comp"
  '               "01-01-0000 thru 01-01-0000
  Print #RptFile, "Emp #      Name                       Gross       Gross    Overtime       Code"
  Print #RptFile, Dash
  LineCnt = 5
  Return
  
SumEmpGross:
  EGross# = OldRound#(EGross# + TransHRec(1).GrossPay)
  EOTWage# = OldRound#(EOTWage# + TransHRec(1).TotOTWage)
Return
  
PrintEmpGrossTotals:
  EmpInfo$ = Space$(80)
  CompCode$ = Space$(10)
  LSet EmpInfo$ = QPTrim$(Emp2Rec.EmpNo)
  Mid$(EmpInfo$, 11) = QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
  Mid$(EmpInfo$, 34) = Using(Image2$, EGross#)
  Mid$(EmpInfo$, 46) = Using(Image2$, AnnualGross#)
  Mid$(EmpInfo$, 58) = Using(Image2$, EOTWage#)
  RSet CompCode$ = QPTrim$(Emp2Rec.EMPWCCLS)
  Mid$(EmpInfo$, 69) = CompCode$
  ThisCnt = ThisCnt + 1
  Print #RptFile, EmpInfo$
  
  LineCnt = LineCnt + 1

  TEGross# = OldRound#(TEGross# + EGross#)
  TAnnualGross# = OldRound#(TAnnualGross# + AnnualGross#)
  TEOTWage# = OldRound#(TEOTWage# + EOTWage#)

  EGross# = 0
  EOTWage# = 0
  AnnualGross# = 0
Return
  
PrintGrossGTotals:
  Print #RptFile, FF$
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(80)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 65) = "Page:" + Pg(1)
  Print #RptFile, UTemp$
  Print #RptFile, "Annual Workers Comp Report Summary"
  Print #RptFile, MakeRegDate(LowDate) + " thru " + MakeRegDate(HiDate)
  Print #RptFile, Dash
  Print #RptFile, " Fiscal Gross: "; Tab(22); Using(Image3$, TEGross#)
  Print #RptFile, " Proj Annual Gross: "; Tab(22); Using(Image3$, TAnnualGross#)
  Print #RptFile, " Fiscal Overtime: "; Tab(22); Using(Image3$, TEOTWage#)
  Print #RptFile, FF$
Return

MakeCompIndex:
  ReDim CompSort(1 To NumOfRecs) As CompSortType
  
  Found = 1
  For RecNo = 1 To NumOfRecs
    Get DHandle, CLng(RecNo), Emp2Rec
    Found = Found + 1
    RSet CompSort(RecNo).CompCode = QPTrim$(Emp2Rec.EMPWCCLS)
    CompSort(RecNo).RecNo = RecNo
  Next
  
  FrmShowPctComp.Label1 = "Annual Workers Comp Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  Do  'Sort the employee Names
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumOfRecs - 1
      If Val(CompSort(x).CompCode) > Val(CompSort(x + 1).CompCode) Then
        LSet TempWIdxRec = CompSort(x)
        LSet CompSort(x) = CompSort(x + 1)
        LSet CompSort(x + 1) = TempWIdxRec
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder

Return

End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdEscape.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub
