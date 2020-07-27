VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmGrossWageReport 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Gross Wage Report"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmGrossWageReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint2 
      Height          =   4860
      Left            =   2112
      TabIndex        =   3
      Top             =   2004
      Width           =   7404
      _Version        =   196609
      _ExtentX        =   13060
      _ExtentY        =   8572
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmGrossWageReport.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3555
         TabIndex        =   2
         Top             =   2925
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
         ColDesigner     =   "frmGrossWageReport.frx":08E6
      End
      Begin EditLib.fpDateTime fptxtStartDate 
         Height          =   348
         Left            =   3552
         TabIndex        =   0
         Top             =   1680
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
      Begin EditLib.fpDateTime fptxtEndDate 
         Height          =   348
         Left            =   3552
         TabIndex        =   1
         Top             =   2304
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
         Left            =   4224
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the employee gross wage report."
         Top             =   3792
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
         ButtonDesigner  =   "frmGrossWageReport.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1344
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   3792
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
         ButtonDesigner  =   "frmGrossWageReport.frx":0DBC
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
         Left            =   1680
         TabIndex        =   7
         Top             =   3030
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1392
         Top             =   432
         Width           =   4716
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   348
         Left            =   1200
         TabIndex        =   6
         Top             =   2352
         Width           =   2028
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   5
         Top             =   1776
         Width           =   2028
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Gross Wage Report"
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
         Height          =   444
         Left            =   1632
         TabIndex        =   4
         Top             =   624
         Width           =   4188
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5160
      Left            =   1980
      Top             =   1860
      Width           =   7692
   End
End
Attribute VB_Name = "frmGrossWageReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

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
  Call LoadUnitFile
  Me.HelpContextID = hlpGrossWage
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub PrintGraphics()
  
  Dim UsingThisOne As Boolean
  Dim DHandle As Integer
  Dim THandle As Integer
  Dim LowDate As Long
  Dim HiDate As Long
  Dim Image1$, EmpData2FileHandle As Integer
  Dim Image2$, RecNo As Integer
  Dim Image3$, x As Integer, UTemp$
  Dim TransRecNum&
  Dim EQWL1Len As Long
  Dim EmpRecSize As Long
  Dim TRecSize As Long, cnt As Integer
  Dim IdxRecLen As Integer
  Dim IdxFileSize&, RHandle As Integer
  Dim NumOfRecs As Long
  Dim EmpIdxLNameHandle As Integer
  Dim UnitHandle As Integer
  Dim City$, RptTitle$, RptName$
  Dim Emp2Rec As EmpData2Type
  Dim dlm$
  '-------------------01/04------------------
  Dim FEDTAX As FederalTaxRecType
  Dim FedTaxHandle As Integer
  Dim FedSSMax As Double
  Dim ThisDate As Integer
  Dim TotSocGross As Double
  Dim BegDate$
  Dim SSMaxReachedFlag As Boolean
  Dim ThisDif As Double
  Dim NextDate As Integer
  Dim StopDate As Integer
  Dim SSMaxCode As Integer
  Dim StopDateFlag As Boolean
  Dim z As Integer
  Dim SocGross#
  Dim TSocGross#
  Dim SSTotal As Double
  Dim SOCMat As Double
  Dim ThisCnt As Integer
  
  OpenFedTaxFile FedTaxHandle
  Get FedTaxHandle, 1, FEDTAX
  Close FedTaxHandle
  FedSSMax = FEDTAX.FTMSSMW
  dlm$ = "~"
  
  If CheckValDate(fptxtStartDate.Text) = False Then
    MsgBox "The Start Date is not valid"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  
  If CheckValDate(fptxtEndDate.Text) = False Then
     MsgBox "The End Date is not valid"
     fptxtEndDate.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStartDate.Text)
  HiDate = Date2Num(fptxtEndDate.Text)
  
  If LowDate > HiDate Then
     MsgBox "The Start Date is later than the End Date"
     fptxtStartDate.SetFocus
     Exit Sub
  End If
  
  If fptxtStartDate.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStartDate.SetFocus
     GoTo EndTran
  End If

  If fptxtEndDate.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEndDate.SetFocus
     GoTo EndTran
  End If
  
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  Image3$ = "######0.00"
  
  ReDim DashLine(1) As String * 133
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  
  ReDim GrsRpt(1 To 3) As GrossWageRptType
  
  ReDim EQWL1(1) As EQWRptRecType
  
  EQWL1Len = Len(EQWL1(1))
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If

  FrmShowPctComp.Label1 = "Gross Wage Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  City = QPTrim$(Unit(1).UFEMPR)
  Close UnitHandle
   
  RptTitle$ = "Employee Gross Wage Report"
  RptName$ = "PRRPTS\GROSWAGEG.RPT"
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  OpenEmpData3File EHandle
  
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm4
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
'---------------New as of 01/12/04----------------
      SSMaxReachedFlag = False
      ThisDif = 0
      SSMaxCode = 1

      If Mid(fptxtStartDate.Text, 7, 4) = Mid(fptxtEndDate.Text, 7, 4) Then
        BegDate = ("01/01/" + Mid(fptxtEndDate.Text, 7, 4))
        ThisDate = Date2Num(BegDate)
        TotSocGross = 0
        Do
          Get THandle, TransRecNum&, TransHRec(1)
          If (TransHRec(1).CheckDate >= ThisDate) And (TransHRec(1).CheckDate <= HiDate) Then
            TotSocGross = OldRound(TotSocGross + TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
          End If
          If TransHRec(1).PrevTransRec > 0 Then
            TransRecNum& = TransHRec(1).PrevTransRec
          Else
            Exit Do
          End If
        Loop
        If TotSocGross > FedSSMax Then
          SSMaxReachedFlag = True
          GoSub MaxSSWage
        End If
      End If
      TransRecNum& = Emp2Rec.LastTransRec
'---------------New as of 01/12/04----------------
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
SkipEm4:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(LowDate); dlm; MakeRegDate(HiDate); dlm; ""; dlm;
  Print #RHandle, ""; dlm; ""; dlm;
  Print #RHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  Print #RHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  '                     14                          15                     16                      17
  Print #RHandle, GrsRpt(3).GrossPay; dlm; GrsRpt(3).FedGross; dlm; GrsRpt(3).FEDTAX; dlm; GrsRpt(3).SocGross; dlm;
  '                     18                       19                     20                   21                     22               23
  Print #RHandle, GrsRpt(3).SOCTAX; dlm; GrsRpt(3).MedGross; dlm; GrsRpt(3).MEDTAX; dlm; GrsRpt(3).EIC; dlm; GrsRpt(3).STATAX; dlm; "END"
  Close DHandle
  Close THandle
  Close RHandle
  
  If ThisCnt = 0 Then
    EnableCloseButton Me.hwnd, True
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  arGrossWage.Show
  frmLoadingRpt.Show
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Gross Wage Report was processed.")
  
  Exit Sub
  
SumEmpGross:
  ThisCnt = ThisCnt + 1
'  SocMat = OldRound(SocMat + TransHRec(1).MatchSocAmt)
  GrsRpt(1).GrossPay = OldRound(GrsRpt(1).GrossPay + TransHRec(1).GrossPay)
  GrsRpt(1).FedGross = OldRound(GrsRpt(1).FedGross + TransHRec(1).FedGrossPay)
  GrsRpt(1).FEDTAX = OldRound(GrsRpt(1).FEDTAX + TransHRec(1).FedTaxAmt)
'  GrsRpt(1).SocGross = OldRound(GrsRpt(1).SocGross + TransHRec(1).SocGrossPay)
  GrsRpt(1).SOCTAX = OldRound(GrsRpt(1).SOCTAX + TransHRec(1).SocTaxAmt)
  GrsRpt(1).MedGross = OldRound(GrsRpt(1).MedGross + TransHRec(1).MedGrossPay)
  GrsRpt(1).MEDTAX = OldRound(GrsRpt(1).MEDTAX + TransHRec(1).MedTaxAmt)
  GrsRpt(1).EIC = OldRound(GrsRpt(1).EIC + TransHRec(1).EICAmt)
  GrsRpt(1).STATAX = OldRound(GrsRpt(1).STATAX + TransHRec(1).StaTaxAmt)
  If TransHRec(1).TaxFring <> 0 Then 'added 7/22/03
    GrsRpt(1).GrossPay = OldRound(GrsRpt(1).GrossPay + TransHRec(1).TaxFring) '12/30/04
    GrsRpt(1).FedGross = OldRound(GrsRpt(1).FedGross + TransHRec(1).TaxFring)
'    GrsRpt(1).SocGross = OldRound(GrsRpt(1).SocGross + TransHRec(1).TaxFring)
    GrsRpt(1).MedGross = OldRound(GrsRpt(1).MedGross + TransHRec(1).TaxFring)
  End If
  '------------New as of 01/12/2004-----------------
  If SSMaxReachedFlag = False Then 'SSMaxReachedFlag is set
  'as a new employee is introduced to the process
    GrsRpt(1).SocGross = OldRound(GrsRpt(1).SocGross + TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
    GrsRpt(3).SocGross = OldRound#(GrsRpt(3).SocGross + TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
  End If
  
  '------------New as of 01/12/2004-----------------

Return

PrintEmpGrossTotals:
  
  'sum to quarter grand totals
  GrsRpt(3).GrossPay = OldRound(GrsRpt(3).GrossPay + GrsRpt(1).GrossPay)
  GrsRpt(3).FedGross = OldRound(GrsRpt(3).FedGross + GrsRpt(1).FedGross)
  GrsRpt(3).FEDTAX = OldRound(GrsRpt(3).FEDTAX + GrsRpt(1).FEDTAX)
' GrsRpt(3).SocGross is now being collected where GrsRpt(1).SocGross is being collected
'  GrsRpt(3).SocGross = OldRound(GrsRpt(3).SocGross + GrsRpt(1).SocGross)
  GrsRpt(3).SOCTAX = OldRound(GrsRpt(3).SOCTAX + GrsRpt(1).SOCTAX)
  GrsRpt(3).MedGross = OldRound(GrsRpt(3).MedGross + GrsRpt(1).MedGross)
  GrsRpt(3).MEDTAX = OldRound(GrsRpt(3).MEDTAX + GrsRpt(1).MEDTAX)
  GrsRpt(3).EIC = OldRound(GrsRpt(3).EIC + GrsRpt(1).EIC)
  GrsRpt(3).STATAX = OldRound(GrsRpt(3).STATAX + GrsRpt(1).STATAX)
  
  '                          0                         1                            2                        3
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(LowDate); dlm; MakeRegDate(HiDate); dlm; QPTrim$(Emp2Rec.EmpNo); dlm;
  '                          4                                                                5
  Print #RHandle, QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName); dlm; GrsRpt(1).GrossPay; dlm;
  '                      6                        7                       8                        9
  Print #RHandle, GrsRpt(1).FedGross; dlm; GrsRpt(1).FEDTAX; dlm; GrsRpt(1).SocGross; dlm; GrsRpt(1).SOCTAX; dlm;
  '                       10                      11                   12                    13
  Print #RHandle, GrsRpt(1).MedGross; dlm; GrsRpt(1).MEDTAX; dlm; GrsRpt(1).EIC; dlm; GrsRpt(1).STATAX; dlm;
  '                     14                          15                     16                      17
  Print #RHandle, GrsRpt(3).GrossPay; dlm; GrsRpt(3).FedGross; dlm; GrsRpt(3).FEDTAX; dlm; GrsRpt(3).SocGross; dlm;
  '                     18                       19                     20                   21                     22               23
  Print #RHandle, GrsRpt(3).SOCTAX; dlm; GrsRpt(3).MedGross; dlm; GrsRpt(3).MEDTAX; dlm; GrsRpt(3).EIC; dlm; GrsRpt(3).STATAX; dlm; "OK"
  'clear for the next transaction record
  GrsRpt(1) = GrsRpt(2)
  
  Return
  
MaxSSWage:
  
  StopDateFlag = False
  TransRecNum& = Emp2Rec.LastTransRec
  ReDim SSTempDates(1 To 1) As Integer
  ReDim SSTempAmts(1 To 1) As Double
  Do
    Get THandle, TransRecNum&, TransHRec(1)
    If (TransHRec(1).CheckDate >= ThisDate) And (TransHRec(1).CheckDate <= HiDate) Then
      NextDate = NextDate + 1
      ReDim Preserve SSTempDates(1 To NextDate) As Integer
      ReDim Preserve SSTempAmts(1 To NextDate) As Double
      SSTempDates(NextDate) = TransHRec(1).CheckDate
      SSTempAmts(NextDate) = OldRound(TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
    End If
    If TransHRec(1).PrevTransRec > 0 Then
      TransRecNum& = TransHRec(1).PrevTransRec
    Else
      Exit Do
    End If
  Loop
    
  ReDim SSAmts(1 To NextDate) As Double
  ReDim SSDates(1 To NextDate) As Integer
  'assign dates/amts in forward order
  z = NextDate
  For x = 1 To NextDate
    SSDates(x) = SSTempDates(z)
    SSAmts(x) = SSTempAmts(z)
    z = z - 1
  Next x
  
  TotSocGross = 0
    
  For x = 1 To NextDate
    TotSocGross = OldRound(TotSocGross + SSAmts(x))
    If SSDates(x) >= LowDate Then SSTotal = SSTotal + SSAmts(x)
    If TotSocGross > FedSSMax Then
      If StopDateFlag = False Then
        StopDateFlag = True
        StopDate = SSDates(x)
        If NextDate > 1 Then
          ThisDif = OldRound(TotSocGross - FedSSMax)
          Exit For
        Else
          SSMaxCode = 4
          Exit For
        End If
      End If
    End If
  Next x
  
  If SSMaxCode = 4 Then
    GrsRpt(1).SocGross = FedSSMax
    GrsRpt(3).SocGross = GrsRpt(3).SocGross + FedSSMax
    Return
  End If
  
  If StopDate < LowDate Then
    GrsRpt(1).SocGross = 0
    GrsRpt(3).SocGross = GrsRpt(3).SocGross
  Else
    GrsRpt(1).SocGross = SSTotal - ThisDif
    GrsRpt(3).SocGross = GrsRpt(3).SocGross + GrsRpt(1).SocGross
  End If
  
  Return
  
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."


EndTran:
End Sub

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmGrossWageReport
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
Private Sub LoadUnitFile()
   Dim Today As String * 10
   Today = Date '$
   fptxtEndDate.Text = Today
   fptxtStartDate.Text = "01-01-" + Mid(Today, 7, 4)
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmGrossWageReport.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim UsingThisOne As Boolean
  Dim DHandle As Integer
  Dim THandle As Integer
  Dim LowDate As Long
  Dim HiDate As Long
  Dim Image1$, EmpData2FileHandle As Integer
  Dim Image2$, RecNo As Integer
  Dim Image3$, x As Integer, UTemp$
  Dim MaxLines As Integer, TransRecNum&
  Dim LineCnt As Integer, Page As Integer
  Dim EQWL1Len As Long
  Dim EmpRecSize As Long, FF$
  Dim TRecSize As Long, cnt As Integer
  Dim IdxRecLen As Integer
  Dim IdxFileSize&, RHandle As Integer
  Dim NumOfRecs As Long
  Dim EmpIdxLNameHandle As Integer
  Dim UnitHandle As Integer
  Dim City$, RptTitle$, RptName$
  Dim Emp2Rec As EmpData2Type
  '-------------------01/04------------------
  Dim FEDTAX As FederalTaxRecType
  Dim FedTaxHandle As Integer
  Dim FedSSMax As Double
  Dim ThisDate As Integer
  Dim TotSocGross As Double
  Dim BegDate$
  Dim SSMaxReachedFlag As Boolean
  Dim ThisDif As Double
  Dim NextDate As Integer
  Dim StopDate As Integer
  Dim SSMaxCode As Integer
  Dim StopDateFlag As Boolean
  Dim z As Integer
  Dim SocGross#
  Dim TSocGross#
  Dim SSTotal As Double
  Dim ThisCnt As Integer
  
  OpenFedTaxFile FedTaxHandle
  Get FedTaxHandle, 1, FEDTAX
  Close FedTaxHandle
  FedSSMax = FEDTAX.FTMSSMW
  
  FF$ = Chr$(12)
  
  If CheckValDate(fptxtStartDate.Text) = False Then
    MsgBox "The Start Date is not valid"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  
  If CheckValDate(fptxtEndDate.Text) = False Then
     MsgBox "The End Date is not valid"
     fptxtEndDate.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStartDate.Text)
  HiDate = Date2Num(fptxtEndDate.Text)
  
  If LowDate > HiDate Then
     MsgBox "The Start Date is later than the End Date"
     fptxtStartDate.SetFocus
     Exit Sub
  End If
  
  If fptxtStartDate.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStartDate.SetFocus
     GoTo EndTran
  End If

  If fptxtEndDate.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEndDate.SetFocus
     GoTo EndTran
  End If
  
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  Image3$ = "######0.00"
  
  ReDim DashLine(1) As String * 133
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  
  ReDim GrsRpt(1 To 3) As GrossWageRptType
  
  ReDim EQWL1(1) As EQWRptRecType
  
  MaxLines = 57
  LineCnt = 0
  DashLine(1) = String$(131, "-")
  EQWL1Len = Len(EQWL1(1))
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If

  FrmShowPctComp.Label1 = "Gross Wage Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  City = QPTrim$(Unit(1).UFEMPR)
  Close UnitHandle
   
  RptTitle$ = "Employee Gross Wage Report"
  RptName$ = "PRRPTS\GROSWAGE.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 5, RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  
  GoSub PrintGrossHeader
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm4
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
'---------------New as of 01/12/04----------------
      SSMaxReachedFlag = False
      ThisDif = 0
      SSMaxCode = 1

      If Mid(fptxtStartDate.Text, 7, 4) = Mid(fptxtEndDate.Text, 7, 4) Then
        BegDate = ("01/01/" + Mid(fptxtEndDate.Text, 7, 4))
        ThisDate = Date2Num(BegDate)
        TotSocGross = 0
        Do
          Get THandle, TransRecNum&, TransHRec(1)
          If (TransHRec(1).CheckDate >= ThisDate) And (TransHRec(1).CheckDate <= HiDate) Then
            TotSocGross = OldRound(TotSocGross + TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
          End If
          If TransHRec(1).PrevTransRec > 0 Then
            TransRecNum& = TransHRec(1).PrevTransRec
          Else
            Exit Do
          End If
        Loop
        If TotSocGross > FedSSMax Then
          SSMaxReachedFlag = True
          GoSub MaxSSWage
        End If
      End If
      TransRecNum& = Emp2Rec.LastTransRec
'---------------New as of 01/12/04----------------
    
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
            Print #RHandle, FF$
            GoSub PrintGrossHeader              'write title lines
          End If
        End If
        Exit Do
      Else      'there are more trans to check
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEm4:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  GoSub PrintGrossGTotals
  Close DHandle
  Close THandle
  RPTSetupPRN 123, RHandle
  Close RHandle
  
  If ThisCnt = 0 Then
    EnableCloseButton Me.hwnd, True
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$, True
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Gross Wage Report was processed.")
  Exit Sub
  
PrintGrossHeader:
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(133)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 124) = "Page:" + Pg(1)
  
  Print #RHandle, UTemp$
  Print #RHandle, "Gross Wage Report"
  Print #RHandle, MakeRegDate(LowDate) + " thru " + MakeRegDate(HiDate)
  Print #RHandle, "Emp #      Name                   Gross Pay   FedGross     FedTax   SocGross     SocTax   MedGross     MedTax        EIC     StaTax" ' + CrLf$
  Print #RHandle, DashLine(1)
  LineCnt = 5
  Return
  
SumEmpGross:
  ThisCnt = ThisCnt + 1
  GrsRpt(1).GrossPay = OldRound(GrsRpt(1).GrossPay + TransHRec(1).GrossPay)
  GrsRpt(1).FedGross = OldRound(GrsRpt(1).FedGross + TransHRec(1).FedGrossPay)
  GrsRpt(1).FEDTAX = OldRound(GrsRpt(1).FEDTAX + TransHRec(1).FedTaxAmt)
'  GrsRpt(1).SocGross = OldRound(GrsRpt(1).SocGross + TransHRec(1).SocGrossPay)
  GrsRpt(1).SOCTAX = OldRound(GrsRpt(1).SOCTAX + TransHRec(1).SocTaxAmt)
  GrsRpt(1).MedGross = OldRound(GrsRpt(1).MedGross + TransHRec(1).MedGrossPay)
  GrsRpt(1).MEDTAX = OldRound(GrsRpt(1).MEDTAX + TransHRec(1).MedTaxAmt)
  GrsRpt(1).EIC = OldRound(GrsRpt(1).EIC + TransHRec(1).EICAmt)
  GrsRpt(1).STATAX = OldRound(GrsRpt(1).STATAX + TransHRec(1).StaTaxAmt)
  If TransHRec(1).TaxFring <> 0 Then 'added 7/22/03
    GrsRpt(1).GrossPay = OldRound(GrsRpt(1).GrossPay + TransHRec(1).TaxFring) '12/30/04
    GrsRpt(1).FedGross = OldRound(GrsRpt(1).FedGross + TransHRec(1).TaxFring)
'    GrsRpt(1).SocGross = OldRound(GrsRpt(1).SocGross + TransHRec(1).TaxFring)
    GrsRpt(1).MedGross = OldRound(GrsRpt(1).MedGross + TransHRec(1).TaxFring)
  End If
  '------------New as of 01/12/2004-----------------
  If SSMaxReachedFlag = False Then 'SSMaxReachedFlag is set
  'as a new employee is introduced to the process
    GrsRpt(1).SocGross = OldRound#(GrsRpt(1).SocGross + TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
    GrsRpt(3).SocGross = OldRound#(GrsRpt(3).SocGross + TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
  End If
  '------------New as of 01/12/2004-----------------

Return

PrintEmpGrossTotals:
  LSet EQWL1(1).ENumb = QPTrim$(Emp2Rec.EmpNo)
  LSet EQWL1(1).EName = QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
  RSet EQWL1(1).GPay = Using(Image3$, GrsRpt(1).GrossPay)
  RSet EQWL1(1).FedGr = Using(Image3$, GrsRpt(1).FedGross)
  RSet EQWL1(1).FedTx = Using(Image3$, GrsRpt(1).FEDTAX)
  RSet EQWL1(1).SocGr = Using(Image3$, GrsRpt(1).SocGross)
  RSet EQWL1(1).SocTx = Using(Image3$, GrsRpt(1).SOCTAX)
  RSet EQWL1(1).MedGr = Using(Image3$, GrsRpt(1).MedGross)
  RSet EQWL1(1).MedTx = Using(Image3$, GrsRpt(1).MEDTAX)
  RSet EQWL1(1).EIC = Using(Image3$, GrsRpt(1).EIC)
  RSet EQWL1(1).StaTx = Using(Image3$, GrsRpt(1).STATAX)
  
  Print #RHandle, EQWL1(1).ENumb; EQWL1(1).EName; EQWL1(1).GPay;
  Print #RHandle, EQWL1(1).FedGr; EQWL1(1).FedTx; EQWL1(1).SocGr;
  Print #RHandle, EQWL1(1).SocTx; EQWL1(1).MedGr; EQWL1(1).MedTx;
  Print #RHandle, EQWL1(1).EIC; ; EQWL1(1).StaTx
  LineCnt = LineCnt + 1
  
  'sum to quarter grand totals
  GrsRpt(3).GrossPay = OldRound(GrsRpt(3).GrossPay + GrsRpt(1).GrossPay)
  GrsRpt(3).FedGross = OldRound(GrsRpt(3).FedGross + GrsRpt(1).FedGross)
  GrsRpt(3).FEDTAX = OldRound(GrsRpt(3).FEDTAX + GrsRpt(1).FEDTAX)
' GrsRpt(3).SocGross is now being collected where GrsRpt(1).SocGross is being collected
'  GrsRpt(3).SocGross = OldRound(GrsRpt(3).SocGross + GrsRpt(1).SocGross)
  GrsRpt(3).SOCTAX = OldRound(GrsRpt(3).SOCTAX + GrsRpt(1).SOCTAX)
  GrsRpt(3).MedGross = OldRound(GrsRpt(3).MedGross + GrsRpt(1).MedGross)
  GrsRpt(3).MEDTAX = OldRound(GrsRpt(3).MEDTAX + GrsRpt(1).MEDTAX)
  GrsRpt(3).EIC = OldRound(GrsRpt(3).EIC + GrsRpt(1).EIC)
  GrsRpt(3).STATAX = OldRound(GrsRpt(3).STATAX + GrsRpt(1).STATAX)
  
  'clear for the next transaction record
  GrsRpt(1) = GrsRpt(2)
  
  Return
  
PrintGrossGTotals:
  LSet EQWL1(1).ENumb = ""
  LSet EQWL1(1).EName = "        Totals:"
  RSet EQWL1(1).GPay = Using(Image3$, GrsRpt(3).GrossPay)
  RSet EQWL1(1).FedGr = Using(Image3$, GrsRpt(3).FedGross)
  RSet EQWL1(1).FedTx = Using(Image3$, GrsRpt(3).FEDTAX)
  RSet EQWL1(1).SocGr = Using(Image3$, GrsRpt(3).SocGross)
  RSet EQWL1(1).SocTx = Using(Image3$, GrsRpt(3).SOCTAX)
  RSet EQWL1(1).MedGr = Using(Image3$, GrsRpt(3).MedGross)
  RSet EQWL1(1).MedTx = Using(Image3$, GrsRpt(3).MEDTAX)
  RSet EQWL1(1).EIC = Using(Image3$, GrsRpt(3).EIC)
  RSet EQWL1(1).StaTx = Using(Image3$, GrsRpt(3).STATAX)
  
  Print #RHandle, DashLine(1)
  Print #RHandle, EQWL1(1).ENumb; EQWL1(1).EName; EQWL1(1).GPay;
  Print #RHandle, EQWL1(1).FedGr; EQWL1(1).FedTx; EQWL1(1).SocGr;
  Print #RHandle, EQWL1(1).SocTx; EQWL1(1).MedGr; EQWL1(1).MedTx;
  Print #RHandle, EQWL1(1).EIC; EQWL1(1).StaTx
  Print #RHandle, FF$
  
  Return

MaxSSWage:
  
  StopDateFlag = False
  TransRecNum& = Emp2Rec.LastTransRec
  ReDim SSTempDates(1 To 1) As Integer
  ReDim SSTempAmts(1 To 1) As Double
  Do
    Get THandle, TransRecNum&, TransHRec(1)
    If (TransHRec(1).CheckDate >= ThisDate) And (TransHRec(1).CheckDate <= HiDate) Then
      NextDate = NextDate + 1
      ReDim Preserve SSTempDates(1 To NextDate) As Integer
      ReDim Preserve SSTempAmts(1 To NextDate) As Double
      SSTempDates(NextDate) = TransHRec(1).CheckDate
      SSTempAmts(NextDate) = OldRound(TransHRec(1).SocGrossPay + TransHRec(1).TaxFring)
    End If
    If TransHRec(1).PrevTransRec > 0 Then
      TransRecNum& = TransHRec(1).PrevTransRec
    Else
      Exit Do
    End If
  Loop
    
  ReDim SSAmts(1 To NextDate) As Double
  ReDim SSDates(1 To NextDate) As Integer
  'assign dates/amts in forward order
  z = NextDate
  For x = 1 To NextDate
    SSDates(x) = SSTempDates(z)
    SSAmts(x) = SSTempAmts(z)
    z = z - 1
  Next x
  
  TotSocGross = 0
    
  For x = 1 To NextDate
    TotSocGross = OldRound(TotSocGross + SSAmts(x))
    If SSDates(x) >= LowDate Then SSTotal = SSTotal + SSAmts(x)
    If TotSocGross > FedSSMax Then
      If StopDateFlag = False Then
        StopDateFlag = True
        StopDate = SSDates(x)
        If NextDate > 1 Then
          ThisDif = OldRound(TotSocGross - FedSSMax)
          Exit For
        Else
          SSMaxCode = 4
          Exit For
        End If
      End If
    End If
  Next x
  
  If SSMaxCode = 4 Then
    GrsRpt(1).SocGross = FedSSMax
    GrsRpt(3).SocGross = GrsRpt(3).SocGross + FedSSMax
    Return
  End If
  
  If StopDate < LowDate Then
    GrsRpt(1).SocGross = 0
    GrsRpt(3).SocGross = GrsRpt(3).SocGross
  Else
    GrsRpt(1).SocGross = SSTotal - ThisDif
    GrsRpt(3).SocGross = GrsRpt(3).SocGross + GrsRpt(1).SocGross
  End If
  
  Return
  


EndTran:

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

