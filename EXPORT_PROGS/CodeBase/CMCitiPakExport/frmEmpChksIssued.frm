VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpChksIssued 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Checks Issued Report"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmEmpChksIssued.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8839.999
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4815
      Left            =   2160
      TabIndex        =   4
      Top             =   2018
      Width           =   7350
      _Version        =   196609
      _ExtentX        =   12965
      _ExtentY        =   8493
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmEmpChksIssued.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3405
         TabIndex        =   3
         Top             =   3150
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         ColDesigner     =   "frmEmpChksIssued.frx":08E6
      End
      Begin VB.CheckBox chkVoidOnly 
         BackColor       =   &H00D0D0D0&
         Caption         =   "List Voided Checks Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   2640
         Width           =   3135
      End
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   375
         Left            =   3795
         TabIndex        =   0
         Top             =   1290
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   667
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
         Height          =   375
         Left            =   3795
         TabIndex        =   1
         Top             =   1920
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   667
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
         ToolTipText     =   "Press F10 to generate a report detailing payroll checks issued to employees."
         Top             =   3720
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
         ButtonDesigner  =   "frmEmpChksIssued.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1320
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   3720
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
         ButtonDesigner  =   "frmEmpChksIssued.frx":0DBC
      End
      Begin VB.Label Label6 
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
         Left            =   1650
         TabIndex        =   8
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Checks Issued Report"
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
         Height          =   495
         Left            =   1350
         TabIndex        =   7
         Top             =   555
         Width           =   4815
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
         Height          =   255
         Left            =   2010
         TabIndex        =   6
         Top             =   1440
         Width           =   1260
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
         Left            =   2115
         TabIndex        =   5
         Top             =   2010
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1290
         Top             =   405
         Width           =   4905
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5085
      Left            =   1980
      Top             =   1898
      Width           =   7695
   End
End
Attribute VB_Name = "frmEmpChksIssued"
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
   Unload frmEmpChksIssued
   MainLog ("Checks Issued Report screen exited.")
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
  Call LoadEmpChksIssScreen
  Me.HelpContextID = hlpChecksIssued
  MainLog ("Checks Issued Report accessed.")
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

  Dim LowDate As Long, HiDate As Long
  Dim Image1$, Image2$, Chknum$
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, cnt As Integer
  Dim UnitFileName As Integer
  Dim RptName$, THandle As Integer, RHandle As Integer
  Dim DHandle As Integer, UsingThisOne As Boolean
  Dim RecNo As Long
  Dim City$, UnitHandle As Integer, x As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim EmpIdxLNameHandle As Integer, EmpHistoryRpt$
  Dim TransRecLen As Long, TransRecNum&
  Dim Emp2Rec As EmpData2Type
  Dim TotNumOfChks As Integer
  Dim TotAmtOfChks As Double
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  Dim Image$
  Dim SumTotal As Boolean
  Dim ThisTotal As Double
  Dim ThisCnt As Integer
  Dim TransCnt As Integer, dlm$
  
  dlm$ = "~"
  SumTotal = False
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  If HiDate < LowDate Then
    MsgBox "ERROR: The End Date is earlier than the Start Date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  Image$ = "$###,###,##0.00"
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle

  EmpRecSize = Len(Emp2Rec)
  
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  FrmShowPctComp.Label1 = "Employee Checks Issued Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  RptName$ = "PRRPTS\CHKISSUEG.RPT"
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  OpenEmpData2File DHandle

  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    If RecNo = NumOfRecs Then SumTotal = True '8/13
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm4
    End If
    TransRecNum& = Emp2Rec.LastTransRec

    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HiDate
        If chkVoidOnly.Value = 1 And TransHRec(1).NetPay > 0 Then GoTo VoidOnly
        GoSub PrintChkInfo
        UsingThisOne = True
      Case Else
      End Select
VoidOnly:
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
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
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  Close DHandle
  Close THandle
  Close RHandle
  If ThisCnt = 0 Then
    MsgBox "No checks were found that fit the parameters entered."
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    Exit Sub
  End If
  
  arEmpChksIssued.Show
  frmLoadingRpt.Show
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  MainLog ("Checks Issued Report processed.")
  Exit Sub

PrintChkInfo:
  ThisCnt = ThisCnt + 1
  TotNumOfChks = TotNumOfChks + 1
  ThisTotal = ThisTotal + TransHRec(1).NetPay
  TotAmtOfChks = TotAmtOfChks + TransHRec(1).NetPay
  '                 0                 1
  Print #RHandle, City; dlm; MakeRegDate(LowDate) + " through " + MakeRegDate(HiDate); dlm;
  '                               2                                  3
  Print #RHandle, MakeRegDate(TransHRec(1).CheckDate); dlm; TransHRec(1).CheckNum; dlm;
  '                                 4
  Print #RHandle, QPTrim$(Emp2Rec.EmpFName) + "    " + QPTrim$(Emp2Rec.EmpLName); dlm;
  '                                   5                                      6
  Print #RHandle, Using("#,###,##0.00", TransHRec(1).NetPay); dlm; Using("#,##0", ThisCnt); dlm;
  '                               7                            8                         9
  Print #RHandle, Using("#,###,##0.00", ThisTotal); dlm; TotNumOfChks; dlm; Using(Image, TotAmtOfChks)
Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."
End Sub
Private Sub LoadEmpChksIssScreen()
   Dim Today As String * 10
   Today = Date '$
   fptxtStart.Text = "01-01-" + Mid(Today, 7, 4)
   fptxtEnd.Text = Today
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
   chkVoidOnly.Value = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpChksIssued.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim LowDate As Long, HiDate As Long
  Dim Image1$, Image2$, ToPrint$, Chknum$
  Dim MaxLines As Integer, LineCnt As Integer
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, cnt As Integer
  Dim UnitFileName As Integer, RptTitle$, UTemp$
  Dim RptName$, THandle As Integer, RHandle As Integer
  Dim DHandle As Integer, UsingThisOne As Boolean
  Dim RecNo As Long, Page As Integer, CrLf$
  Dim Dash As String * 80 ', OKFlag As Boolean
  Dim City$, UnitHandle As Integer, x As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim EmpIdxLNameHandle As Integer, EmpHistoryRpt$
  Dim TransRecLen As Long, TransRecNum&, FF$
  Dim Emp2Rec As EmpData2Type
  Dim TotNumOfChks As Integer
  Dim TotAmtOfChks As Double
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  Dim Image$
  Dim SumTotal As Boolean
  Dim ThisTotal As Double
  Dim ThisCnt As Integer
  Dim TransCnt As Integer
  
  SumTotal = False
  FF$ = Chr(12)
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  If HiDate < LowDate Then
    MsgBox "ERROR: The End Date is earlier than the Start Date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  Image$ = "$###,###,##0.00"
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle

  ToPrint$ = Space$(78)
  Chknum$ = Space$(12)

  EmpRecSize = Len(Emp2Rec)
  
  MaxLines = 57
  LineCnt = 0
  Dash = String$(78, "-") + CrLf$

  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  FrmShowPctComp.Label1 = "Employee Checks Issued Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  RptTitle$ = "Employee Checks Issued Report"
  RptName$ = "PRRPTS\CHKISSUE.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 10, RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  OpenEmpData2File DHandle

  GoSub PrintIssuedHeader

  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    If RecNo = NumOfRecs Then SumTotal = True '8/13
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm4
    End If
    TransRecNum& = Emp2Rec.LastTransRec

    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HiDate
        If chkVoidOnly.Value = 1 And TransHRec(1).NetPay > 0 Then GoTo VoidOnly
        GoSub PrintChkInfo
        UsingThisOne = True
      Case Else
      End Select
VoidOnly:
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
        If UsingThisOne Then    'if there were any at all
          GoSub PrintSubTotals '8/13
          Print #RHandle, CrLf$
          LineCnt = LineCnt + 1
          If LineCnt > MaxLines Then            'bottom of page?
            Print #RHandle, FF$   'yes; form feed
            GoSub PrintIssuedHeader
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
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  Print #RHandle, FF$
  GoSub PrintIssuedHeader
  GoSub PrintChkTotals
  Print #RHandle, Dash
  Print #RHandle, FF$
  RPTSetupPRN 123, RHandle '7/24 revised 8/15/02
  Close DHandle
  Close THandle
  Close RHandle
  
  If ThisCnt = 0 Then
    MsgBox "No checks were found that fit the parameters entered."
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  MainLog ("Checks Issued Report processed.")
  Exit Sub

PrintIssuedHeader:
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(80)
  LSet UTemp$ = City
  Mid$(UTemp$, 71) = "Page:" + Pg(1) + CrLf$
  Print #RHandle, UTemp$
  Print #RHandle, "Checks Issued Report" + CrLf$
  Print #RHandle, MakeRegDate(LowDate) + " thru " + MakeRegDate(HiDate) + CrLf$
  If SumTotal = False Then
    Print #RHandle, "Check Date    Check Number        Employee Name                       Net Pay" + CrLf$
    Print #RHandle, Dash
  End If
  LineCnt = 5
  Return

PrintChkInfo:
  LSet ToPrint$ = MakeRegDate(TransHRec(1).CheckDate)
  RSet Chknum$ = TransHRec(1).CheckNum
  ThisCnt = ThisCnt + 1
  TotNumOfChks = TotNumOfChks + 1
  Mid$(ToPrint$, 15) = Chknum$
  Mid$(ToPrint$, 35) = QPTrim$(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName)
  Mid$(ToPrint$, 65) = Using("$#,###,##0.00", TransHRec(1).NetPay) '7/26
  ThisTotal = ThisTotal + TransHRec(1).NetPay
  TotAmtOfChks = TotAmtOfChks + TransHRec(1).NetPay
  Print #RHandle, ToPrint$ + CrLf$
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then            'bottom of page?
    Print #RHandle, FF$   'yes; form feed
    GoSub PrintIssuedHeader
  End If
Return

PrintSubTotals: '8/13 entire sub
  LSet ToPrint$ = ""
  Print #RHandle,
  Mid$(ToPrint$, 2) = "Total Number of Checks: " + Using("#,##0", ThisCnt)
  Mid$(ToPrint$, 41) = "Total Amount of Checks: " + Using("$#,###,##0.00", ThisTotal)
  Print #RHandle, ToPrint$ + CrLf$
  LineCnt = LineCnt + 2
  ThisTotal = 0
'  ThisCnt = 0
Return

PrintChkTotals:
  Print #RHandle, "                  Number of Checks                                    Net Pay"
  Print #RHandle, Dash
  Print #RHandle, "Totals"; Tab(23); TotNumOfChks; Tab(63); Using(Image, TotAmtOfChks)
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


