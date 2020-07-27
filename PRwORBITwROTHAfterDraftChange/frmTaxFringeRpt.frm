VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxFringeRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Taxable Fringe Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxFringeRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4770
      Left            =   2115
      TabIndex        =   0
      Top             =   1980
      Width           =   7410
      _Version        =   196609
      _ExtentX        =   13070
      _ExtentY        =   8414
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483627
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxFringeRpt.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3450
         TabIndex        =   1
         Top             =   2970
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
         ColDesigner     =   "frmTaxFringeRpt.frx":08E6
      End
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   1635
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3228
         _ExtentY        =   653
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
         Left            =   3600
         TabIndex        =   3
         Top             =   2310
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3228
         _ExtentY        =   653
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
         Left            =   4170
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to create the 'Taxable Fringe' report."
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
         ButtonDesigner  =   "frmTaxFringeRpt.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1290
         TabIndex        =   5
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
         ButtonDesigner  =   "frmTaxFringeRpt.frx":0DBC
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   828
         Left            =   912
         Top             =   432
         Width           =   5628
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
         Left            =   1920
         TabIndex        =   9
         Top             =   2385
         Width           =   1215
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
         Left            =   1875
         TabIndex        =   8
         Top             =   1725
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable Fringe Report"
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
         Left            =   1056
         TabIndex        =   7
         Top             =   624
         Width           =   5388
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
         Left            =   1350
         TabIndex        =   6
         Top             =   3075
         Width           =   1500
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5040
      Left            =   1973
      Top             =   1845
      Width           =   7695
   End
End
Attribute VB_Name = "frmTaxFringeRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'8/13 created this new report
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload Me
   MainLog ("Taxable Fringe Report screen exited.")
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
  Call LoadMe
  Me.HelpContextID = hlpTaxFringe
  MainLog ("Taxable Fringe Report accessed.")
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub PrintGraphics()
  Dim LowDate As Long, HiDate As Long
  Dim NumOfRecs As Long, cnt As Integer
  Dim UnitFileName As Integer, RptTitle$, UTemp$
  Dim RptName$, THandle As Integer, RHandle As Integer
  Dim DHandle As Integer, UsingThisOne As Boolean
  Dim RecNo As Long
  Dim City$, UnitHandle As Integer, x As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim EmpIdxLNameHandle As Integer, EmpHistoryRpt$
  Dim Emp2Rec As EmpData2Type
  Dim TransRec As TransRecType
  Dim Unit As UnitFileRecType
  Dim ThisCnt As Integer
  Dim TransCnt As Integer
  ReDim EmpIdx(1 To 1) As String
  Dim dlm$
  Dim ThisTRec As Long
  Dim First As Boolean
  Dim EmpFTotal As Double
  Dim gfTotal As Double
  Dim EmpGTotal As Double
  Dim GGTotal As Double
  Dim NumOfTransRecs As Double
  
  On Error GoTo ErrorHandler
  dlm$ = "~"
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (MM-DD-YYYY)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (MM-DD-YYYY)"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  If HiDate < LowDate Then
    MsgBox "ERROR: The End Date is earlier than the Start Date"
    fptxtStart.SetFocus
    Exit Sub
  End If

  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If

  ReDim IdxBuff(1 To NumOfRecs) As Integer

  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  RptTitle$ = "Tax Fringe Report"
  RptName$ = "PRRPTS\TAXFRING.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  
  OpenTransHistFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  If NumOfTransRecs = 0 Then
    MsgBox "There are no transactions saved."
    Close
    Exit Sub
  End If
  OpenEmpData2File EHandle
  FrmShowPctComp.Label1 = "Tax Fringe Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  For x = 1 To NumOfRecs
    First = True
    EmpFTotal = 0
    EmpGTotal = 0
    Get EHandle, IdxBuff(x), Emp2Rec
    ThisTRec = Emp2Rec.LastTransRec
    If ThisTRec > 0 Then
      Do While ThisTRec > 0
        Get THandle, ThisTRec, TransRec
            If TransRec.CheckDate >= LowDate And TransRec.CheckDate <= HiDate And TransRec.TaxFring > 0 Then
              ThisCnt = ThisCnt + 1
              '                 1                  2                              3                                                           4                                     5                            6
              Print #RHandle, City; dlm; QPTrim$(Emp2Rec.EmpNo); dlm; QPTrim$(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName); dlm; MakeRegDate(TransRec.CheckDate); dlm; TransRec.TaxFring; dlm; TransRec.GrossPay
            End If
          ThisTRec = TransRec.PrevTransRec
      Loop
    End If
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  
  Close RHandle

  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp

  If ThisCnt = 0 Then
    MsgBox "No taxable fringe amounts were found that fit the parameters entered."
    Exit Sub
  End If

  arTaxFringRpt.Show

  MainLog ("Tax Fringe Report (graphics) processed.")

  Exit Sub


ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload frmLoadingRpt
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub
Private Sub LoadMe()
   Dim Today As String * 10
'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   fptxtStart.Text = "01-01-" + Mid(Today, 7, 4)
   fptxtEnd.Text = Today
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
      MainLog ("Payroll.exe terminated via menu bar on frmTaxFringeRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim LowDate As Long, HiDate As Long
  Dim MaxLines As Integer, LineCnt As Integer
  Dim NumOfRecs As Long, cnt As Integer
  Dim UnitFileName As Integer, RptTitle$
  Dim RptName$, THandle As Integer, RHandle As Integer
  Dim DHandle As Integer, UTemp$
  Dim RecNo As Long, Page As Integer
  Dim Dash As String * 80
  Dim City$, UnitHandle As Integer, x As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim EmpIdxLNameHandle As Integer
  Dim TransRecLen As Long, TransRecNum&, FF$
  Dim Emp2Rec As EmpData2Type
  Dim EHandle As Integer
  Dim TransRec As TransRecType
  Dim NumOfTransRecs As Double
  
  Dim Unit As UnitFileRecType
  ReDim Pg(1) As String * 3
  Dim ThisCnt As Integer
  Dim TransCnt As Integer
  ReDim EmpIdx(1 To 1) As String
  Dim ThisTRec As Long
  Dim First As Boolean
  Dim EmpFTotal As Double
  Dim gfTotal As Double
  Dim EmpGTotal As Double
  Dim GGTotal As Double
  
  On Error GoTo ErrorHandler
  FF$ = Chr(12)
  MaxLines = 55
  LineCnt = 0
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (MM-DD-YYYY)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (MM-DD-YYYY)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  If HiDate < LowDate Then
    MsgBox "ERROR: The End Date is earlier than the Start Date"
    fptxtStart.SetFocus
    Exit Sub
  End If

  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
  
  MaxLines = 57
  LineCnt = 0
  Dash = String$(78, "-")

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  RptTitle$ = "Taxable Fringe Report"
  RptName$ = "PRRPTS\TAXFRING.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  
  OpenTransHistFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  If NumOfTransRecs = 0 Then
    MsgBox "There are no transactions saved."
    Close
    Exit Sub
  End If

  GoSub PrintHeader
  
  OpenEmpData2File EHandle
  FrmShowPctComp.Label1 = "Tax Fringe Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  For x = 1 To NumOfRecs
    First = True
    EmpFTotal = 0
    EmpGTotal = 0
    Get EHandle, IdxBuff(x), Emp2Rec
    ThisTRec = Emp2Rec.LastTransRec
    If ThisTRec > 0 Then
      Do While ThisTRec > 0
        Get THandle, ThisTRec, TransRec
          If TransRec.CheckDate >= LowDate And TransRec.CheckDate <= HiDate Then
            If TransRec.TaxFring > 0 And First = True Then
              First = False
              If LineCnt >= MaxLines Then
                Print #RHandle, FF$
                GoSub PrintHeader
              End If
              GoSub PrintEmp
              ThisCnt = ThisCnt + 1
              EmpFTotal = OldRound(EmpFTotal + TransRec.TaxFring)
              EmpGTotal = OldRound(EmpGTotal + TransRec.GrossPay)
              Print #RHandle, Tab(27); MakeRegDate(TransRec.CheckDate); Tab(49); Using$("$##,##0.00", TransRec.TaxFring); Tab(69); Using$("$##,##0.00", TransRec.GrossPay)
              LineCnt = LineCnt + 1
            ElseIf TransRec.TaxFring > 0 Then
              ThisCnt = ThisCnt + 1
              If LineCnt >= MaxLines Then
                Print #RHandle, FF$
                GoSub PrintHeader
                GoSub PrintEmp
              End If
              EmpFTotal = OldRound(EmpFTotal + TransRec.TaxFring)
              EmpGTotal = OldRound(EmpGTotal + TransRec.GrossPay)
              Print #RHandle, Tab(27); MakeRegDate(TransRec.CheckDate); Tab(49); Using$("$##,##0.00", TransRec.TaxFring); Tab(69); Using$("$##,##0.00", TransRec.GrossPay)
              LineCnt = LineCnt + 1
            End If
          End If
          ThisTRec = TransRec.PrevTransRec
      Loop
    End If
    If EmpFTotal > 0 Then
      If LineCnt >= MaxLines - 4 Then
        Print #RHandle, FF$
        GoSub PrintHeader
        GoSub PrintEmp
      End If
      Print #RHandle,
      Print #RHandle, "Totals For: " + QPTrim$(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName); Tab(49); Using$("$##,##0.00", EmpFTotal); Tab(68); Using$("$###,##0.00", EmpGTotal)
      Print #RHandle, Dash$
      Print #RHandle,
      LineCnt = LineCnt + 4
    End If
    gfTotal = OldRound(gfTotal + EmpFTotal)
    GGTotal = OldRound(GGTotal + EmpGTotal)
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
   
  Print #RHandle, Dash
  Print #RHandle, "Grand Totals"; Tab(49); Using$("$##,##0.00", gfTotal); Tab(66); Using$("$#,###,##0.00", GGTotal)
      
  Print #RHandle, FF$
  Close
  
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  
  If ThisCnt = 0 Then
    MsgBox "No taxable fringe amounts were found that fit the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$
  MainLog ("Taxable Fringe Report processed.")
  
  Exit Sub
  
PrintEmp:
  Print #RHandle,
  Print #RHandle, "Employee #: " + QPTrim$(Emp2Rec.EmpNo); Tab(23);
  Print #RHandle, "            " + QPTrim$(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName)
  Print #RHandle, Dash
  LineCnt = LineCnt + 3
  Return
  
  
PrintHeader:
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(80)
  LSet UTemp$ = City
  Mid$(UTemp$, 71) = "Page:" + Pg(1)
  Print #RHandle, UTemp$
  Print #RHandle, "Taxable Fringe Report"
  Print #RHandle, MakeRegDate(LowDate) + " thru " + MakeRegDate(HiDate)
  Print #RHandle,
  Print #RHandle, "                             Date        Tax Fringe Amount           Gross Pay"
  Print #RHandle, Dash
  LineCnt = 6
  Return
  
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload frmLoadingRpt
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."


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
      fptxtStart.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub


