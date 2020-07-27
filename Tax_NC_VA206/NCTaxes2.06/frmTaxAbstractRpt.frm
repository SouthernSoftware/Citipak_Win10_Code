VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxAbstractRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abstract Report"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmTaxAbstractRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4944
      Left            =   1920
      TabIndex        =   3
      Top             =   1914
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   8721
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxAbstractRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbPageBreak 
         Height          =   384
         Left            =   3240
         TabIndex        =   1
         Top             =   2280
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         BackColor       =   16777215
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
         MaxEditLen      =   5
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmTaxAbstractRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   2928
         TabIndex        =   2
         Top             =   2916
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         BackColor       =   16777215
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
         MaxEditLen      =   5
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmTaxAbstractRpt.frx":0C51
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   0
         Top             =   1608
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmTaxAbstractRpt.frx":0FBC
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   2040
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   3930
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxAbstractRpt.frx":1327
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   4275
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   $"frmTaxAbstractRpt.frx":1505
         Top             =   3930
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxAbstractRpt.frx":15B0
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Report Type:"
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
         Left            =   1272
         TabIndex        =   9
         Top             =   1716
         Width           =   1500
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1530
         Top             =   315
         Width           =   4905
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Abstract Report"
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
         Height          =   390
         Left            =   1800
         TabIndex        =   8
         Top             =   450
         Width           =   4335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Order:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1476
         TabIndex        =   7
         Top             =   3000
         Width           =   1308
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2220
         Left            =   1008
         Top             =   1368
         Width           =   5976
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Page Break Y/N?:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   2364
         Width           =   1908
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   5196
      Left            =   1800
      Top             =   1782
      Width           =   8052
   End
End
Attribute VB_Name = "frmTaxAbstractRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Break As Integer
  Dim ThisOpt$
Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    If Mid(fpcmbPageBreak.Text, 1, 1) = "1" Then
      Break = 1
    ElseIf Mid(fpcmbPageBreak.Text, 1, 1) = "2" Then
      Break = 2
    Else
      Break = 3
    End If
    Call PrintGraphicsAbRpt
  Else
    frmTaxMsg.Label1.Caption = "Pitch 17 is recommended for this printout."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    If Mid(fpcmbPageBreak.Text, 1, 1) = "1" Then
      Break = 1
    ElseIf Mid(fpcmbPageBreak.Text, 1, 1) = "2" Then
      Break = 2
    Else
      Break = 3
    End If
    Call PrintTextAbRpt
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpMasterCustomer
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxAbstractRpt.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Integer
  
  'on error goto ERRORSTUFF
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  fpcmbPrintOrder.AddItem "Search Name"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem ThisOpt + " Order"
  End If
  
  fpcmbPageBreak.Text = "1. Page Break Each Customer"
  fpcmbPageBreak.AddItem "1. Page Break Each Customer"
  fpcmbPageBreak.AddItem "2. Page Break Each Property"
  fpcmbPageBreak.AddItem "3. Normal Page Break"
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCustListRpt", "LoadMe", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
End Sub

Private Sub PrintGraphicsAbRpt()
  Dim x As Long
  Dim y As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PropRec As PropertyRecType
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim NextRec As Long
  Dim RealVal As Double
  Dim PersVal As Double
  Dim TotVal As Double
  Dim PrintHeader As Boolean
  Dim RealCnt As Integer
  Dim PersCnt As Integer
  Dim Town$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrintDesc As Boolean
  Dim dlm$
  Dim PCnt As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim CustRec As Long
  
  'on error goto ERRORSTUFF
  
  dlm$ = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
  End If
  
  RptFile$ = "TAXRPTS\ABSTLIST.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  frmTaxShowPctComp.Label1 = "Gathering Property Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    '  If InStr(UCase(TaxCust.CustName), "VIOLA") > 0 Then Stop
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If

    If TaxCust.Active = "N" Then GoTo SkipIt
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    RealVal = 0
    PersVal = 0
    TotVal = 0
    RealCnt = 0
    PersCnt = 0
    'look for valid property for this customer
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo NotThisOne
        RealCnt = RealCnt + 1
NotThisOne:
        NextRec = RealRec.NextRec
      Loop
    End If
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo NotThisPers
        PersCnt = PersCnt + 1
NotThisPers:
        NextRec = PersRec.NextRec
      Loop
    End If
    
    If RealCnt = 0 Then GoTo NoReal
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo RealDeleted
        GoSub PrintReal
RealDeleted:
        NextRec = RealRec.NextRec
      Loop
    End If
NoReal:
    If PersCnt > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo PersDeleted
        GoSub PrintPers
PersDeleted:
        NextRec = PersRec.NextRec
      Loop
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no properties to report.")
    Exit Sub
  End If
  
  arTaxAbstractRpt.Show
  DoEvents
  
  Exit Sub
  
PrintReal:
  '                   0                     1                   2
  Print #RptHandle, Town$; dlm; QPTrim$(TaxCust.CustName); dlm; CustRec; dlm;
  '                             3                           4
  Print #RptHandle, QPTrim$(TaxCust.Addr1); dlm; QPTrim$(TaxCust.Addr2); dlm;
  '                                                 5
  Print #RptHandle, QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip); dlm;
  '                   6                      7                         8
  Print #RptHandle, "REAL"; dlm; QPTrim$(RealRec.RealPin); dlm; RealRec.PROPVALU; dlm;
  '                             9                             10
  Print #RptHandle, QPTrim$(RealRec.PropAddr); dlm; QPTrim$(RealRec.Map) + "/" + QPTrim$(RealRec.BLOCK) + "/" + QPTrim$(RealRec.LOTNUMB); dlm;
  '                        11                      12                   13
  Print #RptHandle, RealRec.PROPNOT1; dlm; RealRec.PROPNOT2; dlm; RealRec.PROPNOT3; dlm;
  '                 14       15       16
  Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  '                 17       18
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 19       20
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 21       22
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 23       24        25
  Print #RptHandle, ""; dlm; ""; dlm; Break
  
  PCnt = PCnt + 1
  
  Return
  
PrintPers:
  '                   0                     1                   2
  Print #RptHandle, Town$; dlm; QPTrim$(TaxCust.CustName); dlm; CustRec; dlm;
  '                             3                           4
  Print #RptHandle, QPTrim$(TaxCust.Addr1); dlm; QPTrim$(TaxCust.Addr2); dlm;
  '                                                 5
  Print #RptHandle, QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip); dlm;
  '                     6                       7                   8
  Print #RptHandle, "PERSONAL"; dlm; QPTrim$(PersRec.PropPin); dlm; ""; dlm;
  '                 9        10
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 11       12       13
  Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  '                       14                    15
  Print #RptHandle, PersRec.PersVal; dlm; PersRec.CVALUE; dlm;
  '                        16                   17
  Print #RptHandle, PersRec.MHVALUE; dlm; PersRec.MTVALUE; dlm;
  '                       18                   19
  Print #RptHandle, PersRec.MCVALUE; dlm; PersRec.DESC1; dlm;
  '                      20                   21
  Print #RptHandle, PersRec.DESC2; dlm; PersRec.DESC3; dlm;
  '                             22                  23               24         25
  Print #RptHandle, QPTrim$(PersRec.Desc4); dlm; PersRec.Desc5; dlm; ""; dlm; Break
  
  PCnt = PCnt + 1
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReportsMenu", "PrintGraphicsAbRpt", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
End Sub

Private Sub PrintTextAbRpt()
  Dim x As Long
  Dim y As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PropRec As PropertyRecType
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim FF$
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim NextRec As Long
  Dim RealVal As Double
  Dim PersVal As Double
  Dim TotVal As Double
  Dim PrintHeader As Boolean
  Dim RealCnt As Integer
  Dim PersCnt As Integer
  Dim Town$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrintDesc As Boolean
  Dim PCnt As Long
  Dim NewOne As Boolean
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim CustRec As Long
  
  'on error goto ERRORSTUFF
  
  NewOne = False
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  FF$ = Chr(12)
  MaxLines = 58
  
  RptFile$ = "TAXRPTS\ABSTLIST.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  frmTaxShowPctComp.Label1 = "Gathering Property Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Active = "N" Then GoTo SkipIt
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    RealVal = 0
    PersVal = 0
    TotVal = 0
    RealCnt = 0
    PersCnt = 0
    NewOne = True
    'look for valid property for this customer
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo NotThisOne
        RealCnt = RealCnt + 1
NotThisOne:
        NextRec = RealRec.NextRec
      Loop
    End If
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo NotThisPers
        PersCnt = PersCnt + 1
NotThisPers:
        NextRec = PersRec.NextRec
      Loop
    End If
    If RealCnt > 0 Or PersCnt > 0 Then
      If LineCnt <> 7 Then
'        If Break = 3 Then GoTo GoSkip
        If Break <> 3 Then Print #RptHandle, FF$
'        LineCnt = 0
        GoSub PrintCustHeader
      End If
    End If
GoSkip:
    If RealCnt = 0 Then GoTo NoReal
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        If NewOne = False And Break = 2 Then
          Print #RptHandle, FF$
          LineCnt = 0
          GoSub PrintCustHeader
        End If
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo RealDeleted
        GoSub PrintReal
RealDeleted:
        NewOne = False
        NextRec = RealRec.NextRec
      Loop
    End If
NoReal:
     If PersCnt > 0 Then
       NextRec = TaxCust.FirstPersRec
       Do While NextRec > 0
         If NewOne = False And Break = 2 Then
           Print #RptHandle, FF$
           LineCnt = 0
           GoSub PrintCustHeader
         End If
         Get PHandle, NextRec, PersRec
         If PersRec.Deleted = -1 Then GoTo PersDeleted
         GoSub PrintPers
PersDeleted:
         NextRec = PersRec.NextRec
         NewOne = False
       Loop
     End If
     If PersCnt > 0 Or RealCnt > 0 Then
       Print #RptHandle, String$(80, "=")
       LineCnt = LineCnt + 1
       If LineCnt > MaxLines Then
         Print #RptHandle, FF$
         If Break = 3 Then
'           LineCnt = 0
           GoSub PrintCustHeader
         Else
           LineCnt = 0
         End If
       End If
     End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Print #RptHandle, FF$
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no properties to report.")
    Exit Sub
  End If
  
  ViewPrint RptFile, "Property Listing", True
  
  Exit Sub
  
PrintCustHeader:
  If Break = 3 Then
    If LineCnt + 9 >= MaxLines Then
      LineCnt = 0
      Print #RptHandle, FF$
    End If
  End If
  Print #RptHandle, "Abstract Listing: " + Town$
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Account of: "; Tab(15); QPTrim$(TaxCust.CustName); Tab(67); "Acct #: " + Using$("####0", CustRec)
  Print #RptHandle, Tab(15); QPTrim$(TaxCust.Addr1)
  Print #RptHandle, Tab(15); QPTrim$(TaxCust.Addr2)
  Print #RptHandle, Tab(15); QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip)
  Print #RptHandle, Tab(5); String(69, "-")
  If Break = 3 Then
    LineCnt = LineCnt + 7
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 7
    End If
  Else
    LineCnt = 7
  End If
  
  Return
  
PrintReal:
  PrintDesc = False
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "*** REAL PROPERTY ***"
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "Address:"; Tab(22); QPTrim$(RealRec.PropAddr)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  
  Print #RptHandle, Tab(10); "PIN #"; Tab(22); QPTrim$(RealRec.RealPin); Tab(45); "MAP/BLOCK/LOT: "; Tab(62); QPTrim$(RealRec.Map) + "/" + QPTrim$(RealRec.BLOCK) + "/" + QPTrim$(RealRec.LOTNUMB)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "REAL VALUE:"; Tab(22); Using$("$###,###,##0.00", RealRec.PROPVALU)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  If QPTrim$(RealRec.PROPNOT1) = "" And QPTrim$(RealRec.PROPNOT2) = "" And QPTrim$(RealRec.PROPNOT3) = "" Then
    Print #RptHandle, Tab(10); "DESC:"; Tab(22); "NO DESCRIPTION AVAILABLE"
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
    GoTo NoRealDesc
  End If
  If QPTrim$(RealRec.PROPNOT1) <> "" Then
    Print #RptHandle, Tab(10); "DESC:"; Tab(22); RealRec.PROPNOT1
    PrintDesc = True
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
    
  If QPTrim$(RealRec.PROPNOT2) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESC:"; Tab(22); RealRec.PROPNOT2
      PrintDesc = True
    Else
      Print #RptHandle, Tab(22); RealRec.PROPNOT2
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
  
  If QPTrim$(RealRec.PROPNOT3) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESC:"; Tab(22); RealRec.PROPNOT3
      PrintDesc = True
    Else
      Print #RptHandle, Tab(22); RealRec.PROPNOT3
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
  
  Print #RptHandle,
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
NoRealDesc:

  PCnt = PCnt + 1
  
  Return
  
PrintPers:
  PrintDesc = False
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "*** PERSONAL PROPERTY ***"
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "PIN #"; Tab(22); QPTrim$(PersRec.PropPin)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "VALUE AMOUNTS"
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "PERSONAL:"; Tab(30); Using$("$###,###,##0.00", PersRec.PersVal)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "FARM EQUIPMENT:"; Tab(30); Using$("$###,###,##0.00", PersRec.CVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "MOBILE HOMES:"; Tab(30); Using$("$###,###,##0.00", PersRec.MHVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "MACHINE/TOOLS:"; Tab(30); Using$("$###,###,##0.00", PersRec.MTVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "MERCHANT CAPITAL:"; Tab(30); Using$("$###,###,##0.00", PersRec.MCVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
  
  If QPTrim$(PersRec.DESC1) = "" And QPTrim$(PersRec.DESC2) = "" And QPTrim$(PersRec.DESC3) = "" And QPTrim$(PersRec.Desc4) = "" And QPTrim$(PersRec.Desc5) = "" Then
    Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); "NO DESCRIPTION AVAILABLE"
    GoTo NoPersDesc
  End If
  
  If QPTrim$(PersRec.DESC1) <> "" Then
    PrintDesc = True
    Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.DESC1
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
    
  If QPTrim$(PersRec.DESC2) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.DESC2
      LineCnt = 0
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.DESC2
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
  
  If QPTrim$(PersRec.DESC3) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.DESC3
      LineCnt = 0
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.DESC3
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
  
  If QPTrim$(PersRec.Desc4) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.Desc4
      LineCnt = 0
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.Desc4
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
  
  If QPTrim$(PersRec.Desc5) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.Desc5
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.Desc5
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      LineCnt = 0
      GoSub PrintCustHeader
      'If Break <> 3 Then LineCnt = 0
    End If
  End If
  
  Print #RptHandle,
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    LineCnt = 0
    GoSub PrintCustHeader
    'If Break <> 3 Then LineCnt = 0
  End If
NoPersDesc:

  PCnt = PCnt + 1

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReportsMenu", "PrintTextAbRpt", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
End Sub

Private Sub fpcmbPageBreak_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPageBreak.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPageBreak.ListIndex = -1
  End If
  If fpcmbPageBreak.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPageBreak.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

