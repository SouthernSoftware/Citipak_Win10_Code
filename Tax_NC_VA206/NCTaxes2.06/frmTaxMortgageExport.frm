VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxMortgageExport 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bills File Creation for Mortgage Companies"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxMortgageExport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7125
      Left            =   1194
      TabIndex        =   0
      Top             =   786
      Width           =   9345
      _Version        =   196609
      _ExtentX        =   16484
      _ExtentY        =   12568
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxMortgageExport.frx":08CA
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   384
         Left            =   3600
         TabIndex        =   1
         Top             =   1680
         Width           =   3492
         _Version        =   196608
         _ExtentX        =   6159
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
         ColDesigner     =   "frmTaxMortgageExport.frx":08E6
      End
      Begin LpLib.fpList fpList 
         Height          =   2448
         Left            =   960
         TabIndex        =   3
         Tag             =   $"frmTaxMortgageExport.frx":0CC1
         Top             =   2280
         Width           =   7452
         _Version        =   196608
         _ExtentX        =   13144
         _ExtentY        =   4318
         TextAlias       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         Columns         =   4
         Sorted          =   0
         LineWidth       =   1
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   1
         ColumnWidthScale=   2
         RowHeight       =   -1
         MultiSelect     =   1
         WrapList        =   0   'False
         WrapWidth       =   0
         SelMax          =   -1
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
         ColumnHeaderShow=   -1  'True
         ColumnHeaderHeight=   -1
         GrpsFrozen      =   0
         BorderGrayAreaColor=   -2147483637
         ExtendRow       =   0
         DataField       =   ""
         OLEDragMode     =   0
         OLEDropMode     =   0
         Redraw          =   -1  'True
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         ColDesigner     =   "frmTaxMortgageExport.frx":0E3A
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   3696
         TabIndex        =   2
         ToolTipText     =   "Select a print option for the report that prints out after the export files are created."
         Top             =   5400
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
         ColDesigner     =   "frmTaxMortgageExport.frx":128E
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   492
         Left            =   1536
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "Press 'Exit' to return to the main Customer Maintenance menu."
         Top             =   6096
         Width           =   1692
         _Version        =   131072
         _ExtentX        =   2984
         _ExtentY        =   868
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
         ButtonDesigner  =   "frmTaxMortgageExport.frx":1669
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   495
         Left            =   6210
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   $"frmTaxMortgageExport.frx":1847
         Top             =   6090
         Width           =   1695
         _Version        =   131072
         _ExtentX        =   2990
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmTaxMortgageExport.frx":18E2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   492
         Left            =   3600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   6096
         Width           =   2292
         _Version        =   131072
         _ExtentX        =   4043
         _ExtentY        =   868
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
         ButtonDesigner  =   "frmTaxMortgageExport.frx":1AC1
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
         Height          =   345
         Left            =   2040
         TabIndex        =   9
         Top             =   5505
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3825
         Left            =   480
         Top             =   1395
         Width           =   8415
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Mortgage Company Tax Bill Export File Creation"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   450
         Width           =   6825
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1110
         Top             =   315
         Width           =   7185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Range:"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   1800
         Width           =   2175
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7440
      Left            =   954
      Top             =   648
      Width           =   9732
   End
End
Attribute VB_Name = "frmTaxMortgageExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim MortRecs() As Long
  Dim MortCnt As Integer

Private Sub cmdClear_Click()
  fpList.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  If Exist("frombillpost.dat") Then
    Call frmTaxBillingMenu.cmdPost_Click
    KillFile "frombillpost.dat"
    Exit Sub
  End If
  
  frmTaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim AllFlag As Boolean
  Dim SpreadCnt As Integer
  Dim MortRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim x As Integer
  Dim NumOfMCRecs As Integer
  
  'on error goto ERRORSTUFF
  
  OpenMortCodeFile MCHandle, NumOfMCRecs
  Close MCHandle
  
  AllFlag = True
  If fpcmbRange.Text <> "ALL MORTGAGORS" Then
    AllFlag = False
  End If
  ReDim SpreadIdx(1 To 1) As Long
  SpreadCnt = 0
  If AllFlag = False Then
    For x = 0 To NumOfMCRecs - 1
      fpList.Row = x
      If fpList.Selected = True Then
        fpList.ListIndex = x
        fpList.Col = 3
        SpreadCnt = SpreadCnt + 1
        ReDim Preserve SpreadIdx(1 To SpreadCnt) As Long
        SpreadIdx(SpreadCnt) = CInt(fpList.ColText)
      End If
    Next x
    If SpreadCnt = 0 Then
      Call TaxMsg(900, "Please make a selection from the list.")
      Close
      Exit Sub
    End If
  Else
    ReDim SpreadIdx(1 To MortCnt) As Long
    For x = 1 To MortCnt
      SpreadIdx(x) = MortRecs(x)
    Next x
    SpreadCnt = MortCnt
  End If
  
  Call PrintMortFiles(SpreadIdx(), SpreadCnt)
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMortgageExport", "cmdProcess_Click", Erl)
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
    Case vbKeyF3:
      SendKeys "%S"
      Call cmdClear_Click
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
  Me.HelpContextID = hlpCreateMortgage
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxMortgageExport.")
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
  Dim MortRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim x As Integer
  Dim NumOfMCRecs As Integer
  
  MortCnt = 0
  OpenMortCodeFile MCHandle, NumOfMCRecs
  For x = 1 To NumOfMCRecs
    Get MCHandle, x, MortRec
    If MortRec.Deleted <> 0 Then GoTo SkipIt
    MortCnt = MortCnt + 1
    ReDim Preserve MortRecs(1 To MortCnt) As Long
    MortRecs(MortCnt) = x
    fpList.InsertRow = QPTrim$(MortRec.MORTCODE) + Chr(9) + QPTrim$(MortRec.BName) + Chr(9) + QPTrim$(MortRec.XFileNme) + Chr(9) + CStr(x)
SkipIt:
  Next x
  fpList.ListIndex = 0
  
  fpcmbRange.Text = "ALL MORTGAGORS"
  fpcmbRange.AddItem "ALL MORTGAGORS"
  fpcmbRange.AddItem "SELECT FROM LIST"
    
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
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
      fpcmbRange.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbRange_Change()
  If fpcmbRange.Text = "ALL MORTGAGORS" Then
    fpList.Action = ActionDeselectAll
    fpList.Enabled = False
  Else
    fpList.Enabled = True
  End If

End Sub

Private Sub PrintMortFiles(SpreadIdx() As Long, ListCnt As Integer)
  Dim MortRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim x As Integer, y As Long
  Dim NumOfMCRecs As Integer
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim TaxCust As TaxCustType
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim FF$
  Dim FileName$
  Dim RptOut$
  Dim ThisCode$
  Dim CustCnt As Long
  Dim NextRec
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisYear$, GCustCnt As Long
  Dim dlm$, ThisXName$
  Dim MCCnt As Integer
  Dim BillNum$
  Dim CustName$
  Dim CustPin$
  Dim CustAdd1$
  Dim CustAdd2$
  Dim CustAdd3$
  Dim ZIPCODE$
  Dim LOTACRE$
  Dim PersPin$
  Dim RealPin$
  Dim TaxTotal As Double
  Dim One As Integer
  Dim AHandle As Integer
  
  'on error goto ERRORSTUFF
  MCCnt = 0
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  ThisYear$ = CStr(TaxMasterRec.TaxYear)
  
  dlm$ = "~"
  ReportFile$ = "MortExpt.PRN"   'Report File Name
  FF$ = Chr$(12)
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenMortCodeFile MCHandle, NumOfMCRecs
  RptHandle = FreeFile
  GCustCnt = 0
  Open ReportFile$ For Output As #RptHandle
  ReDim MortCodeCos(1 To 1) As String
  ReDim MortCodeXName(1 To 1) As String
  ReDim MortCustCnt(1 To 1) As Long
  ReDim MortTaxTot(1 To 1) As Double
  frmTaxShowPctComp.Label1 = "Creating Mortgage Company Export Files"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  For x = 1 To ListCnt
    Get MCHandle, SpreadIdx(x), MortRec
    ThisCode$ = QPTrim$(MortRec.MORTCODE) 'get mort codes
    ThisXName$ = QPTrim$(MortRec.XFileNme)
    If ThisXName = "EMPTY" Then
      If Exist(QPTrim$("TaxMortExp\" + MortRec.MORTCODE) + ThisYear + ".TXT") Then
        KillFile "TaxMortExp\" + QPTrim$(MortRec.MORTCODE) + ThisYear + ".TXT"
      End If
      FileName = "TaxMortExp\" + QPTrim$(MortRec.MORTCODE) + ThisYear + ".TXT"
    Else
      If Exist("TaxMortExp\" + QPTrim$(MortRec.XFileNme) + ThisYear + ".TXT") Then
        KillFile "TaxMortExp\" + QPTrim$(MortRec.XFileNme) + ThisYear + ".TXT"
      End If
      FileName$ = "TaxMortExp\" + QPTrim$(MortRec.XFileNme) + ThisYear + ".TXT"
    End If
    RptOut = FreeFile
    Open FileName$ For Output As #RptOut
    Print #RptOut, QPTrim$(MortRec.BName)
    Print #RptOut, QPTrim$(MortRec.Add1)
    Print #RptOut, QPTrim$(MortRec.Add2)
    Print #RptOut, QPTrim$(MortRec.Add3)
    Print #RptOut,
    Print #RptOut, "Bill #"; dlm; "Cust Name"; dlm; "Cust Pin #"; dlm; "Addr 1"; dlm; "Addr 2"; dlm; "Addr 3"; dlm; "Zip"; dlm;
    Print #RptOut, "Exemptions"; dlm; "Lot Size"; dlm; "Late Tax Due"; dlm; "Lot Or Acre"; dlm;
    Print #RptOut, "Real Pin #"; dlm; "Real Tax Due"; dlm; "Real Rate"; dlm; "Real Value"; dlm;
    Print #RptOut, "Tax Year"
    Print #RptOut,
    CustCnt = 0
    TaxTotal = 0
    For y = 1 To NumOfTBRecs 'look thru bill file
      Get TBHandle, y, TBRec
      If QPTrim$(TBRec.MORTCODE) = ThisCode$ Then
        CustName$ = QPTrim$(TBRec.CustName)
        If InStr(CustName, "~") Then
          CustName = ReplaceString(CustName, "~", "")
        End If
        CustAdd1$ = QPTrim$(TBRec.CustAdd1)
        If InStr(CustAdd1, "~") Then
          CustAdd1 = ReplaceString(CustAdd1, "~", "")
        End If
        CustAdd2$ = QPTrim$(TBRec.CustAdd2)
        If InStr(CustAdd2, "~") Then
          CustAdd2 = ReplaceString(CustAdd2, "~", "")
        End If
        CustAdd3$ = QPTrim$(TBRec.CustAdd3)
        If InStr(CustAdd3, "~") Then
          CustAdd3 = ReplaceString(CustAdd3, "~", "")
        End If
        ZIPCODE$ = QPTrim$(TBRec.CustZip)
        If InStr(ZIPCODE$, "~") Then
          ZIPCODE = ReplaceString$(ZIPCODE, "~", "")
        End If
        LOTACRE$ = QPTrim$(TBRec.LotOrAcre)
        If InStr(LOTACRE, "~") Then
          LOTACRE = ReplaceString(LOTACRE, "~", "")
        End If
        RealPin$ = QPTrim$(TBRec.RealPin)
        If InStr(RealPin$, "~") Then
          RealPin = ReplaceString(RealPin, "~", "")
        End If
        CustCnt = CustCnt + 1
        GCustCnt = GCustCnt + 1
        Print #RptOut, CStr(TBRec.BillNumber); dlm; CustName; dlm; CStr(TBRec.CustPin); dlm;
        Print #RptOut, CustAdd1; dlm; CustAdd2; dlm; CustAdd3; dlm;
        Print #RptOut, ZIPCODE; dlm; Using$("$###,###,##0.00", TBRec.ExptValue); dlm;
        Print #RptOut, QPTrim$(TBRec.LASize); dlm; Using$("$###,###,##0.00", TBRec.LateTaxDue); dlm;
        Print #RptOut, LOTACRE; dlm; RealPin; dlm; Using$("$###,###,##0.00", TBRec.RealTaxDue); dlm;
        Print #RptOut, Using$("##0.00", TBRec.RealTaxRate) + "%"; dlm; Using$("$###,###,##0.00", TBRec.RealValue); dlm;
        Print #RptOut, Using("###0", TBRec.TaxYear)
        TaxTotal = OldRound(TaxTotal + TBRec.RealTaxDue)
      End If
    Next y
    If CustCnt = 0 Then
      Close RptOut
      If Exist(FileName) Then
        KillFile FileName
      End If
    Else
      MCCnt = MCCnt + 1
      ReDim Preserve MortCodeCos(1 To MCCnt) As String
      MortCodeCos(MCCnt) = QPTrim$(ThisCode$)
      ReDim Preserve MortCodeXName(1 To MCCnt) As String
      MortCodeXName(MCCnt) = FileName
      ReDim Preserve MortCustCnt(1 To MCCnt) As Long
      MortCustCnt(MCCnt) = CustCnt
      ReDim Preserve MortTaxTot(1 To MCCnt) As Double
      MortTaxTot(MCCnt) = TaxTotal
      Close RptOut
    End If
    frmTaxShowPctComp.ShowPctComp x, ListCnt
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
  If GCustCnt = 0 Then
    Call TaxMsg(900, "No mortgage codes could be matched up for printing.")
    If Exist("frombillpost.dat") Then
      frmTaxBillingMenu.Show
      Call frmTaxBillingMenu.cmdPost_Click
      DoEvents
      Unload Me
    End If
    Exit Sub
  Else
    One = 1
    AHandle = FreeFile
    Open "mortx" + ThisYear + ".dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
    Call Savemsg(900, "Mortgage code tax bill files have been created successfully.")
  End If
  
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics(MortCodeCos(), MortCodeXName(), MortCustCnt(), MortTaxTot(), MCCnt)
  Else
    frmTaxMsg.Label1.Caption = "Pitch 12 is recommended for this printout."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Call PrintText(MortCodeCos(), MortCodeXName(), MortCustCnt(), MortTaxTot(), MCCnt)
  End If
  
'  If Exist("frombillpost.dat") Then Unload Me
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMortgageExport", "PrintMortFiles", Erl)
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

Private Sub fpcmbRange_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbRange.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRange.ListIndex = -1
  End If
  If fpcmbRange.ListDown <> True Then
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

Private Sub PrintGraphics(MortCodeCos() As String, MortCodeXName() As String, MortCustCnt() As Long, MortTaxTot() As Double, MCCnt As Integer)
  Dim x As Long, y As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim dlm$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TownName$
  Dim MortRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumOfMCRecs As Integer
  Dim Code$
  
  'on error goto ERRORSTUFF
  
  dlm = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\MORTXPRT.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenMortCodeFile MCHandle, NumOfMCRecs
  
  For x = 1 To MCCnt
    Code$ = QPTrim$(MortCodeCos(x))
    For y = 1 To NumOfMCRecs
      Get MCHandle, y, MortRec
      If QPTrim(MortRec.MORTCODE) = Code$ Then
        If QPTrim$(MortRec.PHONE) = "(" Then MortRec.PHONE = ""
        '                    0                      1
        Print #RptHandle, TownName; dlm; QPTrim$(MortRec.BName); dlm;
        '                          2                               3
        Print #RptHandle, QPTrim$(MortRec.Add1); dlm; QPTrim$(MortRec.Add2); dlm;
        '                          4                        5
        Print #RptHandle, QPTrim$(MortRec.Add3); dlm; MortCodeXName(x); dlm;
        '                            6                             7
        Print #RptHandle, QPTrim$(MortRec.Contact); dlm; QPTrim$(MortRec.PHONE); dlm;
        '                       8                   9                10          11
        Print #RptHandle, MortCustCnt(x); dlm; MortTaxTot(x); dlm; Code$; dlm; MCCnt
      End If
    Next y
  Next x
  
  Close
  
  arTaxMortXRpt.Show
  
  Exit Sub

ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMortgageExport", "PrintGraphics", Erl)
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

Private Sub PrintText(MortCodeCos() As String, MortCodeXName() As String, MortCustCnt() As Long, MortTaxTot() As Double, MCCnt As Integer)
  Dim x As Long, y As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TownName$
  Dim MortRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumOfMCRecs As Integer
  Dim Code$
  Dim FF$
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  
  'on error goto ERRORSTUFF
  
  FF$ = Chr(12)
  MaxLines = 58
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TownName = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\MORTXPRT.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  OpenMortCodeFile MCHandle, NumOfMCRecs
  For x = 1 To MCCnt
    Code$ = QPTrim$(MortCodeCos(x))
    For y = 1 To NumOfMCRecs
      Get MCHandle, y, MortRec
      If QPTrim(MortRec.MORTCODE) = Code$ Then
        If LineCnt >= MaxLines - 8 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        If QPTrim$(MortRec.PHONE) = "(" Then MortRec.PHONE = ""
        Print #RptHandle, "Export File Name: " + QPTrim$(MortCodeXName(x)); Tab(42); " Mortgage Code: " + QPTrim$(MortRec.MORTCODE)
        Print #RptHandle, QPTrim$(MortRec.BName); Tab(42); "       Contact: " + QPTrim$(MortRec.Contact)
        Print #RptHandle, QPTrim$(MortRec.Add1); Tab(42); "       Phone #: " + QPTrim$(MortRec.PHONE)
        Print #RptHandle, QPTrim$(MortRec.Add2); Tab(42); "# of Customers: " + Using$("##,##0", MortCustCnt(x))
        Print #RptHandle, QPTrim$(MortRec.Add3); Tab(42); " Total Tax Due: " + Using$("$###,###,##0.00", MortTaxTot(x))
        Print #RptHandle, String(83, "-")
        Print #RptHandle,
        LineCnt = LineCnt + 7
      End If
    Next y
  Next x
  
  Print #RptHandle, FF$
  Close
  
  ViewPrint RptFile, "Mortgage Company Tax Bill Export Report", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(24); "Mortgage Company Tax Bill Export Report"
  Print #RptHandle, TownName
  Print #RptHandle, "# Of Export Files: " + CStr(MCCnt)
  Print #RptHandle, "Report Date: " + CStr(Date); Tab(65); "Page #: " + CStr(Page)
  Print #RptHandle, String(83, "-")
  Print #RptHandle,
  LineCnt = 6
  
  Return

ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMortgageEXport", "PrintText", Erl)
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
