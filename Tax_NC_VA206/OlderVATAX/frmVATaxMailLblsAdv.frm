VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxMailLblsAdv 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Mailing Labels"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxMailLblsAdv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6420
      Left            =   1193
      TabIndex        =   0
      Top             =   1035
      Width           =   9345
      _Version        =   196609
      _ExtentX        =   16484
      _ExtentY        =   11324
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxMailLblsAdv.frx":08CA
      Begin LpLib.fpList fpList 
         Height          =   1584
         Left            =   960
         TabIndex        =   6
         Tag             =   $"frmVATaxMailLblsAdv.frx":08E6
         Top             =   3240
         Width           =   7452
         _Version        =   196608
         _ExtentX        =   13144
         _ExtentY        =   2794
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
         ColDesigner     =   "frmVATaxMailLblsAdv.frx":0A5F
      End
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   384
         Left            =   4080
         TabIndex        =   4
         Top             =   2520
         Width           =   2412
         _Version        =   196608
         _ExtentX        =   4254
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
         ColDesigner     =   "frmVATaxMailLblsAdv.frx":0E41
      End
      Begin LpLib.fpCombo fpcmbLabel 
         Height          =   384
         Left            =   3120
         TabIndex        =   1
         Top             =   1848
         Width           =   4476
         _Version        =   196608
         _ExtentX        =   7895
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
         ColDesigner     =   "frmVATaxMailLblsAdv.frx":11A8
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
         Height          =   492
         Left            =   4812
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmVATaxMailLblsAdv.frx":150F
         Top             =   5520
         Width           =   1548
         _Version        =   131072
         _ExtentX        =   2730
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
         ButtonDesigner  =   "frmVATaxMailLblsAdv.frx":15EE
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   492
         Left            =   960
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "Press 'Exit' to return to the main Customer Maintenance menu."
         Top             =   5520
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
         ButtonDesigner  =   "frmVATaxMailLblsAdv.frx":17CA
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   492
         Left            =   6600
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   $"frmVATaxMailLblsAdv.frx":19A8
         Top             =   5520
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
         ButtonDesigner  =   "frmVATaxMailLblsAdv.frx":1A43
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   492
         Left            =   2880
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5520
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
         ButtonDesigner  =   "frmVATaxMailLblsAdv.frx":1C22
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
         Left            =   1800
         TabIndex        =   5
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1590
         Top             =   315
         Width           =   6225
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Mailing Labels For Advertising Charges"
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
         TabIndex        =   3
         Top             =   450
         Width           =   5865
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label Type:"
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
         Left            =   1635
         TabIndex        =   2
         Top             =   1950
         Width           =   1350
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
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6675
      Left            =   953
      Top             =   900
      Width           =   9735
   End
End
Attribute VB_Name = "frmVATaxMailLblsAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim CustRecs() As Long
  Dim CustCnt As Long

Private Sub cmdAlign_Click()
  Dim TaxRpt As Integer
  Dim LType As Integer
  Dim cnt As Integer
  Dim Align$
  Dim ReportFile$
  
  On Error GoTo ERRORSTUFF
  
  ReDim OSet(1 To 4) As Integer
  
  Align$ = String$(34, "X")
  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110
  
  If fpcmbLabel.Text = "2) 1 X 2 5/8 1 Wide Text" Then
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 2 5/8 3 Wide Text" Then
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 2 5/8 4 Wide Text" Then
    LType = 3
  Else
    Call TaxMsg(900, "Please make a Label Type selection.")
    fpcmbLabel.SetFocus
    Exit Sub
  End If
  
  TaxRpt = FreeFile
  Open "TAXLABEL.RPT" For Output As TaxRpt
  Select Case LType
  Case 1
    For cnt = 1 To 5
      Print #TaxRpt, Align$
    Next
    Print #TaxRpt,
  Case 2
    For cnt = 1 To 5
      Print #TaxRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$
    Next
    Print #TaxRpt,
  Case 3
    For cnt = 1 To 5
      Print #TaxRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$; Tab(OSet(4)); Align$
    Next
    Print #TaxRpt,
  End Select

  Close TaxRpt

  ViewPrint "TAXLABEL.RPT", "Mailing Labels Alignment", True
  
  MainLog ("Mailing labels 'Align' feature for advertising charges used.")
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMailLblsAdv", "cmdAlign_Click", Erl)
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

Private Sub cmdClear_Click()
  fpList.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  frmVATaxAdvColMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If InStr(fpcmbLabel.Text, "Graphical") Then
    Call PrintGraphics
  ElseIf InStr(fpcmbLabel.Text, "Text") Then
    Call TaxMsg(900, "Pitch 10 is recommended for this report.")
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim LType As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Zip$
  Dim DidCnt As Integer
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim CustPCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim CityName As String * 18
  Dim Address As String * 23
  Dim AllFlag As Boolean
  Dim ValidCnt As Integer
  Dim SpreadCnt As Long
  Dim NumOfCustIdx As Long
  
  On Error GoTo ERRORSTUFF

  AllFlag = True
  If fpcmbRange.Text <> "ALL" Then
    AllFlag = False
  End If

  If fpcmbLabel.Text = "2) 1 X 2 5/8 1 Wide Text" Then
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 2 5/8 3 Wide Text" Then
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 2 5/8 4 Wide Text" Then
    LType = 3
  Else
    Call TaxMsg(900, "Please make a Label Type selection.")
    fpcmbLabel.SetFocus
    Exit Sub
  End If

  OpenAdvColRecFile ATHandle, NumOfATRecs
  Close ATHandle
  
  ReDim SpreadIdx(1 To 1) As Long
  SpreadCnt = 0
  If AllFlag = False Then
    For x = 0 To NumOfATRecs - 1
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
    ReDim SpreadIdx(1 To CustCnt) As Long
    For x = 1 To CustCnt
      SpreadIdx(x) = CustRecs(x)
    Next x
    SpreadCnt = CustCnt
  End If
  
  ReportFile$ = "TAXLABEL.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  ReDim ToPrint(1 To 5, 1 To 5) As String

  If SpreadCnt > 0 Then
    NumOfCustIdx = SpreadCnt
    If SpreadCnt < 11 Then
      If TaxMsgWOpts(800, "Printing so few labels wastes mailing label paper. Press F10 if you wish to continue anyway. Otherwise, press ESC to stop the printing process.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        fpcmbRange.SetFocus
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
      End If
    End If
  End If

  OpenTaxCustFile TCHandle, NumOfTCRecs 'TaxCustRec
  
  For cnt = 1 To NumOfCustIdx
    If SpreadCnt > 0 Then
      Get TCHandle, SpreadIdx(cnt), TaxCustRec
    End If

    CustPCnt = CustPCnt + 1

    If Mid(TaxCustRec.Zip, 7, 1) <> " " Then
      Zip$ = TaxCustRec.Zip
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(TaxCustRec.Zip, 5)
      Zip$ = QPTrim$(Zip$)
    End If

    Select Case LType
    Case 1
      Print #RptHandle, "Cust # " + CStr(SpreadIdx(cnt))
      Print #RptHandle, Left$(QPTrim$(TaxCustRec.CustName), 23)
      Print #RptHandle, Left$(QPTrim$(TaxCustRec.Addr1), 23)
      If Len(QPTrim$(TaxCustRec.Addr2)) > 0 Then
        Print #RptHandle, Left$(QPTrim$(TaxCustRec.Addr2), 23)
        Print #RptHandle, Left$(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim(Zip$)
      Else
        Print #RptHandle, Left$(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
        Print #RptHandle,
      End If
      Print #RptHandle,
      ValidCnt = ValidCnt + 1
    Case 2
      LabelCnt = LabelCnt + 1 'this requires a line to be printed
      'in columns of 3 with each column containing data gathered
      'from different customers...it also must limit the size of some
      'variables to accommodate the limitations of a mailing label's
      'size
      ToPrint(1, LabelCnt) = "Cust # " + CStr(SpreadIdx(cnt))
      ToPrint(2, LabelCnt) = Left(QPTrim$(TaxCustRec.CustName), 23)
      ToPrint(3, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr1), 23)

      If Len(QPTrim$(TaxCustRec.Addr2)) > 0 Then
          ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr2), 23)
          ToPrint(5, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
      Else
          ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
      End If

      If LabelCnt = 3 Then
        For PCnt = 1 To 5
          Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3)
          ToPrint(PCnt, 1) = ""
          ToPrint(PCnt, 2) = ""
          ToPrint(PCnt, 3) = ""
          ToPrint(PCnt, 4) = ""
          ToPrint(PCnt, 5) = ""
        Next
        Print #RptHandle,
        LabelCnt = 0
      End If
      ValidCnt = ValidCnt + 1

    Case 3
      LabelCnt = LabelCnt + 1
      ToPrint(1, LabelCnt) = "Cust # " + CStr(SpreadIdx(cnt))
      ToPrint(2, LabelCnt) = Left(QPTrim$(TaxCustRec.CustName), 23)
      ToPrint(3, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr1), 23)

      If Len(QPTrim$(TaxCustRec.Addr2)) > 0 Then
        ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr2), 23)
        ToPrint(5, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim(Zip$)
      Else
        ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
      End If

      If LabelCnt = 4 Then
        For PCnt = 1 To 5
          Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3); Tab(110); ToPrint(PCnt, 4)
          ToPrint(PCnt, 1) = ""
          ToPrint(PCnt, 2) = ""
          ToPrint(PCnt, 3) = ""
          ToPrint(PCnt, 4) = ""
          ToPrint(PCnt, 5) = ""
        Next
        Print #RptHandle,
        LabelCnt = 0
      End If
      ValidCnt = ValidCnt + 1
    End Select

NextLabel:
  Next

  If LType = 2 Or LType = 3 Then 'this catches the last
  'line of a print job if the last line contains fewer than
  'the number required to trigger another print line
    If LabelCnt > 0 Then
      For PCnt = 1 To 5
        Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3)
      Next
      Print #RptHandle,
    End If
  End If
  PCnt = 0

  Close

'  If ValidCnt = 0 Then
'    frmBLMessageBoxJr.Label1.Caption = "There are no customers within the parameters entered on this screen. Mailing labels not printed."
'    frmBLMessageBoxJr.Label1.Top = 700
'    frmBLMessageBoxJr.Show vbModal
'    Close
'    Exit Sub
'  End If
  
  ViewPrint ReportFile$, "Mailing Labels for Advertising Charges", True

  KillFile ReportFile$

  MainLog ("Mailing labels for advertising charges processed in text format.")

  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMailLblsAdv", "PrintText", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpAdvertisingMailing
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
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
    Case vbKeyF3:
      SendKeys "%L"
      Call cmdClear_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlign_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "C:\CPWork\mllbls.dat"
      ClearInUse PWcnt
      MainLog ("CitiTax.exe terminated via menu bar on frmVATaxMailLblsAdv.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim AdvTrans As InterestRecType
  Dim ARHandle As Integer
  Dim NumOfARRecs As Long
  Dim x As Long
  
  On Error Resume Next
  OpenAdvColRecFile ARHandle, NumOfARRecs
  ReDim CustRecs(1 To 1) As Long
  CustCnt = 0
  fpList.Enabled = False
  For x = 1 To NumOfARRecs
    Get ARHandle, x, AdvTrans
    If AdvTrans.DelFlag = True Then GoTo SkipIt
    CustCnt = CustCnt + 1
    ReDim Preserve CustRecs(1 To CustCnt) As Long
    CustRecs(CustCnt) = AdvTrans.CustRec
    fpList.InsertRow = "  " & Using$("#####", AdvTrans.CustRec) & Chr$(9) & "  " & QPTrim$(AdvTrans.CustName) & Chr$(9) & "  " & QPTrim(AdvTrans.InfoTxt) & Chr$(9) & CStr(AdvTrans.CustRec)
SkipIt:
  Next x
  
  fpList.ListIndex = 0
  
  Close
  
  fpcmbLabel.Text = "1) 1 X 2 5/8 3 Wide Graphical"
  fpcmbLabel.AddItem "1) 1 X 2 5/8 3 Wide Graphical"
  fpcmbLabel.AddItem "2) 1 X 2 5/8 1 Wide Text"
  fpcmbLabel.AddItem "3) 1 X 2 5/8 3 Wide Text"
  fpcmbLabel.AddItem "4) 1 X 2 5/8 4 Wide Text"
  
  fpcmbRange.Text = "ALL"
  fpcmbRange.AddItem "ALL"
  fpcmbRange.AddItem "SELECT FROM LIST"
  
End Sub

Private Sub fpcmbLabel_Change()
  If QPTrim$(fpcmbLabel.Text) = "" Then
    fpcmbLabel.Text = "1) 1 X 2 5/8 3 Wide Graphical"
  End If
  
  If InStr(fpcmbLabel.Text, "Text") Then
    cmdAlign.Enabled = True
  Else
    cmdAlign.Enabled = False
  End If
End Sub

Private Sub fpcmbLabel_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLabel.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLabel.ListIndex = -1
  End If
  If fpcmbLabel.ListDown <> True Then
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
  If fpcmbRange.Text = "ALL" Then
    fpList.Action = ActionDeselectAll
    fpList.Enabled = False
  Else
    fpList.Enabled = True
  End If
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
      fpcmbLabel.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintGraphics()
  Dim LType As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Zip$
  Dim DidCnt As Integer
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim CustPCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim CityName As String * 18
  Dim Address As String * 23
  Dim AllFlag As Boolean
  Dim dlm$
  Dim ValidCnt As Integer
  Dim SpreadCnt As Long
  Dim NumOfCustIdx As Long

  On Error GoTo ERRORSTUFF

  AllFlag = True
  If fpcmbRange.Text <> "ALL" Then
    AllFlag = False
  End If
  
  dlm = "~"
  OpenAdvColRecFile ATHandle, NumOfATRecs
  Close ATHandle

  ReDim SpreadIdx(1 To 1) As Long
  SpreadCnt = 0
  If AllFlag = False Then
    For x = 0 To NumOfATRecs - 1
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
    ReDim SpreadIdx(1 To CustCnt) As Long
    For x = 1 To CustCnt
      SpreadIdx(x) = CustRecs(x)
    Next x
    SpreadCnt = CustCnt
  End If
  
  ReportFile$ = "TAXRPTS\TAXLABEL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  ReDim ToPrint(1 To 5, 1 To 5) As String

  If SpreadCnt > 0 Then
    NumOfCustIdx = SpreadCnt
    If SpreadCnt < 11 Then
      If TaxMsgWOpts(800, "Printing so few labels wastes mailing label paper. Press F10 if you wish to continue anyway. Otherwise, press ESC to stop the printing process.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        fpcmbRange.SetFocus
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
      End If
    End If
  End If

  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For cnt = 1 To NumOfCustIdx
    If SpreadCnt > 0 Then
      Get TCHandle, SpreadIdx(cnt), TaxCustRec
    End If
    CustPCnt = CustPCnt + 1
    ValidCnt = ValidCnt + 1

    If Mid(TaxCustRec.Zip, 7, 1) <> " " Then
      Zip$ = TaxCustRec.Zip
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(TaxCustRec.Zip, 5)
      Zip$ = QPTrim$(Zip$)
    End If
    LabelCnt = LabelCnt + 1 'this requires a line to be printed
    'in columns of 3 with each column containing data gathered
    'from different customers...it also must limit the size of some
    'variables to accommodate the limitations of a mailing label's
    'size
    ToPrint(1, LabelCnt) = QPTrim$(TaxCustRec.CustName)
    ToPrint(2, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr1), 23)
    ToPrint(3, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr2), 23)
    ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)

    If LabelCnt = 3 Then 'got a complete line
      For PCnt = 1 To 4
        '                       0                    1                      2
        Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3); dlm; ' ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5); dlm;
        ToPrint(PCnt, 1) = ""
        ToPrint(PCnt, 2) = ""
        ToPrint(PCnt, 3) = ""
        ToPrint(PCnt, 4) = ""
        ToPrint(PCnt, 5) = ""
      Next
      Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3) '; dlm; ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5)
      ToPrint(PCnt, 1) = ""
      ToPrint(PCnt, 2) = ""
      ToPrint(PCnt, 3) = ""
      ToPrint(PCnt, 4) = ""
      ToPrint(PCnt, 5) = ""
      LabelCnt = 0
    End If

NextLabel:
  Next

  'this catches the last
  'line of a print job if the last line contains fewer than
  'the number required to trigger another print line
  For PCnt = 1 To 4
    '                   0,3,6,9,12            1,4,7,10,13            2,5,8,11,14
    Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3); dlm; ' ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5); dlm;
  Next
  Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3) '; dlm; ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5)

  PCnt = 0

  Close

  arVATaxMailLabels.Show

  MainLog ("Mailing labels processed in graphics format for advertising charges.")

  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMailLblsAdv", "PrintGraphics", Erl)
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



