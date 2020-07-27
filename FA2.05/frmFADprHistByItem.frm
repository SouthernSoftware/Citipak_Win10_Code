VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFADprHistByItem 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depreciation History By Item"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11580
   Icon            =   "frmFADprHistByItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5220
      Left            =   1560
      TabIndex        =   3
      Top             =   1812
      Width           =   8460
      _Version        =   196609
      _ExtentX        =   14922
      _ExtentY        =   9208
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFADprHistByItem.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3450
         TabIndex        =   2
         ToolTipText     =   "Select the method you wish to have this report displayed."
         Top             =   3060
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
         ColDesigner     =   "frmFADprHistByItem.frx":08E6
      End
      Begin EditLib.fpText fptxtFirst 
         Height          =   396
         Left            =   3600
         TabIndex        =   0
         ToolTipText     =   "Enter the first item tag number from which you wish to begin this report."
         Top             =   1656
         Width           =   2172
         _Version        =   196608
         _ExtentX        =   3831
         _ExtentY        =   698
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         Text            =   ""
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - A L a l"
         MaxLength       =   20
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtLast 
         Height          =   396
         Left            =   3600
         TabIndex        =   1
         ToolTipText     =   "Enter the last item tag number to which you wish to end this report."
         Top             =   2352
         Width           =   2172
         _Version        =   196608
         _ExtentX        =   3831
         _ExtentY        =   698
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         Text            =   ""
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - A L a l"
         MaxLength       =   20
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1812
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFADprHistByItem.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4776
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFADprHistByItem.frx":0DB9
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdTag 
         Height          =   1785
         Left            =   6150
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a complete list of all fixed assets."
         Top             =   1635
         Width           =   915
         _Version        =   131072
         _ExtentX        =   1614
         _ExtentY        =   3149
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFADprHistByItem.frx":0F98
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "First Tag Number"
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
         Left            =   1392
         TabIndex        =   7
         Top             =   1764
         Width           =   2040
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1824
         Top             =   432
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciation History By Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   1875
         TabIndex        =   6
         Top             =   570
         Width           =   4815
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
         Height          =   348
         Left            =   1776
         TabIndex        =   5
         Top             =   3120
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Tag Number"
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
         Left            =   1392
         TabIndex        =   4
         Top             =   2460
         Width           =   2040
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   5532
      Left            =   1404
      Top             =   1668
      Width           =   8772
   End
End
Attribute VB_Name = "frmFADprHistByItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim First$
  Dim Last$

Private Sub cmdExit_Click()
  frmFAReportMenu.Show
  Close
  KillFile ("dprhistbyitemrpt.dat") 'these files are created as identifiers
  'so other programs know where to send data when called by this form
  KillFile "taglistopen.dat"
  DoEvents
  Unload frmFADprHistByItem

End Sub

Private Sub cmdProcess_Click()
  If Exist("taglistopen.dat") Then 'one of the tag values probably
  'came from tag list
    KillFile "taglistopen.dat"
  End If
  If Check4ValidTag(ReplaceString$(fptxtFirst.Text, "-", "")) = False Then
    MsgBox "The value entered for First Tag Number is not valid. Please try again."
    fptxtFirst.SetFocus
    Exit Sub
  ElseIf Check4ValidTag(ReplaceString(fptxtLast.Text, "-", "")) = False Then
    MsgBox "The value entered for Last Tag Number is not valid. Please try again."
    fptxtLast.SetFocus
    Exit Sub
  End If
  
  If Val(fptxtFirst.Text) > Val(fptxtLast.Text) Then
    MsgBox "The First Number is greater than the Last Number. Please correct this error."
    fptxtFirst.SetFocus
    Close
    Exit Sub
  End If
  
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    MsgBox "Pitch 12 is recommended for this report."
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim First$
  Dim Last$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DeptNumber As Integer
  Dim x As Long
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim Page As Integer
  Dim DHHandle As Integer
  Dim DprRec As DprHistType
  Dim DprHistCnt As Long
  Dim DeptDescHeader$, DeptDescription$
  Dim ThisTag$
  Dim ThisRec As Long
  Dim LastYrCnt As Integer
  Dim Y As Integer
  Dim YrCompare$
  Dim DisposalYear$
  
  On Error GoTo ERRORSTUFF
  ReportFile$ = "FADEPRRPT.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)

  MaxLines = 57
  LineCnt& = 0
  ItemCnt& = 0
  First$ = ReplaceString(fptxtFirst.Text, "-", "")
  Last$ = ReplaceString(fptxtLast.Text, "-", "")
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  GoSub PrintMasterHeader1

  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle


  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  
  OpenDprHistFile DHHandle
  DprHistCnt = (LOF(DHHandle) / Len(DprRec))
  If DprHistCnt = 0 Then
    MsgBox "No Depreciation History Records on file."
    Close
    Exit Sub
  End If
  LastYrCnt = 1
  ReDim YearCnt(1 To LastYrCnt) As String * 4
  Get DHHandle, 1, DprRec
  
  YearCnt(LastYrCnt) = QPTrim$(DprRec.DprYear)
  
  For x = 2 To DprHistCnt 'determine the years that
  'have been depreciated and store them in YearCnt()
    Get DHHandle, x, DprRec
    YrCompare = QPTrim$(DprRec.DprYear)
    If YrCompare = "" Then GoTo SameYear
    For Y = 1 To LastYrCnt
      If YrCompare = YearCnt(Y) Then
        GoTo SameYear
      End If
    Next Y
    If Y = LastYrCnt + 1 Then
      LastYrCnt = LastYrCnt + 1 'counts the number of depreciated years
      ReDim Preserve YearCnt(1 To LastYrCnt)
      YearCnt(LastYrCnt) = DprRec.DprYear
    End If
SameYear:
  Next x
  
  ReDim YrItems(1 To LastYrCnt) As Integer
  ReDim PurPr(1 To LastYrCnt) As Double
  ReDim ThisYrDpr(1 To LastYrCnt) As Double
  ReDim AllDprToDate(1 To LastYrCnt) As Double
  ReDim CurrYrVal(1 To LastYrCnt) As Double
  
  frmFAShowPctComp.Label1 = "Gathering Depreciation Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  Nextx = 1
  
  For cnt& = 1 To NumOfFARecs
    Get FAHandle, TagIdxRecs(Nextx), FAItemRec 'retrieve fixed assets in numerical order
    ThisTag = ReplaceString(FAItemRec.ItemTag, "-", "")
    If Val(ThisTag) < Val(First) Or Val(ThisTag) > Last Then 'only include tag parameters
    'entered on the screen
      GoTo NoData
    End If
    ThisRec = FAItemRec.LastDprRec 'start here to retrieve depreciation data for this asset
    If ThisRec = 0 Then GoTo NoData
    DisposalYear = ""
    If FAItemRec.DsplFlag = 2 Then DisposalYear = "*D*" 'flags assets that have been disposed
    Print #RptHandle, QPTrim$(FAItemRec.ItemTag); Tab(21); DisposalYear; Tab(25); QPTrim$(FAItemRec.IDESC1);
    Print #RptHandle, Tab(58); FAItemRec.IDEPT; Tab(65); MakeRegDate(FAItemRec.AQURDATE); Tab(75); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST)) ';
    Print #RptHandle, String$(88, "-")
    Print #RptHandle, Tab(5); "Year"; Tab(18); "Life/Left"; Tab(32); "Deprec For Year"; Tab(55); "Deprec To Date"; Tab(75); "Year End Value"
    LineCnt = LineCnt + 3
    
    Do
      Get DHHandle, ThisRec, DprRec 'now start printing out this asset's depreciation data
      If LineCnt& >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader1
      End If
      Print #RptHandle, Tab(5); DprRec.DprYear;
      Print #RptHandle, Tab(19); Using("##0", DprRec.Life); "/"; Using("#0", DprRec.LifeLeft);
      Print #RptHandle, Tab(33); Using("###,###,##0.00", CStr(DprRec.DprAmt));
      Print #RptHandle, Tab(55); Using("###,###,##0.00", CStr(DprRec.DprToDate));
      Print #RptHandle, Tab(75); Using("###,###,##0.00", CStr(DprRec.BookTotal))
      For x = 1 To LastYrCnt 'collect each year's data for totals
        If QPTrim$(DprRec.DprYear) = YearCnt(x) Then
          YrItems(x) = YrItems(x) + 1
          PurPr(x) = PurPr(x) + DprRec.OrigCost
          ThisYrDpr(x) = ThisYrDpr(x) + DprRec.DprAmt
          AllDprToDate(x) = AllDprToDate(x) + DprRec.DprToDate
          CurrYrVal(x) = CurrYrVal(x) + DprRec.BookTotal
          Exit For 'stop when you come up on the matching year
        End If
      Next x
      
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      If DprRec.PrevDprRec <= 0 Then
        Exit Do
      Else
        ThisRec = DprRec.PrevDprRec
      End If

    Loop
    
    Print #RptHandle, String$(88, "=")
    Print #RptHandle,
    LineCnt& = LineCnt& + 2
NoData:
    frmFAShowPctComp.ShowPctComp Nextx, NumOfFARecs
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    Nextx = Nextx + 1
  Next cnt&

  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True

  GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi

  Close         'Close all open files now

  ViewPrint ReportFile$, "Depreciation Listing", True

  KillFile (ReportFile$)
  
  Exit Sub

PrintMasterHeader1:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Master Asset Listing : Depreciation History by Item"
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "*D* = This Item Has Been Disposed"
  Print #RptHandle, "Tag Number"; Tab(29); "Description"; Tab(58); "Dept"; Tab(64); "Acquired On"; Tab(78); "Purch Price"
  Print #RptHandle, String$(88, "=")
  LineCnt& = 5
  
  Return

PrintMasterValueEnding1:
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle,
  Print #RptHandle, Tab(2); "Year"; Tab(10); "# Items"; Tab(20); "Purchase Price"; Tab(38); "Depr For Year"; Tab(58); "Depr To Date"; Tab(77); "Book Value"
  Print #RptHandle, String$(88, "=")
  
  For x = LastYrCnt To 1 Step -1
    Print #RptHandle, Tab(2); YearCnt(x); Tab(11); YrItems(x);
    Print #RptHandle, Tab(20); Using("###,###,##0.00", CStr(PurPr(x)));
    Print #RptHandle, Tab(37); Using("###,###,##0.00", CStr(ThisYrDpr(x)));
    Print #RptHandle, Tab(56); Using("###,###,##0.00", CStr(AllDprToDate(x)));
    Print #RptHandle, Tab(73); Using("###,###,##0.00", CStr(CurrYrVal(x)))
  Next x
  Print #RptHandle, FF$

  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADprHistByItem", "PrintText", Erl)
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
    Unload Me
End Sub

Private Sub cmdTag_Click()
  frmFATagList.Show vbModal
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
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
      SendKeys "%E"
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      cmdProcess_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%T"
      Call cmdTag_Click
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
      KillFile ("dprhistbyitemrpt.dat")
      KillFile "taglistopen.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFADprHistByItem.")
      Call Terminate
      End
    End If
  End If
End Sub


Private Sub LoadMe()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  Dim One As Integer
  Dim FileHandle As Integer
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  
  Get TagIdxHandle, 1, TagIdx
  fptxtFirst = QPTrim$(TagIdx.TagNumb) 'load the first employee number
  First = QPTrim$(TagIdx.TagNumb)
  Get TagIdxHandle, TagIdxCnt, TagIdx
  Close TagIdxHandle
  fptxtLast = QPTrim$(TagIdx.TagNumb) 'load the last employee number
  Last = QPTrim$(TagIdx.TagNumb)
  One = 1
  FileHandle = FreeFile
  Open "dprhistbyitemrpt.dat" For Output As FileHandle Len = 2 'tells
  'dept list that the request for a dept number from the list comes
  'from here
  Print #FileHandle, One
  Close FileHandle
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  
End Sub

Private Function Check4ValidTag(ThisTag$) As Boolean
  Dim x As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TIdxHandle As Integer
  Dim TIdxRecNums As Long
  
'  Call CreateTagIdx
  On Error GoTo ERRORSTUFF
  'can't investigate the depreciation history on a bogus tag number
  Check4ValidTag = True 'assume what is entered is valid
  OpenTagIdxFile TIdxHandle
  TIdxRecNums = LOF(TIdxHandle) \ Len(TagIdx)
  If TIdxRecNums = 0 Then
    MsgBox "No tag numbers saved in index."
    Close
    Check4ValidTag = False 'gotta save something first...anytime a new
    'asset is saved it is indexed
    Exit Function
  End If

  For x = 1 To TIdxRecNums
    Get TIdxHandle, x, TagIdx
    If ThisTag$ = ReplaceString(TagIdx.TagNumb, "-", "") Then
      Close 'number entered is now validated
      Exit Function 'Check4ValidTag is true because
      'when the function opened it was set to true
    End If
  Next x

  Check4ValidTag = False
  Close
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADprHistByItem", "Check4ValidTag", Erl)
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
    Unload Me
End Function

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdExit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtFirst_LostFocus()
  If QPTrim$(fptxtFirst) = "" Then
    fptxtFirst.Text = First
    Exit Sub
  End If
  First = QPTrim$(fptxtFirst.Text)
End Sub

Private Sub fptxtLast_LostFocus()
  If QPTrim$(fptxtLast) = "" Then
    fptxtLast.Text = Last
    Exit Sub
  End If

  Last = QPTrim$(fptxtLast.Text)
End Sub

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim ItemCnt&
  Dim First$
  Dim Last$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim x As Long
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim DHHandle As Integer
  Dim DprRec As DprHistType
  Dim DprHistCnt As Long
  Dim ThisTag$
  Dim ThisRec As Long
  Dim LastYrCnt As Integer
  Dim Y As Integer
  Dim YrCompare$
  Dim SubReportFile$
  Dim SubRptHandle As Integer
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim Employer$, dlm$
  Dim DisposalDate$
  Dim PrintCnt As Integer
  
  On Error GoTo ERRORSTUFF
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Employer = QPTrim$(FASetUpRec.TownName)
  Close FASHandle
  dlm$ = "~"
  ReportFile$ = "FARPTS\FAITEMHIST.RPT"  'Report File Name
  SubReportFile$ = "FARPTS\SUBITEMHIST.RPT"
  ItemCnt& = 0
  First$ = ReplaceString(fptxtFirst.Text, "-", "")
  Last$ = ReplaceString(fptxtLast.Text, "-", "")

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load array with assets in numerical order
  Next x
  Close TagIdxHandle

  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)

  OpenDprHistFile DHHandle
  DprHistCnt = (LOF(DHHandle) / Len(DprRec))
  If DprHistCnt = 0 Then
    MsgBox "No Depreciation History Records on file."
    Close
    Exit Sub
  End If

  LastYrCnt = 1
  ReDim YearCnt(1 To LastYrCnt) As String * 4
  Get DHHandle, 1, DprRec
  YearCnt(LastYrCnt) = QPTrim$(DprRec.DprYear) 'YearCnt(1) assigned the first year in the file
  For x = 2 To DprHistCnt 'trying to boil out the years that have been depreciated
  'and how many years have been depreciated
    Get DHHandle, x, DprRec
    YrCompare = QPTrim$(DprRec.DprYear) 'get a year from the file
    If YrCompare = "" Then GoTo SameYear
    For Y = 1 To LastYrCnt
      If YrCompare = YearCnt(Y) Then 'this year has already been found
        GoTo SameYear
      End If
    Next Y
    If Y = LastYrCnt + 1 Then 'now we've been through all dates saved so far
    'and none has matched with DprRec.DprYear
      LastYrCnt = LastYrCnt + 1 'increment valid year count
      ReDim Preserve YearCnt(1 To LastYrCnt) 'enlarge array
      YearCnt(LastYrCnt) = QPTrim$(DprRec.DprYear) 'assign last element the latest valid year
    End If
SameYear:
  Next x

  ReDim YrItems(1 To LastYrCnt) As Integer
  ReDim PurPr(1 To LastYrCnt) As Double
  ReDim ThisYrDpr(1 To LastYrCnt) As Double
  ReDim AllDprToDate(1 To LastYrCnt) As Double
  ReDim CurrYrVal(1 To LastYrCnt) As Double

  frmFAShowPctComp.Label1 = "Gathering Depreciation Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
  Nextx = 1
  For cnt& = 1 To NumOfFARecs
    Get FAHandle, TagIdxRecs(Nextx), FAItemRec 'retrieve assets in numerical order
    ThisTag = ReplaceString(FAItemRec.ItemTag, "-", "")
    If Val(ThisTag) < Val(First) Or Val(ThisTag) > Last Then 'only include assets allowed
    'according to parameters set on screen
      GoTo NoData
    End If
    ThisRec = FAItemRec.LastDprRec
    If ThisRec = 0 Then GoTo NoData
    DisposalDate = ""
    If FAItemRec.DsplFlag = 2 Then DisposalDate = MakeRegDate(FAItemRec.DispDate)
    Do
      '                     0                         1                              2
      Print #RptHandle, Employer$; dlm; QPTrim$(FAItemRec.ItemTag); dlm; QPTrim$(FAItemRec.IDESC1); dlm;
      '                        3                       4
      Print #RptHandle, FAItemRec.IDEPT; dlm; FAItemRec.ORGCOST; dlm;
      Get DHHandle, ThisRec, DprRec
      If DprRec.SoSoftFlag = True Then 'This year's depreciation has been reversed by Southern Software
      'and has not been re-depreciated
        '                        5
        Print #RptHandle, DprRec.DprYear + "Rvrs"; dlm;
      Else
        '                        5
        Print #RptHandle, DprRec.DprYear; dlm;
      End If
      '                        6                  7
      Print #RptHandle, DprRec.Life; dlm; DprRec.LifeLeft; dlm;
      '                        8
      Print #RptHandle, DprRec.DprAmt; dlm;
      '                        9
      Print #RptHandle, DprRec.DprToDate; dlm;
      '                       10                            11
      Print #RptHandle, DprRec.BookTotal; dlm; MakeRegDate(FAItemRec.AQURDATE); dlm;
      '                       12
      Print #RptHandle, DisposalDate
      
      For x = 1 To LastYrCnt 'collect data for year totals
        If QPTrim$(DprRec.DprYear) = YearCnt(x) Then
          YrItems(x) = YrItems(x) + 1
          PurPr(x) = PurPr(x) + DprRec.OrigCost
          ThisYrDpr(x) = ThisYrDpr(x) + DprRec.DprAmt
          AllDprToDate(x) = AllDprToDate(x) + DprRec.DprToDate
          CurrYrVal(x) = CurrYrVal(x) + DprRec.BookTotal
          Exit For
        End If
      Next x
      PrintCnt = PrintCnt + 1
      ItemCnt& = ItemCnt& + 1
      If DprRec.PrevDprRec <= 0 Then
        Exit Do
      Else
        ThisRec = DprRec.PrevDprRec
      End If
  Loop
NoData:
    Nextx = Nextx + 1
    frmFAShowPctComp.ShowPctComp cnt, NumOfFARecs
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
  Next cnt&

  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True

  GoSub PrintMasterValueEnding1

  Close         'Close all open files now
  
  If PrintCnt = 0 Then
    MsgBox "No depreciation records saved for this range of fixed assets."
    Exit Sub
  End If
  
  arFAItemHistRpt.Show
  frmFALoadReport.Show

  Exit Sub

PrintMasterValueEnding1:
  SubRptHandle = FreeFile
  Open SubReportFile$ For Output As #SubRptHandle

  For x = LastYrCnt To 1 Step -1 'this data was loaded from the first year to the last but needs
  'to be displayed from last year to first
    If YrItems(x) = 0 Then GoTo NoItems
    '                        0                1                2
    Print #SubRptHandle, YearCnt(x); dlm; YrItems(x); dlm; PurPr(x); dlm;
    '                        3                     4                   5
    Print #SubRptHandle, ThisYrDpr(x); dlm; AllDprToDate(x); dlm; CurrYrVal(x)
NoItems:
  Next x
  
  Close SubRptHandle
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADprHistByItem", "PrintGraphics", Erl)
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
    Unload Me
  
  
End Sub

