VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxLateNoticeReprint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Late Notice Reprints"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxLateNoticeReprint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6525
      Left            =   1193
      TabIndex        =   0
      Top             =   1050
      Width           =   9345
      _Version        =   196609
      _ExtentX        =   16484
      _ExtentY        =   11509
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxLateNoticeReprint.frx":08CA
      Begin LpLib.fpList fpList 
         Height          =   2232
         Left            =   960
         TabIndex        =   10
         Tag             =   $"frmVATaxLateNoticeReprint.frx":08E6
         Top             =   2520
         Width           =   7452
         _Version        =   196608
         _ExtentX        =   13144
         _ExtentY        =   3937
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
         ColDesigner     =   "frmVATaxLateNoticeReprint.frx":0A5F
      End
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   384
         Left            =   4200
         TabIndex        =   1
         Top             =   1920
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
         ColDesigner     =   "frmVATaxLateNoticeReprint.frx":0E3E
      End
      Begin EditLib.fpText fptxtCurrForm 
         Height          =   390
         Left            =   3960
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Late notices are selected on the System Setup screen."
         Top             =   1200
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5027
         _ExtentY        =   688
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
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   1
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
         ControlType     =   1
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   50
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
      Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
         Height          =   492
         Left            =   4812
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmVATaxLateNoticeReprint.frx":11A5
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
         ButtonDesigner  =   "frmVATaxLateNoticeReprint.frx":1284
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   492
         Left            =   1200
         TabIndex        =   7
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
         ButtonDesigner  =   "frmVATaxLateNoticeReprint.frx":1460
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   492
         Left            =   6456
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmVATaxLateNoticeReprint.frx":163E
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
         ButtonDesigner  =   "frmVATaxLateNoticeReprint.frx":16D9
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   492
         Left            =   3000
         TabIndex        =   9
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
         ButtonDesigner  =   "frmVATaxLateNoticeReprint.frx":18B8
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Form In Use:"
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
         Left            =   2040
         TabIndex        =   5
         Top             =   1260
         Width           =   1665
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
         Caption         =   "Reprints For Late Notices"
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3345
         Left            =   480
         Top             =   1755
         Width           =   8415
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
         Left            =   1920
         TabIndex        =   2
         Top             =   2040
         Width           =   2175
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6900
      Left            =   953
      Top             =   915
      Width           =   9735
   End
End
Attribute VB_Name = "frmVATaxLateNoticeReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim TownName$
  Dim TownAdd1$
  Dim TownAdd2$
  Dim TownCSZ$
  Dim CustRecs() As Long
  Dim CustCnt As Long
  Dim GTaxYear As Integer

Private Sub cmdClear_Click()
  fpList.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  frmVATaxLateNoticeMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim LLRec As LateListPrintType
  Dim LLHandle As Integer
  Dim NumOfLLRecs As Long
  Dim LtrType$
  Dim x As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenLatePrnFile LLHandle, NumOfLLRecs
  For x = 1 To NumOfLLRecs
    Get LLHandle, x, LLRec
    If LLRec.LtrType = "G" Then
      If QPTrim$(fptxtCurrForm.Text) = "SELF EDIT #1" Then
        Close
        Call PrintGraphicsSelfEdit1
        Exit For
      Else
        Exit Sub
      End If
    Else
      If QPTrim$(fptxtCurrForm.Text) = "SELF EDIT #1" Then
        Close
        Call PrintTextSelfEdit1
        Exit For
      Else
        Exit Sub
      End If
    End If
    Exit For
  Next x
  If x > NumOfLLRecs Then
    Close
    Call TaxMsg(900, "The late notice letter format (text or graphics) could not be determined. Please try again.")
    Exit Sub
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxLateNoticeReprint", "cmdProcess_Click", Erl)
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
      SendKeys "%E"
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
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpReprintLate
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxLateNoticeReprint.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim LLRec As LateListPrintType
  Dim LLHandle As Integer
  Dim NumOfLLRecs As Long
  Dim x As Long, y As Integer
  Dim YrCnt As Integer
  Dim ThisRec As Integer
  Dim ThatRec As Integer
  
  On Error Resume Next
  OpenLatePrnFile LLHandle, NumOfLLRecs
  ReDim CustRecs(1 To 1) As Long
  CustCnt = 0
  fpList.Enabled = False
  YrCnt = 0
  
  ReDim Years(1 To 1) As Integer
  For x = 1 To NumOfLLRecs
    Get LLHandle, x, LLRec
    If x = 1 Then
      ThatRec = LLRec.CustAcct
      ThisRec = LLRec.CustAcct
      YrCnt = YrCnt + 1
      ReDim Preserve Years(1 To YrCnt) As Integer
      Years(YrCnt) = LLRec.TaxYear
    Else
      ThisRec = LLRec.CustAcct
      If ThatRec <> ThisRec Then
        YrCnt = 1
        ReDim Years(1 To 1) As Integer
        Years(YrCnt) = LLRec.TaxYear
        ThatRec = ThisRec
      Else
        For y = 1 To YrCnt
          If LLRec.TaxYear = Years(y) Then
            GoTo SkipIt
          End If
        Next y
        If y > YrCnt Then
          YrCnt = YrCnt + 1
          ReDim Preserve Years(1 To YrCnt) As Integer
          Years(YrCnt) = LLRec.TaxYear
        End If
      End If
    End If
    CustCnt = CustCnt + 1
    ReDim Preserve CustRecs(1 To CustCnt) As Long
    CustRecs(CustCnt) = x
    fpList.InsertRow = "  " & Using$("#####", LLRec.CustAcct) & Chr$(9) & "  " & QPTrim$(LLRec.CustName) & Chr$(9) & Using$("$###,###,##0.00", LLRec.TotBal) & Chr$(9) & CStr(x)
SkipIt:
  Next x
  
  fpList.ListIndex = 0
  
  Close LLHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Select Case TaxMasterRec.LateForm
    Case 0:
      fptxtCurrForm.Text = "None Saved"
    Case 1:
      fptxtCurrForm.Text = "SELF EDIT #1"
    Case Else
  End Select
  TownName = QPTrim$(TaxMasterRec.Name)
  TownAdd1 = QPTrim$(TaxMasterRec.Add1)
  TownAdd2 = QPTrim$(TaxMasterRec.Add2)
  TownCSZ$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TaxSt) + "  " + QPTrim$(TaxMasterRec.Zip)
  GTaxYear = CInt(TaxMasterRec.RTaxYear)
  
  fpcmbRange.Text = "ALL"
  fpcmbRange.AddItem "ALL"
  fpcmbRange.AddItem "SELECT FROM LIST"
  
End Sub

Private Sub fpcmbRange_Change()
  If fpcmbRange.Text = "ALL" Then
    fpList.Action = ActionDeselectAll
    fpList.Enabled = False
  Else
    fpList.Enabled = True
  End If

End Sub

Private Sub PrintGraphicsSelfEdit1()
  Dim dlm$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim LLRec As LateListPrintType
  Dim LLHandle As Integer
  Dim NumOfLLRecs As Long
  Dim x As Long, y As Integer
  Dim AllFlag As Boolean
  Dim LLtrRec As TAXLateLetterType
  Dim SpreadCnt As Long
  Dim ListCnt As Long
  Dim LtrDate$
  
  On Error GoTo ERRORSTUFF
  
  AllFlag = True
  If fpcmbRange.Text <> "ALL" Then
    AllFlag = False
  End If
  dlm$ = "~"
  OpenLateLtrFile LLHandle 'letter format data
  Get LLHandle, 1, LLtrRec
  Close LLHandle
  
  ReDim SpreadIdx(1 To 1) As Long
  SpreadCnt = 0
  ListCnt = fpList.ListCount
  If AllFlag = False Then
    For x = 0 To ListCnt - 1
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
  
  RptFile$ = "TAXRPTS\LATENOTICE.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenLatePrnFile LLHandle, NumOfLLRecs
  For x = 1 To NumOfLLRecs
    Get LLHandle, x, LLRec
    If LLRec.LtrDate > 0 Then
      LtrDate = MakeRegDate(LLRec.LtrDate)
      Exit For
    End If
  Next x
  If x > NumOfLLRecs Then
    Close
    Call TaxMsg(800, "The late letter date could not be determined. Please rerun late notice letters making sure the letter date is accurate.")
    Exit Sub
  End If
  If NumOfLLRecs = 0 Then
    Call TaxMsg(900, "There are no late notices necessary for the parameters entered.")
    Close
    Exit Sub
  End If
  For x = 1 To SpreadCnt
    Get LLHandle, SpreadIdx(x), LLRec
    '                          0                           1                       2
    Print #RptHandle, QPTrim$(LLRec.Addr1); dlm; QPTrim$(LLRec.Addr2); dlm; LLRec.AdvBal; dlm;
    '                            3                             4                            5
    Print #RptHandle, MakeRegDate(LLRec.AdvDate); dlm; QPTrim$(LLRec.City); dlm; QPTrim$(LLRec.CustName); dlm;
    '                       6                    7                      8                     9
    Print #RptHandle, LLRec.IntBal; dlm; LLRec.LateListBal; dlm; LLRec.LateSeqNum; dlm; LLRec.Opt1Bal; dlm;
    '                       10                  11                       12                      13
    Print #RptHandle, LLRec.Opt2Bal; dlm; LLRec.Opt3Bal; dlm; MakeRegDate(LLRec.PayDate); dlm; LLRec.PersExemp; dlm;
    '                       14                  15                     16                     17
    Print #RptHandle, LLRec.PersValue; dlm; LLRec.PrincBal; dlm; LLRec.RealExemp; dlm; LLRec.RealValue; dlm;
    '                       18                   19                           20                         21
    Print #RptHandle, QPTrim$(LLRec.State); dlm; LLRec.TaxYear; dlm; QPTrim$(LLRec.TownName); dlm; QPTrim$(LLRec.Zip); dlm;
    '                       22                   23                           24                25
    Print #RptHandle, QPTrim$(TownAdd1); dlm; QPTrim$(TownAdd2); dlm; QPTrim$(TownCSZ); dlm; LtrDate; dlm;
    '                       26                   27                28              29                30
    Print #RptHandle, LLtrRec.Head1; dlm; LLtrRec.Head2; dlm; LLtrRec.Head3; dlm; LLtrRec.Head4; dlm; LLtrRec.Head5; dlm;
    
    For y = 1 To 20
      '31 - 50
      Print #RptHandle, LLtrRec.Body(y); dlm;
    Next y
    '                     51                  52                   53                  54                55              56
    Print #RptHandle, LLRec.TotBal; dlm; LLRec.CurrBal; dlm; LLRec.PrevBal; dlm; LLRec.CustAcct; dlm; GTaxYear; dlm; LLRec.NegYN
  Next x
  
  Close
  
  arVATaxLateLetter.Show
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxLateNoticeReprint", "PrintGraphicsSelfEdit1", Erl)
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
'  If fpcmbRange.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      If fpList.Enabled = True Then
'        fpList.SetFocus
'      End If
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If

End Sub

Private Sub fptxtCurrForm_Change()
  If fptxtCurrForm.Text = "SELF EDIT #1" Then
    cmdAlign.Enabled = True
  Else
    cmdAlign.Enabled = False
  End If
End Sub
Private Sub PrintTextSelfEdit1()
  Dim RptFile$
  Dim RptHandle As Integer
  Dim LLRec As LateListPrintType
  Dim LLHandle As Integer
  Dim NumOfLLRecs As Long
  Dim x As Long, y As Integer
  Dim LLtrRec As TAXLateLetterType
  Dim AllFlag As Boolean
  Dim FF$
  Dim HdrLen As Integer
  Dim Start1 As Integer
  Dim Start2 As Integer
  Dim Start3 As Integer
  Dim Start4 As Integer
  Dim Start5 As Integer
  Dim SpreadCnt As Long
  Dim ListCnt As Long
  Dim LtrDate$
  
  On Error GoTo ERRORSTUFF
  
  AllFlag = True
  If fpcmbRange.Text <> "ALL" Then
    AllFlag = False
  End If
  FF$ = Chr(12)
  OpenLateLtrFile LLHandle
  Get LLHandle, 1, LLtrRec
  Close LLHandle
  ReDim SpreadIdx(1 To 1) As Long
  SpreadCnt = 0
  ListCnt = fpList.ListCount
  If AllFlag = False Then
    For x = 0 To ListCnt - 1
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
  HdrLen = Len(QPTrim$(LLtrRec.Head1))
  HdrLen = HdrLen / 2
  Start1 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head2))
  HdrLen = HdrLen / 2
  Start2 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head3))
  HdrLen = HdrLen / 2
  Start3 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head4))
  HdrLen = HdrLen / 2
  Start4 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head5))
  HdrLen = HdrLen / 2
  Start5 = 40 - HdrLen
  
  RptFile$ = "TAXRPTS\LATENOTICE.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenLatePrnFile LLHandle, NumOfLLRecs
  If NumOfLLRecs = 0 Then
    Call TaxMsg(900, "There are no late notices necessary for the parameters entered.")
    Close
    Exit Sub
  End If
  For x = 1 To SpreadCnt
    Get LLHandle, SpreadIdx(x), LLRec
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(Start1); QPTrim$(LLtrRec.Head1)
    Print #RptHandle, Tab(Start2); QPTrim$(LLtrRec.Head2)
    Print #RptHandle, Tab(Start3); QPTrim$(LLtrRec.Head3)
    Print #RptHandle, Tab(Start4); QPTrim$(LLtrRec.Head4)
    Print #RptHandle, Tab(Start5); QPTrim$(LLtrRec.Head5)
    Print #RptHandle,
    Print #RptHandle, MakeRegDate(LLRec.LtrDate)
    Print #RptHandle,
    Print #RptHandle, QPTrim$(LLRec.CustName)
    Print #RptHandle, QPTrim$(LLRec.Addr1)
    Print #RptHandle, QPTrim$(LLRec.Addr2)
    Print #RptHandle, QPTrim$(LLRec.City) + ", " + QPTrim$(LLRec.State) + "  " + QPTrim$(LLRec.Zip)
    Print #RptHandle,
    
    For y = 1 To 10
      Print #RptHandle, LLtrRec.Body(y)
    Next y
    
    Print #RptHandle,
    If LLRec.NegYN = "N" Then
      If LLRec.TaxYear = GTaxYear Then
        Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Prev Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
      Else
        Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Other Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
      End If
      Print #RptHandle, Tab(5); "Tax Year: "; Tab(25); Using("###0", LLRec.TaxYear); Tab(33); "Curr Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.CurrBal)
      Print #RptHandle, Tab(33); "Total Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.TotBal)
      Print #RptHandle,
      For y = 11 To 20
        Print #RptHandle, LLtrRec.Body(y)
      Next y
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle, FF$
    ElseIf LLRec.NegYN = "Y" Then
      If LLRec.PrevBal >= 0 Then
        If LLRec.TaxYear = GTaxYear Then
          Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Prev Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
        Else
          Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Other Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
        End If
        Print #RptHandle, Tab(5); "Tax Year: "; Tab(25); Using("###0", LLRec.TaxYear); Tab(33); "Curr Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.CurrBal)
        Print #RptHandle, Tab(33); "Total Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.TotBal)
        Print #RptHandle,
        For y = 11 To 20
          Print #RptHandle, LLtrRec.Body(y)
        Next y
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle, FF$
      ElseIf LLRec.PrevBal < 0 Then
        Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct)
        Print #RptHandle, Tab(5); "Tax Year: "; Tab(25); Using("###0", LLRec.TaxYear); Tab(33); "Total Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.TotBal)
        Print #RptHandle,
        For y = 11 To 20
          Print #RptHandle, LLtrRec.Body(y)
        Next y
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle, FF$
      End If
    End If
  Next x
  
  Close
  
  ViewPrint RptFile, "Printing Late Notice Letters", True
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxLateNoticeReprint", "PrintTextSelfEdit1", Erl)
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

