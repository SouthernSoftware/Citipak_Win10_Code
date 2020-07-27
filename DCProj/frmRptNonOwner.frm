VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptNonOwner 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NonOwner Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptNonOwner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5256
      TabIndex        =   2
      Top             =   3732
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      Columns         =   1
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
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
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptNonOwner.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5256
      TabIndex        =   3
      Top             =   4272
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      Columns         =   1
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
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      ColDesigner     =   "frmRptNonOwner.frx":0BED
   End
   Begin LpLib.fpCombo fpDecalCat 
      Height          =   348
      Left            =   5256
      TabIndex        =   1
      Top             =   3168
      Width           =   3996
      _Version        =   196608
      _ExtentX        =   7048
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      Columns         =   3
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   1
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
      AutoSearchFillDelay=   500
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptNonOwner.frx":0F1B
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   492
      Left            =   10080
      TabIndex        =   5
      Top             =   7560
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
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
      Height          =   492
      Left            =   8400
      TabIndex        =   4
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "5:07 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "11/14/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtOpnDate 
      Height          =   348
      Left            =   5256
      TabIndex        =   0
      Top             =   2592
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      Text            =   "11/06/2001"
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   0
      Left            =   3024
      TabIndex        =   11
      Top             =   3168
      Width           =   2076
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acct Open Date As Of:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   2448
      TabIndex        =   10
      Top             =   2640
      Width           =   2676
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Order:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   7
      Left            =   3408
      TabIndex        =   9
      Top             =   3756
      Width           =   1716
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2820
      Left            =   2364
      Top             =   2184
      Width           =   7284
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2808
      TabIndex        =   8
      Top             =   4296
      Width           =   2388
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3192
      Top             =   600
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print NonOwner Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3624
      TabIndex        =   7
      Top             =   768
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3192
      Top             =   480
      Width           =   5772
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
Attribute VB_Name = "frmRptNonOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmDCReportsMenu.Show
  Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via RptTransJournal by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtOpnDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpDecalCat.SetFocus
  End If
End Sub

Private Sub fpDecalCat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpDecalCat.ListDown = True
  End If
  If fpDecalCat.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpDecalCat.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPrintOrder.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Function ValidDate()
  If CheckValDate(txtOpnDate) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    ValidDate = True
  End If
End Function

Private Sub cmdPrint_Click()
  If ValidDate Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 1 Then
      OwnerListing
      ActivateControls Me, True
    ElseIf fpcboRptType.ListIndex = 0 Then
      OwnerListing
    Else
      ActivateControls Me, True
    End If
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  txtOpnDate.Text = Format(Now, "mm/dd/yyyy")
  FillCatCMBOAll fpDecalCat
  fpDecalCat.ListIndex = 0
  Me.HelpContextID = hlpNonOwnerReport
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub OwnerListing()
  Dim DCCustRecLen As Integer, Page As Integer
  Dim UsingName As Boolean, TotalBal As Double, TCnt As Long
  Dim CustomerCnt As Long, UsingAcct As Boolean, Dothiscust As Boolean
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, DCRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfDCRecs As Long, AcctNo As Long
  Dim DCFile As Integer, ToPrint As String, Header As String
  Dim CatSel As String, Catdo As Boolean, DCvFile As Integer
  Dim RptType As Integer, DCCFile As Integer, NumOfVRecs As Long
  Dim Dash80 As String, CustCnt As Long, VehRecord As Long
  Dim cnt As Long, RptHandle As Integer, Category As String
  Dim ReportFile As String, DCVehReclen As Integer, ODate As Integer
  Dim DCCodeRecLen As Integer, NumCodeRecs As Integer, OPdate As String
  ReDim DCCustREc(1) As DCCustRecType
  ReDim DCCodeRec(1) As DCCatCodeRecType
  On Local Error GoTo ERRORSTUFF
  Dothiscust = False
  RptType = fpcboRptType.ListIndex
  DCCodeRecLen = Len(DCCodeRec(1))
  DCCustRecLen = Len(DCCustREc(1))
  NumCodeRecs = FileSize(DCPath + "DCCODE.DAT") \ DCCodeRecLen
  Dash120$ = String$(121, "-")
  ODate = Date2Num(txtOpnDate)
  FrmShowPctComp.Label1 = "Creating NonOwner List"
  FrmShowPctComp.Show , Me

  If fpDecalCat.ListIndex <> 0 Then
    fpDecalCat.col = 1
    CatSel = QPTrim(fpDecalCat.ColText)
    Catdo = True
    fpDecalCat.col = 2
    Category$ = CatSel + " " + QPTrim(fpDecalCat.ColText)
  Else
    Category$ = "ALL"
    CatSel = "ALL"
    Catdo = False
  End If

  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    If Not Exist(DCPath$ + "DCCUST.IDX") Then
      Unload FrmShowPctComp
      MsgBox "Missing Name Index, Please Reindex then try again.", vbOKOnly, "Missing File"
      Close
      ActivateControls Me, True
      Exit Sub
    End If
    IndexName$ = DCPath$ + "DCCUST.IDX"
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case Else
  End Select
  ReportFile$ = "DCNOWNRP.PRN" 'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  Linecnt = 0
  CustCnt = 0
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  'PRINT #RptHandle, CHR$(27); CHR$(58); ' oki 320 12 cpi
  GoSub PrintOwnRepRptHeader
  ' Print Main Body
  If UsingName = True Then
    NumOfDCRecs = FileSize(IndexName$) \ 4
    ReDim IndexArray(1 To NumOfDCRecs) As DCTempIDXRecType
    'FGetAH IndexName$, IndexArray(1), , NumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = 4
    For cnt& = 1 To NumOfDCRecs
      Get #Handle, cnt&, IndexArray(cnt&)
    Next
    Close Handle
  Else
    NumOfDCRecs = FileSize(DCPath$ + "DCCUST.DAT") \ DCCustRecLen
  End If
  'Open Vehicle File
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  
  DCFile = FreeFile
  Open DCPath + "DCCUST.DAT" For Random Shared As DCFile Len = DCCustRecLen
  For cnt = 1 To NumOfDCRecs
  'If cnt = NumOfRecs Then Stop
    FrmShowPctComp.ShowPctComp cnt, NumOfDCRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitPrint
    End If

    If UsingAcct Then
      AcctNo& = cnt
    ElseIf UsingName Then
      AcctNo& = IndexArray(cnt).IDXRECORD
    End If
    Get DCFile, AcctNo&, DCCustREc(1)
    If DCCustREc(1).Deleted <> "Y" Then
      If Linecnt >= MaxLines Then
        If RptType = 1 Then Print #RptHandle, FF$
        GoSub PrintOwnRepRptHeader
      End If
      If DCCustREc(1).Owner <> "N" Then
        GoTo SkipRecord
      End If
      If Not DCCustREc(1).DATEOPED <= ODate Then
        GoTo SkipRecord
      End If
      OPdate = Num2Date$(DCCustREc(1).DATEOPED)
      VehRecord = DCCustREc(1).FirstCar
      If VehRecord > 0 Then Dothiscust = True
      While VehRecord <> 0
        Get DCvFile, VehRecord, DCVRec(1)
        If RTrim$(DCVRec(1).DecalCat) = CatSel Or CatSel = "ALL" Then
          If DCVRec(1).Active <> "N" Then
            GoSub ProcessThisOne
          End If
        End If
        If VehRecord <> DCVRec(1).NextRec Then
          VehRecord = DCVRec(1).NextRec
        Else
          VehRecord = 0
        End If
      Wend
ExitWhile:
   End If
SkipRecord:
  Next cnt
    GoSub PrintOwnRepRptEnding
    If RptType <> 0 Then Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
    Close         'Close all open files now
    If RptType = 1 Then
      ViewPrint ReportFile$, Header$
      Kill ReportFile$
    Else
      Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmRptNonOwner
      ARptOwnerList.txtDate = Now
      ARptOwnerList.txtTown = TOWNNAME$
      ARptOwnerList.lblCat = "Category -" + Category$
      ARptOwnerList.Title = "Customer NonOwner Report"
      ARptOwnerList.totcust = CustCnt
      ARptOwnerList.GetName ReportFile$
      ARptOwnerList.startrpt
    End If
ExitPrint:
  Exit Sub
ProcessThisOneCust:
    If RptType = 1 Then
      Print #RptHandle, String$(80, "-")
      Print #RptHandle, Str(AcctNo&);
      Print #RptHandle, Tab(10); QPTrim$(DCCustREc(1).BILLNAME); Tab(50); DCCustREc(1).SOSEC; Tab(65); OPdate$
    End If
    CustCnt = CustCnt + 1
    Dothiscust = False
Return
ProcessThisOne:
  If Dothiscust Then GoSub ProcessThisOneCust
    If RptType = 1 Then
      Print #RptHandle, RTrim$(DCVRec(1).makemodel) + " " + RTrim$(DCVRec(1).Desc); Tab(60); "Tag #"; QPTrim$(DCVRec(1).StateTag)
      Linecnt = Linecnt + 1
    Else
      ToPrint$ = Str(AcctNo&) + "~" + QPTrim$(DCCustREc(1).BILLNAME)
      ToPrint$ = ToPrint$ + "~" + DCCustREc(1).SOSEC + "~" + OPdate
      ToPrint$ = ToPrint$ + "~" + RTrim$(DCVRec(1).makemodel)
      ToPrint$ = ToPrint$ + " " + RTrim$(DCVRec(1).Desc) + "~" + QPTrim$(DCVRec(1).StateTag)
      Print #RptHandle, ToPrint$
    End If
Return
PrintOwnRepRptHeader:
 If RptType <> 0 Then
  Page = Page + 1
  Print #RptHandle, Tab(22); "Va. Decal System : Customer NonOwner Listing"
  Print #RptHandle, "Date Opened Cutoff: "; txtOpnDate.Text
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "Category: "; Category$
  Print #RptHandle, "Acct #"; Tab(10); "Customer Name"; Tab(50); "Soc Sec #"; Tab(64); "Date Opened"
  Print #RptHandle, String$(80, "=")
  Linecnt = 5
 End If
Return

PrintOwnRepRptEnding:
If RptType <> 0 Then
  Print #RptHandle, String$(80, "=")
  Print #RptHandle, "Total Customers Printed: "; Using("#####", CustCnt)
  Print #RptHandle,
  Print #RptHandle, FF$
End If
Return
ERRORSTUFF:
  Unload FrmShowPctComp
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "RptNonOwner", "Calc Report", Erl)
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
      ActivateControls Me, True
End Sub
