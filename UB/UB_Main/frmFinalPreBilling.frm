VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFinalPreBilling 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Final Pre-Billing"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmFinalPreBilling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboApplyDep 
      Height          =   375
      Left            =   7350
      TabIndex        =   0
      Top             =   3150
      Width           =   825
      _Version        =   196608
      _ExtentX        =   1455
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmFinalPreBilling.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   5595
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ColDesigner     =   "frmFinalPreBilling.frx":0CA4
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   375
      Left            =   5385
      TabIndex        =   2
      Top             =   5070
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ColDesigner     =   "frmFinalPreBilling.frx":107E
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
      Left            =   9648
      TabIndex        =   5
      Top             =   7512
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
      Left            =   7968
      TabIndex        =   4
      Top             =   7512
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "9:28 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/12/2009"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpText fptxtAdjustment 
      Height          =   348
      Left            =   5280
      TabIndex        =   1
      Top             =   4068
      Width           =   732
      _Version        =   196608
      _ExtentX        =   1291
      _ExtentY        =   614
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment Factor:"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   4128
      Width           =   2076
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the fuel adjustment amount for your Electric Service."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6132
      TabIndex        =   11
      Top             =   3960
      Width           =   3060
   End
   Begin VB.Line Line1 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   3396
      X2              =   8796
      Y1              =   4632
      Y2              =   4632
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Final Pre-Billing Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   10
      Top             =   1176
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   936
      Width           =   5772
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apply Deposit to Final Billing?"
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
      Index           =   2
      Left            =   3408
      TabIndex        =   9
      Top             =   3168
      Width           =   3804
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
      Left            =   3588
      TabIndex        =   8
      Top             =   5112
      Width           =   1716
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
      Left            =   2964
      TabIndex        =   7
      Top             =   5652
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3444
      Left            =   2724
      Top             =   2856
      Width           =   6780
   End
   Begin VB.Line Line2 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   3384
      X2              =   8784
      Y1              =   3792
      Y2              =   3792
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   816
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
Attribute VB_Name = "frmFinalPreBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Dim Grpt As Boolean
Private Sub cmdExit_Click()
  frmUBFinalBillMenu.Show
  Unload frmFinalPreBilling
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via FinalPreBilling by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fpcboApplyDep_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboApplyDep.ListDown = True
  End If
  If fpcboApplyDep.ListDown <> True Then
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
        fpcboApplyDep.SetFocus
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


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub cmdPrint_Click()
  If fpcboRptType.ListIndex = 0 Then
    DeActivateControls Me, True 'do graphic report
    Grpt = True
    PreBillReport
  ElseIf fpcboRptType.ListIndex = 1 Then
    DeActivateControls Me, True
    Grpt = False
    PreBillReport
    ActivateControls Me, True
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
  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Grpt = False
  fpcboApplyDep.AddItem "No"
  fpcboApplyDep.AddItem "Yes"
  fpcboApplyDep.ListIndex = 0
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  'fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.AddItem "Postal Carrier Route Order"
  fpcboPrintOrder.AddItem "ZipCode Order"
  GetPreBillReady
  fpcboPrintOrder.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpPreBillingReport
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub GetPreBillReady()
  Dim DoFuel As Boolean, cnt As Integer, TempRev As String
  Dim UBSetupLen As Integer, NumofRevs As Integer
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupLen

  For cnt = 1 To MaxRevsCnt     'find last active revenue
    TempRev$ = UCase$(QPTrim$(UBSetUp(1).Revenues(cnt).RevName))
    If Len(TempRev$) = 0 Then
      'NumOfRevs = cnt - 1       'set actual number of revenues
      'Exit For
    Else        'build revenue description lines
      If InStr(TempRev$, "ELECTRIC") Then
        DoFuel = True
      End If
    End If
  Next
  If DoFuel = True Then
    fptxtAdjustment.Enabled = True
    fptxtAdjustment = " "
  Else
    fptxtAdjustment.Enabled = False
    fptxtAdjustment = " "
  End If
  If UBSetUp(1).UseSeq = "Y" Then
    fpcboPrintOrder.AddItem "Sequence Number Order"
  End If
  
End Sub
Private Sub PreBillReport()
  Dim Temp2 As String, Temp1 As String
  Dim NumofRevs As Integer, NumOfRates As Integer
  Dim UBRateTblRecLen As Integer, RateFile As Integer, cnt As Long
  Dim UBSetupLen As Integer, MowFlag As Boolean, TennFlag As Boolean
  Dim TempRev As String, DoFuelAdjFlag As Boolean, SkipInactive As Boolean
  Dim Choice As Integer, CustDepAmt As Double, ReadErr As Boolean
  Dim FuelAdjAmt As Double, NCCnt As Long, BalanceAmt As Double
  Dim IndexName As String, UsingAcct As Boolean, IdxTypeText As String
  Dim AbortFlag As Boolean, TheDate As String, UBCustRecLen As Integer
  Dim UBBillRecLen As Integer, TBooks As Integer, NumOfRecs As Long
  Dim Handle As Integer, IdxRecLen As Integer, lcnt As Long
  Dim UBBill As Integer, UBCust As Integer, UBRpt As Integer
  Dim ThisCustRec As Long, BillTo As String, BadBookFlag As Boolean
  Dim WhatBook As Integer, FRCnt As Integer, WhatService As Integer
  Dim Multi As Integer, FlatAmt As Double, WhatRate As Integer
  Dim DoneOne As Boolean, TRevCnt As Integer, IFlag As Boolean
  Dim TRateCnt As Integer, MINAMT As Long, PrintedRevAmt As Boolean
  Dim MCCnt As Integer, CubMtr As Boolean, LocMeterType As String
  Dim MeterMulti As Long, MeterNum As String, ReadAmt As Long
  Dim MaxMeterAmt As Long, Consump As Long, ThisMeterUseCnt As Integer
  Dim AvgUse As Long, HiConsump As Long, LowConsump As Long
  Dim TTRevCnt As Integer, CurReadAmt As Long, PreReadAmt As Long
  Dim ConsumpFlag As Boolean, ConsumpAmt As Long, NONRateCnt As Integer
  Dim NONRate As Integer, CTaxAmt As Double, TXCnt As Integer
  Dim Bills2Print As Integer, AcctBalance As Double, WhatPump As Integer
  Dim TAcctBalance As Double, HasAPumpCode As Boolean, MPCnt As Integer
  Dim PumpMtrOK As Boolean, TotalFlatAmt As Double, TotalRevAmt As Double
  Dim TotalTaxAmt As Double, RaCnt As Integer, TestTot As Double
  Dim ZCnt As Integer, Book As String, TBookAmt As Double, TPumps As Integer
  Dim TBTaxAmt As Double, RCnt As Integer, TBookGTot As Double
  Dim TMMConsump As Double, RptText As String, TBCnt As Integer
  Dim CustPump As String, ThisPump As String, ReportFile As String
  Dim Pumptest As Double, Dash80 As String, DepositFlag As Integer
  Dim DepFile As Integer, linechk As Integer
  UBLog "IN: Final Prebilling Report"
  Pumptest = 0
  PageNo = 0
  linechk = 0
   ' Dash80$ = String$(80, "-")
    Temp1$ = Space$(10)
    Temp2$ = Space$(12)

    NumofRevs = MaxRevsCnt      'assume max munber of revenue sources
    '111698 Prorate
    ReDim ProrateServ(1 To 15) As Integer

    ReDim UBSetUpRec(1) As UBSetupRecType
    LoadUBSetUpFile UBSetUpRec(), UBSetupLen
    TOWNNAME$ = UBSetUpRec(1).UTILNAME

    If InStr(TOWNNAME$, "MOWAS") > 0 Then
      MowFlag = True
    End If
    ReDim RevDesc(1 To MaxRevsCnt) As String * 12
    For cnt = 1 To MaxRevsCnt   'find last active revenue
      TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName)
      If Len(TempRev$) = 0 Then
        NumofRevs = cnt - 1     'set actual number of revenues
        Exit For
      Else      'build revenue description lines
        LSet RevDesc(cnt) = UCase$(TempRev$)
        If InStr(RevDesc(cnt), "ELECTRIC") Then
          DoFuelAdjFlag = True
        End If
      End If
    Next

    '111398 Prorate
    For cnt = 1 To MaxRevsCnt
      If UBSetUpRec(1).Revenues(cnt).ProRate = "Y" Then
        ProrateServ(cnt) = True
      End If
    Next
    NumOfRates = GetNumRateRecs%
    ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
    ReDim RateConsump(1 To NumOfRates) As Long

    UBRateTblRecLen = Len(UBRateTbls(1))

    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
    For cnt = 1 To NumOfRates
      Get RateFile, cnt, UBRateTbls(cnt)
    Next
    Close

ReStart:
  Choice = fpcboPrintOrder.ListIndex + 1

  Select Case Choice
  Case 0
    'ExitFlag = True
  Case 1        'Name
    IndexName$ = NameIndexFile
    'OkFlag = True
  Case 2        'Acct
    IndexName$ = ""
    UsingAcct = True
    'OkFlag = True
'  Case 3        'Location
'    IndexName$ = BookIndexFile
'    'OkFlag = True
  Case 3 '4        'Postal Route
    IdxTypeText$ = "Postal Route"
    MakePostalIndex IdxTypeText$
    IndexName$ = TempIndexName
    'OkFlag = True
  Case 4 '5        'ZipCode
    IdxTypeText$ = "Zip-Code"
    'this mowflag for zip index doesn't matter cause both index
    'routines do same thing now.
    If MowFlag Then
      MakeMowZipCodeIndex IdxTypeText$
    Else
      MakeZipCodeIndex IdxTypeText$
    End If
    IndexName$ = TempIndexName
    'OkFlag = True
  Case 5 '6        'Sequence number
    IdxTypeText$ = "Sequence Number"
    MakeSequenceIndex IdxTypeText$, Me
    IndexName$ = TempIndexName
    'OkFlag = True
  End Select

 If AbortFlag Then GoTo ExitPreReport
  FrmShowPctComp.Label1 = "Creating Final PreBilling Report"
  FrmShowPctComp.Show , Me

  If Grpt Then
    MaxLines = 50
  Else
    MaxLines = 53
  End If
  ReDim fmt$(0 To 6)
  fmt$(0) = String$(80, "-")
  fmt$(1) = "#########.##"
  fmt$(2) = "#########"
  fmt$(3) = "######.##"
  fmt$(4) = "###########"
  fmt$(5) = "$###,###,###.##"
  fmt$(6) = "$#,###,###.##"

    TheDate$ = "Date: " + Date$
    DepositFlag = fpcboApplyDep.ListIndex

    If DepositFlag = -1 Then
      GoTo ReStart
    Else
      DepFile = FreeFile
      Open UBPath$ + "UBDEPFLG.DAT" For Random Shared As DepFile Len = 2
      Put DepFile, , DepositFlag
      Close DepFile
    End If
    If DoFuelAdjFlag Then
      FuelAdjAmt# = Val(fptxtAdjustment)
      UBLog "Fuel adjustment factor:" + Str$(FuelAdjAmt#)
    Else
      FuelAdjAmt# = 0
    End If

    If FuelAdjAmt# = -10000 Then GoTo ReStart

    UBLog "Calculating utility charges."
    MakeFinalBillFile AbortFlag, FuelAdjAmt#

    If AbortFlag Then
      UBLog "ABORTED: CALCULATIONS"
    Else
      UBLog "Finished utility calculations."
    End If

    If AbortFlag Then GoTo ExitPreReport

    ReDim UBCustRec(1 To 2) As NewUBCustRecType
    UBCustRecLen = Len(UBCustRec(1))

    ReDim UBBillRec(1) As UBTransRecType
    UBBillRecLen = Len(UBBillRec(1))
    ReDim RevTotals(1 To NumofRevs) As Double   'holds revenues total amt
    '052097 added tax by revenue totals
    ReDim TaxTotals(1 To NumofRevs) As Double   'holds revenues tax total amt
    ReDim ConsumpTot(1 To NumofRevs, 1 To 2) As Double          'holds each re
    ReDim RateConsump(1 To NumOfRates) As Long
    ReDim RateTotals(1 To NumOfRates) As Double 'holds each Rates $totals
    '052097 added tax by rate code totals
    ReDim RTaxTot(1 To NumOfRates) As Double    'holds each Rates Tax totals
    '052097 added tax by book totals to type def

    '012698 Added bill count by rate code
    ReDim RateCount(1 To NumOfRates) As Long
    ReDim PumpConsump(0 To 1) As PumpConsumpType                'Consumption b

    TBooks = 0

    If UsingAcct Then
      NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
    Else        'load the index
      UBLog "Loading index file: " + IndexName$
      IdxRecLen = 4
      NumOfRecs = FileSize(IndexName$) \ 4
      ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
      Handle = FreeFile
      Open IndexName$ For Random Shared As Handle Len = IdxRecLen
      For lcnt& = 1 To NumOfRecs
        'ReDim Preserve IndexArray(1 To lcnt&) As UBCustIndexRecType
        Get #Handle, lcnt&, IndexArray(lcnt&)
      Next
      Close Handle
      'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
    End If

    ReportFile$ = UBPath$ + UBFinPreRptFile
    UBBill = FreeFile
    Open UBPath$ + UBFinBillsFile For Random Shared As UBBill Len = UBBillRecLen

    UBCust = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

    UBRpt = FreeFile
    Open ReportFile$ For Output As UBRpt

    'ShowProcessingScrn "Processing Pre-Billing Report"

    UBLog "Writing prebilling report to disk."

    GoSub PrintPreHeader

    For cnt = 1 To NumOfRecs
      FrmShowPctComp.ShowPctComp cnt, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        ActivateControls Me, True
        GoTo ExitPreReport
      End If

      If UsingAcct Then
        ThisCustRec& = cnt
      Else
        ThisCustRec& = IndexArray(cnt).RecNum
      End If

      Get UBCust, ThisCustRec&, UBCustRec(1)
      If UBCustRec(1).BillTo = "O" Then
        BillTo$ = " O"
      Else
        BillTo$ = " C"
      End If

      If UBCustRec(1).DelFlag Then
        GoTo SkipEm
      End If

      If UBCustRec(1).Status <> "F" Then
        GoTo SkipEm
      End If
      Get UBBill, ThisCustRec&, UBBillRec(1)

      'IF ThisCustRec& = 190 THEN
      '  STOP
      'END IF

      If DepositFlag Then
        If UBCustRec(1).DepositAmt > 0 Then
          CustDepAmt# = UBCustRec(1).DepositAmt
          UBBillRec(1).TaxAmt(15) = CustDepAmt#
          Put UBBill, ThisCustRec&, UBBillRec(1)
        End If
      End If
      If LineCnt >= MaxLines - 8 Then
        Print #UBRpt, FF$
        GoSub PrintPreHeader
      End If

      If UBBillRec(1).ActiveFlag <> 0 Then
        Print #UBRpt, UBCustRec(1).Status; Tab(4); Using("######", ThisCustRec&);
        Print #UBRpt, "      "; Left$(QPTrim$(UBCustRec(1).CustName), 30); Tab(48); Left$(UBCustRec(1).ServAddr, 23);
        Print #UBRpt, Using("   ###", UBBillRec(1).ProRatePCT); "%"; 'Using("   ###%", UBBillRec(1).ProRatePCT);
        Print #UBRpt, BillTo$
        LineCnt = LineCnt + 1
      End If

      WhatRate = 0

      If LineCnt >= MaxLines Then
        Print #UBRpt, FF$
        GoSub PrintPreHeader
      End If

      DoneOne = False
      For TRevCnt = 1 To NumofRevs
        If LineCnt >= MaxLines Then
          Print #UBRpt, FF$
          GoSub PrintPreHeader
        End If

        If UBBillRec(1).RevAmt(TRevCnt) <> 0 Then
          DoneOne = False
          Print #UBRpt, RevDesc(TRevCnt);
          For TRateCnt = 1 To NumOfRates
            If UBRateTbls(TRateCnt).Ratecode = UBCustRec(1).serv(TRevCnt).Ratecode Then
              MINAMT& = UBRateTbls(TRateCnt).MINUNITS
              WhatRate = TRateCnt
              Exit For
            Else
            WhatRate = 0
            End If
          Next

          If UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" Then
            RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            '02-20-97 Add new cash totals by rate
            If WhatRate > 0 Then
              RateCount(WhatRate) = RateCount(WhatRate) + 1
              RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
              RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
            End If

            For MCCnt = 1 To 7
              CubMtr = False
              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
              MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti

              If UBCustRec(1).LocMeters(MCCnt).MtrUnit = "C" Then
                CubMtr = True
              End If

              If MeterMulti& <= 0 Then MeterMulti& = 1
              If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).serv(TRevCnt).RMtrType) Then
                DoneOne = True
                MeterNum$ = QPTrim$(UBCustRec(1).serv(TRevCnt).Ratecode)
                'use the Meternum$ to hold the rate code temporarily
                If Len(MeterNum$) > 0 Then
                  If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
                    MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                  End If
                  RSet Temp2$ = MeterNum$
                  Print #UBRpt, Tab(14); Temp2$;
                Else
                  RSet Temp2$ = "RATE ERROR"
                  Print #UBRpt, Tab(14); Temp2$;
                End If
                '052797  Read Error
                ReadErr = False
                If UBBillRec(1).CurRead(MCCnt) < 0 Then
                  ReadErr = True
                  Print #UBRpt, Tab(30); Using("**#######", 0);
                Else
                  Print #UBRpt, Tab(30); Using("#########", UBBillRec(1).CurRead(MCCnt));
                End If

                If UBBillRec(1).PrevRead(MCCnt) < 0 Then
                  ReadErr = True
                  Print #UBRpt, Tab(42); Using("**#######", 0);
                Else
                  Print #UBRpt, Tab(42); Using("#########", UBBillRec(1).PrevRead(MCCnt));
                End If

                If ReadErr Then
                  ReadAmt& = 0
                Else
                  ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
                End If
                If ReadAmt& < 0 Then
                  'Meter has rolled over or, been misread
                  MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MCCnt))) - 1)
                  ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
                End If
                If CubMtr Then
                  ReadAmt& = ReadAmt& * 7.481
                End If

                RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
                ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)

                If ReadErr Then
                  Print #UBRpt, Tab(56); Using("**#######", ReadAmt& * MeterMulti&);
                Else
                  Print #UBRpt, Tab(56); Using("#########", ReadAmt& * MeterMulti&);
                End If
                Consump& = ReadAmt& * MeterMulti&
                ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
                If ThisMeterUseCnt <= 0 Then ThisMeterUseCnt = 1
                '***
                AvgUse& = Round#((UBCustRec(1).LocMeters(MCCnt).AvgUse / ThisMeterUseCnt) + 0#)

                If AvgUse& > 0 Then
                  LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * 0.01))
                  HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * 0.01))
                End If

                If UBCustRec(1).EstFlag = "E" Then
                  Print #UBRpt, " E";           'Est. Reading
                ElseIf Consump& < MINAMT& Then
                  Print #UBRpt, " M";           'Minium Usage
                ElseIf Consump& < LowConsump& Then
                  Print #UBRpt, " L";           'Low reading
                ElseIf Consump& > HiConsump& Then
                  Print #UBRpt, " H";           'High Reading
                End If
                If UBBillRec(1).RevAmt(TRevCnt) > 0 Then
                  Print #UBRpt, Tab(69); Using("######.##", UBBillRec(1).RevAmt(TRevCnt));
                  If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
                    Print #UBRpt, "*";
                  End If
                End If
                Print #UBRpt,
                LineCnt = LineCnt + 1
                If UBBillRec(1).TaxAmt(TRevCnt) > 0 Then
                  TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
                  Print #UBRpt, "*Tax"; Tab(69); Using("######.##", UBBillRec(1).TaxAmt(TRevCnt))
                  LineCnt = LineCnt + 1
                End If
              End If
            Next

            '071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
            '       a metered service but no meter on a flat rate charge. Rev was added
            '       to total, but didn't show on prebilling report.
            If Not DoneOne Then
              DoneOne = True
              Print #UBRpt, Tab(69); Using("######.##", UBBillRec(1).RevAmt(TRevCnt));
              If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
                Print #UBRpt, "*";
              End If
              Print #UBRpt,
              LineCnt = LineCnt + 1
            End If
            '*****************************************************************

          Else  'it's a nonmetered service
            ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
            If WhatRate > 0 Then
              RateCount(WhatRate) = RateCount(WhatRate) + 1
              RateConsump(WhatRate) = RateConsump(WhatRate) + 1
              RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
              RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
            Else
              'STOP
            End If
            If LineCnt >= MaxLines Then
              Print #UBRpt, FF$
              GoSub PrintPreHeader
            End If

            RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
            Print #UBRpt, Tab(69); Using("######.##", UBBillRec(1).RevAmt(TRevCnt));
            If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
              Print #UBRpt, "*";
            End If
            If UBBillRec(1).TaxAmt(TRevCnt) > 0 Then
              TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
              Print #UBRpt,
              Print #UBRpt, "*Tax"; Tab(69); Using("######.##", UBBillRec(1).TaxAmt(TRevCnt));
              LineCnt = LineCnt + 1
            End If
          End If
          If Not DoneOne Then
            Print #UBRpt,
            LineCnt = LineCnt + 1
          End If
        End If
        If (TRevCnt = NumofRevs) And UBBillRec(1).Transamt = 0 Then
          'CONSUMPTION inactive account
          If UBBillRec(1).Transamt = 0 Then
            For TTRevCnt = 1 To NumofRevs
              For MCCnt = 1 To 7
                CubMtr = False
                MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
                If UBCustRec(1).LocMeters(MCCnt).MtrUnit = "C" Then
                  CubMtr = True
                End If
                If MeterMulti& <= 0 Then MeterMulti& = 1
                LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
                If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).serv(TTRevCnt).RMtrType) Then
                  If UBBillRec(1).CurRead(MCCnt) < 0 Then
                    UBBillRec(1).CurRead(MCCnt) = 0
                  End If
                  If UBBillRec(1).PrevRead(MCCnt) < 0 Then
                    UBBillRec(1).PrevRead(MCCnt) = 0
                  End If
                  CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
                  PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
                  If CurReadAmt& <> PreReadAmt& Then
                    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'                    If Not ConsumpFlag Then
'                      Print #UBRpt, UBCustRec(1).Status; "  "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; UBCustRec(1).CustName; "  "; QPTrim$(UBCustRec(1).SERVADDR)
'                    End If
                    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                    ConsumpFlag = True
                    MeterNum$ = QPTrim$(UBCustRec(1).serv(TTRevCnt).Ratecode)
                    If Len(MeterNum$) > 0 Then
                      If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
                        MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                      End If
                      RSet Temp2$ = MeterNum$
                    End If

                    ConsumpAmt& = (CurReadAmt& - PreReadAmt&) * MeterMulti&
 
                   If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
                    'For Nonprofits include consumption as normal   'cleveland
                    '040998 Made changes here
                    For NONRateCnt = 1 To NumOfRates
                      If UBRateTbls(NONRateCnt).Ratecode = UBCustRec(1).serv(TTRevCnt).Ratecode Then
                        NONRate = NONRateCnt
                        Exit For
                      End If
                    Next
                    If NONRate > 0 Then
                      RateConsump(NONRate) = RateConsump(NONRate) + ConsumpAmt&
                    End If
                    ConsumpTot(TTRevCnt, 1) = ConsumpTot(TTRevCnt, 1) + ConsumpAmt&
                    
                    'Bookconsump(WhatBook).Consump(TTRevCnt) = Bookconsump(WhatBook).Consump(TTRevCnt) + ConsumpAmt&
                    
                    '040998 Made changes here 'cleveland
                  Else          'add consumption to inactives
                    ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
                  End If
                 
                  If LineCnt >= MaxLines Then
                    Print #UBRpt, FF$
                    GoSub PrintPreHeader
                  End If

                  Print #UBRpt, RevDesc(TTRevCnt); Tab(14); Temp2$; Tab(30); Using(fmt$(2), CurReadAmt&); Tab(42); Using(fmt$(2), PreReadAmt&); Tab(54); Using(fmt$(2), ConsumpAmt&)
                  LineCnt = LineCnt + 1
                    
'                 ConsumpFlag = True
'                 Print #UBRpt, RevDesc(TTRevCnt);
'                 'MeterNum$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRNum)
'                 MeterNum$ = QPTrim$(UBCustRec(1).serv(MCCnt).RATECODE)
'                 If Len(MeterNum$) > 0 Then
'                   If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
'                     MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'                   End If
'                   RSet Temp2$ = MeterNum$
'                   Print #UBRpt, Tab(14); Temp2$;
'                 Else
'                   RSet Temp2$ = "RATE ERROR"
'                   Print #UBRpt, Tab(14); Temp2$;
'                 End If
'                 Print #UBRpt, Tab(30); Using("#########", CurReadAmt&);
'                 Print #UBRpt, Tab(42); Using("#########", PreReadAmt&);
'                 ConsumpAmt& = (CurReadAmt& - PreReadAmt&) * UBCustRec(1).LocMeters(MCCnt).MTRMulti
'                 ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
'                 Print #UBRpt, Tab(56); Using("#########", ConsumpAmt&)
'                 LineCnt = LineCnt + 1
               End If
                End If
              Next
            Next
          End If
          If ConsumpFlag And UBCustRec(1).Status <> "A" Then
            If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
              Print #UBRpt, "*** NON-PROFIT ***"
              Print #UBRpt, fmt$(0)
              LineCnt = LineCnt + 2
            Else
              Print #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
              Print #UBRpt, fmt$(0)
              LineCnt = LineCnt + 2
            End If
            ConsumpFlag = False
            NCCnt = NCCnt + 1
          End If
        ElseIf (TRevCnt = NumofRevs) And UBBillRec(1).Transamt > 0 Then
          AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
          Print #UBRpt, Tab(5); "Current:";
          Print #UBRpt, Using("$#,###,###.##", UBBillRec(1).Transamt);
          If AcctBalance# <> 0 Then
            Print #UBRpt, Tab(30); "Previous:";
            Print #UBRpt, Using("$#,###,###.##", AcctBalance#);
          End If
          Print #UBRpt, Tab(56); "Total:";
          Print #UBRpt, Tab(66); Using("$#,###,###.##", Round#(AcctBalance# + UBBillRec(1).Transamt))
          LineCnt = LineCnt + 1

          If DepositFlag Then
            Print #UBRpt, Tab(49); "Less Deposit:";
            Print #UBRpt, Tab(66); Using("$#,###,###.##", -UBCustRec(1).DepositAmt)
            BalanceAmt# = Round#(AcctBalance# + UBBillRec(1).Transamt - UBCustRec(1).DepositAmt)
            Select Case Sgn(BalanceAmt#)
            Case -1
              Print #UBRpt, Tab(51); "Refund Due:";
              Print #UBRpt, Tab(66); Using("$#,###,###.##", Abs(BalanceAmt#))
            Case 0
              Print #UBRpt, Tab(51); "       Due:";
              Print #UBRpt, Tab(66); Using("$#,###,###.##", 0)
            Case 1
              Print #UBRpt, Tab(50); "Balance Due:";
              Print #UBRpt, Tab(66); Using("$#,###,###.##", BalanceAmt#)
            End Select
            LineCnt = LineCnt + 2
          End If

          '-=-=-=-=-=-=-=-=-=-=-=-=-=
          Print #UBRpt, fmt$(0)
          LineCnt = LineCnt + 1
        End If
      Next

      '020199 Moved pump code processing to here. Stops bug in getting true
      '       meter consumption figures.
      GoSub GetWhatPump
      If HasAPumpCode Then
        PumpConsump(WhatPump).CustCnt = PumpConsump(WhatPump).CustCnt + 1
        For MPCnt = 1 To 7
          PumpMtrOK = False
          CubMtr = False
          LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MPCnt).MTRType)
          Select Case LocMeterType$
          Case "C", "S", "W"
            PumpMtrOK = True
          End Select
          If PumpMtrOK Then
            MeterMulti& = UBCustRec(1).LocMeters(MPCnt).MTRMulti
            If UBCustRec(1).LocMeters(MPCnt).MtrUnit = "C" Then
              CubMtr = True
            End If
            If MeterMulti& <= 0 Then MeterMulti& = 1
            ReadAmt& = UBBillRec(1).CurRead(MPCnt) - UBBillRec(1).PrevRead(MPCnt)
            If ReadAmt& < 0 Then                'Meter rolled over or, been  misread
              MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MPCnt))) - 1)
              ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MPCnt)) + UBBillRec(1).CurRead(MPCnt)
            End If
            If CubMtr Then
              ReadAmt& = ReadAmt& * 7.481
            End If
            PumpConsump(WhatPump).Consump = PumpConsump(WhatPump).Consump + (ReadAmt& * MeterMulti&)
          End If
        Next
      End If

SkipEm:

    Next

    If AbortFlag Then GoTo ExitPreReport

    Print #UBRpt, FF$
    GoSub TitleLine
    Print #UBRpt, "Billing Grand Totals--------------------"
    Print #UBRpt, "                                        Inactive"
    Print #UBRpt, "Revenue/Tax         Consumption        Consumption             Amount"
    Print #UBRpt, fmt$(0)

    TotalRevAmt# = 0
    TotalTaxAmt# = 0
    For RaCnt = 1 To NumofRevs
      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).RevName; Tab(20);
      Print #UBRpt, Tab(20); Using("###########", ConsumpTot(RaCnt, 1));
      Print #UBRpt, Tab(40); Using("###########", ConsumpTot(RaCnt, 2));
      Print #UBRpt, Tab(60); Using("#######.##", RevTotals(RaCnt))
      LineCnt = LineCnt + 1
      TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
      If TaxTotals(RaCnt) > 0 Then
        Print #UBRpt, QPTrim$(UBSetUpRec(1).Revenues(RaCnt).RevName); "*Tax"; Tab(20);
        Print #UBRpt, Tab(60); Using("#######.##", TaxTotals(RaCnt))
        LineCnt = LineCnt + 1
        TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
      End If
    Next
    If LineCnt >= MaxLines Then
      Print #UBRpt, FF$
      GoSub PrintPreHeader
    End If
    Print #UBRpt,
    Print #UBRpt, Tab(38); "TAX TOTAL:"; Tab(52); Using("$#,###,###.##", TotalTaxAmt#)
    Print #UBRpt, Tab(38); "    TOTAL:"; Tab(52); Using("$#,###,###.##", Round#(TotalRevAmt# + TotalTaxAmt#))
    Print #UBRpt,
    
    GoSub RptTotRateHeader

    TotalRevAmt# = 0

    For RaCnt = 1 To NumOfRates
      If RateConsump(RaCnt) <> 0 Or RateTotals(RaCnt) <> 0 Then
        Print #UBRpt, UBRateTbls(RaCnt).Ratecode; "   "; UBRateTbls(RaCnt).RATEDESC;
        Print #UBRpt, Tab(38); Using("###########", RateConsump(RaCnt));
        Print #UBRpt, Tab(54); Using("#######.##", RateTotals(RaCnt));
        Print #UBRpt, Tab(71); Using("#####", RateCount(RaCnt))
        LineCnt = LineCnt + 1
        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
        If RTaxTot(RaCnt) > 0 Then
          Print #UBRpt, "*Tax"; Tab(54); Using("#######.##", RTaxTot(RaCnt))
          LineCnt = LineCnt + 1
        End If
      End If
    Next
    Print #UBRpt,
    Print #UBRpt, Tab(38); "TAX TOTAL:"; Tab(52); Using("$#,###,###.##", TotalTaxAmt#)
    Print #UBRpt, Tab(38); "    TOTAL:"; Tab(52); Using("$#,###,###.##", Round#(TotalRevAmt# + TotalTaxAmt#))
    Print #UBRpt, Chr$(12)
    LineCnt = LineCnt + 3
    If TPumps > 0 Then
      GoSub PumpHeader
      TMMConsump# = 0
      For cnt = 1 To TPumps
        Print #UBRpt, PumpConsump(cnt).PumpCode; Tab(30); Using("###########", PumpConsump(cnt).CustCnt); Tab(60); PumpConsump(cnt).Consump
        LineCnt = LineCnt + 1
        TMMConsump# = TMMConsump# + PumpConsump(cnt).Consump
      Next
      Print #UBRpt, fmt$(0)
      Print #UBRpt, Tab(35); "Pump Code Total:"; Tab(60); Using("###########", TMMConsump#)
      LineCnt = LineCnt + 1
    End If


    Close

    UBLog "Finished writing prebilling report."

    Select Case Choice
    Case 1
      RptText$ = " (Customer Order)"
    Case 2
      RptText$ = "(Account Order)"
    Case 3
      RptText$ = "(Postal RT. Order)"
    Case 4
      RptText$ = "(ZipCode Order)"
    Case 5
      RptText$ = "(Sequence Order)"
    End Select

    Erase UBCustRec, UBBillRec, RevTotals, TaxTotals
    Erase ConsumpTot, RateConsump, UBRateTbls
    Erase RateConsump, UBSetUpRec, RevDesc

    If Not AbortFlag Then
      If Grpt Then
        Load frmLoadingRpt
        frmLoadingRpt.setwherefrom frmFinalPreBilling
        ARptPreBilling.Title = "Utility Final Pre-Billing Report"
        ARptPreBilling.txtDate = Now
        ARptPreBilling.txtTown = TOWNNAME$
        ARptPreBilling.GetName ReportFile$
        ARptPreBilling.startrpt
      Else
        ViewPrint ReportFile$, "Final Pre-Billing Report " + RptText$
      End If
        'PrintRptFile "Pre-Billing Report " + RptText$, UBFinPreRptFile, LPTPort
    End If
    GoTo ExitPreReport

PrintPreHeader:
  GoSub TitleLine
  Print #UBRpt, "Stat  Act.  Locat    Customer Name             Service Address       Prorate%"
  Print #UBRpt, "Revenue            R-Code     Cur Read    Pre Read     Consump        Charges"
  Print #UBRpt, fmt$(0)
  LineCnt = 5
Return

TitleLine:
  If Grpt Then GoTo SkipTitle:
  PageNo = PageNo + 1
  Print #UBRpt, "Utility Final Pre-Billing Report  "; TOWNNAME$; Tab(70); "Page: "; PageNo
  Print #UBRpt, TheDate$
SkipTitle:
  Return

GetWhatPump:
    HasAPumpCode = True         'assume they have a pump code
    WhatPump = 0
    If Len(QPTrim$(UBCustRec(1).PumpCode)) = 0 Then
      If UBCustRec(1).Status = "A" Then
        HasAPumpCode = False    'no pump code
        WhatPump = 0
      End If
      GoTo PumpCodeReturn
    End If

    CustPump$ = UCase$(QPTrim$(UBCustRec(1).PumpCode))
    If Len(CustPump$) > 0 Then
      For TBCnt = 1 To TPumps
        ThisPump$ = QPTrim$(PumpConsump(TBCnt).PumpCode)
        If ThisPump$ = CustPump$ Then
          WhatPump = TBCnt
          Exit For
        End If
      Next
      If WhatPump = 0 Then
        TPumps = TPumps + 1
        ReDim Preserve PumpConsump(0 To TPumps) As PumpConsumpType
        PumpConsump(TPumps).PumpCode = CustPump$
        WhatPump = TPumps
      End If
    Else
      TPumps = TPumps + 1
      PumpConsump(TPumps).PumpCode = CustPump$
      WhatPump = TPumps
    End If

PumpCodeReturn:
    Return

PumpHeader:
  GoSub TitleLine
  Print #UBRpt, "Report Totals by Pump Code:"
  Print #UBRpt,
  Print #UBRpt, "PumpCode                  Customer Count                    Consumption"
  Print #UBRpt, fmt$(0)
  LineCnt = 6
  Return

RptTotRateHeader:
  GoSub TitleLine
  Print #UBRpt,
  Print #UBRpt, "Report Totals by Rate Code:"
  Print #UBRpt,
  Print #UBRpt, "Code      Rate Description            Consumption        Amount       Bills"
  Print #UBRpt, fmt$(0)
  LineCnt = 5
  Return


ExitPreReport:
    UBLog "OUT: Prebilling Report" + CrLf$
End Sub
Private Sub MakeFinalBillFile(AbortFlag, FuelAdjAmt#)
  Dim UBSetupLen As Integer, PrinceFlag As Boolean, YadkinFlag As Boolean
  Dim WadeFlag As Boolean, ElkFlag As Boolean, ScottFlag As Boolean
  Dim DaleFlag As Boolean, SunBchFlag As Boolean, GotCustFlag As Boolean
  Dim NorwoodFlag As Boolean, SkipInactive As Boolean
  Dim BookFlag As Boolean, CycleFlag As Boolean, ThisRevCnt As Integer
  Dim UBBillRecLen As Integer, UBCustRecLen As Integer
  Dim NumOfRates As Integer, UBRateTblRecLen As Integer, RateFile As Integer
  Dim cnt As Integer, lcnt As Long, NumOfCustRecs As Long, BillFile As Integer
  Dim CustFile As Integer, BillCnt As Integer, NumofRevs As Integer
  Dim RCnt As Integer, zz As Integer, GotIRRMtr As Boolean
  Dim IrrConsp As Long, IrrMtr As Integer, IrrMtrNum As String
  Dim MaxMeterAmt As Long, ProRateFlag As Boolean, ProPct As Double
  Dim MeterConsp As Long, TMeterConsp As Long, FRCnt As Integer
  Dim WhatService As Integer, Multi As Integer, FlatAmt As Double
  Dim TaxAmt As Double, MRCnt As Integer, TestAmt As Double
  Dim HowMuch As Double, WhatTbl As Integer, NonMAmt As Double
  Dim MeterType As String, MeterLocNum As Integer, MCnt As Integer
  Dim ThisMeterConsp As Double, AddRevAmt As Double, TMaxAmt As Double
  Dim NumUser As Long, MinimumConsp As Long, ProRevAmt As Double
  Dim ConwayFlag As Boolean, RevAmt As Double, SewCalcConsp As Long
  Dim MeterMulti As Long, FuelAddAmt As Double, TZCnt As Integer
  Dim Ctype As String, Diff As Double, Ratecode As String
  Dim TCnt As Integer, CubMtr As Boolean, AdjRev As Double
  Dim CoveCityFlag As Boolean
  'ShowProcessingScrn "Calculating Utility Charges."

  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupLen

'  If InStr(UBSetUp(1).UTILNAME, "PRINCETON") > 0 Then
'    PrinceFlag = True
'  End If

'  If InStr(UBSetUp(1).UTILNAME, "WADE") > 0 Then
'    WadeFlag = True
'  End If

  If InStr(UBSetUp(1).UTILNAME, "YADKIN") > 0 Then
    YadkinFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "NORWOOD") > 0 Then
    NorwoodFlag = True
  End If

  If InStr(UBSetUp(1).UTILNAME, "CONWAY") > 0 Then
    ConwayFlag = True
  End If

  If InStr(UBSetUp(1).UTILNAME, "ELKTON") > 0 Then
    ElkFlag = True
  End If

  If InStr(UBSetUp(1).UTILNAME, "SCOTTSBURG") > 0 Then
    ScottFlag = True
  End If

  If InStr(UBSetUp(1).UTILNAME, "SUMMERDALE") > 0 Then
    DaleFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "COVE CITY") > 0 Then
    CoveCityFlag = True
  End If
  '111698 Prorate
  ReDim ProrateServ(1 To 15) As Integer
  ReDim ElecRev(1 To 15) As Integer

  'find the electric revenue position
  If FuelAdjAmt# <> 0 Then
    For ThisRevCnt = 1 To 15
      If InStr(UBSetUp(1).Revenues(ThisRevCnt).RevName, "ELECTRIC") Then
        ElecRev(ThisRevCnt) = ThisRevCnt
        'Exit For
      End If
    Next
  Else
    For ThisRevCnt = 1 To 15
      ElecRev(ThisRevCnt) = -1
    Next
  End If

  '111698 Prorate
  For ThisRevCnt = 1 To 15
    If UBSetUp(1).Revenues(ThisRevCnt).ProRate = "Y" Then
      ProrateServ(ThisRevCnt) = True
    End If
  Next
  ReDim UBBillRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType

  UBBillRecLen = Len(UBBillRec(1))
  UBCustRecLen = Len(UBCustRec(1))

  NumOfRates = GetNumRateRecs%

  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTbls(1))

  RateFile = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
  For cnt = 1 To NumOfRates
    Get RateFile, cnt, UBRateTbls(cnt)
  Next
  Close

  NumOfCustRecs& = FileSize&(UBPath$ + "UBCUST.DAT") \ UBCustRecLen

  If Exist(UBPath$ + UBFinBillsFile) Then
    Kill UBPath$ + UBFinBillsFile
  End If

  BillFile = FreeFile
  Open UBPath$ + UBFinBillsFile For Random Shared As BillFile Len = UBBillRecLen

  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen

  BillCnt = 0
  NumofRevs = GetNumOfRevs%

  For lcnt = 1 To NumOfCustRecs&
    Get CustFile, lcnt, UBCustRec(1)

'STOP
'    IF LCnt = 10668 THEN
'      STOP
'    END IF

    ReDim UBBillRec(1) As UBTransRecType
    If UBCustRec(1).Status = "F" Then
      GotCustFlag = True
    Else
      GotCustFlag = False
    End If

    MeterConsp& = 0
    TMeterConsp& = 0

    If Not GotCustFlag Then
      UBBillRec(1).Transamt = 0
      For RCnt = 1 To NumofRevs
        UBBillRec(1).RevAmt(RCnt) = 0
      Next
      UBBillRec(1).ActiveFlag = False
      GoTo NotAFinal
    End If

    '111698 Prorate
    ProRateFlag = False
    ProPct# = 100
    If UBCustRec(1).ProRatePCT < 100 And UBCustRec(1).ProRatePCT > 0 Then
      UBBillRec(1).ProRatePCT = UBCustRec(1).ProRatePCT
      UBLog "MBF: Prorate Account No:" + Str$(lcnt&) + "  @" + QPTrim$(Str$(UBBillRec(1).ProRatePCT)) + "%"
      ProPct# = Round#(UBBillRec(1).ProRatePCT * 0.01)
      ProRateFlag = True
    Else
      UBBillRec(1).ProRatePCT = 100
    End If

    For FRCnt = 1 To 4
      WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
      If UBCustRec(1).FlatRates(FRCnt).FRAMT > 0 And WhatService > 0 Then
        '11/19/96 Fixed Rev. amt. to add to current amt
        '01-09-97 Fixed Multiplier problem in flat rates
        Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
        If Multi < 1 Then Multi = 1
        FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
        '111698 Prorate
        If ProRateFlag And ProrateServ(WhatService) Then
          FlatAmt# = Round#(FlatAmt# * ProPct#)
        End If
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + FlatAmt#)
        UBBillRec(1).Transamt = Round(UBBillRec(1).Transamt + FlatAmt#)
        If UBSetUp(1).Revenues(WhatService).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          UBBillRec(1).TaxAmt(WhatService) = Round(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).Transamt = Round(UBBillRec(1).Transamt + UBBillRec(1).TaxAmt(WhatService))
        End If
      End If
    Next

    For RCnt = 1 To NumofRevs   'look at each rev line
      MeterConsp& = 0
      TMeterConsp& = 0
      GoSub GetWhatRateTable
      'WhatTbl = FindRateTbl(UBCustRec(1).Serv(RCnt).RATECODE, NumOfRates, UBRateTbls())
      If WhatTbl Then
        If UBSetUp(1).Revenues(RCnt).UseMtr = "N" Then
          'if this is a non-metered service
          '111398 Prorate
          NonMAmt# = UBRateTbls(WhatTbl).MINAMT
          If ProRateFlag And ProrateServ(RCnt) Then
            NonMAmt# = Round#(NonMAmt# * ProPct#)
          End If
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + NonMAmt#)
          '11/19/96 Fixed Rev. amt. to add to current amt
          UBBillRec(1).Transamt = Round(UBBillRec(1).Transamt + NonMAmt#)
          GoTo GotAmt
        End If

        MeterType$ = UBCustRec(1).serv(RCnt).RMtrType
        MeterLocNum = 0

        For MCnt = 1 To 7
        If MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType Then

        'For MCnt = 1 To 7
          CubMtr = False
          If MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType Then
            MeterLocNum = MCnt
            If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
              CubMtr = True
            End If
            'Found correct meter
            MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
              MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
            End If
    'Remark out the following 3 lines per Dale 12/11/03
'102505 unremark per Dale for Mowasa problem w/reg and final not same
'remarked again 4/1/2008 when cashion called with final not correct amt
'            If CubMtr Then
'              MeterConsp& = MeterConsp& * 7.481
'            End If
            If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
              MeterConsp& = MeterConsp& * UBCustRec(1).LocMeters(MCnt).MTRMulti
            End If
            UBBillRec(1).CurRead(MCnt) = UBCustRec(1).LocMeters(MCnt).CurRead
            UBBillRec(1).PrevRead(MCnt) = UBCustRec(1).LocMeters(MCnt).PrevRead
            UBBillRec(1).MtrTypes(MCnt) = GetCustMeterType(UBCustRec(), MCnt)
'Added the following if statement on 2/12/2009 to fix issue with Mowasa not calc correctly - this is directly from regular prebill makebillfile.
           If (UBBillRec(1).MtrTypes(MCnt) = 1 Or UBBillRec(1).MtrTypes(MCnt) = 2 Or UBBillRec(1).MtrTypes(MCnt) = 3) And UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
              MeterConsp& = MeterConsp& * 7.481
              'convert units from cubic feet to gallons here
            End If
            'convert units here if necessary
            TMeterConsp& = TMeterConsp& + MeterConsp&
          End If
        End If
        Next
        If MeterLocNum = 0 Then
          Unload FrmShowPctComp
          If ErrorScrn(4, lcnt&) Then
            AbortFlag = True
            GoTo AbortExit
          End If
        End If
'Replaced the code below with code "& ....'& to fix missing code for max amts on rate codes
''        AddRevAmt# = 0

''        If UBCustRec(1).LocMeters(MeterLocNum).NumUser > 1 Then
''          '100798 Corrected NumUser calc bug.  'Hillsville'
''          'adjust min consumption for calc below
''          '032803 corrected bug (incorrect num users after first rev.
''          '''NumUser& = UBCustRec(1).LocMeters(RCnt).NumUser - 1
''          NumUser& = UBCustRec(1).LocMeters(MeterLocNum).NumUser - 1
''          AddRevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
''          MinimumConsp& = NumUser& * UBRateTbls(WhatTbl).MINUNITS
''          TMeterConsp& = TMeterConsp& - MinimumConsp&
''          If (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 Then
''            '062697 fix for min consump test to actual (NumUsers * MINUNITS)
'''071201 Added fix for prorating
'''pro dale
''            ProRevAmt# = 0
''            If ProRateFlag And ProrateServ(RCnt) Then
''              ProRevAmt# = Round#((AddRevAmt# + UBRateTbls(WhatTbl).MINAMT) * ProPct#)
''              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + ProRevAmt#)
''              UBBillRec(1).Transamt = Round#(UBBillRec(1).Transamt + ProRevAmt#)
''            Else
''              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
''              UBBillRec(1).Transamt = Round#(UBBillRec(1).Transamt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
''            End If
''            GoTo GotAmt
''          End If
'''          IF (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 THEN
'''            UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (A
'''            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + (AddRevAmt
'''            GOTO GotAmt
'''          END IF
''
''        Else
''          NumUser& = 1
''        End If

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        AddRevAmt# = 0
        TMaxAmt# = 0
        If UBRateTbls(WhatTbl).MaxAmt > 0 Then
          TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
        End If
        If UBCustRec(1).LocMeters(MeterLocNum).NumUser > 1 Then
          TMaxAmt# = Round#(UBRateTbls(WhatTbl).MaxAmt * UBCustRec(1).LocMeters(MeterLocNum).NumUser)
          'adjust min consumption for calc below
          NumUser& = UBCustRec(1).LocMeters(MeterLocNum).NumUser - 1
          AddRevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
          MinimumConsp& = NumUser& * UBRateTbls(WhatTbl).MINUNITS
          TMeterConsp& = TMeterConsp& - MinimumConsp&
          If (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 Then
            '062697 fix for min consump test to actual (NumUsers * MINUNITS)
            '071201 Added fix for prorating
            ProRevAmt# = 0
            If ProRateFlag And ProrateServ(RCnt) Then
              ProRevAmt# = Round#((AddRevAmt# + UBRateTbls(WhatTbl).MINAMT) * ProPct#)
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + ProRevAmt#)
              UBBillRec(1).Transamt = Round#(UBBillRec(1).Transamt + ProRevAmt#)
            Else
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
              UBBillRec(1).Transamt = Round#(UBBillRec(1).Transamt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            End If
            GoTo GotAmt
          End If
        Else
          NumUser& = 1
          If TMaxAmt# > 0 Then
            TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
          End If
        End If


'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        '033198 Added code to Calc correctly for Conway...
        If ConwayFlag Then
          If TMeterConsp& Mod 1000 Then
            TMeterConsp& = (Int(TMeterConsp& / 1000) + 1)
          Else
            TMeterConsp& = Int(TMeterConsp& / 1000)
          End If
        End If
        '033198 Conway *********

        '052998 Added code to calc correctly for Princeton
        If PrinceFlag Or WadeFlag Or ScottFlag Or DaleFlag Or CoveCityFlag Then
          If TMeterConsp& Mod 1000 Then
            TMeterConsp& = (Int(TMeterConsp& / 1000) + 1)
          Else
            TMeterConsp& = Int(TMeterConsp& / 1000)
          End If
          TMeterConsp& = TMeterConsp& * 1000
        ElseIf YadkinFlag Then
          If TMeterConsp& Mod 1000 Then
            TMeterConsp& = TMeterConsp& / 1000
            TMeterConsp& = TMeterConsp& * 1000
          End If
        End If
        'Princeton*****

        If TMeterConsp& <= UBRateTbls(WhatTbl).MINUNITS Then
          'if we bill the minium
          RevAmt# = Round#(NumUser& * UBRateTbls(WhatTbl).MINAMT)
          'ADDED!!!!!!!!!!!!!
          If ProRateFlag And ProrateServ(RCnt) Then
            RevAmt# = Round#(RevAmt# * ProPct#)
          End If
          'ADDED!!!!!!!!!!!!!
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
          '11/19/96 Fixed Rev. amt. to add to current amt
          UBBillRec(1).Transamt = Round#(UBBillRec(1).Transamt + RevAmt#)
          GoTo GotAmt
        End If

        '04-22-97 Fixed to add to current rev amt
        '05-29-97 Refixed
        RevAmt# = GetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
        RevAmt# = RevAmt# + AddRevAmt#

        '111398 Prorate
        If ProRateFlag And ProrateServ(RCnt) Then
          RevAmt# = Round#(RevAmt# * ProPct#)
        End If
        If TMaxAmt# > 0 Then
          If RevAmt# > TMaxAmt# Then
            RevAmt# = TMaxAmt#
          End If
        End If

        UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
'NORWOOD Minimum Check START
        If RCnt = 2 And NorwoodFlag Then
          If Left$(UBCustRec(1).ZONE, 1) = "I" And UBBillRec(1).RevAmt(RCnt) < 6.77 Then
            UBBillRec(1).RevAmt(RCnt) = 6.77
            RevAmt# = 6.77
          End If
        End If
        If RCnt = 2 And NorwoodFlag Then
          If Left$(UBCustRec(1).ZONE, 1) = "O" And UBBillRec(1).RevAmt(RCnt) < 13.54 Then
            UBBillRec(1).RevAmt(RCnt) = 13.54
            RevAmt# = 13.54
          End If
        End If
'NORWOOD Minimum Check END

        UBBillRec(1).Transamt = Round(UBBillRec(1).Transamt + RevAmt#)
        If RCnt = ElecRev(RCnt) Then
          AdjRev# = Round(UBBillRec(1).RevAmt(RCnt) * FuelAdjAmt#)
          UBBillRec(1).RevAmt(RCnt) = Round(UBBillRec(1).RevAmt(RCnt) + AdjRev#)
          UBBillRec(1).Transamt = Round(UBBillRec(1).Transamt + AdjRev#)
        End If
GotAmt:
        If UBSetUp(1).Revenues(RCnt).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          UBBillRec(1).TaxAmt(RCnt) = Round(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
          UBBillRec(1).Transamt = Round(UBBillRec(1).Transamt + UBBillRec(1).TaxAmt(RCnt))
        End If
      End If
    Next        'loop through all revenue sources

    BillCnt = BillCnt + 1
    UBBillRec(1).ActiveFlag = True
    UBBillRec(1).CustAcctNo = lcnt

'04-08-99  Elkton
    If ElkFlag Then
      If UBBillRec(1).ActiveFlag Then
        For TZCnt = 1 To 15
          If UBBillRec(1).TaxAmt(TZCnt) > 0 Then
            Ctype$ = QPTrim$(UBCustRec(1).CUSTTYPE)
            Select Case Ctype$
            Case "R"
              If UBBillRec(1).TaxAmt(TZCnt) > 2 Then
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 2)
                UBBillRec(1).Transamt = Round#(UBBillRec(1).Transamt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 2
              End If
            Case "C"
              If UBBillRec(1).TaxAmt(TZCnt) > 20 Then
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 20)
                UBBillRec(1).Transamt = Round#(UBBillRec(1).Transamt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 20
              End If
            Case Else
              'BlockClear
'              DisplayUBScrn "ERRSCRN1"
'              QPrintRC "Invalid Customer Type!", 10, 36, -1
'              QPrintRC "ACCOUNT:" + Str$(lcnt), 10, 22, -1
'              QPrintRC "Correct and Print Again.", 13, 28, -1
'              WaitForAction
'              AbortFlag = True
              GoTo AbortExit
            End Select
          End If
        Next
      End If
    End If
'ELKton

NotAFinal:
    'If lcnt = 1309 Then Stop
    Put BillFile, lcnt, UBBillRec(1)
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If
    'NotAFinal:
    'ShowPctComp lcnt, NumOfCustRecs&
  Next
AbortExit:
  Close
    ActivateControls Me, True
  Exit Sub
GetWhatRateTable:
  WhatTbl = 0
  Ratecode$ = QPTrim$(UBCustRec(1).serv(RCnt).Ratecode)
  If Len(Ratecode$) Then        'if this rev has a rate code
    For TCnt = 1 To NumOfRates  'find the right one
      If Ratecode$ = QPTrim$(UBRateTbls(TCnt).Ratecode) Then
        WhatTbl = TCnt
        Exit For
      End If
    Next
  End If
Return

End Sub
