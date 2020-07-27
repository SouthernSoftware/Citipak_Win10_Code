VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPreBilling 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pre-Billing Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPreBilling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5508
      TabIndex        =   2
      Top             =   4392
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmPreBilling.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5508
      TabIndex        =   3
      Top             =   4932
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmPreBilling.frx":0BED
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
      TabIndex        =   11
      Top             =   7080
      Width           =   1332
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
      TabIndex        =   10
      Top             =   7080
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
            TextSave        =   "3:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "6/17/2003"
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5508
      TabIndex        =   0
      Top             =   2820
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      MaxLength       =   2
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
   Begin EditLib.fpText fptxtAdjustment 
      Height          =   348
      Left            =   5508
      TabIndex        =   1
      Top             =   3636
      Width           =   732
      _Version        =   196608
      _ExtentX        =   1291
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
      MaxLength       =   2
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
   Begin VB.Line Line2 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   3504
      X2              =   8904
      Y1              =   4152
      Y2              =   4152
   End
   Begin VB.Line Line1 
      BorderStyle     =   4  'Dash-Dot
      X1              =   3504
      X2              =   8880
      Y1              =   3408
      Y2              =   3408
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Use Book ""99"" for all books."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6288
      TabIndex        =   13
      Top             =   2856
      Width           =   3132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the fuel adjustment amount for your Electric Service."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6360
      TabIndex        =   12
      Top             =   3528
      Width           =   3060
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3252
      Left            =   2460
      Top             =   2400
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
      Left            =   3084
      TabIndex        =   9
      Top             =   4980
      Width           =   2388
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Book Number:"
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
      Left            =   3492
      TabIndex        =   8
      Top             =   2880
      Width           =   1932
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
      Left            =   3708
      TabIndex        =   7
      Top             =   4440
      Width           =   1716
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
      Index           =   2
      Left            =   3348
      TabIndex        =   6
      Top             =   3696
      Width           =   2076
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   720
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Pre-Billing Report"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   600
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
Attribute VB_Name = "frmPreBilling"
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
  frmUBBillingMenu.Show
  Unload frmPreBilling
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
      End If
    End If
  End If
End Sub
Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fptxtAdjustment.Enabled = True Then
      fptxtAdjustment.SetFocus
    End If
  End If
End Sub
Private Sub fptxtAdjustment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboPrintOrder.SetFocus
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
        fptxtAdjustment.SetFocus
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
Private Function alloktogo()
  If Len(fptxtRoute1.Text) <> 0 Then
    alloktogo = True
  Else
    alloktogo = False
  End If
End Function
Private Sub cmdPrint_Click()
  If alloktogo Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
     'do graphic report
      Grpt = True
      PreBillReport
    ElseIf fpcboRptType.ListIndex = 1 Then
      Grpt = False
      PreBillReport
    End If
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
  StatusBar1.Panels.Item(1).Text = TownName$
  Grpt = False
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.AddItem "Postal Carrier Route Order"
  fpcboPrintOrder.AddItem "ZipCode Order"
  GetPreBillReady
  fpcboPrintOrder.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub GetPreBillReady()
  Dim DoFuel As Boolean, cnt As Integer, TempRev As String
  Dim UBSetupLen As Integer, NumOfRevs As Integer
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupLen

  If UBSetUp(1).PreByBook = "Y" Then
     LabelB1.Caption = "Book Number"
     fptxtRoute1 = "99"
  ElseIf UBSetUp(1).BILLCYCL = "Y" Then
     LabelB1.Caption = "Cycle Number"
     Label5.Caption = "Enter Cycle to use for this billing."
     fptxtRoute1 = " "
  ElseIf UBSetUp(1).PreByBook = "N" And UBSetUp(1).BILLCYCL = "N" Then
     LabelB1.Caption = "Book Number"
     fptxtRoute1 = "99"
     fptxtRoute1.Enabled = False
  End If
  For cnt = 1 To MaxRevsCnt     'find last active revenue
    TempRev$ = UCase$(QPTrim$(UBSetUp(1).Revenues(cnt).REVNAME))
    If Len(TempRev$) = 0 Then
      NumOfRevs = cnt - 1       'set actual number of revenues
      Exit For
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
  Dim Temp2 As String, NumOfRevs As Integer, NumOfRates As Integer
  Dim UBRateTblRecLen As Integer, RateFile As Integer, cnt As Long
  Dim UBSetupLen As Integer, MowFlag As Boolean, TennFlag As Boolean
  Dim TempRev As String, DoFuelAdjFlag As Boolean, SkipInactive As Boolean
  Dim SkipSeparator As Boolean, ThisBook As Integer, BookNum As Integer
  Dim BookFlag As Boolean, ThisCycle As Integer, CycleFlag As Boolean
  Dim SeqFlag As String, Choice As Integer, FuelAdjAmt As Double
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
  UBLog "IN: Prebilling Report"
  
'  If Exist("UBBILLS.DAT") And Exist("UBBILLS.PRN") Then
'    UBLog "ERROR: UNPOSTED BILLING DETECTED!"
'    UBLog "ASKING USER WANT TO CONTINUE?"
'    OK = PreBillYouSure%
'    If Not OK Then
'      UBLog "USER ABORTED PREBILLING."
'      AbortFlag = True
'      GoTo ExitPreReport
'    Else
'      UBLog "USER WANTS TO CONTINUE!"
'      KillFile ("UBBILLS.PRN")
'    End If
'  End If
  PageNo = 0
  Temp2$ = Space$(12)
  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
  NumOfRates = GetNumRateRecs%
  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTbls(1))

  ReDim RateConsump(1 To NumOfRates) As Double

  RateFile = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
  For cnt = 1 To NumOfRates
    Get RateFile, cnt, UBRateTbls(cnt)
  Next
  Close
  
  'SortT UBRateTbls(1), NumOfRates, 0, UBRateTblRecLen, 0, 4
  RateQSort UBRateTbls(), 1, NumOfRates
'  SortT MDateIdx(1), FoundCnt, 0, 4, 0, -1
'  'Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
'  QPrintRC "      Writing Index Records      ", 11, 25, -1
'''  IndexName$ = TempIndexName
'''  KillFile IndexName$
'''  IHandle = FreeFile
'''    'FCreate IndexName$
'''  Open IndexName$ For Random Shared As IHandle Len = 4
'''  For cnt = 1 To FoundCnt
'''    CRec =
'''    Put IHandle, cnt, CRec
'''    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
'''  Next
'''  Close IHandle
'''
'''  Erase UBCustRec, MDateIdx

  ReDim ProrateServ(1 To 15) As Integer

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  TownName$ = UBSetUpRec(1).UTILNAME
  If InStr(TownName$, "MOWAS") > 0 Then
    MowFlag = True
  End If

  If UBSetUpRec(1).DEFSTATE = "TN" Then
    TennFlag = True
  End If

  ReDim RevDesc(1 To MaxRevsCnt) As String * 12
  For cnt = 1 To MaxRevsCnt     'find last active revenue
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(cnt).REVNAME)
    If Len(TempRev$) = 0 Then
      NumOfRevs = cnt - 1       'set actual number of revenues
      Exit For
    Else        'build revenue description lines
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

  If UBSetUpRec(1).SkipInactive = "Y" Then
    SkipInactive = True
  End If

  If UBSetUpRec(1).SkipSeparator = "Y" Then
    SkipSeparator = True
  End If

  If UBSetUpRec(1).PreByBook = "Y" Then
    ThisBook = Val(fptxtRoute1)
    If ThisBook = 99 Then
      ThisBook = -1
    End If
    BookNum = ThisBook
    If ThisBook = -1 Then
      BookFlag = False
    ElseIf ThisBook <= 0 Then
      GoTo ExitPreReport
    Else
      BookFlag = True
    End If
  ElseIf UBSetUpRec(1).BILLCYCL = "Y" Then
    ThisCycle = Val(fptxtRoute1)
    If ThisCycle <= 0 Then
      GoTo ExitPreReport
    Else
      CycleFlag = True
    End If
  End If

  If UBSetUpRec(1).UseSeq = "Y" Then
    SeqFlag$ = "Y"
  End If
  FrmShowPctComp.Label1 = "Creating PreBilling Report"
  FrmShowPctComp.Show , Me

Restart:
  Choice = fpcboPrintOrder.ListIndex + 1
  'GetPreBillOrder Choice, ExitFlag, SeqFlag$

 'If ExitFlag Then GoTo ExitPreReport
  If DoFuelAdjFlag Then
    FuelAdjAmt# = Val(fptxtAdjustment)
    UBLog "Fuel adjustment factor:" + Str$(FuelAdjAmt#)
  Else
    FuelAdjAmt# = 0
  End If

  If FuelAdjAmt# = -10000 Then GoTo Restart

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
  Case 3        'Location
    IndexName$ = BookIndexFile
    'OkFlag = True
  Case 4        'Postal Route
    IdxTypeText$ = "Postal Route"
    MakePostalIndex IdxTypeText$
    IndexName$ = TempIndexName
    'OkFlag = True
  Case 5        'ZipCode
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
  Case 6        'Sequence number
    IdxTypeText$ = "Sequence Number"
    MakeSequenceIndex IdxTypeText$, Me
    IndexName$ = TempIndexName
    'OkFlag = True
  End Select
  MakeBillFile AbortFlag, FuelAdjAmt#, ThisCycle, ThisBook

  If AbortFlag Then GoTo ExitPreReport
  If Grpt Then
    MaxLines = 48
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

  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBBillRec(1) As UBTransRecType
  UBBillRecLen = Len(UBBillRec(1))

  ReDim FlatTotals(1 To NumOfRevs) As Double
  '021998 added flat revenue totals
  ReDim RevTotals(1 To NumOfRevs) As Double     'Revenue total amts
  '052097 added tax by revenue totals
  ReDim TaxTotals(1 To NumOfRevs) As Double     'Tax total amts
  ReDim ConsumpTot(1 To NumOfRevs, 1 To 2) As Double            'Consumption total amts
  ReDim RateConsump(1 To NumOfRates) As Double
  '012698 Added bill count by rate code
  ReDim RateCount(1 To NumOfRates) As Long
  ReDim RateTotals(1 To NumOfRates) As Double   'Rates total amts
  '052097 added tax by rate code totals
  ReDim RTaxTot(1 To NumOfRates) As Double      'Rates Tax total amts
  '052097 added tax by book totals to type def
  ReDim Bookconsump(0 To 1) As BookConsumpType  'Consumption by book
  ReDim PumpConsump(0 To 1) As PumpConsumpType  'Consumption by pump code
  ReDim TaxExmp(0 To NumOfRevs) As Double

  TBooks = 0
  If UsingAcct Then
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  Else          'load the index
    UBLog "Loading index file: " + IndexName$
    IdxRecLen = 4
    NumOfRecs = FileSize(IndexName$) \ 4
    ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For lcnt& = 1 To NumOfRecs
      Get #Handle, lcnt&, IndexArray(lcnt&)
    Next
    Close Handle
    'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  End If
  
  UBBill = FreeFile
  Open UBBillsFile For Random Shared As UBBill Len = UBBillRecLen
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  ReportFile$ = UBPath$ + "UBPREBIL.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  'BlockClear
  'ShowProcessingScrn "Processing Pre-Billing Report"
  UBLog "Writing prebilling report to disk."

  GoSub PrintPreHeader
  For cnt = 1 To NumOfRecs
    If UsingAcct Then
      ThisCustRec& = cnt
    Else
      ThisCustRec& = IndexArray(cnt).RecNum
    End If

    Get UBCust, ThisCustRec&, UBCustRec(1)

    If UBCustRec(1).DelFlag Then
      GoTo SkipEm
    End If

    If SkipInactive And UBCustRec(1).Status <> "A" Then
      GoTo SkipEm
    ElseIf UBCustRec(1).Status = "F" Then       'skip over final's
      GoTo SkipEm
    ElseIf UBCustRec(1).Status = "B" Then       'skip over B-Status
      GoTo SkipEm
    End If
    If BookFlag Then
      If Val(UBCustRec(1).Book) <> ThisBook Then
        GoTo SkipEm
      End If
    End If

    If CycleFlag Then
      If UBCustRec(1).BILLCYCL <> ThisCycle Then
        GoTo SkipEm
      End If
    End If

    Get UBBill, ThisCustRec&, UBBillRec(1)

    If Linecnt >= MaxLines Then
      Print #UBRpt, FF$
      GoSub PrintPreHeader
    End If

    If UBBillRec(1).ActiveFlag <> 0 Then
      If UBCustRec(1).BillTo = "O" Then
        BillTo$ = " O"
      Else
        BillTo$ = " C"
      End If
      GoSub GetWhatBook
      If BadBookFlag Then
        Unload FrmShowPctComp
        If ErrorScrn(2, ThisCustRec&) Then
          AbortFlag = True
          Exit For
        End If
      End If
      Bookconsump(WhatBook).CustCnt = Bookconsump(WhatBook).CustCnt + 1
      Print #UBRpt, UBCustRec(1).Status; Using("  #####  ", ThisCustRec&);
      Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; Left$(UBCustRec(1).CustName, 25); " "; Left$(UBCustRec(1).SERVADDR, 22); " ";
      Print #UBRpt, Using("   ###", UBBillRec(1).ProRatePCT); "%";
      Print #UBRpt, BillTo$
      Linecnt = Linecnt + 1
      For FRCnt = 1 To 4
        WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
        If UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 And WhatService > 0 Then
          Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
          If Multi < 1 Then Multi = 1
          FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
          '021998 Added flat rate summaries
          FlatTotals(WhatService) = Round#(FlatTotals(WhatService) + FlatAmt#)
        End If
      Next
      '102798 Added to skip accts that don't have a book/seq no. "J.R."
    ElseIf Len(QPTrim$(UBCustRec(1).Book)) = 0 And Len(QPTrim$(UBCustRec(1).SEQNUMB)) = 0 Then
      GoTo SkipEm
    End If
    WhatRate = 0
    DoneOne = False
    For TRevCnt = 1 To NumOfRevs
      If TRevCnt = 2 And UBBillRec(1).PenAtBill = -1 Then
        IFlag = True
      Else
        IFlag = False
      End If
      WhatRate = 0
      If UBBillRec(1).RevAmt(TRevCnt) <> 0 Then
        DoneOne = False
        Print #UBRpt, RevDesc(TRevCnt);
        '102198 Moved out of meter loop, Stoped multi meter tax report bug
        If UBBillRec(1).TaxAmt(TRevCnt) > 0 Then
          TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
        End If
        For TRateCnt = 1 To NumOfRates
          If UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE Then
            MINAMT& = UBRateTbls(TRateCnt).MINUNITS
            WhatRate = TRateCnt
            '102198 Moved from meter loop, Stops multi meter tax report bug
            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
            Exit For
          End If
        Next
        If UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" Then
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))

          '02-20-97 Add revenue totals by rate code
          If WhatRate > 0 Then
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
          End If
          PrintedRevAmt = False
          For MCCnt = 1 To 7
            CubMtr = False
            LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
            MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
            '063098 Added adjustment for cubic meters in consumption totals
            If UBCustRec(1).LocMeters(MCCnt).MTRUnit = "C" Then
              CubMtr = True
            End If
            If MeterMulti& <= 0 Then MeterMulti& = 1
            If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMtrType) Then
              DoneOne = True
              MeterNum$ = QPTrim$(UBCustRec(1).Serv(TRevCnt).RATECODE)
              'use the Meternum$ to hold the rate code temporarily
              If Len(MeterNum$) > 0 Then
                If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
                  MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                End If
                RSet Temp2$ = MeterNum$
              End If
              ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
              If ReadAmt& < 0 Then              'Meter rolled over or, been misread
                MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MCCnt))) - 1)
                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
              End If
              If CubMtr Then
                ReadAmt& = ReadAmt& * 7.481
              End If
              RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
              RateCount(WhatRate) = RateCount(WhatRate) + 1
              Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
              ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
              Consump& = ReadAmt& * MeterMulti&
              ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
              If ThisMeterUseCnt <= 0 Then ThisMeterUseCnt = 1
              AvgUse& = UBCustRec(1).LocMeters(MCCnt).AvgUse
              If AvgUse& > 0 Then
                HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * 0.01))
                LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * 0.01))
              End If
              Print #UBRpt, Tab(14); Temp2$; Tab(30); Using(fmt$(2), UBBillRec(1).CurRead(MCCnt)); Tab(42); UBBillRec(1).PrevRead(MCCnt); Tab(54); ReadAmt& * MeterMulti&;
              If UBCustRec(1).EstFlag = "E" Then
                Print #UBRpt, " E";             'Est. Reading
              ElseIf Consump& < LowConsump& Then
                Print #UBRpt, " L";             'Low reading
              ElseIf Consump& > HiConsump& Then
                Print #UBRpt, " H";             'High Reading
              End If
              If Consump& < MINAMT& Then
                Print #UBRpt, " M";             'Minium Usage
              End If
              If UBBillRec(1).RevAmt(TRevCnt) > 0 And PrintedRevAmt = False Then
                PrintedRevAmt = True
                Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
                If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
                  Print #UBRpt, "*";
                End If
                If IFlag Then
                  Print #UBRpt, " IR";
                End If

              End If
              Print #UBRpt,
              Linecnt = Linecnt + 1
            End If
          Next
          '071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
          '      a metered service but no meter on a flat rate charge. Charge was added
          '      to total, but didn't show on prebilling report.
          If Not DoneOne Then
            DoneOne = True
            Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
            If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
              Print #UBRpt, "*";
            End If
            'THIS WAS REMARKED OUT, I DON'T KNOW WHY?
            Print #UBRpt,
            ''''''''''''''''''''''''''''''''''''''
            Linecnt = Linecnt + 1
          End If
        Else    'it's a nonmetered service
          ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
          If WhatRate > 0 Then
            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
            RateCount(WhatRate) = RateCount(WhatRate) + 1
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
          End If
          Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + 1
          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
          If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
            Print #UBRpt, "*";
          End If
        End If
        If Not DoneOne Then
          Print #UBRpt,
          Linecnt = Linecnt + 1
        End If
      End If
      If (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt = 0 Then
        If UBBillRec(1).TransAmt = 0 Then       'CONSUMPTION inactive account
          For TTRevCnt = 1 To NumOfRevs
            For MCCnt = 1 To 7
              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
              If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMtrType) Then
                If UBBillRec(1).CurRead(MCCnt) < 0 Then
                  UBBillRec(1).CurRead(MCCnt) = 0
                End If
                If UBBillRec(1).PrevRead(MCCnt) < 0 Then
                  UBBillRec(1).PrevRead(MCCnt) = 0
                End If
                CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
                PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
                If CurReadAmt& <> PreReadAmt& Then
                  If Not ConsumpFlag Then
                    Print #UBRpt, UBCustRec(1).Status; Using("     #####   ", ThisCustRec&);
                    Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "   "; Left$(UBCustRec(1).CustName, 25); "  "; Left$(UBCustRec(1).SERVADDR, 25)
                    Linecnt = Linecnt + 1
                  End If
                  ConsumpFlag = True
                  MeterNum$ = QPTrim$(UBCustRec(1).Serv(TTRevCnt).RATECODE)
                  If Len(MeterNum$) > 0 Then
                    If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
                      MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                    End If
                    RSet Temp2$ = MeterNum$
                  End If
                  ConsumpAmt& = CurReadAmt& - PreReadAmt&
                  '103098 Added meter roll over check to inactive consumption
                  If ConsumpAmt& < 0 Then       'Meter rolled over or, been misread
                    MaxMeterAmt& = 10& ^ (Len(Str$(PreReadAmt&)) - 1)
                    ConsumpAmt& = (MaxMeterAmt& - PreReadAmt&) + CurReadAmt&
                  End If
                  If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
                    'For Nonprofits include consumption as normal   'cleveland
                    '040998 Made changes here
                    For NONRateCnt = 1 To NumOfRates
                      If UBRateTbls(NONRateCnt).RATECODE = UBCustRec(1).Serv(TTRevCnt).RATECODE Then
                        NONRate = NONRateCnt
                        Exit For
                      End If
                    Next
                    If NONRate > 0 Then
                      RateConsump(NONRate) = RateConsump(NONRate) + ConsumpAmt&
                    End If
                    ConsumpTot(TTRevCnt, 1) = ConsumpTot(TTRevCnt, 1) + ConsumpAmt&
                    Bookconsump(WhatBook).Consump(TTRevCnt) = Bookconsump(WhatBook).Consump(TTRevCnt) + ConsumpAmt&
                    '040998 Made changes here 'cleveland
                  Else          'add consumption to inactives
                    ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
                  End If
                  Print #UBRpt, RevDesc(TTRevCnt); Tab(14); Temp2$; Tab(30); Using(fmt$(2), CurReadAmt&); Tab(42); Using(fmt$(2), PreReadAmt&); Tab(54); Using(fmt$(2), ConsumpAmt&)
                  Linecnt = Linecnt + 1
                End If
              End If
            Next
          Next
        End If
        If ConsumpFlag And UBCustRec(1).Status <> "A" Then
          ConsumpFlag = False
          Print #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
          Linecnt = Linecnt + 1
          If Not SkipSeparator Then
            Print #UBRpt, fmt$(0)
            Linecnt = Linecnt + 1
          End If
        ElseIf ConsumpFlag Then
          'Customer Status is "A"
          'This happens when a cust has consumption and there rate code
          'has a zero calc amount. "i.e. a Church or other nonprofit"
          If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
            Print #UBRpt, "*** NON-PROFIT ***"
            Linecnt = Linecnt + 1
          End If
          ConsumpFlag = False
          If Not SkipSeparator Then
            Print #UBRpt, fmt$(0)
            Linecnt = Linecnt + 1
          End If
        End If
      ElseIf (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt > 0 Then
        '102998  Moved tax printing to here "now prints one tax line per customer
        CTaxAmt# = 0
        For TXCnt = 1 To 15
          If UBBillRec(1).TaxAmt(TXCnt) > 0 Then
            CTaxAmt# = Round#(CTaxAmt# + UBBillRec(1).TaxAmt(TXCnt))
          End If
        Next
        If CTaxAmt# > 0 Then
          Print #UBRpt, " Tax"; Tab(69); Using(fmt$(3), CTaxAmt#)
          Linecnt = Linecnt + 1
        End If
        Bills2Print = Bills2Print + 1
        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
        Print #UBRpt, Tab(5); "Current:"; Using(fmt$(6), UBBillRec(1).TransAmt);
        If AcctBalance# <> 0 Then
          Print #UBRpt, Tab(30); "Previous:"; Using(fmt$(6), AcctBalance#);
          TAcctBalance# = Round#(TAcctBalance# + AcctBalance#)
        End If
        Print #UBRpt, Tab(55); "Total:"; Tab(66); Using(fmt$(6), Round#(AcctBalance# + UBBillRec(1).TransAmt))
        Linecnt = Linecnt + 1
        If Not SkipSeparator Then
          Print #UBRpt, fmt$(0)
          Linecnt = Linecnt + 1
        End If
      End If
      If UBBillRec(1).TaxExempt = "Y" Then
        TaxExmp(TRevCnt) = Round#(TaxExmp(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
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
        Case "C", "S", "W", "T"
          PumpMtrOK = True
        End Select
        If PumpMtrOK Then
          MeterMulti& = UBCustRec(1).LocMeters(MPCnt).MTRMulti
          If UBCustRec(1).LocMeters(MPCnt).MTRUnit = "C" Then
            CubMtr = True
          End If
          If MeterMulti& <= 0 Then MeterMulti& = 1
          ReadAmt& = UBBillRec(1).CurRead(MPCnt) - UBBillRec(1).PrevRead(MPCnt)
          If ReadAmt& < 0 Then  'Meter rolled over or, been misread
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
'    If AskAbandonPrint% Then
'      UBLog "ABORTED: Prebilling report"
'      UBLog "Closing files."
'      Close
'      AbortFlag = True
'      Exit For
'    End If
'    ShowPctComp cnt, NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitPreReport
    End If

  Next
  If AbortFlag Then GoTo ExitPreReport

  Print #UBRpt, FF$

  GoSub TitleLine
  Print #UBRpt, "Billing Grand Totals"
  If TennFlag Then
    Print #UBRpt, "                                Inactive          Taxed      NONTax     FlatRate"
    Print #UBRpt, "Revenue/Tax        Consump       Consump         Amount      Amount      Amount"
  Else
    Print #UBRpt, "                                 Inactive                             Flat Rate"
    Print #UBRpt, "Revenue/Tax    Consumption      Consumption            Amount           Amount"
  End If
  Print #UBRpt, fmt$(0)

  TotalFlatAmt# = 0
  TotalRevAmt# = 0
  TotalTaxAmt# = 0

  For RaCnt = 1 To NumOfRevs
    If TennFlag Then
      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(fmt$(4), ConsumpTot(RaCnt, 1)); Tab(30); Using(fmt$(4), ConsumpTot(RaCnt, 2));
      If TaxTotals(RaCnt) > 0 Then
        Print #UBRpt, Tab(44); Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt) - TaxExmp(RaCnt))); Tab(56); Using(fmt$(1), TaxExmp(RaCnt)); Tab(68); Using(fmt$(1), FlatTotals(RaCnt))
      Else
        Print #UBRpt, Tab(44); Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt))); Tab(68); Using(fmt$(1), FlatTotals(RaCnt))
      End If
    Else
      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(fmt$(4), ConsumpTot(RaCnt, 1)); Tab(33); Using(fmt$(4), ConsumpTot(RaCnt, 2));
      Print #UBRpt, Tab(50); Using(fmt$(1), RevTotals(RaCnt) - FlatTotals(RaCnt)); Tab(67); Using(fmt$(1), FlatTotals(RaCnt))
    End If
    TotalFlatAmt# = Round#(TotalFlatAmt# + FlatTotals(RaCnt))
    TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
    If TaxTotals(RaCnt) > 0 Then
      If TennFlag Then
        Print #UBRpt, " Tax"; Tab(44); Using(fmt$(1), TaxTotals(RaCnt))
      Else
        Print #UBRpt, " Tax"; Tab(50); Using(fmt$(1), TaxTotals(RaCnt))
      End If
      TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
    End If
  Next
  Print #UBRpt, fmt$(0)
  Print #UBRpt, "  PREVIOUS: "; Using(fmt$(6), TAcctBalance#);
  Print #UBRpt, Tab(33); "REVENUE TOTAL: "; Using(fmt$(5), Round#(TotalRevAmt# - TotalFlatAmt#))
  Print #UBRpt, "BILL COUNT: "; Using(fmt$(2), Bills2Print);
  Print #UBRpt, Tab(33); "   FLAT TOTAL: "; Using(fmt$(5), TotalFlatAmt#)
  Print #UBRpt, Tab(33); "    TAX TOTAL: "; Using(fmt$(5), TotalTaxAmt#)
  Print #UBRpt, Tab(33); "BILLING TOTAL: "; Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
  Print #UBRpt, FF$

  TotalRevAmt# = 0

  GoSub RptTotRateHeader

  For RaCnt = 1 To NumOfRates
    If (RateTotals(RaCnt) <> 0) Or (RateConsump(RaCnt) <> 0) Then
      If Len(QPTrim$(UBRateTbls(RaCnt).RATECODE)) > 0 Then
        Print #UBRpt, UBRateTbls(RaCnt).RATECODE; "    "; UBRateTbls(RaCnt).RATEDESC; Tab(39); Using(fmt$(4), RateConsump(RaCnt));
        Print #UBRpt, Tab(55); Using(fmt$(1), RateTotals(RaCnt));
        Print #UBRpt, Tab(69); Using(fmt$(2), RateCount(RaCnt))
        Linecnt = Linecnt + 1
        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
        If RTaxTot(RaCnt) > 0 Then
          Print #UBRpt, " Tax"; Tab(55); Using(fmt$(1), RTaxTot(RaCnt))
          Linecnt = Linecnt + 1
        End If
        If Linecnt >= MaxLines Then
          Print #UBRpt, FF$
          GoSub RptTotRateHeader
        End If
      End If
    End If
  Next

  Print #UBRpt, fmt$(0)
  Print #UBRpt, Tab(36); "TAX TOTAL:"; Tab(53); Using(fmt$(5), TotalTaxAmt#)
  Print #UBRpt, Tab(40); "TOTAL:"; Tab(53); Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
  Print #UBRpt, FF$
  'SortT BookConsump(1), TBooks, 0, Len(BookConsump(1)), 0, -1
  BookCQSort Bookconsump(), 1, TBooks
  GoSub BookHeader

  For cnt = 1 To TBooks
    TestTot# = 0
    For ZCnt = 1 To NumOfRevs
      TestTot# = Round#(TestTot# + Bookconsump(cnt).RevAmt(ZCnt))
    Next
    If TestTot# <> 0 Then
      If Bookconsump(cnt).Book < 10 Then
        Book$ = "0" + QPTrim$(Str$(Bookconsump(cnt).Book))
      Else
        Book$ = QPTrim$(Str$(Bookconsump(cnt).Book))
      End If
      Print #UBRpt, "Book: "; Book$; "    Customers:"; Bookconsump(cnt).CustCnt
'*******
      Linecnt = Linecnt + 1
      TBookAmt# = 0
      TBTaxAmt# = 0
      For RCnt = 1 To NumOfRevs
        Print #UBRpt, RevDesc(RCnt); Tab(30); Using(fmt$(4), Bookconsump(cnt).Consump(RCnt));
        Print #UBRpt, Tab(59); Using("##########.##", Bookconsump(cnt).RevAmt(RCnt))
        TBookAmt# = Round#(TBookAmt# + Bookconsump(cnt).RevAmt(RCnt))
        If Bookconsump(cnt).TaxAmt(RCnt) > 0 Then
          Print #UBRpt, " Tax"; Tab(60); Using(fmt$(1), Bookconsump(cnt).TaxAmt(RCnt))
          TBTaxAmt# = Round#(TBTaxAmt# + Bookconsump(cnt).TaxAmt(RCnt))
          Linecnt = Linecnt + 1
        End If
        Linecnt = Linecnt + 1
      Next
      TBookGTot# = Round#(TBookGTot# + TBookAmt# + TBTaxAmt#)
      Print #UBRpt, Tab(42); "Book Total:"; Tab(58); Using(fmt$(5), Round#(TBookAmt# + TBTaxAmt#))
      Linecnt = Linecnt + 1
      If cnt < TBooks Then
        Print #UBRpt, fmt$(0)
'******
        Linecnt = Linecnt + 1
      End If
      
    End If
    If ((Linecnt + NumOfRevs) >= MaxLines) And (cnt < TBooks) Then
      Print #UBRpt, FF$
      GoSub BookHeader
    End If

SkipThisBook:
  Next

  Print #UBRpt, fmt$(0)
  Print #UBRpt, Tab(35); "Books GRAND Total:"; Tab(58); Using(fmt$(5), TBookGTot#)
  Print #UBRpt, FF$

  If TPumps > 0 Then
    GoSub PumpHeader
    TMMConsump# = 0
    For cnt = 1 To TPumps
      Print #UBRpt, PumpConsump(cnt).PumpCode; Tab(30); Using("###########", PumpConsump(cnt).CustCnt); Tab(60); PumpConsump(cnt).Consump
      TMMConsump# = TMMConsump# + PumpConsump(cnt).Consump
    Next
    Print #UBRpt, fmt$(0)
    Print #UBRpt, Tab(35); "Pump Code Total:"; Tab(60); Using("###########", TMMConsump#)
  End If

  Close

  UBLog "Finished writing prebilling report."
  Select Case Choice
  Case 1
    RptText$ = "(Customer"
  Case 2
    RptText$ = "(Account"
  Case 3
    RptText$ = "(Location"
  Case 4
    RptText$ = "(Postal RT."
  Case 5
    RptText$ = "(ZipCode"
  Case 6
    RptText$ = "(Sequence"
  End Select
  RptText$ = RptText$ + " Order)"

  Erase UBSetUpRec, RevDesc, UBRateTbls, RateConsump
  Erase fmt$, UBCustRec, UBBillRec, FlatTotals
  Erase RevTotals, TaxTotals, ConsumpTot
  Erase RateTotals, RTaxTot, Bookconsump, IndexArray
  Erase RateCount, ProrateServ
  Erase PumpConsump, TaxExmp

  If Not AbortFlag Then
    If Grpt Then
      Load frmLoadingRpt
      ARptPreBilling.Title = "Utility Pre-Billing Report"
      ARptPreBilling.txtDate = Now
      ARptPreBilling.txtTown = TownName$
      ARptPreBilling.GetName ReportFile$
      ARptPreBilling.startrpt
    Else
      ViewPrint ReportFile$, "Pre-Billing Report " + RptText$
    'PrintRptFile "Pre-Billing Report " + RptText$, "UBPREBIL.RPT", LPTPort, RetCode, EntryPoint
    End If
    If BookFlag Then
      Kill UBBillsFile
    End If
  End If

  GoTo ExitPreReport

PrintPreHeader:
  GoSub TitleLine
  Print #UBRpt, "Stat  Act.  Locat    Customer Name             Service Address       Prorate%"
  Print #UBRpt, "Revenue            R-Code     Cur Read    Pre Read     Consump        Charges"
  Print #UBRpt, fmt$(0)
  Linecnt = 5
Return

GetWhatBook:
  BadBookFlag = False
 WhatBook = 0
 If Len(QPTrim$(UBCustRec(1).Book)) = 0 Then
   If UBCustRec(1).Status = "A" Then
     BadBookFlag = True
     'testing vvv
     WhatBook = 0
   End If
   GoTo ErrorBookExit
 End If

 ThisBook = Val(UBCustRec(1).Book)
 If TBooks > 0 Then
   For TBCnt = 1 To TBooks
     If Bookconsump(TBCnt).Book = ThisBook Then
       WhatBook = TBCnt
       Exit For
     End If
   Next
   If WhatBook = 0 Then
     TBooks = TBooks + 1
     ReDim Preserve Bookconsump(0 To TBooks) As BookConsumpType
      Bookconsump(TBooks).Book = ThisBook
      WhatBook = TBooks
    End If
  Else
    TBooks = TBooks + 1
    Bookconsump(TBooks).Book = ThisBook
    WhatBook = TBooks
  End If

ErrorBookExit:
  Return

GetWhatPump:
  HasAPumpCode = True           'assume they have a pump code
  WhatPump = 0
  If Len(QPTrim$(UBCustRec(1).PumpCode)) = 0 Then
    If UBCustRec(1).Status = "A" Then
      HasAPumpCode = False      'no pump code
      WhatPump = 0
    End If
    GoTo PumpCodeReturn
  End If

  CustPump$ = UCase$(QPTrim$(UBCustRec(1).PumpCode))

  'IF CustPump$ = "34" THEN STOP

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

RptTotRateHeader:
  GoSub TitleLine
  Print #UBRpt,
  Print #UBRpt, "Report Totals by Rate Code"
  Print #UBRpt,
  Print #UBRpt, "Code      Rate Description            Consumption           Amount      Bills"
  Print #UBRpt, fmt$(0)
  Linecnt = 5
  Return

BookHeader:
  GoSub TitleLine
  Print #UBRpt, "Report Totals by Book"
  Print #UBRpt,
  Print #UBRpt, "Book"
  Print #UBRpt, "Revenue                      Consumption                         Amount"
  Print #UBRpt, fmt$(0)
  Linecnt = 7
  Return

PumpHeader:
  GoSub TitleLine
  Print #UBRpt, "Report Totals by Pump Code"
  Print #UBRpt,
  Print #UBRpt, "PumpCode                  Customer Count                    Consumption"
  Print #UBRpt, fmt$(0)
  Linecnt = 6
  Return

TitleLine:
  If Grpt Then GoTo SkipTitle:
  PageNo = PageNo + 1
  Print #UBRpt, "Utility Pre-Billing Report.  "; TownName$; Tab(70); "Page: "; PageNo
  Print #UBRpt, TheDate$
SkipTitle:
  Return
ErrorAbortExit:
  Close

ExitPreReport:
  UBLog "OUT: Prebilling Report" + CrLf$

End Sub
Public Sub MakeBillFile(AbortFlag, FuelAdjAmt#, ThisCycle%, ThisBook%)
  Dim UBSetupLen As Integer, PrinceFlag As Boolean, YadkinFlag As Boolean
  Dim WadeFlag As Boolean, ElkFlag As Boolean, ScottFlag As Boolean
  Dim DaleFlag As Boolean, SunBchFlag As Boolean, SkipInactive As Boolean
  Dim BookFlag As Boolean, CycleFlag As Boolean, ThisRevCnt As Integer
  Dim ElecRev As Integer, UBBillRecLen As Integer, UBCustRecLen As Integer
  Dim NumOfRates As Integer, UBRateTblRecLen As Integer, RateFile As Integer
  Dim cnt As Integer, lcnt As Long, NumCustRec As Long, BillFile As Integer
  Dim CustFile As Integer, BillCnt As Integer, NumOfRevs As Integer
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
  Dim Ctype As String, Diff As Double, RATECODE As String
  Dim TCnt As Integer
 ' BlockClear
 ' ShowProcessingScrn "Calculating Utility Charges."

  UBLog "IN: MakeBillFile."
  UBLog "MBF: Calculating charges."

  ReDim ProrateServ(1 To 15) As Integer

  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupLen

  If InStr(UBSetUp(1).UTILNAME, "PRINCETON") > 0 Then
    PrinceFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "YADKIN") > 0 Then
    YadkinFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "WADE") > 0 Then                'OR INSTR(UBSetUp(1).UTILNAME, "WADE") THEN
    WadeFlag = True
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
  If InStr(UBSetUp(1).UTILNAME, "SUNSET") > 0 Then
    SunBchFlag = True
  End If

  If UBSetUp(1).SkipInactive = "Y" Then
    SkipInactive = True
  End If

  If UBSetUp(1).PreByBook = "Y" And ThisBook > 0 Then
    BookFlag = True
  ElseIf UBSetUp(1).BILLCYCL = "Y" Then
    CycleFlag = True
  End If

  'find the electric revenue position
  For ThisRevCnt = 1 To 15
    If InStr(UBSetUp(1).Revenues(ThisRevCnt).REVNAME, "ELECTRIC") Then
      ElecRev = ThisRevCnt
      Exit For
    End If
  Next

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
  Close RateFile

  NumCustRec& = FileSize&("UBCUST.DAT") \ UBCustRecLen

  If Exist(UBBillsFile) Then
    KillFile UBBillsFile
  End If
  BillFile = FreeFile
  Open UBBillsFile For Random Shared As BillFile Len = UBBillRecLen

  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen

  BillCnt = 0
  NumOfRevs = GetNumOfRevs%

  For lcnt& = 1 To NumCustRec&
    ReDim UBBillRec(1) As UBTransRecType        'clear bill rec for this custo
    Get CustFile, lcnt&, UBCustRec(1)
    'IF LCnt& = 2 THEN STOP
    If UBCustRec(1).DelFlag <> 0 Then
      UBBillRec(1).TransAmt = 0
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = lcnt&
      GoTo MSkipEm
    End If

    If SkipInactive And UBCustRec(1).Status = "I" Then
      UBBillRec(1).TransAmt = 0
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = lcnt&
      GoTo MSkipEm
    End If

    If BookFlag Then
      If Val(UBCustRec(1).Book) <> ThisBook Then
        UBBillRec(1).TransAmt = 0
        For RCnt = 1 To NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        Next
        For zz = 1 To 7
          UBBillRec(1).CurRead(zz) = 0
          UBBillRec(1).PrevRead(zz) = 0
        Next
        UBBillRec(1).ActiveFlag = False
        GoTo MSkipEm
      End If
    End If
    If CycleFlag Then
      If UBCustRec(1).BILLCYCL <> ThisCycle Then
        UBBillRec(1).TransAmt = 0
        For RCnt = 1 To NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        Next
        For zz = 1 To 7
          UBBillRec(1).CurRead(zz) = UBCustRec(1).LocMeters(zz).CurRead
          UBBillRec(1).PrevRead(zz) = UBCustRec(1).LocMeters(zz).PrevRead
          UBBillRec(1).MtrTypes(zz) = GetCustMeterType(UBCustRec(), zz)
        Next
        UBBillRec(1).ActiveFlag = False
        GoTo MSkipEm
      End If
    End If

    If UBCustRec(1).Status <> "A" Then
      UBBillRec(1).TransAmt = 0
      For RCnt = 1 To NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
        UBBillRec(1).TaxAmt(RCnt) = 0
      Next
      For zz = 1 To 7
        UBBillRec(1).CurRead(zz) = UBCustRec(1).LocMeters(zz).CurRead
        UBBillRec(1).PrevRead(zz) = UBCustRec(1).LocMeters(zz).PrevRead
        UBBillRec(1).MtrTypes(zz) = GetCustMeterType(UBCustRec(), zz)
      Next
      UBBillRec(1).ActiveFlag = False
      GoTo MSkipEm
    End If
    '052698 Added tax exempt flag to bill rec
    UBBillRec(1).TaxExempt = UBCustRec(1).TAXEXPT

'SunnyBeach 091701
    If SunBchFlag Then
      GotIRRMtr = False
      IrrConsp& = 0
      For IrrMtr = 1 To 7
        IrrMtrNum$ = QPTrim$(UBCustRec(1).LocMeters(IrrMtr).MtrNum)
        If Left$(IrrMtrNum$, 3) = "IRR" Then
          GotIRRMtr = True
          IrrConsp& = UBCustRec(1).LocMeters(IrrMtr).CurRead - UBCustRec(1).LocMeters(IrrMtr).PrevRead
          If IrrConsp& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(UBCustRec(1).LocMeters(IrrMtr).PrevRead)) - 1)
            IrrConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(IrrMtr).PrevRead) + UBCustRec(1).LocMeters(IrrMtr).CurRead
          End If
          Exit For
        End If
      Next
    End If

    '111398 Prorate
    ProRateFlag = False
    ProPct# = 100
    If UBCustRec(1).ProRatePCT < 100 And UBCustRec(1).ProRatePCT > 0 Then
      UBBillRec(1).ProRatePCT = UBCustRec(1).ProRatePCT
      UBLog "MBF: Prorated Account No:" + Str$(lcnt&) + "  @" + QPTrim$(Str$(UBBillRec(1).ProRatePCT)) + "%"
      ProPct# = Round#(UBBillRec(1).ProRatePCT * 0.01)
      ProRateFlag = True
    Else
      UBBillRec(1).ProRatePCT = 100
    End If
    MeterConsp& = 0
    TMeterConsp& = 0

    'look at flat rates
    For FRCnt = 1 To 4
      WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
      If WhatService > NumOfRevs Then
        Unload FrmShowPctComp
        If ErrorScrn(6, lcnt&) Then
          AbortFlag = True
          GoTo AbortExit
        End If
      End If
      If UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 And WhatService > 0 Then
        '11/19/96 Fixed Rev. amt. to add to current rev amt
        If UBCustRec(1).FlatRates(FRCnt).FRAMT < -1000000 Then
          Unload FrmShowPctComp
          If ErrorScrn(6, lcnt&) Then
            AbortFlag = True
            GoTo AbortExit
          End If
        End If
        '01-09-97 Fixed Multiplier bug in flat rates
        Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
        If Multi < 1 Then Multi = 1
        FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
        '111398 Prorate
        If ProRateFlag And ProrateServ(WhatService) Then
          FlatAmt# = Round#(FlatAmt# * ProPct#)
        End If
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + FlatAmt#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FlatAmt#)
        If UBSetUp(1).Revenues(WhatService).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          TaxAmt# = Round#(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TaxAmt(WhatService) = TaxAmt#
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(WhatService))
        End If
      End If
    Next
    'end of flat rates
    '12-6-96  Monthly Billed amounts
    For MRCnt = 1 To 2
      WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
      If WhatService > NumOfRevs Or WhatService < 0 Then
        'IF ErrorScrn(7, LCnt&) THEN
        '  AbortFlag = True
        '  GOTO AbortExit
        'END IF
      End If

      If UBCustRec(1).Monthly(MRCnt).PayAmt > 0 And WhatService > 0 Then
        TestAmt# = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + UBCustRec(1).Monthly(MRCnt).PayAmt)
        If TestAmt# > UBCustRec(1).Monthly(MRCnt).AMTOWED Then
          HowMuch# = Round#(UBCustRec(1).Monthly(MRCnt).AMTOWED - UBCustRec(1).Monthly(MRCnt).TotAmtPD)
        Else
          HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
        End If
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + HowMuch#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + HowMuch#)
        If UBSetUp(1).Revenues(WhatService).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          TaxAmt# = Round#(HowMuch# * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TaxAmt(WhatService) = Round#(UBBillRec(1).TaxAmt(WhatService) + TaxAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + TaxAmt#)
        End If
      End If
    Next

    For RCnt = 1 To NumOfRevs   'look at each rev line
      MeterConsp& = 0
      TMeterConsp& = 0
      GoSub GetWhatRateTable
      If WhatTbl Then
        If UBSetUp(1).Revenues(RCnt).UseMtr = "N" Then
          'if this is a non-metered service
          '02-05-97 added fix add to current rev amt
          If UBRateTbls(WhatTbl).MINAMT > -1000000 Then
            NonMAmt# = UBRateTbls(WhatTbl).MINAMT
            If ProRateFlag And ProrateServ(RCnt) Then
              NonMAmt# = Round#(NonMAmt# * ProPct#)
            End If
            UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + NonMAmt#)
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + NonMAmt#)
          Else
            Unload FrmShowPctComp
            RateCodeErrScrn UBRateTbls(WhatTbl).RATECODE
            AbortFlag = True
            GoTo AbortExit
          End If
          GoTo GotAmt
        End If
        'it's metered
        MeterType$ = UBCustRec(1).Serv(RCnt).RMtrType
        MeterLocNum = 0
        For MCnt = 1 To 7
          If MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType Then
            MeterLocNum = MCnt
            UBBillRec(1).CurRead(MCnt) = UBCustRec(1).LocMeters(MCnt).CurRead
            UBBillRec(1).PrevRead(MCnt) = UBCustRec(1).LocMeters(MCnt).PrevRead
            UBBillRec(1).MtrTypes(MCnt) = GetCustMeterType(UBCustRec(), MCnt)
            'Found correct meter
            '052797 Added to stop overflow error.
            If (UBCustRec(1).LocMeters(MCnt).CurRead < 0) Or (UBCustRec(1).LocMeters(MCnt).PrevRead < 0) Then
              Unload FrmShowPctComp
              If ErrorScrn(1, lcnt&) Then
                AbortFlag = True
                GoTo AbortExit
              End If
              MeterConsp& = 0
            Else
              MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            End If
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
              MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
            End If
            If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
              ThisMeterConsp# = (0# + MeterConsp&) * UBCustRec(1).LocMeters(MCnt).MTRMulti
              '                  ^This forces basic to convert to a Double
              '                   before calculation, traps overflow errors
              If ThisMeterConsp# > 2147483647 Then
                '                  ^Max long integer value
                Unload FrmShowPctComp
                If ErrorScrn(1, lcnt&) Then
                  AbortFlag = True
                  GoTo AbortExit
                End If
              End If
              MeterConsp& = ThisMeterConsp#
            End If
            If (UBBillRec(1).MtrTypes(MCnt) = 1 Or UBBillRec(1).MtrTypes(MCnt) = 2 Or UBBillRec(1).MtrTypes(MCnt) = 3) And UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
              MeterConsp& = MeterConsp& * 7.481
              'convert units from cubic feet to gallons here
            End If
            TMeterConsp& = TMeterConsp& + MeterConsp&
          End If
        Next
        If MeterLocNum = 0 Then
          Unload FrmShowPctComp
          If ErrorScrn(4, lcnt&) Then
            AbortFlag = True
            GoTo AbortExit
          End If
        End If
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
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + ProRevAmt#)
            Else
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            End If
            GoTo GotAmt
          End If
        Else
          NumUser& = 1
          If TMaxAmt# > 0 Then
            TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
          End If
          '033198 Added code to Calc correctly for Conway...
          If ConwayFlag Then
            If TMeterConsp& Mod 1000 Then
              TMeterConsp& = (Int(TMeterConsp& / 1000) + 1)
            Else
              TMeterConsp& = Int(TMeterConsp& / 1000)
            End If
          End If
          '033198 Conway *********
          '052998 Added code for calc method Princeton
          'summerdale
          If PrinceFlag Or WadeFlag Or ScottFlag Or DaleFlag Then
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
            If UBRateTbls(WhatTbl).MINAMT > -1000000 Then
              RevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
              If ProRateFlag And ProrateServ(RCnt) Then
                RevAmt# = Round#(RevAmt# * ProPct#)
              End If
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
              GoTo GotAmt
            End If
          End If
        End If
        '01-20-97 Added Fix for minium units test for equal to also.
        '04-23-1997 'Fixed to ADD TO TOTAL
        If RCnt = 2 And SunBchFlag And GotIRRMtr Then
          SewCalcConsp& = TMeterConsp& - IrrConsp&
          RevAmt# = GetRevCharge#(UBRateTbls(WhatTbl), SewCalcConsp&, MeterMulti&)
          UBBillRec(1).PenAtBill = -1
        Else
          RevAmt# = GetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
        End If
        RevAmt# = Round(RevAmt# + AddRevAmt#)

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
        If RCnt = ElecRev Then
          FuelAddAmt# = Round#(FuelAdjAmt# * TMeterConsp&)
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + FuelAddAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FuelAddAmt#)
        End If
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
GotAmt:
        If UBSetUp(1).Revenues(RCnt).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          TaxAmt# = Round#(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
          UBBillRec(1).TaxAmt(RCnt) = Round#(UBBillRec(1).TaxAmt(RCnt) + TaxAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + TaxAmt#)
        End If
      Else
        If Len(QPTrim$(UBCustRec(1).Serv(RCnt).RMtrType)) > 0 Then
          Unload FrmShowPctComp
          If ErrorScrn(3, lcnt&) Then
            AbortFlag = True
            GoTo AbortExit
          End If
        End If
      End If
    Next        'loop through all revenue sources
    If UBCustRec(1).Status = "I" And UBBillRec(1).TransAmt > 0 Then
      UBBillRec(1).TransAmt = 0
      For RCnt = 1 To NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
      Next
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = lcnt&
    Else
      UBBillRec(1).ActiveFlag = False
    End If

    'Mod for cleveland***
    If UBCustRec(1).CUSTTYPE = "NON" Then
      UBBillRec(1).CustAcctNo = lcnt&
      UBBillRec(1).NONProfit = "Y"
    End If
    '********************

    If UBBillRec(1).TransAmt > 0 Then
      BillCnt = BillCnt + 1
      UBBillRec(1).ActiveFlag = True
      UBBillRec(1).CustAcctNo = lcnt&
    End If
    '0727 Added NEW trap for a meter defined with no rate code.
    '    FOR MTstCnt = 1 TO 7
    '      IF LEN(QPTrim$(UBCustRec(1).LocMeters(MTstCnt).MTRType)) > 0 THEN
    '        FOR MTCnt = 1 TO 7
    '          IF UBBillRec(1).MtrTypes(MTCnt) > 0 THEN
    '            GOTO ThereOK
    '          END IF
    '          IF ErrorScrn(8, LCnt&) THEN
    '            AbortFlag = True
    '            GOTO AbortExit
    '          END IF
    '        NEXT
    '      END IF
    '    NEXT

    '04-07-99 Added special tax calc for elkton
    If ElkFlag Then
      If UBBillRec(1).ActiveFlag Then
        For TZCnt = 1 To 15
          If UBBillRec(1).TaxAmt(TZCnt) > 0 Then
            Ctype$ = QPTrim$(UBCustRec(1).CUSTTYPE)
            Select Case Ctype$
            Case "R"
              If UBBillRec(1).TaxAmt(TZCnt) > 2 Then
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 2)
                UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 2
              End If
            Case "C"
              If UBBillRec(1).TaxAmt(TZCnt) > 20 Then
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 20)
                UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 20
              End If
            Case Else
              Unload FrmShowPctComp
              If ErrorScrn(9, lcnt&) Then
                AbortFlag = True
                GoTo AbortExit
              End If
            End Select
          End If
        Next
      End If
    End If

MSkipEm:
    Put BillFile, lcnt&, UBBillRec(1)
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If
'    ShowPctComp lcnt&, NumCustRec&
  Next

AbortExit:

  Close BillFile, CustFile

  If AbortFlag Then
    UBLog "MBF: ABORTED!"
  Else
    UBLog "MBF: Finished calculations."
  End If
  UBLog "OUT: MakeBillFile."
  
  Erase UBBillRec, UBCustRec, UBSetUp, UBRateTbls
  Exit Sub
  '*******************************

GetWhatRateTable:
  WhatTbl = 0
  RATECODE$ = QPTrim$(UBCustRec(1).Serv(RCnt).RATECODE)
  If Len(RATECODE$) Then        'if this rev has a rate code
    For TCnt = 1 To NumOfRates  'find the right one
      If RATECODE$ = QPTrim$(UBRateTbls(TCnt).RATECODE) Then
        WhatTbl = TCnt
        Exit For
      End If
    Next
  End If

  Return

End Sub
