VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptMeterNotes 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meter Reading Notes"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptMeterNotes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5232
      TabIndex        =   3
      Top             =   4944
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ColDesigner     =   "frmRptMeterNotes.frx":08CA
   End
   Begin LpLib.fpCombo fpcboNotes 
      Height          =   348
      Left            =   5232
      TabIndex        =   2
      Top             =   4416
      Width           =   4068
      _Version        =   196608
      _ExtentX        =   7175
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
      ColDesigner     =   "frmRptMeterNotes.frx":0BF8
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
      Left            =   9234
      TabIndex        =   5
      Top             =   7464
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
      Left            =   7554
      TabIndex        =   4
      Top             =   7464
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
            TextSave        =   "11:06 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "4/1/2004"
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   5226
      TabIndex        =   1
      Top             =   3900
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5226
      TabIndex        =   0
      Top             =   3384
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
   Begin VB.Label Exit2Flag 
      Height          =   276
      Left            =   8208
      TabIndex        =   12
      Top             =   1392
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   588
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Notes Only:"
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
      Height          =   324
      Left            =   3030
      TabIndex        =   11
      Top             =   4440
      Width           =   2028
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Meter Reading Notes"
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
      TabIndex        =   10
      Top             =   1368
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1128
      Width           =   5772
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Route:"
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
      Height          =   324
      Left            =   3126
      TabIndex        =   9
      Top             =   3948
      Width           =   1932
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Route:"
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
      Height          =   324
      Left            =   3156
      TabIndex        =   8
      Top             =   3420
      Width           =   1908
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2844
      Left            =   2418
      Top             =   2928
      Width           =   7356
   End
   Begin VB.Label Label2 
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
      Left            =   2760
      TabIndex        =   7
      Top             =   4968
      Width           =   2388
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   1008
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
Attribute VB_Name = "frmRptMeterNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim Grpt As Boolean

Private Sub cmdExit_Click()
  Select Case Me.Exit2Flag
  Case "1"
    Load frmUBMeterMenu
    frmUBMeterMenu.Show
  Case "2"
    Load frmUBHHMenu
    frmUBHHMenu.Show
  End Select
  DoEvents
  Unload frmRptMeterNotes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptMeterNotes by " + PWUser$
        CitiTerminate
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
        fpcboNotes.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboNotes_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboNotes.ListDown = True
  End If
  If fpcboNotes.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Route Selection, The Beginning Route Should Be Less or Equal to Ending Route.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
    End If
  Else
    MsgBox "Route Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function
Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub
Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboNotes.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub cmdPrint_Click()
  Grpt = False
  If ValidRoutes Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
      Grpt = True
      PrintMeterNotes
    ElseIf fpcboRptType.ListIndex = 1 Then
      Grpt = False
      PrintMeterNotes
      ActivateControls Me, True
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
  fptxtRoute1 = "01"
  fptxtRoute2 = "99"
  fpcboNotes.AddItem "YES (Will Clear New Notes Flag!)"
  fpcboNotes.AddItem "NO  (Will NOT Clear New Notes Flag!)"
  fpcboNotes.ListIndex = 1
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub PrintMeterNotes()
  Dim UBCustRecLen As Integer, ReportFile As String
  Dim UBSetupLen As Integer, IdxName As String, lcnt As Long
  Dim NumOfRecs As Long, IdxNumOfRecs As Long, Handle As Integer
  Dim cnt As Long, VacFlag As Boolean, SAddr As String
  Dim Header As String, CustName As String, NumOfCust As Long
  Dim PrintedOne As Boolean, MeterType As String, Book As Integer
  Dim ToPrint As String, IdxRecLen As Long, RptHandle As Integer
  Dim MaxMeterAmt As Long, Page As Integer, IdxFileSize As Long
  Dim DidOne As Boolean, MtrCnt As Long, TempRev As String
  Dim NewFlag As Boolean, UBCust As Integer
  FrmShowPctComp.Label1 = "Creating Meter Notes Report"
  FrmShowPctComp.Show , Me
  
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  MaxLines = 40
  FF$ = Chr$(12)

  'Open Report File
  ReportFile$ = UBPath$ + "UBMTRNOT.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  CustName$ = Space$(30)
  If fpcboNotes.ListIndex = 0 Then
    NewFlag = True
  Else
    NewFlag = False
  End If
  ' Location Order ********************************************************
  IdxName$ = UBPath$ + "UBCUSTBK.IDX"
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&(IdxName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  Handle = FreeFile
  Open IdxName$ For Random Shared As Handle Len = IdxRecLen
  For cnt& = 1 To IdxNumOfRecs
    Get #Handle, cnt&, IdxBuff(cnt&)
  Next
  Close Handle

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  cnt& = 1
  GoSub PrintNotesHeading
 ' ShowProcessingScrn "Reading Meter Information"
  For lcnt& = 1 To IdxNumOfRecs
    FrmShowPctComp.ShowPctComp lcnt, IdxNumOfRecs&
    If FrmShowPctComp.Out = True Then
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitHere
    End If
 
    Get #UBCust, IdxBuff(lcnt&).RecNum, UBCustRec(1)
    Book = Val(UBCustRec(1).Book)
    If Book >= BegRoute And Book <= EndRoute Then
      If UBCustRec(1).DelFlag = 0 Then
        If NewFlag Then
          If UBCustRec(1).NewNotes <> True Then
            GoTo SkipNotes
          Else
            UBCustRec(1).NewNotes = False
            Put #UBCust, IdxBuff(lcnt&).RecNum, UBCustRec(1)
          End If
        End If
        LSet CustName$ = UBCustRec(1).CustName
        If Len(QPTrim$(UBCustRec(1).HHMSG1)) > 0 Or Len(QPTrim$(UBCustRec(1).HHMSG2)) > 0 Or Len(QPTrim$(UBCustRec(1).HHMSG3)) > 0 Then
          If Not Grpt Then
            Print #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; CustName$; Left$(UBCustRec(1).ServAddr, 30)
            Print #RptHandle, Tab(3); QPTrim(UBCustRec(1).HHMSG1); Tab(26); QPTrim(UBCustRec(1).HHMSG2); Tab(49); QPTrim(UBCustRec(1).HHMSG3)
            Print #RptHandle, String$(79, "-")
            LineCnt = LineCnt + 3
          Else
            ToPrint$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + CustName$ + "~" + Left$(UBCustRec(1).ServAddr, 30) + "~"
            ToPrint$ = ToPrint$ + QPTrim(UBCustRec(1).HHMSG1) + "~" + QPTrim(UBCustRec(1).HHMSG2) + "~" + QPTrim(UBCustRec(1).HHMSG3)
            Print #RptHandle, ToPrint$
            ToPrint$ = ""
          End If
        End If
      End If
    End If
    If LineCnt >= MaxLines And Not Grpt Then
      Print #RptHandle, FF$
      GoSub PrintNotesHeading
    End If
SkipNotes:
  Next
  If Not Grpt Then
    Print #RptHandle, FF$
  End If
  Close
  Erase IdxBuff
  Header$ = "Meter Reading Notes"
  'PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  If Not Grpt Then
    ViewPrint ReportFile$, Header$
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptMeterNotes
    ARptMeterNotes.Title = Header$
    ARptMeterNotes.txtDate = Now
    ARptMeterNotes.txtTown = TOWNNAME$
    ARptMeterNotes.GetName ReportFile$
    ARptMeterNotes.startrpt
  End If
  Exit Sub

PrintNotesHeading:
If Not Grpt Then
  Page = Page + 1
  Print #RptHandle, Tab(30); "Meter Reading Notes"; Tab(65); "Date: "; Date$
  Print #RptHandle, "Beginning Route: "; BegRoute
  Print #RptHandle, "   Ending Route: "; EndRoute; Tab(70); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, "Location   Customer Name"; Tab(41); "Service Address"
  Print #RptHandle, Tab(3); "NOTE1"; Tab(26); "NOTE2"; Tab(49); "NOTE3"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
End If
  Return
ExitHere:
  Close
  Erase IdxBuff
  
End Sub