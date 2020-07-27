VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmDraftAccountsRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts to Draft"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmDraftAccountsRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5424
      TabIndex        =   1
      Top             =   5064
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmDraftAccountsRpt.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5424
      TabIndex        =   2
      Top             =   5616
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmDraftAccountsRpt.frx":0BED
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
      Left            =   9456
      TabIndex        =   4
      Top             =   7392
      Width           =   1332
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F10 &Ok"
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
      Left            =   7776
      TabIndex        =   3
      Top             =   7392
      Width           =   1332
   End
   Begin EditLib.fpText fptxtCycleSel 
      Height          =   348
      Left            =   6120
      TabIndex        =   0
      Top             =   2592
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
            TextSave        =   "8:36 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "1/2/2008"
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
   Begin EditLib.fpText fptxtcycle 
      Height          =   372
      Left            =   3000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4152
      Width           =   6252
      _Version        =   196608
      _ExtentX        =   11028
      _ExtentY        =   656
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AutoAdvance     =   0   'False
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
      Left            =   3456
      TabIndex        =   13
      Top             =   5112
      Width           =   1716
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press F10 to process selections."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3648
      TabIndex        =   12
      Top             =   3408
      Width           =   3468
   End
   Begin VB.Line Line1 
      X1              =   3984
      X2              =   9648
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Cycles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   3816
      Width           =   2076
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   2460
      Left            =   2424
      Top             =   2280
      Width           =   7284
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1524
      Left            =   2424
      Top             =   4728
      Width           =   7284
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE:  Enter a '0' for all Cycles, or leave blank if do not bill by cycle."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   3000
      TabIndex        =   9
      Top             =   3192
      Width           =   6588
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
      Left            =   2856
      TabIndex        =   8
      Top             =   5640
      Width           =   2388
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts to Draft Report"
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
      Left            =   3636
      TabIndex        =   6
      Top             =   960
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   720
      Width           =   5772
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Cycle:"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   1932
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   600
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmDraftAccountsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Dim Grpt As Boolean, CycleCnt As Integer
Dim Cycle(1 To 16) As Integer

Private Sub cmdExit_Click()
  frmUBDraftMenu.Show
  Unload frmDraftAccountsRpt
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Close via DraftAcctsRpt by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fptxtCycleSel_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtCycleSel_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cnt As Integer
  If KeyCode = vbKeyReturn Then
    If Len(fptxtCycleSel.Text) <> 0 Then
      getcyclelist
    Else
      cmdOk.SetFocus
    End If
  End If
End Sub

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdOk.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPrintOrder.SetFocus
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
        cmdOk.SetFocus
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
Private Sub cmdOk_Click()
  Dim Grpt As Boolean
    
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
     'do graphic report
      Grpt = True
    ElseIf fpcboRptType.ListIndex = 1 Then
      Grpt = False
    End If
    UBAcctsToDraft Grpt
 
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
      cmdOk_Click
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
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  fpcboPrintOrder.AddItem "Bank Order"
  fpcboPrintOrder.AddItem "Customer Account Number Order"
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.ListIndex = 0
  CycleCnt = 0
  Me.HelpContextID = hlpAccountsToDraft
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub getcyclelist()
Dim TCyc As String, ThisCycle As Integer, cnt As Integer
  TCyc$ = QPTrim$(fptxtCycleSel.Text)
  If TCyc$ = "0" Then
    fptxtcycle.Text = ""
    CycleCnt = 0
    Erase Cycle
    cmdOk.SetFocus
  Else
    If Len(TCyc$) > 0 Then
      ThisCycle = Val(fptxtCycleSel.Text)
      For cnt = 1 To 16
        If ThisCycle = Cycle(cnt) Then
          GoTo DupeExit
        End If
      Next
      CycleCnt = CycleCnt + 1
      If CycleCnt > 16 Then
        CycleCnt = 16
        GoTo DupeExit
      End If
      Cycle(CycleCnt) = ThisCycle
      fptxtcycle.Text = ""
      For cnt = 1 To CycleCnt
        If cnt = CycleCnt Then
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt)
        Else
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt) & ","
        End If
      Next
    End If
  End If
DupeExit:
  fptxtCycleSel.Text = ""
End Sub

Private Sub UBAcctsToDraft(Grpt As Boolean)
  Dim Dash80 As String, UBSetupLen As Integer, IndexName As String
  Dim OKFlag As Boolean, UBCustRecLen As Integer, UBCust As Integer
  Dim NumOfRecs As Long, UBRpt As Integer, cnt As Long
  Dim CustCycle As Integer, CustOk As Boolean, CCnt As Integer
  Dim CstCnt As Long, llow As Long, hhigh As Long, BankCnt As Integer
  Dim PrevBank As String, GTotal As Double, TabOffSet As Integer
  Dim ReportFile As String, bnameorder As Boolean, GATot As Double
  Dim ToPrint As String, ReportSum As String, SumRpt As Integer
  Dim Dosome As Integer, UsingName As Boolean, IdxRecLen As Integer
  Dim Handle As Integer, lcnt As Long, num As Long
  ToPrint$ = ""
  Dash80$ = String$(80, "-")
  UsingName = False
  ReDim UBSetUpRec(1) As UBSetupRecType
  ReDim BankTotals(1 To 1) As BankTotalsType

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen          'load setup file
  TOWNNAME$ = UBSetUpRec(1).UTILNAME

  FrmShowPctComp.Label1 = "Creating Accounts to Draft Listing"
  FrmShowPctComp.Show , Me

'*********************************
  If fpcboPrintOrder.ListIndex = 0 Then
    bnameorder = True
  End If
  MaxLines = 58
  PageNo = 0
  If fpcboPrintOrder.ListIndex = 1 Then
    IndexName$ = ""
  ElseIf fpcboPrintOrder.ListIndex = 2 Then
    UsingName = True
    IndexName$ = NameIndexFile
  End If
  OKFlag = True
  
  ReDim DFTRec(1) As DraftRptType
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim BDCust(1 To 16) As BDRptType
  UBCustRecLen = Len(UBCustRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs& = LOF(UBCust) \ UBCustRecLen
  If UsingName Then
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
  End If


  UBRpt = FreeFile
  ReportFile$ = UBPath$ + "UBANKDFT.RPT"
  Open ReportFile$ For Output As UBRpt
  
  'ShowProcessingScrn "Processing Bank Draft Report"
  If Not Grpt Then GoSub PrintBankDFTHeader

  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo AbortExit
    End If
    If UsingName Then
      num& = IndexArray(cnt&).RecNum
    Else
      num& = cnt&
    End If
    Get UBCust, num&, UBCustRec(1)
    CustCycle = UBCustRec(1).BILLCYCL
    CustOk = False
    If CycleCnt > 0 Then
      For CCnt = 1 To CycleCnt
        If Cycle(CCnt) = 0 Then
          CustOk = True
        ElseIf CustCycle = Cycle(CCnt) Then
          CustOk = True
          Exit For
        End If
      Next
    Else
      CustOk = True
    End If

    If CustOk Then
      If UBCustRec(1).Status = "A" Or UBCustRec(1).Status = "B" Then
        If (UBCustRec(1).USEDRAFT = "Y") Then  'And (Len(QPTrim$(UBCustRec(1).BankName)) > 0)
          If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) > 0 Then
            If LineCnt > MaxLines Then
              If Not Grpt Then
                Print #UBRpt, Chr$(12)
                GoSub PrintBankDFTHeader
              End If
            End If
            CstCnt = CstCnt + 1
            ReDim Preserve BDCust(1 To CstCnt) As BDRptType
            BDCust(CstCnt).BankName = QPTrim$(UBCustRec(1).BankName)
            BDCust(CstCnt).CustRec = num&
            BDCust(CstCnt).TransRec = num&
          End If
        End If
      End If
    End If
'    If AskAbandonPrint% Then
'      ABExit = True
'      GoTo NON2PrintExit:
'    End If
'    ShowPctCompL cnt&, NumOfRecs&
DFTskipem:
  Next

  If CstCnt <= 0 Then
    If Not Grpt Then
      Print #UBRpt, "No Bills found to Draft."
      Print #UBRpt, Dash80$
    Else
      MsgBox "No Bills found to Draft.", vbOKOnly, "No Drafts"
    End If
    fptxtcycle = ""
    ActivateControls Me, True
    GoTo NON2PrintExit
  End If
  llow = LBound(BDCust)
  hhigh = UBound(BDCust)
  If bnameorder Then
    BDSort BDCust(), llow, hhigh
  End If
  'SortT BDCust(1), CstCnt, 0, 20, 0, 14
  BankCnt = 1
  Get UBCust, BDCust(1).CustRec, UBCustRec(1)

  PrevBank$ = QPTrim$(BDCust(1).BankName)
  BankTotals(BankCnt).BankName = QPTrim$(BDCust(1).BankName)

  For cnt = 1 To CstCnt
    Get UBCust, BDCust(cnt).CustRec, UBCustRec(1)
    If bnameorder Then
      Dosome = 1
      If PrevBank$ <> QPTrim$(BDCust(cnt).BankName) Then
        BankCnt = BankCnt + 1
        ReDim Preserve BankTotals(1 To BankCnt) As BankTotalsType
        BankTotals(BankCnt).BankName = QPTrim$(BDCust(cnt).BankName)
        PrevBank$ = QPTrim$(BDCust(cnt).BankName)
      End If
      BankTotals(BankCnt).Amount = Round#(BankTotals(BankCnt).Amount + UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    Else
      Dosome = 2
    End If
    If LineCnt > MaxLines And Not Grpt Then
      Print #UBRpt, Chr$(12)
      GoSub PrintBankDFTHeader
    End If
    LSet DFTRec(1).TRANSIT = QPTrim$(UBCustRec(1).TRANSIT)
    LSet DFTRec(1).BankName = QPTrim$(UBCustRec(1).BankName)
    RSet DFTRec(1).CustAcct = QPTrim$(Str$(BDCust(cnt).CustRec))
    LSet DFTRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
    LSet DFTRec(1).AcctType = QPTrim$(UBCustRec(1).AcctType)
    LSet DFTRec(1).BankAcct = QPTrim$(UBCustRec(1).BankAcct)
    LSet DFTRec(1).BillAmt = Using$("#####.##", Str$(Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)))
    If Not Grpt Then
      Print #UBRpt, DFTRec(1).TRANSIT; " "; DFTRec(1).BankName;
      Print #UBRpt, DFTRec(1).CustAcct; "  "; DFTRec(1).CustName;
      Print #UBRpt, " "; DFTRec(1).BillAmt; " "; DFTRec(1).AcctType; "  "; DFTRec(1).BankAcct
    Else
      ToPrint$ = QPTrim(DFTRec(1).TRANSIT) + "~" + QPTrim(DFTRec(1).BankName)
      ToPrint$ = ToPrint$ + "~" + Str(DFTRec(1).CustAcct)
      ToPrint$ = ToPrint$ + "~" + QPTrim(DFTRec(1).CustName)
      ToPrint$ = ToPrint$ + "~" + Str(DFTRec(1).BillAmt)
      ToPrint$ = ToPrint$ + "~" + QPTrim(DFTRec(1).AcctType) + "~" + QPTrim$(DFTRec(1).BankAcct)
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
    End If
    GATot# = GATot# + DFTRec(1).BillAmt
    LineCnt = LineCnt + 1
'    If AskAbandonPrint% Then
'      ABExit = True
'      GoTo NON2PrintExit:
'    End If
'    ShowPctComp cnt, CstCnt
  Next
  If Not Grpt Then
    Print #UBRpt, Chr$(12)
    PageNo = PageNo + 1
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, "Utility Billing Bank Draft Register.                "; QPTrim$(TOWNNAME$)
    Print #UBRpt, "Date: "; Date$; Tab(72); "Page: "; PageNo
    If bnameorder Then
      Print #UBRpt, "Bank Name                 Bank Total."
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, Dash80$
  End If
  GTotal# = 0
  If bnameorder Then
    If Not Grpt Then
      For cnt = 1 To BankCnt
        Print #UBRpt, BankTotals(cnt).BankName; Tab(30); Using$("#####.##", Str$(BankTotals(cnt).Amount))
        GTotal# = Round#(GTotal# + BankTotals(cnt).Amount)
      Next
    Else
      ReportSum$ = UBPath$ + "UBSum.RPT"
      SumRpt = FreeFile
      Open ReportSum$ For Output As SumRpt
      For cnt = 1 To BankCnt
        ToPrint$ = BankTotals(cnt).BankName + "~" + Using$("#####.##", Str$(BankTotals(cnt).Amount))
        Print #SumRpt, ToPrint$
        ToPrint$ = ""
        GTotal# = Round#(GTotal# + BankTotals(cnt).Amount)
      Next
    End If
  End If
  If Not Grpt Then
    Print #UBRpt,
    Print #UBRpt, "   Draft Total:"; Tab(30); Using$("#####.##", Str$(GATot#))
    Print #UBRpt,
    Print #UBRpt, "Customer Count:"; Tab(30); CstCnt
    Print #UBRpt,
    Print #UBRpt, "    Cycle List:"
    TabOffSet = 5
    For cnt = 1 To CycleCnt
      Print #UBRpt, Tab(TabOffSet); Using$("#####", Str$(Cycle(cnt)));
      TabOffSet = TabOffSet + 8
      If TabOffSet > 70 Then
        Print #UBRpt,
        TabOffSet = 5
      End If
    Next
    Print #UBRpt,
    Print #UBRpt, Dash80$
    Print #UBRpt, Chr$(12)
  End If
  Close
  CycleCnt = 0
  Erase UBSetUpRec, DFTRec, UBCustRec, Cycle
  If Grpt Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmDraftAccountsRpt
    ARptDraftAccounts.Title = "Bank Draft Register Report"
    ARptDraftAccounts.txtDate = Now
    ARptDraftAccounts.txtTown = TOWNNAME$
    ARptDraftAccounts.totCust = CstCnt
    ARptDraftAccounts.totDraft = Using$("#####.##", Str$(GATot#))
    ARptDraftAccounts.txtCycles = fptxtcycle.Text
    ARptDraftAccounts.GetName ReportFile$, ReportSum$, Dosome
    ARptDraftAccounts.startrpt
  Else
    ViewPrint ReportFile$, "Bank Draft Register Report"
    ActivateControls Me, True
  End If
  'LPTPort = 1
  'If Not AbortFlag Then
  '  PrintRptFile "Bank Draft Register Report", "UBANKDFT.RPT", LPTPort, RetCod
 ' End If

AbortExit:
  Erase UBSetUpRec, DFTRec, UBCustRec, Cycle
  ActivateControls Me, True
  Exit Sub

'CheckInfo1:
'  If CycleCnt > 0 Then
'    InfoOK = True
'  Else
''    SaveScrn TempScrn()
''    DisplayUBScrn "ERRSCRN1"
''    QPrintRC "Invalid Cycle Selection", 10, 28, -1
''    QPrintRC "Press any Key to Continue.", 13, 27, -1
''    WaitForAction
''    RestScrn TempScrn()
''    Erase TempScrn
'  End If
'Return

PrintBankDFTHeader:
  PageNo = PageNo + 1
  Print #UBRpt, " "
  Print #UBRpt, " "
  Print #UBRpt, "Utility Billing Bank Draft Register.                "; QPTrim$(TOWNNAME$)
  Print #UBRpt, "Date: "; Date$; Tab(72); "Page: "; PageNo
  Print #UBRpt, "Bank No.  Bank Name  Acct No.  Customer Name               Amt  Type  Acct"

  Print #UBRpt, Dash80$
  LineCnt = 7

  Return
NON2PrintExit:
  Close

End Sub
