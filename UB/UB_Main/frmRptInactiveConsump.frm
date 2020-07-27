VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmRptInactiveConsump 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inactive Consumption Report"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptInactiveConsump.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   6330
      TabIndex        =   3
      Top             =   4875
      Width           =   1920
      _Version        =   196608
      _ExtentX        =   3387
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
      ColDesigner     =   "frmRptInactiveConsump.frx":08CA
   End
   Begin LpLib.fpCombo fpcboVacant 
      Height          =   375
      Left            =   6330
      TabIndex        =   1
      Top             =   3810
      Width           =   840
      _Version        =   196608
      _ExtentX        =   1482
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
      ColDesigner     =   "frmRptInactiveConsump.frx":0BF8
   End
   Begin LpLib.fpCombo fpcboNoLocation 
      Height          =   375
      Left            =   6330
      TabIndex        =   0
      Top             =   3270
      Width           =   840
      _Version        =   196608
      _ExtentX        =   1482
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
      ColDesigner     =   "frmRptInactiveConsump.frx":0F26
   End
   Begin LpLib.fpCombo fpcboSort 
      Height          =   375
      Left            =   6330
      TabIndex        =   2
      Top             =   4350
      Width           =   3930
      _Version        =   196608
      _ExtentX        =   6932
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
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   0
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
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   0   'False
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptInactiveConsump.frx":1254
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
      TabIndex        =   5
      Top             =   7152
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
      Left            =   9234
      TabIndex        =   4
      Top             =   7152
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "10:21 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "10/22/2007"
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
   Begin VB.Label Label4 
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
      Left            =   4320
      TabIndex        =   11
      Top             =   4392
      Width           =   1788
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Skip Accounts With No Location Number:"
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
      Left            =   1224
      TabIndex        =   10
      Top             =   3312
      Width           =   4908
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
      Left            =   3816
      TabIndex        =   9
      Top             =   4896
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2940
      Left            =   1512
      Top             =   2784
      Width           =   8940
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   912
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Inactive Consumption"
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
      Left            =   3624
      TabIndex        =   8
      Top             =   1152
      Width           =   5004
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vacant Accounts Only:"
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
      Left            =   3336
      TabIndex        =   7
      Top             =   3864
      Width           =   2796
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   792
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
Attribute VB_Name = "frmRptInactiveConsump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Grpt As Boolean
Private Sub cmdExit_Click()
  frmUBMeterMenu.Show
  Unload frmRptInactiveConsump
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptInactiveConsump by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fpcboSort_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboSort.ListDown = True
  End If
  If fpcboSort.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboVacant.SetFocus
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
        fpcboSort.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboVacant_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVacant.ListDown = True
  End If
  If fpcboVacant.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboSort.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboNoLocation.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboNoLocation_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboNoLocation.ListDown = True
  End If
  If fpcboNoLocation.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboVacant.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
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
  Grpt = False
  
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
      Grpt = True
        '
      InactiveConsReport
    ElseIf fpcboRptType.ListIndex = 1 Then
      Grpt = False
      InactiveConsReport
      ActivateControls Me, True
    Else
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
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fpcboNoLocation.AddItem "Yes"
  fpcboNoLocation.AddItem "No"
  fpcboNoLocation.ListIndex = 1
  fpcboVacant.AddItem "Yes"
  fpcboVacant.AddItem "No"
  fpcboVacant.ListIndex = 1
  fpcboSort.AddItem "Customer Name Order"
  fpcboSort.AddItem "Location Number Order"
  fpcboSort.AddItem "Read Sequence Number Order"
  fpcboSort.AddItem "Account Number Order"
  fpcboSort.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpInactive
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub InactiveConsReport()
  Dim UBCustRecLen As Integer, ReportFile As String, RptHandle As Integer
  Dim UBSetupLen As Integer, IdxName As String, lcnt As Long
  Dim NumOfRecs As Long, IdxNumOfRecs As Long, Handle As Integer
  Dim cnt As Long, VacFlag As Boolean, SAddr As String, IdxRecLen As Integer
  Dim Title As String, CustName As String, NumOfCust As Long
  Dim PrintedOne As Boolean, MeterType As String, AcctNumber As Long
  Dim PrintMrtFlag As Boolean, Multi As Double, MeterConsp As Double
  Dim ToPrint As String, ToPrintN As String, NoLocFlag As Boolean
  Dim MaxMeterAmt As Long, Page As Integer, UBCust As Integer
  Dim DidOne As Boolean, MtrCnt As Long, TempRev As String, UsingName As Boolean
  Dim UsingAcct As Boolean, UsingRead As Boolean, UsingBook As Boolean
  Dim IdxFileSize As Long
  FrmShowPctComp.Label1 = "Creating Inactive Consumption Report"
  FrmShowPctComp.Show , Me
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  MaxLines = 52
  FF$ = Chr$(12)
  If fpcboNoLocation.ListIndex = 0 Then
    NoLocFlag = True
  Else
    NoLocFlag = False
  End If

  If fpcboVacant.ListIndex = 0 Then
    VacFlag = True
  Else
    VacFlag = False
  End If
  'this will evaluate to true only if
  'the vac only flag is "Y"

  'Open Report File
  ReportFile$ = UBPath$ + "UBINCONS.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  CustName$ = Space$(25)
  SAddr$ = Space$(25)
  
  IdxRecLen = 4 'we are using a long integer

  Select Case fpcboSort.ListIndex
  Case 0
    UsingName = True
    IdxName$ = UBPath$ + "UBCUSTNM.IDX"
    Title$ = "Inactive Consumption Listing by Name."
  Case 1
    UsingBook = True
    IdxName$ = UBPath$ + "UBCUSTBK.IDX"
    Title$ = "Inactive Consumption Listing by Location."
  Case 2
    UsingRead = True
    MakeSequenceIndex "Sequence Number", frmRptInactiveConsump
    IdxName$ = TempIndexName
    Title$ = "Inactive Consumption Listing by Sequence No."
  Case 3
    UsingAcct = True
    IdxName$ = ""
    Title$ = "Inactive Consumption Listing by Account."
  End Select
  If Not UsingAcct Then
    IdxFileSize& = FileSize(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    NumOfCust& = IdxNumOfRecs
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  Else
    NumOfCust& = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  GoSub PrintInactHeading
 ' ShowProcessingScrn "Reading Meter Information"

  For lcnt& = 1 To NumOfCust&
    FrmShowPctComp.ShowPctComp lcnt, NumOfCust&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      Exit Sub
    End If
    If Not UsingAcct Then
      AcctNumber = IdxBuff(lcnt&).RecNum
    Else
      AcctNumber = lcnt&
    End If

    Get #UBCust, AcctNumber, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      If InStr(UBCustRec(1).HHMSG1, "NOREAD") = 0 Then
        LSet CustName$ = UBCustRec(1).CustName

        LSet SAddr$ = Left$(UBCustRec(1).ServAddr, 30)
        If UBCustRec(1).Status = "I" Or UBCustRec(1).Status = "P" Then '<> "A") AND (UBCustRec(1).Status <> "F") THEN

          If InStr(CustName$, "VACANT") = 0 Then  'When Cust not vacant and vacant flag
            If VacFlag Then                       'only flag is set skip all but vacant ones.
              GoTo VacFlagSkip
            End If
          End If
          If UBCustRec(1).Book = "  " And UBCustRec(1).SEQNUMB = "      " Then
          '  'This if want to leave off the blank loc cust
            If NoLocFlag Then
              GoTo VacFlagSkip
            End If
          End If

          DidOne = False
          For MtrCnt& = 1 To 7  'find last active meter
            TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRType)
            If Len(TempRev$) > 0 Then
              GoSub GetInactMtrTypePrint
              If PrintMrtFlag Then
                Multi# = UBCustRec(1).LocMeters(MtrCnt&).MTRMulti
                If Multi# = 0 Then Multi# = 1
                If UBCustRec(1).LocMeters(MtrCnt&).CurRead < 0 Or UBCustRec(1).LocMeters(MtrCnt&).PrevRead < 0 Then
                  MeterConsp# = 0
                Else
                  MeterConsp# = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
                End If
                If MeterConsp# < 0 Then
                  MaxMeterAmt& = 10& ^ (Len(Str$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                  MeterConsp# = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
                End If
                MeterConsp# = Round#(MeterConsp# * Multi#)

                  If MeterConsp# > 0 Then
                    If Not Grpt Then
                      If Not DidOne = True Then
                        Print #RptHandle, " "; UBCustRec(1).Status; "   "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; Using("  #####", AcctNumber);
                        Print #RptHandle, Tab(25); CustName$; Tab(50); SAddr$
                        DidOne = True
                        PrintedOne = True
                      End If
                      Print #RptHandle, UBCustRec(1).LocMeters(MtrCnt&).MtrNum;
                      Print #RptHandle, Tab(14); MeterType$;
                      Print #RptHandle, Tab(24); Using("#####", Multi#);
                      Print #RptHandle, Tab(31); Using("##########", UBCustRec(1).LocMeters(MtrCnt&).CurRead);
                      Print #RptHandle, Tab(42); Using("##########", UBCustRec(1).LocMeters(MtrCnt&).PrevRead);
                      Print #RptHandle, Tab(53); Using("##########", MeterConsp#);
                      Print #RptHandle, Tab(67); Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
                      LineCnt = LineCnt + 1
                    Else
                      If Not DidOne = True Then
                        ToPrintN$ = UBCustRec(1).Status + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + Using("  #####", AcctNumber)
                        ToPrintN$ = ToPrintN$ + "~" + CustName$ + "~" + SAddr$
                        DidOne = True
                        PrintedOne = True
                      End If
                      ToPrint$ = UBCustRec(1).LocMeters(MtrCnt&).MtrNum
                      ToPrint$ = ToPrint$ + "~" + MeterType$
                      ToPrint$ = ToPrint$ + "~" + Using("#####", Multi#)
                      ToPrint$ = ToPrint$ + "~" + Using("##########", UBCustRec(1).LocMeters(MtrCnt&).CurRead)
                      ToPrint$ = ToPrint$ + "~" + Using("##########", UBCustRec(1).LocMeters(MtrCnt&).PrevRead)
                      ToPrint$ = ToPrint$ + "~" + Using("##########", MeterConsp#)
                      ToPrint$ = ToPrint$ + "~" + Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
                      Print #RptHandle, ToPrintN$ + "~" + ToPrint$ + "~" + Str(AcctNumber)
                    End If
                  End If
              End If
  
            End If
          Next MtrCnt&
          DidOne = False
          If PrintedOne Then
            If Not Grpt Then
              Print #RptHandle, String$(79, "-")
              LineCnt = LineCnt + 1
            End If
            PrintedOne = False
          End If
        End If
      End If
    End If
VacFlagSkip:
    If LineCnt >= MaxLines Then
      If Not Grpt Then
        Print #RptHandle, FF$
        GoSub PrintInactHeading
      End If
    End If
    'ShowPctComp lcnt&, NumOfCust&
  Next
  If Not Grpt Then
    Print #RptHandle, FF$
  End If
  Close

  'Header$ = "Inactive Consumption Report"
  'PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  If Not Grpt Then
    ViewPrint ReportFile$, Title$
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptInactiveConsump
    ARptInactiveConsump.Title = Title$
    ARptInactiveConsump.txtDate = Now
    ARptInactiveConsump.txtTown = TOWNNAME$
    ARptInactiveConsump.GetName ReportFile$
    ARptInactiveConsump.startrpt
  End If
InactConExit:
  Exit Sub


PrintInactHeading:
  If Not Grpt Then
    Page = Page + 1
    Print #RptHandle, Tab(30); "Inactive Consumption Report"
    Print #RptHandle, "Date: "; Date$; Tab(70); "Page #"; Page
    Print #RptHandle, ""
    Print #RptHandle, "Status  Loc.   Acount     Customer Name"; Tab(51); "Service Address"
    Print #RptHandle, " Mtr No.   Mtr Type    Multi     Current   Previous    Consump     Read Date"
    Print #RptHandle, String$(80, "=")
    LineCnt = 7
  End If
Return

GetInactMtrTypePrint:
  PrintMrtFlag = False
  Select Case UBCustRec(1).LocMeters(MtrCnt&).MTRType
  Case "C"
    MeterType$ = "Wat/Sew"
    PrintMrtFlag = True
  Case "W"
    MeterType$ = "Water"
    PrintMrtFlag = True
  Case "S"
    MeterType$ = "Sewer"
    PrintMrtFlag = True
  Case "T"
    MeterType$ = "T-Read"
    PrintMrtFlag = True
  Case "E", "D"
    MeterType$ = "Elec"
    PrintMrtFlag = True
  Case "I"
    MeterType$ = "Irreg"
    PrintMrtFlag = True
  Case "G"
    MeterType$ = "Gas"
    PrintMrtFlag = True
  Case Else
    MeterType$ = "Undef"
    PrintMrtFlag = True
  End Select
Return

End Sub
