VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptMeterSheets 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meter Reading Sheets"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptMeterSheets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   6384
      TabIndex        =   4
      Top             =   4392
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ColDesigner     =   "frmRptMeterSheets.frx":08CA
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
      TabIndex        =   6
      Top             =   7416
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
      TabIndex        =   5
      Top             =   7416
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
            TextSave        =   "10:24 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "2/16/2006"
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
      Left            =   6384
      TabIndex        =   2
      Top             =   3840
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
      Left            =   6384
      TabIndex        =   0
      Top             =   3336
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
   Begin EditLib.fpText fptxtSeq1 
      Height          =   348
      Left            =   7128
      TabIndex        =   1
      Top             =   3336
      Width           =   1236
      _Version        =   196608
      _ExtentX        =   2180
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
      MaxLength       =   6
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
   Begin EditLib.fpText fptxtSeq2 
      Height          =   348
      Left            =   7128
      TabIndex        =   3
      Top             =   3840
      Width           =   1236
      _Version        =   196608
      _ExtentX        =   2180
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
      MaxLength       =   6
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
      BorderWidth     =   2
      X1              =   6960
      X2              =   7176
      Y1              =   4032
      Y2              =   4032
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6984
      X2              =   7140
      Y1              =   3528
      Y2              =   3540
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
      Left            =   3906
      TabIndex        =   11
      Top             =   4416
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2340
      Left            =   3006
      Top             =   2856
      Width           =   6180
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Book:"
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
      Left            =   4302
      TabIndex        =   10
      Top             =   3372
      Width           =   1908
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Book:"
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
      Left            =   4278
      TabIndex        =   9
      Top             =   3900
      Width           =   1932
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1080
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Meter Reading Sheets"
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
      TabIndex        =   8
      Top             =   1320
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   960
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
Attribute VB_Name = "frmRptMeterSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim Grpt As Boolean
Dim WhiteLakeFlag As Integer
Private Sub cmdExit_Click()
  frmUBMeterMenu.Show
  Unload frmRptMeterSheets
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptMeterSheets by " + PWUser$
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
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
      ValidRoutes = True
'      If Chk4BookSeqNum(BegRoute, fptxtSeq1) <> 0 Then
'        If Chk4BookSeqNum(EndRoute, fptxtSeq2) <> 0 Then
'          ValidRoutes = True
'        Else
'          ValidRoutes = False
'        End If
'      Else
'        ValidRoutes = False
'      End If
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
    fptxtSeq1.SetFocus
  End If
End Sub
Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtSeq2.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtSeq1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtSeq1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub

Private Sub fptxtSeq2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtSeq2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
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
  If ValidRoutes Then
    DeActivateControls Me, True
   ' If fpcboRptType.ListIndex = 0 Then
   '   Grpt = True
   
   ' ElseIf fpcboRptType.ListIndex = 1 Then
   '   Grpt = False
      PrintMeterSheets
   ' End If
    ActivateControls Me, True
 ' Else
 '   MsgBox "Invalid Number", vbOKOnly, "Please Retry"
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
  fptxtSeq1 = "000000"
  fptxtSeq2 = "999999"
  'fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpPrintMeterReadingSheets
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PrintMeterSheets()
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
  Dim UseSeq As Boolean, BookNum2 As String, SeqNumb2 As String
  Dim BookNum1 As String, SeqNumb1 As String, Book1 As Long
  Dim Book2 As Long, Sequ1 As Long, Sequ2 As Long, WatRead As Long
  Dim AcctNumber As Long, Sequ As Long, ZONE As String
  Dim CustT As String, zz As Integer, WatSer As String
  Dim EleFlag As Boolean, ECode As String, WatFlag As Boolean
  Dim WatMin As Integer, SewFlag As Boolean, SewMin As Integer
  Dim SecFlag As Boolean, SecCnt As Integer, TrashFlag As Boolean
  Dim TCode As String, EleMin As Integer, EleSer As String
  Dim EleRead As Long, FRCnt As Integer
  Dim RptHand As Integer, RptMsk As String
  Dim TempType As String, Cedarflag As Boolean
  FrmShowPctComp.Label1 = "Creating Inactive Consumption Report"
  FrmShowPctComp.Show , Me

  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  If InStr(UBSetUpRec(1).UTILNAME, "WHITE LAKE") Then
     WhiteLakeFlag = 1
  Else
     WhiteLakeFlag = 0
  End If
  If InStr(UBSetUpRec(1).UTILNAME, "CEDAR BLUFF") Then
     Cedarflag = True
  Else
     Cedarflag = False
  End If

  If UBSetUpRec(1).UseSeq = "Y" Then
    MakeSequenceIndex "Sequence Number", Me
    IdxRecLen = 4               'we are using a long integer
    IdxName$ = UBPath$ + "UBTEMP.IDX"
    IdxFileSize& = FileSize&(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

    UseSeq = True
  Else
    'ShowProcessingScrn "Scanning Accounts"
    IdxRecLen = 4               'we are using a long integer
    IdxName$ = UBPath + "UBCUSTBK.IDX"
    IdxFileSize& = FileSize&(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

    UBCust = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
    Get UBCust, IdxBuff(IdxNumOfRecs).RecNum, UBCustRec(1)
    BookNum2$ = UBCustRec(1).Book
    SeqNumb2$ = UBCustRec(1).SEQNUMB
    For cnt& = 1 To IdxNumOfRecs
      Get UBCust, IdxBuff(cnt&).RecNum, UBCustRec(1)

      If Len(QPTrim$(UBCustRec(1).Book)) > 0 Then
        BookNum1$ = UBCustRec(1).Book
        SeqNumb1$ = UBCustRec(1).SEQNUMB
        Exit For
      End If
      'ShowPctComp cnt&, IdxNumOfRecs
    Next
    Close UBCust
    UseSeq = False
  End If

  ReportFile$ = UBPath$ + "UBMTRSHT.RPT"

  If UseSeq Then
    GoTo SeqJump
  End If


'    Case F10Key
'      'Check for valid Order of Route Questions
'      In1 = True
'      GoSub CheckBookSequence
'      If OkFlag Then
'        In1 = False
'        GoSub CheckBookSequence
'      End If
'      If OkFlag Then
        Book1& = Val(fptxtRoute1)
        Sequ1& = Val(fptxtSeq1)
        Book2& = Val(fptxtRoute2)
        Sequ2& = Val(fptxtSeq2)
'        Done = True
'      End If
'
'    Case F5KEY
'      GoSub DoSheetMask
'
'    Case EscKey
'      GoTo ExitPrintSheets
'
'    End Select

'  Loop Until Done
'  'Free Up Some Memory
'  Erase Form$, Fld, frm

SeqJump:

  ' Location Order ********************************************************

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle     'Open Report File

 ' ShowProcessingScrn "Reading Meter Information"

  For lcnt& = 1 To IdxNumOfRecs
    Get #UBCust, IdxBuff(lcnt&).RecNum, UBCustRec(1)
    AcctNumber& = IdxBuff(lcnt&).RecNum
    If UBCustRec(1).DelFlag <> 0 Then
      GoTo SkipSheet
    End If
    If UseSeq = False Then
      Book = Val(UBCustRec(1).Book)
      Sequ& = Val(UBCustRec(1).SEQNUMB)
      If Book < Book1& Or Book > Book2& Then
        GoTo SkipSheet
      End If
      If Sequ& < Sequ1& Or Sequ& > Sequ2& Then
        GoTo SkipSheet
      End If
      GoSub PrintEm
    Else
      GoSub PrintEm
    End If

SkipSheet:
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If
    FrmShowPctComp.ShowPctComp lcnt&, IdxNumOfRecs
   Next
  Close

  'If AbortFlag Then GoTo ExitPrintSheets
  GoSub DoSheetMask
  Header$ = "Meter Reading Sheets"

 ' PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  ViewPrint ReportFile$, Header$, , , True, RptMsk$
  GoTo ExitPrintSheets

PrintEm:
  GoSub GetMeterFlags
 ' GoSub LookForSecLights
 ' GoSub LookForTrash
  If Cedarflag Then
   GoSub PrintSkipHeader
  End If
'
'  ZONE$ = QPTrim$(UBCustRec(1).ZONE)
'  Select Case Left$(QPTrim$(UBCustRec(1).CUSTTYPE), 1)
'  Case "B"
'    CustT$ = "Commerical"
'  Case "R"
'    CustT$ = "Residential"
'  Case Else
'    CustT$ = "??????????"
'  End Select
'
'
  If WhiteLakeFlag = 1 Then
   Print #RptHandle, " "
   Print #RptHandle, " "
   Print #RptHandle, Tab(15); "Acct #"; AcctNumber&
   Print #RptHandle, " "
   Print #RptHandle, Tab(15); Tab(15); UBCustRec(1).CustName
   Print #RptHandle, " "
   Print #RptHandle, " "
   Print #RptHandle, Tab(15); UBCustRec(1).ADDR1
   Print #RptHandle, Tab(15); UBCustRec(1).ADDR2
   Print #RptHandle, Tab(15); QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
   Print #RptHandle, " "
   Print #RptHandle, " "
   '''Print #RptHandle, " "
   '''Print #RptHandle, " "
   Print #RptHandle, Tab(15); "SERVICE AT :"; UBCustRec(1).ServAddr
   Print #RptHandle, " "
   Print #RptHandle, Tab(14); Right$(Date$, 2); Tab(43); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
   For zz = 16 To 43
   Print #RptHandle, " "
   Next zz
   Print #RptHandle, Tab(15); WatRead&
   Print #RptHandle, " "
   Print #RptHandle, " "
   Print #RptHandle, " "
   Print #RptHandle, " "
   Print #RptHandle, " "
   Print #RptHandle, Tab(15); WatSer$
   For zz = 51 To 59
    Print #RptHandle, " "
   Next zz
    Print #RptHandle, " "
    Return
 ElseIf Cedarflag Then
    If EleFlag Then
      Print #RptHandle, " Electric"; "  "; ECode$
    Else
      Print #RptHandle, " "
    End If
    If WatFlag Then
      Print #RptHandle, " Water"; "  "; ZONE$; "  Min ="; WatMin
    Else
      Print #RptHandle, " "
    End If
    If SewFlag Then
      Print #RptHandle, " Sewer"; "  "; ZONE$; "  Min ="; SewMin
    Else
      Print #RptHandle, " "
    End If
  
    If SecFlag Then
      Print #RptHandle, " Security Lights   #"; SecCnt
    Else
      Print #RptHandle, " "
    End If
    If TrashFlag Then
      Print #RptHandle, " Trash   "; TCode$
    Else
      Print #RptHandle, " "
    End If
  
    Print #RptHandle, Tab(15); EleMin
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, EleSer$; Tab(24); WatSer$
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, Tab(6); Using("#########", EleRead&); Tab(24); Using("#########", WatRead&)
    For zz = 21 To 43
      Print #RptHandle, " "
    Next
    Print #RptHandle, UBCustRec(1).CustName
    Print #RptHandle, UBCustRec(1).ADDR1
    Print #RptHandle, UBCustRec(1).ServAddr
    Print #RptHandle, QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
    Print #RptHandle, " "
    Print #RptHandle, Tab(11); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #RptHandle, "~"
Else
'This added for Deep Run 1/12/05
  For MtrCnt = 1 To 7
    TempType$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
'
    If Len(TempType$) <> 0 Then
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      If Len(WatSer$) = 0 Then
        WatSer$ = "?????"
      End If
      Print #RptHandle, "~"
      For zz = 2 To 55
       Print #RptHandle, " "
      Next zz
      Print #RptHandle, Tab(10); WatRead&
      Print #RptHandle, " "
      Print #RptHandle, " "
      Print #RptHandle, " "
      Print #RptHandle, Tab(10); QPTrim$(UBCustRec(1).CustName);
      Print #RptHandle, Tab(52); QPTrim$(UBCustRec(1).ServAddr)
      Print #RptHandle, Tab(52); QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE
      Print #RptHandle,
      Print #RptHandle, Tab(52); "Location# "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
      Print #RptHandle, Tab(59); WatSer$
      Print #RptHandle,
      Print #RptHandle, "~"
    End If
  Next
 End If
Return

LookForSecLights:
  SecFlag = False
  For FRCnt = 1 To 4
    If InStr(UBCustRec(1).FlatRates(FRCnt).FRDESC, "SECUR") Then
      SecFlag = True
      SecCnt = UBCustRec(1).FlatRates(FRCnt).NumMin
      Exit For
    End If
  Next
  Return

LookForTrash:
  TrashFlag = False
  If Len(QPTrim$(UBCustRec(1).serv(9).Ratecode)) > 0 Then
    TrashFlag = True
    TCode$ = UBCustRec(1).serv(9).Ratecode
  End If
  Return

PrintSkipHeader:
  Print #RptHandle, "~"
  For zz = 1 To 8
    Print #RptHandle,
  Next
  Return

GetMeterFlags:
  WatFlag = False: WatMin = 0: WatSer$ = "": WatRead& = 0
  SewFlag = False: SewMin = 0:
  EleFlag = False: EleMin = 0: EleSer$ = "": EleRead& = 0

  For MtrCnt = 1 To 7
    Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
    Case "C"
      WatFlag = True
      SewFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      If Len(WatSer$) = 0 Then
        WatSer$ = "?????"
      End If
    Case "W"
      WatFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      If Len(WatSer$) = 0 Then
        WatSer$ = "?????"
      End If
    Case "S"
      SewFlag = True
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      If Len(WatSer$) = 0 Then
        WatSer$ = "?????"
      End If
    Case "E"
      EleFlag = True
      EleMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      EleSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      EleRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      ECode$ = UBCustRec(1).serv(3).Ratecode
      If Len(EleSer$) = 0 Then
        EleSer$ = "?????"
      End If
    End Select
  Next
  If WatRead& < 0 Then
    WatRead& = 0
  End If
  If EleRead& < 0 Then
    EleRead& = 0
  End If
  Return
DoSheetMask:
  RptHand = FreeFile
  RptMsk$ = "UBMtrS.msk"
  Open RptMsk$ For Output As #RptHand     'Open Report File
  If WhiteLakeFlag = 1 Then
      Print #RptHand, "TOP"
      Print #RptHand,
      Print #RptHand, Tab(15); "Acct # XXXXXX"
      Print #RptHand,
      Print #RptHand, Tab(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #RptHand,
      Print #RptHand,
      Print #RptHand, Tab(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #RptHand, Tab(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #RptHand, Tab(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #RptHand, Tab(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #RptHand,
      Print #RptHand,
      Print #RptHand,
      Print #RptHand, Tab(14); Right$(Date$, 2); Tab(43); "XX-XXXXXXX"
      For zz = 16 To 43
      Print #RptHand,
      Next zz
      Print #RptHand, Tab(15); "XXXXXXXXXX"
      Print #RptHand,
      Print #RptHand,
      Print #RptHand,
      Print #RptHand,
      Print #RptHand,
      Print #RptHand, Tab(15); "XXXXXXXXXXX"
      For zz = 51 To 59
      Print #RptHand,
      Next zz
      Print #RptHand, "BOTTOM"

  ElseIf Cedarflag Then
      Print #RptHand, "TOP"
      For zz = 1 To 8
        Print #RptHand, ""
      Next
      Print #RptHand, " Electric  XXXXXXXXXX"
      Print #RptHand, " Water  X  Min = X"
      Print #RptHand, " Sewer  X  Min = X"
      Print #RptHand, ""
      Print #RptHand, ""
      Print #RptHand, "               X"
      Print #RptHand, ""
      Print #RptHand, ""
      Print #RptHand, "XXXXXXXXX              XXXXXXXXX"
      Print #RptHand, ""
      Print #RptHand, ""
      Print #RptHand, "       XXXXXXX           XXXXXXX"
      For zz = 1 To 23
        Print #RptHand, ""
      Next
      Print #RptHand, "XXXXXXX XXXXXXXXXXX"
      Print #RptHand, "XX XXX XXX"
      Print #RptHand, "XXXXXXXXXXX"
      Print #RptHand, "XXXXXXXXX XX XXXXX"
      Print #RptHand, ""
      Print #RptHand, "          XX-XXXXXX"
      Print #RptHand, "BOTTOM"

  Else
      Print #RptHand, "~"
      For zz = 2 To 55
       Print #RptHand, " "
      Next zz
      Print #RptHand, Tab(10); 1234567
      Print #RptHand, " "
      Print #RptHand, " "
      Print #RptHand, " "
      Print #RptHand, Tab(10); "Bobby Muffin Jones";
      Print #RptHand, Tab(52); "110 Drewery Lane"
      Print #RptHand, Tab(52); "Emerald City"; " "; "NL"
      Print #RptHand,
      Print #RptHand, Tab(52); "Location# "; "10-122311"
      Print #RptHand, Tab(59); "329322"
      Print #RptHand,
      Print #RptHand, "~"

   End If
  Close RptHand

  'Call CursorOff
  'PrintRptFile Header$, ReportFile$, 1, RetCode%, 4
  'RestScrn TempScrn()
  'Action = 1
  'ViewPrint Reportfile$, "Alignment Mask"
  
  Return


ExitPrintSheets:
End Sub
Public Function Chk4BookSeqNum(Book$, SeqNum$)
  Dim TBookSeq As Long, BookSeqLen As Integer, Handle As Integer
  Dim NumBookSeq As Integer, cnt As Integer, NCnt As Integer
  Chk4BookSeqNum = False        'assume not found

  TBookSeq& = Val(Book$ + SeqNum$)

  ReDim UBBookSeq(1) As BookSeqRecType
  BookSeqLen = Len(UBBookSeq(1))
  Handle = FreeFile
  If FileSize(UBPath$ + "UBOOKSEQ.DAT") > 0 Then
    Open UBPath$ + "UBOOKSEQ.DAT" For Random Shared As Handle Len = BookSeqLen            'open data file
    NumBookSeq = LOF(Handle) \ BookSeqLen
    ReDim UBBookSeq(1 To NumBookSeq) As BookSeqRecType
    For NCnt = 1 To NumBookSeq
      Get Handle, NCnt, UBBookSeq(NCnt)       ', 1&, NumBookSeq * BookSeqLen
    Next
    Close Handle

    For cnt = 1 To NumBookSeq
      If UBBookSeq(cnt).BookSeq = TBookSeq& Then
        Chk4BookSeqNum = True   'found this book-seq
        Exit For
      End If
    Next
  End If

End Function



