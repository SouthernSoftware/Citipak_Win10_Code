VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPrintApplications 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Applications"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmPrintApplications.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   4845
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
      ColDesigner     =   "frmPrintApplications.frx":08CA
   End
   Begin LpLib.fpCombo fpcboCategory 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   4335
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
      ColDesigner     =   "frmPrintApplications.frx":0C99
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
      Top             =   7776
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
      Left            =   10080
      TabIndex        =   5
      Top             =   7776
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
            TextSave        =   "12:33 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7/25/2006"
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
   Begin EditLib.fpDateTime txtAppDate 
      Height          =   348
      Left            =   5040
      TabIndex        =   0
      Top             =   3288
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpText fptxtmessage 
      Height          =   372
      Left            =   5040
      TabIndex        =   1
      Top             =   3804
      Width           =   4164
      _Version        =   196608
      _ExtentX        =   7345
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      CharValidationText=   ""
      MaxLength       =   40
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Category:"
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
      Left            =   2568
      TabIndex        =   11
      Top             =   4356
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Printing Information"
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
      Left            =   3252
      TabIndex        =   10
      Top             =   1320
      Width           =   5724
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00FFFFFF&
      Height          =   852
      Left            =   3228
      Top             =   1080
      Width           =   5772
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Application Date:"
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
      Left            =   2544
      TabIndex        =   9
      Top             =   3360
      Width           =   2364
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
      Index           =   3
      Left            =   2592
      TabIndex        =   8
      Top             =   4848
      Width           =   2316
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
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
      Left            =   3192
      TabIndex        =   7
      Top             =   3852
      Width           =   1716
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2676
      Left            =   2496
      Top             =   2904
      Width           =   7236
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   972
      Left            =   3228
      Top             =   960
      Width           =   5772
   End
End
Attribute VB_Name = "frmPrintApplications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CycleFlag As Boolean, OKFlag As Boolean, BadDate As Boolean
Dim ErFlag As Boolean, APType As Integer, LPIFlag As Boolean
Dim OkiMode As Integer
Dim Rteflag As Boolean, AcctBar As Boolean
Private Sub cmdExit_Click()
  Load frmDCCustomerMenu
  DoEvents
  frmDCCustomerMenu.Show
  Unload Me
  DoEvents
End Sub

Private Sub cmdPrint_Click()
  If CheckFields Then
    PrintApplications
  End If
'  CheckFields
'  LPIFlag = False
'  If OKFlag = True Then
''do print stuff here
''depending on which late notice they have selected in setup
'  If LNType = 1 Then
'    frmReportOpt.Show 1
'    DeActivateControls Me
'    If rptopt = 1 Then
'    'do the graphics
'     PrintLateNotices True
'    ElseIf rptopt = 2 Then
'    'do the text
'      PrintLateNotices False
'      ActivateControls Me
'    End If
'  ElseIf LNType = 8 Or LNType = 9 Then
'    PrintLateNotices True
'  Else
'    DeActivateControls Me
'    PrintLateNotices False
'    ActivateControls Me
'  End If
'  End If
End Sub


Private Sub txtAppDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtmessage.SetFocus
  End If
End Sub
Private Sub fptxtmessage_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboCategory.SetFocus
  End If
End Sub
Private Sub fpcboCategory_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCategory.ListDown = True
  End If
  If fpcboCategory.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtmessage.SetFocus
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
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboCategory.SetFocus
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via AppPrinting by " + PWUser$
        CitiTerminate
      End If
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
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdPrint_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim DCSetuplen As Integer, cnt As Integer, DCBillSetuplen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  ReDim DCSetup(1) As DCSetupType
  DCSetuplen = Len(DCSetup(1))
  LoadDCSetUpFile DCSetup(), DCSetuplen
  Me.HelpContextID = hlpPrint
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "ZipCode Order"
  txtAppDate.Text = Format(Now, "mm/dd/yyyy")
  'get late notice type from setup and store integer
  'at same time get the OkiMode 1 is not ibm, 2 is ibm
  FillCatCMBOAll fpcboCategory
  APType = DCSetup(1).AppType
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtAppDate) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    ValidDate = True
  End If
End Function

Private Function CheckFields()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  MsgText(2) = ""
  OKFlag = False
  If ValidDate = False Then
    MsgText(3) = "You Have Not Entered a Valid"
    MsgText(4) = "Date!"
  ElseIf fpcboPrintOrder.ListIndex = -1 Then
    MsgText(3) = "Invalid Printing Order."
    MsgText(4) = "Correct and try again."
  ElseIf fpcboCategory.ListIndex = -1 Then
    MsgText(3) = "Must identify Category."
    MsgText(4) = "Correct and try again."
  Else
    OKFlag = True
  End If
  If Not OKFlag Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
  End If
  CheckFields = OKFlag
End Function

Private Sub PrintApplications()
  Dim PDate As String, NDate As String, NMonth As String
  Dim LongPDate As String, LongNDate As String, DCSetuplen As Integer
  Dim PSAFlag As Boolean, UseCycle As Boolean, FromBC As Integer
  Dim ThruBC As Integer, MinBalance As Double, IndexName As String
  Dim NoIndex As Boolean, DCCustRecLen As Integer, TBooks As Integer
  Dim NumOfRecs As Long, IdxNumOfRecs As Long, Handle As Integer
  Dim cnt As Long, TrHandle As Integer, CarRecord As Long, ExpireDate As String
  Dim Next2Print As Integer, AcctNo As Long, GotWater As Boolean
  Dim CustBC As Integer, Location As String, Acct As String, Zip As String
  Dim ZipLen As Integer, TotalBal As Double, CustBal As Double
  Dim Print1 As Integer, PrnCnt As Integer, NIfile As Integer, DoLogo As Boolean
  Dim fmt1 As String, fmt2 As String, ReportFile As String, PCnt As Long
  Dim LaLe As Integer, lenlate As Integer, cntll As Integer, Ext As String
  Dim DeDate As String, lenNI As Integer, AcctNum As Long, Totalamt As Double
  Dim AcctLen As Integer, Previous As Double, Current As Double, IdxRecLen As Integer
  Dim WRevCnt As Integer, PZip As String, ZDigit As String, ToPrint As String
  Dim CustMsg As String, MPCnt As Integer, tmprev As Double, LNcnt As Integer
  Dim ToPrint2 As String, endit As Boolean, DCRpt As Integer, MaskNotice As String
  Dim Fmt10 As String, Fmt10a As String, Fmt15 As String, Today As String
  Dim msg As String, DCVehReclen As Integer, DCvFile As Integer, ControlNumber As Long
  Dim ExpireDateV As Integer, StateTag As String, CatSel As String, Catdo As Boolean
  FrmShowPctComp.Label1 = "Creating Applications"
  FrmShowPctComp.Show , Me
  endit = False
  'if lntype = 1 then
  If Exist(DCPath$ + "DCApplet.dat") Then
    ReDim APPLet(1) As ApplicationDefaultsType
    LaLe = FreeFile
    lenlate = Len(APPLet(1))
    Open DCPath$ + "DCApplet.dat" For Random Shared As LaLe Len = lenlate
    Get LaLe, 1, APPLet(1)
    If APPLet(1).DoLogo = 1 Then DoLogo = True
    Close
  End If
  If fpcboCategory.ListIndex <> 0 Then
    fpcboCategory.col = 1
    CatSel = QPTrim(fpcboCategory.ColText)
    Catdo = True
  Else
    CatSel = "All"
    Catdo = False
  End If
  MaxLines = 53
  
  fmt1$ = String$(80, "-")
  fmt2$ = "$###,###,###.##"
  PDate$ = txtAppDate
  msg$ = QPTrim$(fptxtmessage)
  LNcnt = 0
  'NMonth$ = Left$(MakeMonth$(NDate$), 3) + "."
  ToPrint$ = ""
  LongPDate$ = FormatDateTime(PDate$, vbLongDate)

  PageNo = 0
  MaxLines = 54
  Linecnt = 0
  Select Case fpcboPrintOrder.ListIndex
  Case 0
    IndexName$ = DCPath$ + "DCCUST.IDX"
  Case 1
    NoIndex = True
  Case 2
    IndexName$ = DCPath$ + "DCTEMP.IDX"
    MakeZipCodeIndex "ZipCode"
  End Select
  ReDim DCCustRec(1) As DCCustRecType     ' open customer file
  DCCustRecLen = Len(DCCustRec(1))
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))

  If NoIndex = False Then
    NumOfRecs = FileSize(IndexName$) \ 4
    ReDim IndexArray(1 To NumOfRecs) As DCTempIDXRecType
    'FGetAH IndexName$, IndexArray(1), , NumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = 4
    For cnt& = 1 To NumOfRecs
      Get #Handle, cnt&, IndexArray(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(DCPath$ + "DCCUST.DAT") \ DCCustRecLen
  End If
'''
  TrHandle = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = DCCustRecLen
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  
  ReportFile$ = "DCAPPS.PRN"    'Report File Name
  DCRpt = FreeFile
  Open ReportFile$ For Output As #DCRpt
 
  FF$ = Chr$(12)

  Next2Print = 1

  For cnt = 1 To NumOfRecs
  'If cnt = NumOfRecs Then Stop
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitPrint
    End If

    If NoIndex Then
      AcctNo& = cnt
    Else
      AcctNo& = IndexArray(cnt).IDXRECORD
    End If
    Get TrHandle, AcctNo&, DCCustRec(1)
    If DCCustRec(1).Deleted <> "Y" And Len(QPTrim$(DCCustRec(1).BILLNAME)) > 0 Then

      CarRecord = DCCustRec(1).FirstCar
      While CarRecord > 0
        Get DCvFile, CarRecord, DCVRec(1)
        ExpireDate$ = Num2Date$(DCVRec(1).ExpireDate)
        ExpireDate$ = Right$(ExpireDate$, 4)
        ExpireDateV = Val(ExpireDate$)
        StateTag$ = QPTrim$(DCVRec(1).StateTag)
        If DCVRec(1).Active <> "N" Then
          If Catdo Then
            If CatSel$ <> QPTrim$(DCVRec(1).DecalCat) Then GoTo skipforcat
          End If
          PCnt& = PCnt& + 1
          GoSub PrintThemOne
          ControlNumber = ControlNumber + 1
        End If
        If CarRecord <> DCVRec(1).NextRec Then
          CarRecord = DCVRec(1).NextRec
        ElseIf CarRecord = DCVRec(1).NextRec Then
          'DCCustRec(1).BILLNAME
          'BadCar = BadCar + 1
          CarRecord = 0
        End If
      Wend
    End If
   'IF PCnt& > 1 THEN EXIT FOR
skipforcat:
  Next cnt
 

  Close
  
  NIfile = FreeFile
  ReDim NoticeInfo(1) As NoticeInfoType
  lenNI = Len(NoticeInfo(1))
  Open DCPath$ + "DCAPINFO.DAT" For Random Shared As NIfile Len = lenNI
  
  NoticeInfo(1).NoticeDate = Date2Num(txtAppDate)
  NoticeInfo(1).msgline = QPTrim$(fptxtmessage)
  NoticeInfo(1).PrnCategory = fpcboCategory.ListIndex + 1
  NoticeInfo(1).PrnOrder = fpcboPrintOrder.ListIndex + 1
'  If DoLogo = True Then
'    NoticeInfo(1).DoLogo = 1
'  Else
'    NoticeInfo(1).DoLogo = 0
'  End If
  NoticeInfo(1).PrnCnt = PrnCnt

  Put NIfile, 1, NoticeInfo(1)
  Close
  If APType = 1 Then
    Load frmLoadingRpt
    If DoLogo = True Then
      If Exist(DCPath$ + "DCTNlogo.bmp") Then
        ARptLetterApplication.Image1.Picture = LoadPicture(DCPath$ + "DCTNlogo.bmp")
        ARptLetterApplication.Image1.Visible = True
      End If
    End If
    ARptLetterApplication.Head1 = QPTrim(APPLet(1).Head1)
    ARptLetterApplication.Head2 = QPTrim(APPLet(1).Head2)
    ARptLetterApplication.Head3 = QPTrim(APPLet(1).Head3)
    ARptLetterApplication.Head4 = QPTrim(APPLet(1).Head4)
    ARptLetterApplication.Head5 = QPTrim(APPLet(1).Head5)
    ARptLetterApplication.Pgf1 = QPTrim(APPLet(1).Body(1))
    ARptLetterApplication.Pgf2 = QPTrim(APPLet(1).Body(2))
    ARptLetterApplication.Pgf3 = QPTrim(APPLet(1).Body(3))
    ARptLetterApplication.Pgf4 = QPTrim(APPLet(1).Body(4))
    ARptLetterApplication.Pgf5 = QPTrim(APPLet(1).Body(5))
    ARptLetterApplication.Pgf6 = QPTrim(APPLet(1).Body(6))
    ARptLetterApplication.Pgf7 = QPTrim(APPLet(1).Body(7))
    ARptLetterApplication.Pgf8 = QPTrim(APPLet(1).Body(8))
    ARptLetterApplication.Pgf9 = QPTrim(APPLet(1).Body(9))
    ARptLetterApplication.Pgf10 = QPTrim(APPLet(1).Body(10))
    ARptLetterApplication.Pgf11 = QPTrim(APPLet(1).Body(11))
    ARptLetterApplication.Pgf12 = QPTrim(APPLet(1).Body(12))
    ARptLetterApplication.Pgf13 = QPTrim(APPLet(1).Body(13))
    ARptLetterApplication.Pgf14 = QPTrim(APPLet(1).Body(14))
    ARptLetterApplication.Pgf15 = QPTrim(APPLet(1).Body(15))
    ARptLetterApplication.Pgf16 = QPTrim(APPLet(1).Body(16))
    ARptLetterApplication.Pgf17 = QPTrim(APPLet(1).Body(17))
    ARptLetterApplication.Pgf18 = QPTrim(APPLet(1).Body(18))
    ARptLetterApplication.Pgf19 = QPTrim(APPLet(1).Body(19))
    ARptLetterApplication.Pgf20 = QPTrim(APPLet(1).Body(20))
    frmLoadingRpt.setwherefrom frmApplicationLetter
    ARptLetterApplication.lblDate = PDate$
    ARptLetterApplication.lblMessage = msg$
    ARptLetterApplication.GetName ReportFile$
    ARptLetterApplication.startrpt
  Else
    GoSub DoLNMask

    ViewPrint ReportFile$, "Application Letter"
  End If
  
  GoTo ExitPrint

PrintThemOne:
   Select Case APType
    Case 1:  'Letter Format
      GoSub PrintLetterFormatG
    Case 2:
      GoSub PrintLetterFormatT
    Case Else
   End Select
   
   PrnCnt = PrnCnt + 1
Return

PrintLetterFormatT:   '1 'this is text
  Print #DCRpt, " "
  Print #DCRpt, " "
  Print #DCRpt, " "
  Print #DCRpt, " "
  Print #DCRpt, " "
  Print #DCRpt, " "
  Print #DCRpt, Tab(30); APPLet(1).Head1
  Print #DCRpt, Tab(30); APPLet(1).Head2
  Print #DCRpt, Tab(30); APPLet(1).Head3
  Print #DCRpt, Tab(30); APPLet(1).Head4
  Print #DCRpt, Tab(30); APPLet(1).Head5
  Print #DCRpt, " "
  Print #DCRpt, Tab(5); "Date: " + PDate$; Tab(42); msg$
  Print #DCRpt, Tab(60); Str$(ControlNumber); ""
  Print #DCRpt, " "
  Print #DCRpt, Tab(5); QPTrim(DCCustRec(1).BILLNAME); Tab(52); "Cust Acct#: "; QPTrim$(DCCustRec(1).CUSTNUMB)
  Print #DCRpt, Tab(5); QPTrim(DCCustRec(1).ADDRESS1)
  Print #DCRpt, Tab(5); QPTrim(DCCustRec(1).ADDRESS2)
  Print #DCRpt, Tab(5); QPTrim$(DCCustRec(1).City); ", "; DCCustRec(1).State; " "; DCCustRec(1).ZIPCODE
  Print #DCRpt, " "
  Print #DCRpt, " "
  For cntll = 1 To 10
    Print #DCRpt, Tab(5); RTrim(APPLet(1).Body(cntll))
  Next
  Print #DCRpt, " "
  Print #DCRpt, Tab(10); "        Vehicle Notes: "; QPTrim$(DCVRec(1).Notes)
  Print #DCRpt, Tab(10); " Vehicle Make & Model: "; QPTrim$(DCVRec(1).makemodel)
  Print #DCRpt, Tab(10); "       State License#: "; QPTrim$(DCVRec(1).StateTag)
  Print #DCRpt, Tab(10); "                 VIN#: "; QPTrim$(DCVRec(1).Desc)
  Print #DCRpt, Tab(10); "                  Fee: "; Using$("####.##", DCVRec(1).Fee)
  Print #DCRpt, " "
  For cntll = 11 To 20
    Print #DCRpt, Tab(5); RTrim(APPLet(1).Body(cntll))
  Next
  For cntll = 40 To MaxLines
    Print #DCRpt, " "
  Next
  Print #DCRpt, FF$

Return
PrintLetterFormatG:
    ToPrint$ = Str$(AcctNo&)
    ToPrint$ = ToPrint$ + "~" + QPTrim(DCCustRec(1).BILLNAME)
    ToPrint$ = ToPrint$ + "~" + QPTrim(DCCustRec(1).ADDRESS1)
    ToPrint$ = ToPrint$ + "~" + QPTrim(DCCustRec(1).ADDRESS2)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).City) + ", " + DCCustRec(1).State + " " + DCCustRec(1).ZIPCODE
    ToPrint$ = ToPrint$ + "~" + Str$(CarRecord)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCVRec(1).makemodel)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCVRec(1).StateTag)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCVRec(1).Desc)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCVRec(1).Notes)
    ToPrint$ = ToPrint$ + "~" + Using$("####.##", DCVRec(1).Fee)
    Print #DCRpt, ToPrint$
Return
DoLNMask:
'  UBRptA = FreeFile
'  MaskNotice$ = UBPath$ + "UBLNA.RPT"
'  Open MaskNotice$ For Output As UBRptA
'  Select Case LNType
'  Case 2
'    GoSub PrintNewStandV1Mask
'  Case 3
'    GoSub PrintNewStandBarMask
'  Case 4
'    GoSub PrnStand21LineMask
'  Case 5
'    GoSub PrintNewStandRmStampMask
'  Case 6
'    GoSub PrnStand24L2BxMask
'  Case 7
'    GoSub PrnStand24L3BxMask
'  Case Else
'    'NO MASK
'  End Select
'  Close UBRptA
Return
ExitPrint:

End Sub
'Private Sub PrintApp()
'
''  yr! = Val(Right$(Date$, 2))
''  ControlNumber! = yr! * 10000
''  '  PrintDate$ = RIGHT$(DATE$, 4)
''  '  PrintDate! = VAL(PrintDate$)
''  '  PrintDate! = PrintDate!
'
'  MaxLines = 54
'  Linecnt = 0
'  TotDr# = 0
'  TotCr# = 0
'  ReDim TArray(1 To Size) As Struct
'  'REDIM DCCustRec(1) AS DCCustRecType     ' open customer file
'  CustRecLen = Len(DCCustRec(1))
'  TrHandle = FreeFile
'  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = CustRecLen
'  TrNumRecs = LOF(TrHandle) \ CustRecLen
'  DCVehReclen = Len(DCVRec(1))
'  DCvFile = FreeFile
'  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
'
'  'REDIM DCCustIdxRec(1) AS DCCustIdxRecType     ' open customer file
'  IdxCustRecLen = Len(DCCustIdxRec(1))
'  IdxTrHandle = FreeFile
'  Open "DCCUST.IDX" For Random Access Read Write Shared As IdxTrHandle Len = IdxCustRecLen
'  IdxTrNumRecs = LOF(IdxTrHandle) \ IdxCustRecLen
'
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
'
'  For cnt = 1 To IdxTrNumRecs
'    Get IdxTrHandle, cnt, DCCustIdxRec(1)
'    Get TrHandle, DCCustIdxRec(1).IDXRECORD, DCCustRec(1)
'    If DCCustRec(1).Deleted <> "Y" And Len(QPTrim$(DCCustRec(1).BILLNAME)) > 0 Then
'      CarRecord! = DCCustRec(1).FirstCar
'      While CarRecord! > 0
'        Get DCvFile, CarRecord!, DCVRec(1)
'        ExpireDate$ = Num2Date$(DCVRec(1).ExpireDate)
'        ExpireDate$ = Right$(ExpireDate$, 4)
'        ExpireDate! = Val(ExpireDate$)
'        StateTag$ = QPTrim$(DCVRec(1).StateTag)
'        If DCVRec(1).Active <> "N" Then
'          If PrintDate = ExpireDate Then
'            PCnt& = PCnt& + 1
'            ControlNumber! = ControlNumber! + 1
'
'
'          End If
'        End If
'        If CarRecord! <> DCVRec(1).NextRec Then
'          CarRecord! = DCVRec(1).NextRec
'        ElseIf CarRecord! = DCVRec(1).NextRec Then
'          'DCCustRec(1).BILLNAME
'          BadCar = BadCar + 1
'          CarRecord! = 0
'        End If
'      Wend
'    End If
'   'IF PCnt& > 1 THEN EXIT FOR
'
'  Next cnt
'  GoSub PrintRpt1Ending
'  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
'  Close         'Close all open files now
'
'  ViewPrint ReportFile$, Header$
'  'PrintRptFile , LptPort%, RetCode%, EntryPoint
'
'  'KILL ReportFile$
'
'DoneInHere:
'  Exit Sub
'PrintRpt1Header:
'  Return
'
'PrintRpt1Ending:
'  Return
'
'
'PrintAlignMask:
'
'LPRINT "TOP LINE                           XX-XX-XXXX"
'LPRINT ""
'LPRINT "                                                                XX-XX-XX"
'LPRINT "                                                                XX-XX-XX"
'LPRINT ""
'LPRINT ""
'LPRINT " XX XXXXX XXXXX      VIN#XXXXXXXXXXXXXXXXX"
'LPRINT ""
'LPRINT ""
'LPRINT "                                                                    15.00"
'LPRINT "                                                                     8.00"
'LPRINT ""
'LPRINT "                                                                     1.00"
'LPRINT ""
'LPRINT "                                           XX-XX"
'LPRINT "       XXXXXX XXXXXXXXXX XXXXXXX"
'LPRINT "       XXX XXXXXXX XX"
'LPRINT ""
'LPRINT "       XXXXXXXXXX XX  XXXXX"
'LPRINT ""
'LPRINT ""
'LPRINT "BOTTOM LINE"
'
'  Return
'
'
'
'End Sub
