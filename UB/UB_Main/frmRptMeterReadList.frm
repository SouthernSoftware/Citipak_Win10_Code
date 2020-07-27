VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptMeterReadList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meter Reading List"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptMeterReadList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   6396
      TabIndex        =   3
      Top             =   4944
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ColDesigner     =   "frmRptMeterReadList.frx":08CA
   End
   Begin VB.CheckBox ChkNoPrintCurr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Exclude Current Reading:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3576
      TabIndex        =   4
      Top             =   5448
      Width           =   3012
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
      Top             =   7368
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
      TabIndex        =   6
      Top             =   7368
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
            TextSave        =   "9:30 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/11/2018"
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   6396
      TabIndex        =   1
      Top             =   3804
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   6396
      TabIndex        =   0
      Top             =   3288
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
   Begin EditLib.fpText fptxtCustType 
      Height          =   348
      Left            =   6396
      TabIndex        =   2
      Top             =   4368
      Width           =   1188
      _Version        =   196608
      _ExtentX        =   2096
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
      AutoCase        =   1
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
      CharValidationText=   ""
      MaxLength       =   3
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
      Caption         =   "Customer Type:"
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
      Left            =   4152
      TabIndex        =   12
      Top             =   4428
      Width           =   2076
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
      Left            =   3900
      TabIndex        =   11
      Top             =   4968
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3300
      Left            =   2760
      Top             =   2784
      Width           =   6684
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Route:"
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
      Left            =   4752
      TabIndex        =   10
      Top             =   3348
      Width           =   1476
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Thru Route:"
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
      Left            =   4848
      TabIndex        =   9
      Top             =   3852
      Width           =   1380
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1032
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Meter Reading List"
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
      Top             =   1272
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   912
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
Attribute VB_Name = "frmRptMeterReadList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim Grpt As Boolean
Dim LorR As Integer ' 1 for list , 2 for report
Public Sub GetLorR(x As Integer) 'send appropriate value from menu
  LorR = x
End Sub
Private Sub cmdExit_Click()
  frmUBMeterMenu.Show
  Unload frmRptMeterReadList
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptMeterReadList by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fptxtCustType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
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
        fptxtCustType.SetFocus
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
    fptxtCustType.SetFocus
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
      If LorR = 1 Then
        PrintMeterList
      Else
        PrintMeterReport
      End If
    ElseIf fpcboRptType.ListIndex = 1 Then
      Grpt = False
      If LorR = 1 Then
        PrintMeterList
      Else
        PrintMeterReport
      End If
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
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  If LorR <> 1 Then
    ChkNoPrintCurr.Visible = False
    ChkNoPrintCurr.Enabled = False
  End If
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PrintMeterList()
  Dim UBCustRecLen As Integer, ReportFile As String, RptHandle As Integer
  Dim UBSetupLen As Integer, SeqFlag As Boolean, IdxName As String
  Dim NumOfRecs As Long, IdxNumOfRecs As Long, Handle As Integer
  Dim cnt As Long, lcnt As Long, Prec As Long, process As Boolean
  Dim Header As String, MtrCnt As Integer, ValidCustomer As Boolean
  Dim TempRev As String, MeterStatus As String, MeterType As String
  Dim Page As Integer, RecNo As Long, L2Handle As Integer, UseType As Boolean
  Dim FirstCust As Boolean, WhatBook As Integer, DoHeaderFlag As Boolean
  Dim ToPrint As String, ToPrintN As String, CUSTTYPE As String, ThisType As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  'Open Report File
  ReportFile$ = UBPath$ + "WBMTRLST.RPT"
  FrmShowPctComp.Label1 = "Creating Meter Reading Listing Report"
  'FrmShowPctComp.Show , Me
  CUSTTYPE$ = QPTrim$(fptxtCustType)
  If Len(CUSTTYPE$) > 0 Then
    UseType = True
  Else
    UseType = False
  End If

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  If UBSetUpRec(1).UseSeq = "Y" Then
    SeqFlag = True
    MakeSequenceIndex "Sequence Number", Me
    IdxName$ = UBPath$ + "UBTEMP.IDX"
    FrmShowPctComp.Label1 = "Creating Meter Reading Listing Report"
    FrmShowPctComp.Show , Me
  Else
    IdxName$ = UBPath$ + "UBCUSTBK.IDX"
  End If


  NumOfRecs& = FileSize&(IdxName$) \ 4
  IdxNumOfRecs = NumOfRecs
  ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
  'FGetAH IdxName$, IndexArray(1), 4, NumOfRecs
  Handle = FreeFile
  Open IdxName$ For Random Shared As Handle Len = 4
  For cnt& = 1 To IdxNumOfRecs
    Get #Handle, cnt&, IndexArray(cnt&)
  Next
  Close Handle
  UBCustRecLen = Len(UBCustRec(1))
  L2Handle = FreeFile
  Open UBPath$ + "UBCust.DAT" For Random Shared As L2Handle Len = UBCustRecLen

  Do
    If lcnt < 1 Then lcnt = 1     ' Do Not Allow to Fall Below 1
    'inputting = False           ' Set Edit Finish to No
    Prec& = IndexArray(lcnt).RecNum
    If Not Prec& = 0 Then
      
        RecNo& = Prec&
      '  FOpenS "UBCUST.DAT", L2Handle 'open data file
      '  FGetRTA L2Handle, UBCustRec(1), RecNo&, UBCustRecLen
      '  FClose L2Handle
        
        Get #L2Handle, RecNo&, UBCustRec(1)
        'Close L2Handle
      
       
        If FirstCust Then
          FirstCust = False
          WhatBook = Val(UBCustRec(1).Book)
        End If
      
        If (UBCustRec(1).DelFlag <> 0) Or InStr(UBCustRec(1).HHMSG1, "NOREAD") > 0 Then
          process = False
          GoTo SkiptoHere
        End If
        If UseType Then
          ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
          If ThisType$ <> CUSTTYPE$ Then
            process = False
            GoTo SkiptoHere
          End If
        End If

        If Val(UBCustRec(1).Book) >= BegRoute And Val(UBCustRec(1).Book) <= EndRoute Then
          If Not SeqFlag Then
            If Val(UBCustRec(1).Book) <> WhatBook And Grpt <> True Then
              Print #RptHandle, Chr$(12);
              DoHeaderFlag = True
              LineCnt = 0
              WhatBook = Val(UBCustRec(1).Book)
            End If
          End If
          process = True
        Else
          process = False
        End If
      
        If RecNo& <= 0 Then
          process = False
        End If
SkiptoHere:
      
      If process Then
        If Grpt Then
          GoSub PrintLineG
        Else
          GoSub PrintLine
        End If
      End If
    End If
    lcnt = lcnt + 1
    'ShowPctComp cnt, NumOfRecs
    If Not lcnt > NumOfRecs Then
      FrmShowPctComp.ShowPctComp lcnt, NumOfRecs
    Else
      FrmShowPctComp.ShowPctComp lcnt, lcnt
    End If
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Close
      ActivateControls Me, True
      Exit Sub
    End If

  Loop Until lcnt > NumOfRecs

  Close

'  Select Case dev$
'  Case "S"
'    EntryPoint = 2
'  Case "P"
'    EntryPoint = 5
'  End Select
  Erase IndexArray
  Header$ = "Customer Meter Listing Report"
  If Grpt = True Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptMeterReadList
    ARptMtrReadList.Title = Header$
    ARptMtrReadList.txtDate = Now
    ARptMtrReadList.txtTown = TOWNNAME$
    ARptMtrReadList.GetName ReportFile$
    ARptMtrReadList.startrpt
  Else
    'PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
    ViewPrint ReportFile$, Header$
    ActivateControls Me, True
  End If
  Exit Sub

  'Print the Meter Reading Here *********************************************

PrintLine:
  If LineCnt = 0 Then GoSub PrintHeading

  'Help$ = "Process Location Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
  'PrintHelp Help$

  MtrCnt = 0
  ValidCustomer = False
  Do
    MtrCnt = MtrCnt + 1         'Check For Meter This Customer
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    If Len(TempRev$) <> 0 Then ValidCustomer = True: Exit Do
  Loop Until MtrCnt = 7
  If ValidCustomer = False Then Return
  If LineCnt >= 53 Then
    Print #RptHandle, Chr$(12);
    GoSub PrintHeading
  End If

  GoSub GetMeterStatusPrint
  Print #RptHandle, Left$(UBCustRec(1).CustName, 30);
  If UBSetUpRec(1).UseSeq = "Y" Then
    If UBCustRec(1).Seq < 0 Then UBCustRec(1).Seq = 0
    Print #RptHandle, Tab(32); Using("######", UBCustRec(1).Seq);
  End If
  Print #RptHandle, Tab(40); Left$(UBCustRec(1).ServAddr, 28);
  Print #RptHandle, Tab(70); MeterStatus$
  LineCnt = LineCnt + 1

  For MtrCnt = 1 To 7           'find last active meter
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    If Len(TempRev$) <> 0 Then
      GoSub GetMeterTypePrint
      Print #RptHandle, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB;
      Print #RptHandle, Tab(15); QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum);
      Print #RptHandle, Tab(35); MeterType$;
      If ChkNoPrintCurr = 1 Then
        Print #RptHandle, Tab(55); " "
      Else
        Print #RptHandle, Tab(55); Using("##########", UBCustRec(1).LocMeters(MtrCnt).CurRead);
      End If
      Print #RptHandle, Tab(68); "___________"
      LineCnt = LineCnt + 1
    End If
  Next MtrCnt
  Print #RptHandle, String$(79, "-"): LineCnt = LineCnt + 1
  Return
  ' END OF PRINT ROUTINE *****************************************

PrintHeading:
  Page = Page + 1
  Print #RptHandle, Tab(27); "Meter Reading Listing Report"; Tab(65); "Date: "; Date$
  Print #RptHandle, "Beginning Route: "; BegRoute
  Print #RptHandle, "   Ending Route: "; EndRoute; Tab(65); "Page #"; Page
  Print #RptHandle, " "
  Print #RptHandle, "Customer Name";
  If UBSetUpRec(1).UseSeq = "Y" Then
    Print #RptHandle, Tab(32); "Seq #";
  End If
  Print #RptHandle, Tab(40); "Service Address"; Tab(70); "Status"
  Print #RptHandle, "Location"; Tab(15); "Meter Number"; Tab(35); "Mtr Type"; Tab(55); "Cur Read"; Tab(70); "New Read"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
  Return

GetMeterTypePrint:
  Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
  Case "C"
    MeterType$ = "Water/Sewer"
  Case "W"
    MeterType$ = "Water Only"
  Case "S"
    MeterType$ = "Sewer Only"
  Case "T"
    MeterType$ = "Touch Read"
  Case "E"
    MeterType$ = "Electric"
  Case "D"
    MeterType$ = "Demand"
  Case "I"
    MeterType$ = "Irrigation"
  Case "G"
    MeterType$ = "Gas"
  Case Else
    MeterType$ = "Undefined"
  End Select
  Return
GetMeterStatusPrint:
  Select Case UBCustRec(1).Status
  Case "A"
    MeterStatus$ = "Active"
  Case "F"
    MeterStatus$ = "Final"
  Case "I"
    MeterStatus$ = "Vacant"
  Case Else
    MeterStatus$ = "Undef."
  End Select
  Return
PrintLineG:
 ' If Linecnt = 0 Then GoSub PrintHeading

  'Help$ = "Process Location Record #" + STR$(Cnt) + " of " + STR$(IdxNumOfRecs)
  'PrintHelp Help$

  MtrCnt = 0
  ValidCustomer = False
  Do
    MtrCnt = MtrCnt + 1         'Check For Meter This Customer
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    If Len(TempRev$) <> 0 Then ValidCustomer = True: Exit Do
  Loop Until MtrCnt = 7
  If ValidCustomer = False Then Return
'  If Linecnt >= 53 Then
'    Print #RptHandle, Chr$(12);
'    GoSub PrintHeading
'  End If

  GoSub GetMeterStatusPrint
  ToPrintN$ = UBCustRec(1).Book + "~" + Left$(UBCustRec(1).CustName, 30) + "~"
  If UBSetUpRec(1).UseSeq = "Y" Then
    If UBCustRec(1).Seq < 0 Then UBCustRec(1).Seq = 0
    ToPrintN$ = ToPrintN$ + Using("######", UBCustRec(1).Seq) + "~"
  Else
    ToPrintN$ = ToPrintN$ + " ~"
  End If
  ToPrintN$ = ToPrintN$ + Left$(UBCustRec(1).ServAddr, 28) + "~"
  ToPrintN$ = ToPrintN$ + MeterStatus$ + "~"
  'Linecnt = Linecnt + 1

  For MtrCnt = 1 To 7           'find last active meter
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    If Len(TempRev$) <> 0 Then
      GoSub GetMeterTypePrint
      ToPrint$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~"
      ToPrint$ = ToPrint$ + QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum) + "~"
      ToPrint$ = ToPrint$ + MeterType$ + "~"
      If ChkNoPrintCurr = 1 Then
        ToPrint$ = ToPrint$ + " " + "~"
      Else
        ToPrint$ = ToPrint$ + Using("##########", UBCustRec(1).LocMeters(MtrCnt).CurRead) + "~"
      End If
      ToPrint$ = ToPrint$ + "___________"
      'Linecnt = Linecnt + 1
    Print #RptHandle, ToPrintN$ + ToPrint$
    End If
    
  Next MtrCnt
  'Print #RptHandle, String$(79, "-"): Linecnt = Linecnt + 1
  ToPrint$ = ""
  ToPrintN$ = ""
  Return

End Sub
Private Sub PrintMeterReport()
  Dim UBCustRecLen As Integer, ReportFile As String, RptHandle As Integer
  Dim UBSetupLen As Integer, SeqFlag As Boolean, IdxName As String
  Dim NumOfRecs As Long, IdxNumOfRecs As Long, Handle As Integer
  Dim cnt As Long, lcnt As Long, Prec As Long, process As Boolean
  Dim Header As String, MtrCnt As Long, ValidCustomer As Boolean
  Dim TempRev As String, MeterStatus As String, MeterType As String
  Dim Page As Integer, RecNo As Long, L2Handle As Integer
  Dim IdxFileSize As Long, IdxRecLen As Integer, Book As Integer
  Dim FirstCust As Boolean, DoHeaderFlag As Boolean, PrintMrtFlag As Boolean
  Dim ToPrint As String, ToPrintN As String, CustName As String
  Dim DidOne As Boolean, Multi As Double, MeterConsp As Double
  Dim MaxMeterAmt As Long, UBCust As Integer, IdxNameM As String
  Dim CUSTTYPE As String, ThisType As String, UseType As Boolean
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  FrmShowPctComp.Label1 = "Creating Meter Reading Report"
  FrmShowPctComp.Show , Me

  MaxLines = 52
  FF$ = Chr$(12)
  CUSTTYPE$ = QPTrim$(fptxtCustType)
  If Len(CUSTTYPE$) > 0 Then
    UseType = True
  Else
    UseType = False
  End If

  'Open Report File
  ReportFile$ = UBPath$ + "UBMTRLST.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  CustName$ = Space$(30)

  ' Location Order ********************************************************
  'if UBSetupRec(1).
 '

  IdxName$ = UBPath$ + "UBCUSTBK.IDX"
'
'  IdxRecLen = 4 'we are using a long integer
'  IdxFileSize& = FileSize&(IdxName$)
'  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

'  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
'  Handle = FreeFile
'  Open IdxName$ For Random Shared As Handle Len = IdxRecLen
'  For cnt& = 1 To IdxNumOfRecs
'    Get #Handle, cnt&, IdxBuff(cnt&)
'  Next
'  Close Handle
 ' IdxName$ = UBPath$ + "UBTempBK.IDX"

  NumOfRecs& = FileSize&(IdxName$) \ 4
  IdxNumOfRecs = NumOfRecs
  ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
  'FGetAH IdxName$, IndexArray(1), 4, NumOfRecs
  Handle = FreeFile
  Open IdxName$ For Random Shared As Handle Len = 4
  For cnt& = 1 To IdxNumOfRecs
    Get #Handle, cnt&, IndexArray(cnt&)
  Next
  Close Handle

'
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  cnt& = 1
  GoSub PrintReadHeading
  'ShowProcessingScrn "Reading Meter Information"
  For lcnt& = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp lcnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      Close
      Exit Sub
    End If
    Get #UBCust, IndexArray(lcnt&).RecNum, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then 'And UBCustRec(1).BILLCYCL = 99 Then
        If UseType Then
          ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
          If ThisType$ <> CUSTTYPE$ Then
            process = False
            GoTo SkiptoHere
          End If
        End If

      If InStr(UBCustRec(1).HHMSG1, "NOREAD") = 0 Then
        'Book = QPValI(UBCustRec(1).Book)
        Book = Val(UBCustRec(1).Book)
        If Book >= BegRoute And Book <= EndRoute Then
          LSet CustName$ = QPTrim(UBCustRec(1).CustName)
          If UBCustRec(1).Status > "" Then
            If Not Grpt Then
              Print #RptHandle, UBCustRec(1).Book; "-"; QPTrim(UBCustRec(1).SEQNUMB); "   "; QPTrim(UBCustRec(1).Status); "   "; CustName$; Left$(UBCustRec(1).ServAddr, 30)
            Else
              ToPrintN$ = UBCustRec(1).Book + "-" + QPTrim(UBCustRec(1).SEQNUMB) + "~" + QPTrim(UBCustRec(1).Status) + "~" + CustName$ + "~" + Left$(UBCustRec(1).ServAddr, 30) + "~"
            End If
            'IF LEN(QPTrim$(UBCustRec(1).EstFlag)) > 0 THEN STOP
            LineCnt = LineCnt + 1
            For MtrCnt = 1 To 7                'find last active meter
              TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
              If Len(TempRev$) > 0 Then
                GoSub GetReadMeterTypePrint

                If PrintMrtFlag Then
                  '    IF MtrCnt& > 1 THEN
                  '      PRINT #RptHandle, "HERE DALE"
                  '    END IF
                  'IF TypeFlag AND LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MTRNUM)) > 0 THEN
                  '  GOTO DonotPrintEm
                  'END IF

                  DidOne = True
                  If Not Grpt Then
                    Print #RptHandle, QPTrim(UBCustRec(1).LocMeters(MtrCnt).MtrNum);
                    Print #RptHandle, Tab(14); MeterType$;
                    Multi# = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
                    If Multi# = 0 Then Multi# = 1
                    Print #RptHandle, Tab(24); Using("#####", Multi#);
                    Print #RptHandle, Tab(31); Using("##########", UBCustRec(1).LocMeters(MtrCnt&).CurRead);
                    Print #RptHandle, Tab(42); Using("##########", UBCustRec(1).LocMeters(MtrCnt&).PrevRead);
                  Else
                    ToPrint$ = QPTrim(UBCustRec(1).LocMeters(MtrCnt).MtrNum) + "~"
                    ToPrint$ = ToPrint$ + MeterType$ + "~"
                    Multi# = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
                    If Multi# = 0 Then Multi# = 1
                    ToPrint$ = ToPrint$ + Using("#####", Multi#) + "~" + Using("##########", UBCustRec(1).LocMeters(MtrCnt&).CurRead)
                    ToPrint$ = ToPrint$ + "~" + Using("##########", UBCustRec(1).LocMeters(MtrCnt&).PrevRead)
                  End If
                  If UBCustRec(1).LocMeters(MtrCnt).CurRead < 0 Or UBCustRec(1).LocMeters(MtrCnt&).PrevRead < 0 Then
                    MeterConsp# = 0
                  Else
                    MeterConsp# = UBCustRec(1).LocMeters(MtrCnt).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
                  End If
                  If MeterConsp# < 0 Then
                    MaxMeterAmt& = 10& ^ (Len(Str$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                    MeterConsp# = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
                  End If

                  MeterConsp# = Round#(MeterConsp# * Multi#)
                  If Not Grpt Then
                    Print #RptHandle, Tab(53); Using("##########", MeterConsp#);
                    If UBCustRec(1).LocMeters(MtrCnt&).ReadFlag <> "Y" Then
                      Print #RptHandle, Tab(67); "UNREAD"
                    Else
                      Print #RptHandle, Tab(67); Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
                    End If
                  'RINT #RptHandle, TAB(67); UBCustRec(1).LocMeters(MtrCnt&).MtrMulti
                  'PRINT #RptHandle, " "; UBCustRec(1).LocMeters(MtrCnt&).Readflag
                  LineCnt = LineCnt + 1
                  Else
                    ToPrint$ = ToPrint$ + "~" + Using("##########", MeterConsp#)
                    If UBCustRec(1).LocMeters(MtrCnt&).ReadFlag <> "Y" Then
                      ToPrint$ = ToPrint$ + "~" + "UNREAD"
                    Else
                      ToPrint$ = ToPrint$ + "~" + Num2Date$(UBCustRec(1).LocMeters(MtrCnt&).CurDate)
                    End If
                    Print #RptHandle, ToPrintN$ + ToPrint$
                    ToPrint$ = ""
                  End If
                End If
              End If
DonotPrintEm:
            Next MtrCnt&

            If Not DidOne Then
              If Not Grpt Then
                Print #RptHandle, Tab(14); "NO METERED SERVICE"
                LineCnt = LineCnt + 1
              Else
                ToPrint$ = " **~ ~  NO ~ Metered ~ Service~ ~ ~ ~"
                Print #RptHandle, ToPrintN$ + ToPrint$
                ToPrint$ = ""
              End If
            End If
            DidOne = False
            If Not Grpt Then
              Print #RptHandle, String$(79, "-")
              LineCnt = LineCnt + 1
            End If
          End If
        End If
      End If
    End If
    If LineCnt >= MaxLines And Not Grpt Then
      Print #RptHandle, FF$
      GoSub PrintReadHeading
    End If
   ' ShowPctComp lcnt&, IdxNumOfRecs
   ToPrintN$ = ""
SkiptoHere:
  Next
  If Not Grpt Then
    Print #RptHandle, FF$
  End If
  
  Close

  Header$ = "Customer Meter Reading Report"

  Erase IndexArray
  'PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  If Not Grpt Then
    ViewPrint ReportFile$, Header$
    ActivateControls Me, True
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptMeterReadList
    ARptMtrReadReport.Title = Header$
    ARptMtrReadReport.txtDate = Now
    ARptMtrReadReport.txtTown = TOWNNAME$
    ARptMtrReadReport.GetName ReportFile$
    ARptMtrReadReport.startrpt
  End If
  Exit Sub

PrintReadHeading:
  If Not Grpt Then
    Page = Page + 1
    Print #RptHandle, Tab(30); "Meter Reading Report"; Tab(65); "Date: "; Date$
    Print #RptHandle, "Beginning Route: "; BegRoute
    Print #RptHandle, "   Ending Route: "; EndRoute; Tab(70); "Page #"; Page
    Print #RptHandle, ""
    Print #RptHandle, "Location Status Customer Name"; Tab(41); "Service Address"
    Print #RptHandle, " Mtr No.    Mtr Type    Multi    Current   Previous    Consump     Read Date"
    Print #RptHandle, String$(80, "=")
    LineCnt = 7
  End If
  Return

GetReadMeterTypePrint:
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
