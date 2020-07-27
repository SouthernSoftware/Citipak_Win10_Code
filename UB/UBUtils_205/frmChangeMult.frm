VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmChangeMult 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Multiplier by Rate Code"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmChangeMult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpComboRates 
      Height          =   348
      Left            =   6090
      TabIndex        =   2
      Top             =   3528
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
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
      ColDesigner     =   "frmChangeMult.frx":08CA
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   6
      Top             =   8532
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4:50 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4/28/2006"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   8760
      TabIndex        =   4
      Top             =   6216
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmChangeMult.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   480
      Left            =   7224
      TabIndex        =   3
      Top             =   6216
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmChangeMult.frx":0D9D
   End
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   7530
      TabIndex        =   1
      Top             =   2880
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
      Left            =   5298
      TabIndex        =   0
      Top             =   2880
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
   Begin EditLib.fpText fptxtMult 
      Height          =   348
      Left            =   6120
      TabIndex        =   12
      Top             =   4128
      Width           =   1284
      _Version        =   196608
      _ExtentX        =   2265
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
      MaxLength       =   5
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
      Caption         =   "From Book:"
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
      Left            =   3522
      TabIndex        =   11
      Top             =   2904
      Width           =   1620
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(01 to 99 for All)"
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
      Left            =   8490
      TabIndex        =   10
      Top             =   2904
      Width           =   2124
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "For Customers with rate code:"
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
      Left            =   1554
      TabIndex        =   9
      Top             =   3552
      Width           =   4380
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Book:"
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
      Left            =   6018
      TabIndex        =   8
      Top             =   2904
      Width           =   1380
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   2676
      Left            =   1506
      Top             =   2592
      Width           =   9204
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Assign This Multiplier:"
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
      Height          =   300
      Left            =   2064
      TabIndex        =   7
      Top             =   4104
      Width           =   3864
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Multiplier for Rate Code Criteria"
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
      Left            =   2550
      TabIndex        =   5
      Top             =   1608
      Width           =   7116
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   2922
      Top             =   1368
      Width           =   6396
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   2922
      Top             =   1248
      Width           =   6396
   End
End
Attribute VB_Name = "frmChangeMult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Oper As String
Dim BegRoute As String, EndRoute As String

Private Sub cmdOk_Click()
  SetMult2cust
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'CMLog "Closed via SelectPaySource by " + PWUser$ + " operator-" + Oper$
       ' CitiTerminate
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
      Call fpCmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdOk_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fptxtRoute1 = "00"
  fptxtRoute2 = "99"
  
  Dim RCode As String
  Dim Handle As Integer, cnt As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  ReDim UBRateTblRec(1) As UBRateTblRecType
  
  RCode$ = Space$(10)
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get Handle, cnt, UBRateTblRec(1)
    LSet RCode$ = QPTrim$(UBRateTblRec(1).RATECODE)
    fpComboRates.AddItem RCode$ + QPTrim$(UBRateTblRec(1).RATEDESC)
  Next
  Close
  fptxtMult = "10"
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
'
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub



Private Sub fptxtMult_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtMult.SetFocus
  End If
End Sub

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
    fpComboRates.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpComboRates_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpComboRates.ListDown = True
  End If
  If fpComboRates.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtMult.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fptxtMult_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpCmdExit_Click()
  frmUBEditMenu.Show
  Unload Me
End Sub


Private Sub SetMult2cust()
  Dim UBCustRecLen As Integer, RateRec As Integer, RateNew As Integer
  Dim UBFile As Integer, IdxRecLen As Integer, WhatRate As String
  Dim NumOfRecs As Long, IdxFileSize As Long, NewRate As String
  Dim CCnt As Long, ReSeq As Long, IdxNumOfRecs As Long, RevSource As Integer
  Dim CustLen As Integer, Handle As Integer, ThisCust As Long
  Dim cnt As Long, RCnt As Integer, CustBook As Integer, Mcnt As Integer
  Dim FromBook As Integer, ThruBook As Integer, SCnt As Integer
  Dim RCode As String, ThisRate As String, NuMult As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  ReDim UBRateTblRec(1) As UBRateTblRecType

  FromBook = Val(fptxtRoute1)
  'If FromBook = 0 Then FromBook = 1
  ThruBook = Val(fptxtRoute2)
  If fpComboRates.ListIndex = -1 Then
    MsgBox "Invalid Rate Code.", vbOKOnly, "Invalid Entry"
    Exit Sub
  End If
  RateRec = fpComboRates.ListIndex + 1
  NuMult = Val(fptxtMult.Text)
  If NuMult <= 0 Then
    MsgBox "Invalid Multiplier.", vbOKOnly, "Invalid Entry"
    Exit Sub
  End If
  
  UBRateTblRecLen = Len(UBRateTblRec(1))
  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  Get Handle, RateRec, UBRateTblRec(1)
  WhatRate$ = QPTrim$(UBRateTblRec(1).RATECODE)
  Close Handle
  DeActivateControls frmChangeMult
  FrmShowPctComp.Label1 = "Change Mult ..."
  FrmShowPctComp.Show , Me
  UBLog "SetMult - " & ", book(" & fptxtRoute1 & "-" & fptxtRoute2 & "),Mult-" & fptxtMult.Text
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType

    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open "UBCUSTBK.IDX" For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  UBFile = FreeFile
  Open "UBCUST.dat" For Random Shared As UBFile Len = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  For cnt& = 1 To IdxNumOfRecs&
    FrmShowPctComp.ShowPctComp cnt&, IdxNumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Exit Sub
    End If
    ThisCust& = IdxBuff(cnt&).RecNum
    Get UBFile, ThisCust&, UBCustRec(1)
    CustBook = Val(UBCustRec(1).Book)
    If CustBook >= FromBook And CustBook <= ThruBook Then
      If Not UBCustRec(1).DelFlag Then
        For SCnt = 1 To 15
          ThisRate$ = QPTrim$(UBCustRec(1).serv(SCnt).RATECODE)
          If WhatRate$ = ThisRate$ Then
            For Mcnt = 1 To 7
              If UBCustRec(1).serv(SCnt).RMtrType = UBCustRec(1).LocMeters(Mcnt).MtrType Then
                UBCustRec(1).LocMeters(Mcnt).MTRMulti = NuMult
              End If
            Next
            Put UBFile, ThisCust&, UBCustRec(1)
            Exit For
          End If
        Next
      End If
    End If
  Next
  Close

  Erase UBCustRec, IdxBuff
UBLog "setmult:" & fptxtMult.Text & " Using Util set mult"
MsgBox "Set Mult:" & fptxtMult.Text, vbOKOnly, " Multiplier"
'ReIndexSystem False
ActivateControls frmChangeMult
End Sub


