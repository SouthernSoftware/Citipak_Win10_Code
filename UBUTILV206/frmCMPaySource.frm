VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmUtilDateEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Date Edit"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmCMPaySource.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPaySource 
      Height          =   324
      Left            =   5304
      TabIndex        =   0
      Top             =   3048
      Width           =   3828
      _Version        =   196608
      _ExtentX        =   6752
      _ExtentY        =   572
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCMPaySource.frx":08CA
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   5
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
            TextSave        =   "12:06 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "1/24/2005"
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
      Left            =   8688
      TabIndex        =   3
      Top             =   6888
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
      ButtonDesigner  =   "frmCMPaySource.frx":0C61
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   480
      Left            =   7224
      TabIndex        =   2
      Top             =   6888
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
      ButtonDesigner  =   "frmCMPaySource.frx":0E3D
   End
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5304
      TabIndex        =   1
      Top             =   3584
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Text            =   "10/03/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   5304
      TabIndex        =   9
      Top             =   4680
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Text            =   "10/03/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
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
   Begin EditLib.fpText txtOperator 
      Height          =   348
      Left            =   5304
      TabIndex        =   11
      Top             =   4152
      Width           =   804
      _Version        =   196608
      _ExtentX        =   1418
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
      MaxLength       =   4
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Transaction Date:"
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
      Height          =   396
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      Top             =   4752
      Width           =   3240
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date:"
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
      Height          =   396
      Index           =   1
      Left            =   2952
      TabIndex        =   8
      Top             =   3596
      Width           =   2088
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   2820
      Left            =   1800
      Top             =   2640
      Width           =   8004
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Source:"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   3072
      Width           =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   312
      Index           =   0
      Left            =   2736
      TabIndex        =   6
      Top             =   4216
      Width           =   2304
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date Edit"
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
      Left            =   3288
      TabIndex        =   4
      Top             =   1632
      Width           =   5652
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   1392
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   1272
      Width           =   5772
   End
End
Attribute VB_Name = "frmUtilDateEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Oper As String
Dim BadDate As Boolean, TTCnt As Long
Private Sub CheckDate()
Dim payDate As String, paydate2 As String
  payDate$ = txtDate1.Text
  paydate2$ = txtDate2.Text
  If Val(Left$(payDate$, 2)) < 1 Or Val(Left$(payDate$, 2)) > 12 Then
    If Val(Mid$(payDate$, 4, 2)) < 1 Or Val(Mid$(payDate$, 4, 2)) > 31 Then
      BadDate = True
      Exit Sub
    Else
      BadDate = False
    End If
  Else
    BadDate = False
  End If
  If Val(Left$(paydate2$, 2)) < 1 Or Val(Left$(paydate2$, 2)) > 12 Then
    If Val(Mid$(paydate2$, 4, 2)) < 1 Or Val(Mid$(paydate2$, 4, 2)) > 31 Then
      BadDate = True
      Exit Sub
    Else
      BadDate = False
    End If
  Else
    BadDate = False
  End If

End Sub

Private Sub cmdOK_Click()
  If BadDate = False Then
    If MsgBox("Are You Sure ?  Continue with Date Edit...", vbYesNo, "Continue") = vbYes Then
      TTCnt = 0
      ChgTransDate
      MsgBox "Transactions Changed: " + Str(TTCnt), vbOKOnly, "Completed"
    End If
  Else
    MsgBox "Invalid Date", vbOKOnly, "Invalid"
  End If
End Sub

'Private Sub cmdBLAdj_Click()
'  Dim CMSetuplen As Integer
'  ReDim CMSetUpRec(1) As CMSetupType
'  CMSetuplen = Len(CMSetUpRec(1))
'  LoadCMSetUpFile CMSetUpRec(), CMSetuplen
'  If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
'  'do the password screen
'    frmPassWord.Callingfrm = 3
'    frmPassWord.Show 1
'  ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
'  'get opernum properties if full access then goahead on
'    If LevelPass = 1 Then
'      Load frmBLAdjustBal
'      frmBLAdjustBal.Show
'      Unload Me
'    Else
'      MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'    End If
'  Else 'nobody cares
'    Load frmBLAdjustBal
'    frmBLAdjustBal.Show
'    Unload Me
'  End If
'  Erase CMSetUpRec
'End Sub
'
'Private Sub cmdUtilAdj_Click()
'  Dim CMSetuplen As Integer
'  ReDim CMSetUpRec(1) As CMSetupType
'  CMSetuplen = Len(CMSetUpRec(1))
'  LoadCMSetUpFile CMSetUpRec(), CMSetuplen
'  If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
'  'do the password screen
'    frmPassWord.Callingfrm = 2
'    frmPassWord.Show 1
'  ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
'  'get opernum properties if full access then goahead on
'    If LevelPass = 1 Then
'      Load frmUBAdjustmentEntry
'      frmUBAdjustmentEntry.Show
'      Unload Me
'    Else
'      MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'    End If
'  Else 'nobody cares
'    Load frmUBAdjustmentEntry
'    frmUBAdjustmentEntry.Show
'    Unload Me
'  End If
'  Erase CMSetUpRec
'End Sub
'
'Private Sub cmdVoidPayments_Click()
'  Dim CMSetuplen As Integer
'  ReDim CMSetUpRec(1) As CMSetupType
'  CMSetuplen = Len(CMSetUpRec(1))
'  LoadCMSetUpFile CMSetUpRec(), CMSetuplen
'  If QPTrim(CMSetUpRec(1).Pass4Voids) = "Y" Then
'  'do the password screen
'    frmPassWord.Callingfrm = 1
'    frmPassWord.Show 1
'  ElseIf QPTrim(CMSetUpRec(1).Pass4Voids) = "F" Then
'  'get opernum properties if full access then goahead on
'    If LevelPass = 1 Then
'      Load frmVoidSearch
'      frmVoidSearch.Show
'      Unload Me
'    Else
'      MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'    End If
'  Else 'nobody cares
'    Load frmVoidSearch
'    frmVoidSearch.Show
'    Unload Me
'  End If
'  Erase CMSetUpRec
'End Sub
'
'Private Sub cmdOk_Click()
'  CheckPayDate
'  If BadDate = False Then
'    Select Case fpcboPaySource.ListIndex
'      Case 0:
'        frmPayUtilEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
'        frmPayUtilEntry.Show
'        DoEvents
'        Unload Me
'      Case 1:
'        frmPayDepEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
'        frmPayDepEntry.Show
'        DoEvents
'        Unload Me
'      Case 2:
'        frmPayBLEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
'        frmPayBLEntry.Show
'        DoEvents
'        Unload Me
'      Case 3:
'        frmPayMiscEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
'        frmPayMiscEntry.Show
'        DoEvents
'        Unload Me
'      Case Else:
'        MsgBox "Invalid Selection", vbOKOnly, "Invalid Source"
'    End Select
'  Else
'    MsgBox "Invalid Date", vbOKOnly, "Invalid Entry"
'  End If
'
'End Sub
'
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
      'Call cmdOk_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
'  fpcboPaySource.AddItem "Utility Billing Payment"
'  fpcboPaySource.AddItem "Utility Deposit Entry"
'  fpcboPaySource.AddItem "Business License Payment"
'  fpcboPaySource.AddItem "Miscellaneous Payment"
  fpcboPaySource.AddItem "Utility Transaction"
  fpcboPaySource.ListIndex = 0
  fpcboPaySource.Enabled = False
  txtOperator = ""
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  'lblOperName.Caption = PWUser
  'Oper$ = QPTrim(lblOperator.Caption)
  'CMLog " IN Oper " + Oper$ + "CMPaySource"
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

Private Sub fpcboPaySource_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPaySource.ListDown = True
  End If
  If fpcboPaySource.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtDate1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpCmdExit_Click()
  'frmCMMainMenu.Show
  Unload Me
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    cmdOk.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpcboPaySource.SetFocus
  End If
End Sub
Private Sub ChgTransDate()
  Dim findall As Boolean, opertofind As Integer, UBTranRecLen As Integer
  Dim FromDate As Integer, ToDate As Integer, UBFile As Integer
  Dim TNumOfRecs As Long, cnt As Long
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  findall = False
  FromDate = Date2Num(txtDate1.Text)
  ToDate = Date2Num(txtDate2.Text)
  FrmShowPctComp.Label1 = "Changing Transaction Date"
  FrmShowPctComp.Show , Me

  If QPTrim$(txtOperator) = "0" Or Len(QPTrim$(txtOperator)) = 0 Then
    findall = True
  End If
  opertofind = Val(txtOperator)
  'ShowWarning
  
 ' Print "   Change Transaction Dates"
 ' Print
 ' Print "   From: "; Num2Date$(FromDate); " to "; Num2Date$(ToDate)
'  Ok$ = GetProceed$
 ' Select Case Ok$
'  Case "Y"
 '   Print "Y"
 '   Print
    UBFile = FreeFile
    Open "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
    TNumOfRecs& = LOF(UBFile) / UBTranRecLen
    For cnt& = 1 To TNumOfRecs&
      FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
      Get UBFile, cnt&, UBTranRec(1)
      If UBTranRec(1).TransDate = FromDate Then
        If findall = False Then
          If UBTranRec(1).OperatorNumber = opertofind Then
            TTCnt = TTCnt + 1
            UBTranRec(1).TransDate = ToDate
            Put UBFile, cnt&, UBTranRec(1)
          End If
        Else
          TTCnt = TTCnt + 1
          UBTranRec(1).TransDate = ToDate
          Put UBFile, cnt&, UBTranRec(1)
        End If
      End If
    Next
'    Close
'    Print
'    Print
'    Print "      Changed: "; TTCnt
'  Case Else
'    Print "N"
'    Print
'    Print
'    Print "   NO TRANSACTION DATES CHANGED"
'  End Select
  Erase UBTranRec
End Sub
'Private Sub ShowTransFixDate(Recno&)
'
'  ReDim UBTranRec(1) As UBTransRecType
'  ReDim UBCustRec(1) As NewUBCustRecType
'
'  UBCustRecLen = Len(UBCustRec(1))
'  UBTranRecLen = Len(UBTranRec(1))
'
'  UBFile = FreeFile
'  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
'  Get UBFile, Recno&, UBCustRec(1)
'  Close UBFile
'
'  UBTran = FreeFile
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
'
'  PrevTranRec& = UBCustRec(1).LastTrans
'
'  If PrevTranRec& > 0 Then
'    Do While PrevTranRec& > 0
'      dcnt = dcnt + 1
'      ReDim Preserve MTChoice(1 To dcnt) As FLen2
'      Get UBTran, PrevTranRec&, UBTranRec(1)
'      'IF PrevTranRec& = 11236 THEN
'      '  FOR Rev = 1 TO 15
'      '    UBTranRec(1).RevAmt(Rev) = ABS(UBTranRec(1).RevAmt(Rev))
'      '  NEXT
'      '  UBTranRec(1).TransAmt = ABS(UBTranRec(1).TransAmt)
'      '  PUT UBTran, PrevTranRec&, UBTranRec(1)
'      'END IF
'      If Len(QPTrim$(UBTranRec(1).TransDesc)) = 0 Then
'        UBTranRec(1).TransDesc = "????"
'      End If
'      LSet MTChoice(dcnt).V = Num2Date(UBTranRec(1).TransDate)
'      Mid$(MTChoice(dcnt).V, 15) = Left$(UBTranRec(1).TransDesc, 15)
'      Mid$(MTChoice(dcnt).V, 40) = Using(Str$(UBTranRec(1).Transamt), "#####.##")
'      Mid$(MTChoice(dcnt).V, 49) = Using(Str$(UBTranRec(1).RunBalance), "#####.##")
'      Mid$(MTChoice(dcnt).V, 59) = MKL$(PrevTranRec&)
'      PrevTranRec& = UBTranRec(1).PrevTrans
'    Loop
'    Close UBTran
'
'    MaxLen = 57 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'
'    If Choice < 1 Then
'      Choice = 1                'Pre-load choice to highlight
'    End If
'
'    Title$ = Space$(MaxLen + 4)
'    LSet Title$ = " " + Left$(QPTrim$(UBCustRec(1).CustName), 20)
'    Mid$(Title$, 25) = Left$(QPTrim$(UBCustRec(1).ServAddr), 25)
'    Mid$(Title$, 52, 9) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'
'    '--Find max menu width
'    '--Center Menu within Screen
'
'
'
'EditTransDate:
'
'
''    '----- Set the "Action" flag to force the editor to initialize itself and
''    '      display the data on the form.
'    Action = 1
'    GoSub GetTransInfo
'
'    FirstTime = True
'
'    Do
'
'      EditForm Form$(), Fld(), frm(1), Cnf, Action
'
'      If FirstTime Then
'        FirstTime = False
'        LSet Form$(1, 0) = Num2Date$(TranDate)
'        Action = 1
'      End If
'      Select Case frm(1).KeyCode
'      Case F0Key
'        GoSub ChangeDate
'        Exit Do
'      Case EscKey
'        Exit Do
'      End Select
'    Loop
'
'Return
'
'ChangeDate:
'
'  UBTran = FreeFile
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
'  Get UBTran, TransRec&, UBTranRec(1)
'  UBTranRec(1).TransDate = Date2Num%(Form$(1, 0))
'  Put UBTran, TransRec&, UBTranRec(1)
'  Close
'
'Return
'
'GetTransInfo:
'  TransRec& = CVL(Mid$(MTChoice(Choice).V, 59, 4))
'  UBTran = FreeFile
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
'  Get UBTran, TransRec&, UBTranRec(1)
'  Close
'  TranDate = UBTranRec(1).TransDate
'Return
'
'End Sub
'
