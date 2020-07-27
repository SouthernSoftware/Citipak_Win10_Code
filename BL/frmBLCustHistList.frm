VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCustHistList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Customer List"
   ClientHeight    =   7350
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8805
   Icon            =   "frmBLCustHistList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   3630
      Left            =   1635
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Customer listing."
      Top             =   2070
      Width           =   5535
      _Version        =   196608
      _ExtentX        =   9763
      _ExtentY        =   6403
      TextAlias       =   ""
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      BorderColor     =   0
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
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmBLCustHistList.frx":08CA
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008F8265&
      Caption         =   "Customer Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1050
      Left            =   1608
      TabIndex        =   3
      Top             =   885
      Width           =   5580
      Begin VB.OptionButton optFirst 
         BackColor       =   &H008F8265&
         Caption         =   "First Customer"
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
         Height          =   252
         Left            =   768
         TabIndex        =   1
         ToolTipText     =   "Press F3 for help with this field."
         Top             =   240
         Width           =   1710
      End
      Begin VB.OptionButton optLast 
         BackColor       =   &H008F8265&
         Caption         =   "Last Customer"
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
         Height          =   252
         Left            =   3312
         TabIndex        =   2
         ToolTipText     =   "Press F3 for help with this field."
         Top             =   240
         Width           =   1665
      End
      Begin EditLib.fpText fptxtFirst 
         Height          =   420
         Left            =   570
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Press F3 for help with this field."
         Top             =   525
         Width           =   1890
         _Version        =   196608
         _ExtentX        =   3334
         _ExtentY        =   741
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
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
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
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
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
      Begin EditLib.fpText fptxtLast 
         Height          =   420
         Left            =   3165
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Press F3 for help with this field."
         Top             =   525
         Width           =   1890
         _Version        =   196608
         _ExtentX        =   3334
         _ExtentY        =   741
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
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
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
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
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
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   3828
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6096
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBLCustHistList.frx":0CA2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdApply 
      Height          =   492
      Left            =   5784
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Press to send the data above to the print screen and exit this screen."
      Top             =   6096
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBLCustHistList.frx":0EB8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   1290
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Place the cursor over any field and a pop-up balloon will appear containing help information about that field."
      Top             =   6090
      Width           =   2325
      _Version        =   131072
      _ExtentX        =   4101
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBLCustHistList.frx":10CD
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#5) Press 'F10' and the two numbers are inserted in the number fields on the main screen."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   336
      TabIndex        =   14
      Top             =   6672
      Width           =   8124
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   8112
      X2              =   7344
      Y1              =   6720
      Y2              =   6288
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#4) Next click on a customer in the list. The customer number appears in the 'Last Customer' field."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2556
      Left            =   7344
      TabIndex        =   13
      Top             =   2640
      Width           =   1164
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8400
      X2              =   6960
      Y1              =   2832
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#3) Now click on the 'Last Customer' option button."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1596
      Left            =   7344
      TabIndex        =   12
      Top             =   336
      Width           =   1164
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#2) Next click on a customer in the list. The customer number appears in the 'First Customer' field."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2556
      Left            =   336
      TabIndex        =   11
      Top             =   2640
      Width           =   1164
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   960
      X2              =   1728
      Y1              =   4896
      Y2              =   5664
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#1) Start by clicking on the 'First Customer' option button."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1596
      Left            =   336
      TabIndex        =   10
      Top             =   288
      Width           =   1164
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer List"
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
      Height          =   390
      Left            =   2490
      TabIndex        =   8
      Top             =   360
      Width           =   3900
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6972
      Left            =   144
      Top             =   144
      Width           =   8496
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1248
      X2              =   2256
      Y1              =   768
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   432
      X2              =   1824
      Y1              =   2784
      Y2              =   1920
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   5232
      X2              =   7392
      Y1              =   960
      Y2              =   576
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   7152
      X2              =   7680
      Y1              =   5664
      Y2              =   5136
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   450
      Left            =   2370
      Top             =   285
      Width           =   4050
   End
End
Attribute VB_Name = "frmBLCustHistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim AMeth As Integer

Private Sub cmdApply_Click()
  frmBLCustTransHist.fptxtFirst.Text = QPTrim$(fptxtFirst.Text)
  frmBLCustTransHist.fptxtLast.Text = QPTrim$(fptxtLast.Text)
  
  Unload frmBLCustHistList
End Sub

Private Sub cmdClose_Click()
  Unload frmBLCustHistList
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line5.Visible = True
    Line6.Visible = True
    Line7.Visible = True
  Else
    cmdHelp.Text = "F1 &Turn Help On"
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    If optFirst.Value = True Then
      optLast.Value = True
      optFirstValue = False
      optLast.SetFocus
    ElseIf optLast.Value = True Then
      optFirst.Value = True
      optLast.Value = False
      optFirst.SetFocus
    End If
  End If
  
  Select Case KeyCode
    Case vbKeyReturn
      Call cmdApply_Click
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%A"
      Call cmdApply_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  Dim CIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim x As Integer
  Dim Number$
  Dim Name$
  Dim Found As Boolean
  
  On Error GoTo ERRORSTUFF
  Call FixFonts
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  Label6.Visible = False
  Line1.Visible = False
  Line2.Visible = False
  Line3.Visible = False
  Line4.Visible = False
  Line5.Visible = False
  Line6.Visible = False
  Line7.Visible = False
  optFirst.BackColor = &H80FFFF
  optFirst.ForeColor = &H0&
  fptxtFirst.Text = QPTrim$(frmBLCustTransHist.fptxtFirst.Text)
  fptxtLast.Text = QPTrim$(frmBLCustTransHist.fptxtLast.Text)
  
  OpenCustNumIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) \ Len(CIdxRec)

  If NumOfIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim IdxRec(1 To NumOfIdxRecs) As Integer

  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CIdxRec
    IdxRec(x) = CIdxRec.CustRec
  Next x
  Close IdxHandle

  OpenCustFile CustHandle
  For x = 1 To NumOfIdxRecs
    Get CustHandle, IdxRec(x), CustRec
    fpList1.InsertRow = "  " + QPTrim$(CustRec.CustNumb) + Chr(9) + QPTrim$(CustRec.CustName)
   Next x
  Close CustHandle
   
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmCustHistList", "Form Load", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select

Proc_Exit:
'  --- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
    Unload Me
End Sub

Private Sub fpList1_Click()
  fpList1.Col = 0
  
  If optFirst.Value = True Then
    fptxtFirst.Text = QPTrim$(fpList1.ColText)
  ElseIf optLast.Value = True Then
    fptxtLast.Text = QPTrim$(fpList1.ColText)
  End If
End Sub

Private Sub fptxtLast_GotFocus()
  optLast.Value = True
End Sub

Private Sub fptxtFirst_GotFocus()
  optFirst.Value = True
End Sub

Private Sub optLast_Click()
  optLast.BackColor = &H80FFFF
  optLast.ForeColor = &H0&
  optFirst.BackColor = &H8F8265
  optFirst.ForeColor = &HFFFFFF
End Sub

Private Sub optFirst_Click()
  optFirst.BackColor = &H80FFFF
  optFirst.ForeColor = &H0&
  optLast.BackColor = &H8F8265
  optLast.ForeColor = &HFFFFFF

End Sub

Private Sub FixFonts()
  Select Case ScreenW
    Case 1280
      optFirst.FontSize = 10
      fptxtFirst.FontSize = 10
      optLast.FontSize = 10
      fptxtFirst.FontSize = 10
      Label6.FontSize = 10
      Label5.Height = 2950
      Label3.Height = 2950
    Case 1152
      optFirst.FontSize = 10
      fptxtFirst.FontSize = 10
      optLast.FontSize = 10
      fptxtFirst.FontSize = 10
      Label6.FontSize = 10
      Label5.Height = 2950
      Label3.Height = 2950
    Case 1024
      optFirst.FontSize = 10
      fptxtFirst.FontSize = 10
      Label6.FontSize = 10
      optLast.FontSize = 10
      fptxtFirst.FontSize = 10
    Case 800
      optFirst.FontSize = 9
      fptxtFirst.FontSize = 9
      Label6.FontSize = 10
      optLast.FontSize = 9
      fptxtFirst.FontSize = 9
    Case Else
  End Select
End Sub
