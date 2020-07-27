VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFADispItemList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Building List Of Assets For Disposal"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFADispItemList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListDates 
      Height          =   1200
      Left            =   1290
      TabIndex        =   1
      ToolTipText     =   "Click on a date to bring up data saved for it."
      Top             =   1680
      Width           =   1845
      _Version        =   196608
      _ExtentX        =   3254
      _ExtentY        =   2117
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   0
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
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
      ColDesigner     =   "frmFADispItemList.frx":08CA
   End
   Begin EditLib.fpDateTime fpDateActive 
      Height          =   390
      Left            =   9045
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "This displays the disposal date activated for data entry in the spreadsheet.. "
      Top             =   2760
      Width           =   1455
      _Version        =   196608
      _ExtentX        =   2566
      _ExtentY        =   688
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
      Text            =   "2/28/2003"
      DateCalcMethod  =   0
      DateTimeFormat  =   0
      UserDefinedFormat=   ""
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   0
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fpDateDisp 
      Height          =   456
      Left            =   6228
      TabIndex        =   0
      ToolTipText     =   $"frmFADispItemList.frx":0B56
      Top             =   1668
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   804
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
      BackColor       =   16777215
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
      ButtonStyle     =   2
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
      ControlType     =   0
      Text            =   "2/28/2003"
      DateCalcMethod  =   0
      DateTimeFormat  =   0
      UserDefinedFormat=   ""
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   0
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3825
      Left            =   630
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3555
      Width           =   10395
      _Version        =   196613
      _ExtentX        =   18336
      _ExtentY        =   6747
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   50000
      ProcessTab      =   -1  'True
      RowHeaderDisplay=   0
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmFADispItemList.frx":0BDD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDate 
      Height          =   975
      Left            =   9045
      TabIndex        =   9
      ToolTipText     =   "Press this button after entering the desired disposal date in the field to the immediate left."
      Top             =   1488
      Width           =   1500
      _Version        =   131072
      _ExtentX        =   2646
      _ExtentY        =   1720
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
      ButtonDesigner  =   "frmFADispItemList.frx":1BE35
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   675
      Left            =   1710
      TabIndex        =   10
      ToolTipText     =   "Click this button to delete the active items (denoted with X's)."
      Top             =   7770
      Width           =   2565
      _Version        =   131072
      _ExtentX        =   4524
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmFADispItemList.frx":1C01A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   675
      Left            =   5310
      TabIndex        =   11
      Top             =   7770
      Width           =   1650
      _Version        =   131072
      _ExtentX        =   2910
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmFADispItemList.frx":1C1F8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   675
      Left            =   7185
      TabIndex        =   12
      Top             =   7770
      Width           =   1650
      _Version        =   131072
      _ExtentX        =   2910
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmFADispItemList.frx":1C3D4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail 
      Height          =   675
      Left            =   9075
      TabIndex        =   13
      ToolTipText     =   "Click on a row then click this button to bring up further item details."
      Top             =   7770
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmFADispItemList.frx":1C5B0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   11160
      X2              =   480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spreadsheet Active For Disposal Date Of:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   4440
      TabIndex        =   7
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   7470
      X2              =   7470
      Y1              =   2115
      Y2              =   2355
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   9246
      X2              =   7470
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to Activate"
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
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   3000
      Width           =   2550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   3930
      X2              =   3930
      Y1              =   1300
      Y2              =   3360
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6210
      Left            =   495
      Top             =   1290
      Width           =   10680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Currently Saved Dates:"
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
      Left            =   960
      TabIndex        =   5
      Top             =   1344
      Width           =   2556
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Disposal Date:"
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
      Height          =   360
      Left            =   4440
      TabIndex        =   4
      Top             =   1725
      Width           =   1605
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   336
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Asset Disposal List Building"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   3
      Top             =   480
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   288
      Width           =   8652
   End
End
Attribute VB_Name = "frmFADispItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim ItemCnt As Long
  Dim DateOnScreen As Integer
  Dim DatesOKFlag As Boolean
  Dim UserBailsFlag As Boolean
  Dim DetailFlag As Boolean
  
Private Sub cmdClear_Click()
  Dim x As Long
  Dim DoWhatFlag As CompletelyDeleteOption
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Integer
  Dim Y As Integer
  Dim BigNum As Integer
  Dim Nextx
  Dim SmallNum
  Dim StopSpot As Integer
  Dim HoldSpot As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim TotalAccts As Long
  
  On Error GoTo ERRORSTUFF
  DatesOKFlag = True
  UserBailsFlag = False
  
  If Not Exist("FATEMPDISPDATE.DAT") Then
    MsgBox "Error: No disposal dates saved."
    fpDateDisp.SetFocus
    Exit Sub
  End If
  
  OpenTempDisposedDate GHandle
  DateCnt = LOF(GHandle) / Len(DateRec)
  If DateCnt = 0 Then
    MsgBox "Error: No disposal dates saved."
    fpDateDisp.SetFocus
    Close
    Exit Sub
  End If
  
  'this routine removes the chosen date's disposal data
  DoWhatFlag = PromptCompletelyDelete(Me) 'make sure the user understands that
  'clearing means all data is lost
  Select Case DoWhatFlag
    Case CompletelyDeleteOption.cdExit
      Close
      Exit Sub
    Case CompletelyDeleteOption.cdContinue
    MainLog ("User warned that they are completely removing data saved for items to be disposed on " + MakeRegDate(DateOnScreen) + ". The user elected to delete this data in frmFADispItemList.")
    Case Else
  End Select
 
  For x = 1 To DateCnt
    Get GHandle, x, DateRec
    If DateRec.DsplDate = DateOnScreen Then 'look thru the records until the
    'date selected for deleting is found
      DateRec.DsplDate = 0 'found it and changed it to zero and saved
      Put GHandle, x, DateRec
      Exit For 'job is done...no reason to keep looking
    End If
  Next x

  fpListDates.Clear
  
  OpenTagIdxFile TagIdxHandle
  TotalAccts = LOF(TagIdxHandle) \ Len(TagIdx)
  If TotalAccts = 0 Then Exit Sub
  ReDim TagIdxRecs(1 To TotalAccts) As Integer
  For x = 1 To TotalAccts
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load up array with tag record pointers
  Next x
  Close TagIdxHandle
  
  OpenFAItemFile FAHandle
  For x = 1 To TotalAccts
    Get FAHandle, TagIdxRecs(x), FAItemRec 'retrieve data in order of tag numbers
    If FAItemRec.DispDate = DateOnScreen And FAItemRec.DsplFlag <> 2 Then 'only
    'items that have 0 or 1 as .DsplFlag and the disposal date matches the
    'date selected for disposal
      FAItemRec.DsplFlag = 0 'set value back to zero
      FAItemRec.DispDate = 0 'set value back to zero
      Put FAHandle, TagIdxRecs(x), FAItemRec 'save the reset
    End If
  Next x
  Close FAHandle
  
  GoSub LoadDates
  
  If Exist(PrepostDsplName + CStr(Date2Num(fpDateActive)) + ".DAT") Then
    KillFile PrepostDsplName + CStr(Date2Num(fpDateActive)) + ".DAT"
  End If
  
  MsgBox ("All data for " + fpDateDisp.Text + " has been deleted.")
  
  
  Call LoadMe 'Loadme handles the return to the screen
  Exit Sub
  
LoadDates:
  'disposal dates are saved in a separate file...when one is
  'deleted then the file needs to be re-ordered
  ReDim OrderDate(1 To DateCnt) As Integer
  BigNum = 0
  For x = 1 To DateCnt 'must sort the latest dates
    Get GHandle, x, DateRec
    If DateRec.DsplDate = 0 Then GoTo DateDeleted 'move past dates already deleted
    Y = Y + 1
    OrderDate(x) = DateRec.DsplDate
    If DateRec.DsplDate > BigNum Then
      BigNum = DateRec.DsplDate 'find the latest date to use in sorting
    End If
DateDeleted:
  Next x
  Close GHandle
  
  If Y = 0 Then 'all dates have been deleted so we don't need this file
  'anymore
    KillFile ("FATEMPDISPDATE.DAT")
    GoTo NoMoreDates
  End If
  
  Nextx = 1
  BigNum = BigNum + 1
  SmallNum = BigNum
  Do
    For x = Nextx To DateCnt
      If OrderDate(x) < SmallNum Then
        SmallNum = OrderDate(x)
        StopSpot = x
      End If
    Next x
    'sort dates
    HoldSpot = OrderDate(Nextx)
    OrderDate(Nextx) = SmallNum
    OrderDate(StopSpot) = HoldSpot
    If Nextx = DateCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  For x = 1 To DateCnt
    If OrderDate(x) = 0 Then GoTo DateIsZero
    fpListDates.AddItem (MakeRegDate(OrderDate(x))) 'load active dates list
    'with valid disposal dates
DateIsZero:
  Next x
  
NoMoreDates:
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADispItemList", "cmdClear_Click", Erl)
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
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
    Unload Me

End Sub

Private Sub cmdDate_Click()
  
  If Date2Num(fpDateDisp.Text) <> DateOnScreen Then 'if a user switches
  'active dates then anything entered for the current active date
  'that has not been saved will be destroyed when the change is made
    If MsgBox("All unsaved changes will be lost when dates are changed. Do you wish to continue?", vbYesNo) = vbNo Then
      'user can return to the screen here to save changes
      fpDateDisp.Text = MakeRegDate(DateOnScreen)
      Exit Sub
    End If
  ElseIf Date2Num(fpDateDisp.Text) = DateOnScreen Then
    MsgBox "This date is already activated."
    fpDateDisp.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpDateDisp.Text) = "" Then
    fpDateDisp.Text = MakeRegDate(DateOnScreen) 'DateOnScreen is a
    'global that holds the most recent valid active date...here nothing
    'was in the date field so the program reloads the latest date and
    'exits back to the screen
    Exit Sub
  End If
  
  DateOnScreen = Date2Num(fpDateDisp.Text) 'update global
  Close
  
  Call LoadMe

End Sub

Private Sub cmdDetail_Click()
  Dim ThisNum$
  Dim HoldGRecNum As Long
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim x As Long
  Dim FACnt As Long
  
  On Error Resume Next
  'this routine brings up additional data for an item selected
  'in the spreadsheet
  If GRecNum > 0 Then HoldGRecNum = GRecNum
  OpenFAItemFile FAHandle
  FACnt = LOF(FAHandle) / Len(FAItemRec)
  For x = 1 To FACnt
    vaSpread1.Col = 2
    vaSpread1.Row = x
    If vaSpread1.Row = vaSpread1.ActiveRow Then 'find the item in the
    'item records that corresponds with the item selected
      ThisNum$ = QPTrim$(vaSpread1.Text) 'assign ThisNum with the appropriate tag
      'number found in the spreadsheet
      Exit For 'jump out of loop when match is found
    End If
  Next x
  If x = FACnt + 1 Then 'if x = FACnt + 1 then we know we've been thru the
  'all records without finding a match
    MsgBox "No item row has been selected."
    Close FAHandle
    Exit Sub
  End If
    
  vaSpread1.Col = 2
  vaSpread1.Row = x
  For x = 1 To FACnt
    Get FAHandle, x, FAItemRec
    If ThisNum$ = QPTrim$(FAItemRec.ItemTag) Then 'now look thru the
    'item records for the matching tag number
      GRecNum = x 'we matched it...so assign the global GRecNum it's record number
      Close FAHandle
      Exit For
    End If
  Next x
  
  If x > FACnt Then
    MsgBox "Item match failed. Please try again."
    Close FAHandle
    Exit Sub
  End If
    
  frmFAItemDetail.Show vbModal
  
  If HoldGRecNum > 0 Then 'reassign the global value that
  'it had when this routine started
    GRecNum = HoldGRecNum
  Else
    GRecNum = 0
  End If
End Sub

Private Sub cmdExit_Click()
  frmFADisposalMenu.Show
  Close
  DoEvents
  Unload frmFADispItemList
End Sub
Private Sub cmdSave_Click()
  Dim x As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TotalAccts As Long
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim FAItemCnt As Long
  Dim Nextx As Integer
  Dim ThisDate As Integer
  Dim ThisFileName$
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Integer
  Dim Y As Integer
  Dim z As Integer
  Dim Today As Integer
  Dim ActiveX As Long
  Dim PHandle As Integer
  Dim DsplRec As PrePostDsplType
  Dim InPlayCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  
  ActiveX = 0 'accumulates # of X's on spreadsheet below
  Today = Date2Num(fpDateDisp.Text)
  
  ThisDate = Date2Num(fpDateActive.Text) 'date on screen is marker to this saved data
  If Today <> ThisDate Then
    fpDateActive.BackColor = &HC0FFFF
    fpDateDisp.BackColor = &HC0C0FF
    MsgBox "The activation date(yellow) is not the same as the date displayed(red). Please activate the date displayed."
    fpDateActive.BackColor = &HFFFFFF
    fpDateDisp.BackColor = &HFFFFFF
    Exit Sub
  End If
  
  OpenTagIdxFile TagIdxHandle
  TotalAccts = LOF(TagIdxHandle) \ Len(TagIdx)
  
  If TotalAccts = 0 Then
    Close
    Exit Sub 'no tag data on file
  End If
  
  ReDim TagIdxRecs(1 To TotalAccts) As Integer
  
  For x = 1 To TotalAccts 'get sorted list of tag numbers
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load up array with record pointers
  Next x
  Close TagIdxHandle
  
  OpenFAItemFile FAHandle
  For x = 1 To TotalAccts 'if this disposal date comes before the last time depreciation
  'was processed then the town's depreciation will have been based on an incorrect value...
  'in essence the town could have received the benefit of this item's depreciation value
  'when they didn't own it
     Get FAHandle, x, FAItemRec
     If FAItemRec.CDEPDATE > Today Then
       Exit For
     End If
  Next x
  
  If x <= TotalAccts Then 'for loop above was stopped before it got to the end
  'of the records
    DoEvents
    frmFADsplMess.Label1.Caption = "A depreciation processing date (" + MakeRegDate(FAItemRec.CDEPDATE) + ") comes after the disposal date entered (" + fpDateDisp + "):"
    DoEvents
    frmFADsplMess.Label2.Top = 1500
    frmFADsplMess.Label2.Caption = "1. Fixed assets should not be depreciated after they are disposed of."
    frmFADsplMess.Label3.Top = 3000
    frmFADsplMess.Label3.Caption = "2. Continuing at this point would record assets that are no longer owned as being depreciated."
    frmFADsplMess.Show vbModal
    If frmFADsplMess.fptxtChoice.Text = "abort" Then
      Unload frmFADsplMess
      Close
      Exit Sub
    Else
      Unload frmFADsplMess
      MainLog ("User warned that the disposal date (" + MakeRegDate(Today) + ") is before a later depreciation date (" + MakeRegDate(FAItemRec.CDEPDATE) + ") and saved the depreciation list anyway in frmFADispItemList.")
    End If
  End If
  
  Nextx = 1
  For x = 1 To TotalAccts
    Get FAHandle, TagIdxRecs(x), FAItemRec
    If FAItemRec.DsplFlag = 2 Then GoTo NotAnItem 'dump disposed of items
    
    'whatever X is in the first column gets saved
    vaSpread1.Col = 1
    vaSpread1.Row = Nextx
    If vaSpread1.Text = "X" Then 'user wanted to add this one or it was resaved
    'as selected for disposal
      FAItemRec.DsplFlag = 1
      FAItemRec.DispDate = ThisDate
      ActiveX = ActiveX + 1
    ElseIf FAItemRec.DispDate = ThisDate Then 'this item
    'was on the disposal list because the disposal date is recorded but
    'now the "X" is gone so the user wanted to delist this one
      FAItemRec.DsplFlag = 0
      FAItemRec.DispDate = 0
      FAItemRec.DisposAmt = 0
    End If
    Put FAHandle, TagIdxRecs(x), FAItemRec
    Nextx = Nextx + 1
NotAnItem:
  Next x
  
  'Prepost.. is created when items slated for disposal are
  'assigned methods and disposal prices in the edit screen
  Nextx = 0
  If Exist(PrepostDsplName + CStr(Date2Num(fpDateActive)) + ".DAT") Then
    OpenPrePostDsplData PHandle, Date2Num(fpDateActive)
    InPlayCnt = LOF(PHandle) / Len(DsplRec)
    GoSub CheckInPlay 'sets delete flag to true if an item
    'that has been saved in the disposal edit process is now
    'no longer selected for disposal
  End If
  
  OpenTempDisposedDate GHandle 'contains files of disposed of items
  'by date of their disposal
  DateCnt = LOF(GHandle) / Len(DateRec)
  
  If ActiveX = 0 Then 'used when existing valid dates have had all their
  'items deselected
    For x = 1 To DateCnt 'now update the disposal date records
      Get GHandle, x, DateRec
      If DateRec.DsplDate = ThisDate Then 'no items selected for this date since ActiveX is 0
        MsgBox "No items selected for disposal for this date. " + MakeRegDate(ThisDate) + " has been deleted."
        Close PHandle 'then delete the corresponding edit file
        If Exist(PrepostDsplName + CStr(ThisDate) + ".DAT") Then '8/4/03 also kill any edited file for this date
          KillFile PrepostDsplName + CStr(ThisDate) + ".DAT"
        End If
        'OK...the date with no items selected is gone...now look to see if the
        'overall dates file is still needed
        If DateCnt = 1 Then 'only 1 was on file and we just deleted it's contents so we can
        'now trash the overall dates file
          fpListDates.Clear
          Close
          KillFile (TempDispDateName) 'zero out file because the last date is no longer valid
          Exit Sub
        Else
          DateRec.DsplDate = 0 'multiple valid dates saved (file not killed) so zero out this one and save it
          Put GHandle, x, DateRec 'now saved as deleted
          fpListDates.Clear 'clear the date list and start over
          For z = 1 To DateCnt 'look to see if any of the remaining dates are valid
            Get GHandle, z, DateRec
            If DateRec.DsplDate > 0 Then 'found a valid date so jump out of loop
              Exit For
            End If
          Next z
          If z > DateCnt Then 'no valid dates found (all were 0'd out) so delete file
            Close
            KillFile (TempDispDateName)
            If Exist(PrepostDsplName + CStr(ThisDate) + ".DAT") Then '8/4/03 also kill any edited file for this date
              KillFile PrepostDsplName + CStr(ThisDate) + ".DAT"
            End If
            fpListDates.AddItem "NONE" 'reset list
            Exit Sub 'jump back to screen
          End If
          For z = 1 To DateCnt 'reload list with valid dates
            Get GHandle, z, DateRec
            If DateRec.DsplDate = 0 Then GoTo ZeroDate 'could be some dates embedded that
            'have been deleted mixed with some that are still valid
            fpListDates.AddItem MakeRegDate(DateRec.DsplDate)
ZeroDate:
          Next z
        End If
      End If
    Next x
    'if we get to this point then there are no X's on the spreadsheet and this
    'date was inactive to start with...so nothing to save
    Close
    Exit Sub 'all done...jump back to screen
  End If
  
  If DateCnt = 0 Then 'new file created and saved
    DateRec.DsplDate = Date2Num(fpDateDisp.Text)
    Put GHandle, 1, DateRec 'save as first record because this date is the only one
    Close GHandle
    GoTo NoDateHere
  Else
    For x = 1 To DateCnt
      Get GHandle, x, DateRec
      If DateRec.DsplDate = ThisDate Then 'match date on screen with one on file
        Exit For
      End If
    Next x
  End If
    
  If x > DateCnt Then 'if no date match was found then this is a new active date
    DateRec.DsplDate = Date2Num(fpDateDisp.Text)
    Put GHandle, DateCnt + 1, DateRec 'add to end of list
  Else
    GoSub AddNewItemToExistingEditFile 'we have existing files so check to
    'see if the edit file for this date needs updating either by removing
    'an item that has been delisted or adding an item (with no disposal price or
    'disposal method) to the existing list
  End If
NoDateHere:
  
  Close
  
  MsgBox ("Your temporary item disposal data has been saved for " + MakeRegDate(DateOnScreen) + ". Posting disposal data commits it to memory permanently.")
  
  frmFADisposalMenu.Show
  DoEvents
  Unload frmFADispItemList
StayHere:
  
  Exit Sub
  
CheckInPlay:
  For Y = 1 To InPlayCnt 'looking to see if any of the items
  'that had been edited before are still on this list
  'and if not then delete by setting the deleted flag to true
  '
  'first select each item (Y) on the edit list
    Get PHandle, Y, DsplRec
    For x = 1 To TotalAccts
      'if an item is on this list then it has to have a disposal date set
      'and a disposal flag set to 1 (pending disposal)
      Get FAHandle, TagIdxRecs(x), FAItemRec
        If TagIdxRecs(x) = DsplRec.ThisRec Then 'if a match is found
          If FAItemRec.DispDate = ThisDate And FAItemRec.DsplFlag = 1 Then
          'check to see if it is still valid and if so then move to the next Y
            Exit For
          End If
        End If
    Next x
    If x > TotalAccts Then 'we've been through the whole item list and couldn't
    'find a match...so this one has just been delisted
      DsplRec.Deleted = True 'purge any item that used to have an X but now does not
      Put PHandle, Y, DsplRec
    End If
  Next Y
    
  Return
  
AddNewItemToExistingEditFile: 'done if an edit file exists already for this
'date and the user has added additional items afterwards that are not on the
'list but should be
  Nextx = InPlayCnt + 1
  If Exist(PrepostDsplName + CStr(ThisDate) + ".DAT") Then
    For Y = 1 To TotalAccts
      Get FAHandle, TagIdxRecs(Y), FAItemRec 'get each item record and find any that
      'qualify to be on this date's edit list
        If FAItemRec.DispDate = ThisDate And FAItemRec.DsplFlag = 1 Then
          For x = 1 To InPlayCnt 'now get each edit record to compare against this item
            Get PHandle, x, DsplRec 'go thru edit records saved for this date
            If DsplRec.Deleted = True Then GoTo ThisOneIsGone 'this edit record was already deleted
            'but the .ThisRec still holds a value
            If TagIdxRecs(Y) = DsplRec.ThisRec Then 'already saved in existing edit records
              Exit For 'this one still qualifies so move to the next item (Y)
            End If
ThisOneIsGone:
          Next x
        End If
        If x = InPlayCnt + 1 Then 'been thru all edit records and this
        'particular item wan't on the list although it qualifies to be there...
        'so add it to the list now
          DsplRec.Deleted = False
          DsplRec.DisposAmt = 0
          DsplRec.DsplMethod = ""
          DsplRec.ThisRec = TagIdxRecs(Y)
          Put PHandle, Nextx, DsplRec
          Nextx = Nextx + 1
        End If
        x = 0
    Next Y
  End If
  Close PHandle
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADispItemList", "cmdSave_Click", Erl)
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
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
    Unload Me
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call FixSpread
  fpDateDisp.Text = Date 'loading takes place when this box changes
  DatesOKFlag = True
  UserBailsFlag = False
  Call LoadMe

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%I"
      Call cmdDetail_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      Call cmdClear_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%A"
      Call cmdDate_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFADispItemList.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim FAItemRec As FAItemRecType
  Dim THandle As Integer
  Dim TotalAccts As Long
  Dim x As Integer
  Dim Nextx As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim ThisDate As Integer
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Integer
  Dim Y As Integer
  Dim BigNum As Integer
  Dim SmallNum As Integer
  Dim StopSpot As Integer
  Dim HoldSpot As Integer
  Dim Flag2Cnt As Long
  Dim ActiveDates As Integer
  
'  On Error GoTo ERRORSTUFF
  cmdClear.Visible = False
  If Exist("FATEMPDISPDATE.DAT") Then 'sometimes this file can exist
  'but all dates are zeros
    OpenTempDisposedDate GHandle
    DateCnt = LOF(GHandle) / Len(DateRec) 'even if the dates are zero they are
    GoSub LoadDates
    'still a valid record number
    For x = 1 To DateCnt
      Get GHandle, x, DateRec
      If DateRec.DsplDate = 0 Then
        ActiveDates = ActiveDates + 1 'counts all zeroed out dates
        GoTo NODate
      End If
      'to reduce confusion the delete command button will display the
      'currently active disposal date if this date has valid items
      'saved for disposal on that date
      If Date2Num(fpDateDisp.Text) = DateRec.DsplDate Then
        cmdClear.Visible = True
        cmdClear.Text = "F3 Delete " + fpDateDisp.Text
        Exit For
      End If
NODate:
    Next x
'    'all saved disposal dates
'    'have been examined and the active date field does not hold a date that
'    'is saved as a disposal date
    If ActiveDates = DateCnt Then 'every date saved was zeroed out...this file is
    'now useless
      Close GHandle
      KillFile ("FATEMPDISPDATE.DAT")
    End If
  Else
    fpListDates.Clear
    fpListDates.AddItem ("NONE")
  End If
  
  
  Close GHandle
  
  ThisDate = Date2Num(fpDateDisp.Text)
  DateOnScreen = ThisDate 'set global DateOnScreen
   
  OpenTagIdxFile TagIdxHandle
  TotalAccts = LOF(TagIdxHandle) \ Len(TagIdx)
  If TotalAccts = 0 Then
    Close
    Exit Sub
  End If
  
  ItemCnt = TotalAccts
  ReDim TagIdxRecs(1 To TotalAccts) As Integer
  
  For x = 1 To TotalAccts
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load array with item records
    'in numerical order
  Next x
  Close TagIdxHandle
  
  If Not Exist("FAITEMS.DAT") Then 'jump out of this sub if
  'no item data is saved
    MsgBox "Path to FAITEMS.DAT could not be found"
    Close
    Exit Sub
  End If

  OpenFAItemFile THandle
  
  Nextx = 1
  For x = 1 To TotalAccts 'load spreadsheet
    Get THandle, TagIdxRecs(x), FAItemRec
      If FAItemRec.DsplFlag = 2 Then 'this item has been disposed of
        Flag2Cnt = Flag2Cnt + 1 'count number of disposed of items
        GoTo ItemDisposed 'dump anything that has been disposed
      End If
      vaSpread1.Col = 1
      vaSpread1.Row = Nextx
      If FAItemRec.DispDate = ThisDate And FAItemRec.DsplFlag = 1 Then 'disposal pending assets
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.Text = "X"
        vaSpread1.Lock = False 'we want to be able to edit this column
      ElseIf FAItemRec.DsplFlag = 1 Then 'pending disposal but not for this date
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.Text = MakeRegDate(FAItemRec.DispDate) 'show date to be disposed
        vaSpread1.Lock = True
      Else
        vaSpread1.Text = "" 'pertains to items as yet unselected
        vaSpread1.Lock = False 'editable
      End If
      
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.Col = 2
      vaSpread1.Row = Nextx
      vaSpread1.Text = FAItemRec.ItemTag
      vaSpread1.Col = 3
      vaSpread1.Row = Nextx
      vaSpread1.Text = FAItemRec.IDESC1
      vaSpread1.Col = 4
      vaSpread1.Row = Nextx
      vaSpread1.Text = FAItemRec.IDEPT
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.Col = 5
      vaSpread1.Row = Nextx
      vaSpread1.Text = Using("$###,###,##0.00", FAItemRec.ORGCOST)
      vaSpread1.Col = 6
      vaSpread1.Row = Nextx
      vaSpread1.Text = Using("$###,###,##0.00", FAItemRec.CURRVAL)
      Nextx = Nextx + 1
ItemDisposed:
  Next x
  Close THandle
  vaSpread1.MaxRows = TotalAccts - Flag2Cnt 'limits spreadsheet size to
  '1 row for each active item and no empty rows
  
  fpDateActive = MakeRegDate(DateOnScreen)
  Exit Sub
   
LoadDates:
  
  ReDim OrderDate(1 To DateCnt) As Integer 'DateCnt assigned above
  BigNum = 0
  For x = 1 To DateCnt
    Get GHandle, x, DateRec
    If DateRec.DsplDate = 0 Then GoTo DateDeleted
    Y = Y + 1 'count valid dates
    OrderDate(x) = DateRec.DsplDate
    If DateRec.DsplDate > BigNum Then
      BigNum = DateRec.DsplDate 'end up with BigNum holding
      'the latest date
    End If
DateDeleted:
  Next x
  
  If Y = 0 Then
    fpListDates.AddItem ("NONE")
    Return
  End If
  
  Nextx = 1
  'now sort the dates
  BigNum = BigNum + 1
  SmallNum = BigNum
  Do
    For x = Nextx To DateCnt
      If OrderDate(x) < SmallNum Then
        SmallNum = OrderDate(x)
        StopSpot = x
      End If
    Next x
    HoldSpot = OrderDate(Nextx)
    OrderDate(Nextx) = SmallNum
    OrderDate(StopSpot) = HoldSpot
    If Nextx = DateCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  fpListDates.Clear
  
  For x = 1 To DateCnt
    If OrderDate(x) = 0 Then GoTo DateIsZero 'shouldn't be any zeros
    fpListDates.AddItem (MakeRegDate(OrderDate(x))) 'reload the active date list
    'with updated dates
DateIsZero:
  Next x
  
NoMoreDates:
  Return
    
ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADispItemList", "LoadMe", Erl)
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
  '--- Cleanup code goes here...
  ClearInUse (PWcnt)
  Terminate
  Close
  Unload Me
End Sub

Private Sub fpDateDisp_Change()
  DatesOKFlag = True
  UserBailsFlag = False
End Sub

Private Sub fpListDates_Click()
  Dim ThisDate As Integer
  
  If fpListDates.Text = "NONE" Or QPTrim$(fpListDates.Text) = "" Then GoTo NODate 'no reason to
  'continue if there are no dates to get data for
  
  If Date2Num(fpListDates.Text) <> DateOnScreen Then 'if a user switches
  'active dates then anything entered for the current active date
  'that has not been saved will be destroyed when the change is made
    If MsgBox("All unsaved changes will be lost when dates are changed. Do you wish to continue?", vbYesNo) = vbNo Then
      'user can return to the screen here to save changes
      Exit Sub
    End If
  End If
  
  ThisDate = Date2Num(fpListDates.Text)
  
  If ThisDate <> DateOnScreen Then 'update globals
    DateOnScreen = Date2Num(fpListDates.Text)
    fpDateDisp.Text = MakeRegDate(ThisDate)
    Call LoadMe
  End If
NODate:
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    
  On Error Resume Next
  If DateOnScreen <> Date2Num(fpDateDisp.Text) Then 'user entered a date in the disposal date field
  'but did not activate that date using the activate button
    MsgBox "Please activate the date " + fpDateDisp.Text + " by pressing the Activate Date button or pressing F12."
    Close
    fpDateDisp.SetFocus
    Exit Sub
  End If
  
  If DatesOKFlag = False Then GoTo AlreadyChecked
  UserBailsFlag = False
  If Check4ValidDates = False Then
    If UserBailsFlag = True Then
      DatesOKFlag = True
      Exit Sub
    Else
      DatesOKFlag = False
    End If
  End If
  
AlreadyChecked:
  
  'click on a row and an X will appear in the far left column...
  'click it again and the X disappears
  vaSpread1.Col = 1
  vaSpread1.Row = Row
  If vaSpread1.Lock = True Then Exit Sub
  If vaSpread1.Text = "X" Then
    vaSpread1.Text = " "
  Else
    vaSpread1.Text = "X"
  End If
End Sub

'Private Sub PrintText()
'  Dim ReportFile$
'  Dim DateRec As TempDisposedOfDate
'  Dim DHandle As Integer
'  Dim x As Long, FF$
'  Dim DateCnt As Long
'  Dim ThisDate As Integer
'  Dim RptHandle As Integer
'  Dim MaxLines As Integer
'  Dim LineCnt As Integer
'  Dim FASetUpRec As FASetupRecType
'  Dim FASHandle As Integer
'  Dim FAHandle As Integer
'  Dim FAItemRec As FAItemRecType
'  Dim Employer$, Page As Integer
'  Dim Nextx As Integer
'  Dim TagRec As TagNumbSortIdxType
'  Dim TagHandle As Integer
'  Dim TagCnt As Long
'
'  If QPTrim$(fpDateDisp.Text) = "" Then
'    MsgBox "Please make sure a valid date is entered in the Disposal Date field."
'    fpDateDisp.SetFocus
'    Exit Sub
'  End If
'
'  ThisDate = Date2Num(fpDateDisp.Text)
'  If Exist("FATEMPDISPDATE.DAT") Then
'    OpenTempDisposedDate DHandle
'    DateCnt = LOF(DHandle) / Len(DateRec)
'    If DateCnt = 0 Then
'      MsgBox "No disposal dates have been saved."
'      fpDateDisp.SetFocus
'      Close DHandle
'      Exit Sub
'    End If
'  End If
'
'  For x = 1 To DateCnt
'    Get DHandle, x, DateRec
'    If DateRec.DsplDate = ThisDate Then 'validate the date entered
'      Close DHandle
'      Exit For
'    End If
'  Next x
'
'  'x will be greater than DateCnt only if no match was found in the
'  'for loop above
'  If x > DateCnt Then
'    MsgBox "Nothing is saved for this date. Please choose another date or save data for this date."
'    fpDateDisp.SetFocus
'    Close DHandle
'    Exit Sub
'  End If
'
'  OpenFASetUpFile FASHandle
'  Get FASHandle, 1, FASetUpRec
'  Close FASHandle
'  FF$ = Chr$(12)
'
'  Employer = QPTrim$(FASetUpRec.TownName)
'
'  MaxLines = 57
'  ReportFile$ = "FATEMPDSPLPRINT.PRT"
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
'
'  GoSub PrintHeader
'
'  frmFAShowPctComp.Label1 = "Gathering Disposed Of Item Data"
'  frmFAShowPctComp.Show
'  DoEvents
'  EnableCloseButton Me.hwnd, False
'  Me.cmdExit.Enabled = False
'  Me.cmdClear.Enabled = False
'  Me.cmdSave.Enabled = False
'  Me.cmdPrint.Enabled = False
'
'  OpenTagIdxFile TagHandle
'  TagCnt = LOF(TagHandle) / Len(TagRec)
'  ReDim TagRecNum(1 To TagCnt)
'  For x = 1 To TagCnt
'    Get TagHandle, x, TagRec
'    TagRecNum(x) = TagRec.DataRecNum 'load array with record pointers arranged by
'    'tag numerical order
'  Next x
'  Close TagHandle
'
'  OpenFAItemFile FAHandle
'
'  For x = 1 To TagCnt
'    Get FAHandle, TagRecNum(x), FAItemRec
'    If FAItemRec.DispDate = ThisDate And FAItemRec.DsplFlag = 1 Then
'      Print #RptHandle, FAItemRec.ItemTag; Tab(21); FAItemRec.IDESC1; Tab(53); FAItemRec.IDEPT; Tab(58); Using$("$##,###,##0.00", FAItemRec.ORGCOST); Tab(73); Using$("$##,###,##0.00", FAItemRec.CURRVAL)
'      Print #RptHandle,
'      Print #RptHandle, "Disposal Amount __________________ "; Tab(41); "Method: AUCTION __ SALVAGE __ SOLD__ OTHER __"
'      Print #RptHandle, String$(86, "-")
''      Print #RptHandle,
'      LineCnt = LineCnt + 4
'    End If
'    If LineCnt >= MaxLines Then
'      Print #RptHandle, FF$
'      GoSub PrintHeader
'    End If
'
'    frmFAShowPctComp.ShowPctComp x, TagCnt
'    If frmFAShowPctComp.Out = True Then
'      Close
'      frmFAShowPctComp.Out = False
'      EnableCloseButton Me.hwnd, True
'      Me.cmdExit.Enabled = True
'      Me.cmdClear.Enabled = True
'      Me.cmdSave.Enabled = True
'      Me.cmdPrint.Enabled = True
'      Unload frmFAShowPctComp
'      Exit Sub
'    End If
'  Next x
'
'  EnableCloseButton Me.hwnd, True
'  Me.cmdExit.Enabled = True
'  Me.cmdClear.Enabled = True
'  Me.cmdSave.Enabled = True
'  Me.cmdPrint.Enabled = True
'  Unload frmFAShowPctComp
'  Print #RptHandle, FF$
'  Close RptHandle
'  ViewPrint ReportFile$, "Master Asset Disposed Of Listing", False
'  KillFile (ReportFile$)
'  Exit Sub
'
'PrintHeader:
'  Page = Page + 1
'  Print #RptHandle, Tab(27); "Fixed Asset List of Items For Disposal"
'  Print #RptHandle,
'  Print #RptHandle, Employer; Tab(77); "Page "; Tab(83); Page
'  Print #RptHandle, "Item Disposal Date: "; Tab(22); fpDateDisp.Text
'  Print #RptHandle,
'  Print #RptHandle, Tab(1); "Tag Number"; Tab(25); "Description"; Tab(53); "Dept"; Tab(58); "Purchase Price"; Tab(74); "Current Value"
'  Print #RptHandle, String$(86, "=")
'  LineCnt = 7
'
'  Return
'End Sub
'
'Private Sub PrintGraphics()
'  Dim ReportFile$
'  Dim DateRec As TempDisposedOfDate
'  Dim DHandle As Integer
'  Dim x As Long
'  Dim DateCnt As Long
'  Dim ThisDate As Integer
'  Dim RptHandle As Integer
'  Dim FASetUpRec As FASetupRecType
'  Dim FASHandle As Integer
'  Dim FAHandle As Integer
'  Dim FAItemRec As FAItemRecType
'  Dim Employer$, Page As Integer
'  Dim Nextx As Integer
'  Dim TagRec As TagNumbSortIdxType
'  Dim TagHandle As Integer
'  Dim TagCnt As Long
'  Dim dlm$
'
'  dlm$ = "~"
'  If QPTrim$(fpDateDisp.Text) = "" Then
'    MsgBox "Please enter a valid date in the Disposal Date field."
'    fpDateDisp.SetFocus
'    Exit Sub
'  End If
'
'  ThisDate = Date2Num(fpDateDisp.Text)
'  If Exist("FATEMPDISPDATE.DAT") Then
'    OpenTempDisposedDate DHandle
'    DateCnt = LOF(DHandle) / Len(DateRec)
'    If DateCnt = 0 Then 'nothing saved
'      MsgBox "No disposal dates have been saved."
'      fpDateDisp.SetFocus
'      Close DHandle
'      Exit Sub
'    End If
'  End If
'
'  For x = 1 To DateCnt
'    Get DHandle, x, DateRec 'look for the selected disposal date in the records
'    If DateRec.DsplDate = ThisDate Then
'      Close DHandle
'      Exit For
'    End If
'  Next x
'
'  If x > DateCnt Then 'x will be greater than DateCnt if no match
'  'was found in the for loop above
'    MsgBox "Nothing is saved for this date. Please choose another date or save data for this date."
'    fpDateDisp.SetFocus
'    Close DHandle
'    Exit Sub
'  End If
'
'  OpenFASetUpFile FASHandle
'  Get FASHandle, 1, FASetUpRec
'  Close FASHandle
'
'  Employer = QPTrim$(FASetUpRec.TownName)
'
'  ReportFile$ = "FARPTS\FATEMPDSPLPRINT.RPT"
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
'
'  frmFAShowPctComp.Label1 = "Gathering Disposed Of Item Data"
'  frmFAShowPctComp.Show
'  DoEvents
'  EnableCloseButton Me.hwnd, False
'  Me.cmdExit.Enabled = False
'  Me.cmdClear.Enabled = False
'  Me.cmdSave.Enabled = False
'  Me.cmdPrint.Enabled = False
'
'  OpenTagIdxFile TagHandle
'  TagCnt = LOF(TagHandle) / Len(TagRec)
'  ReDim TagRecNum(1 To TagCnt)
'  For x = 1 To TagCnt
'    Get TagHandle, x, TagRec
'    TagRecNum(x) = TagRec.DataRecNum 'load an array with record pointers
'    'arranged in numerical order
'  Next x
'  Close TagHandle
'
'  OpenFAItemFile FAHandle
'
'  For x = 1 To TagCnt
'    Get FAHandle, TagRecNum(x), FAItemRec
'    If FAItemRec.DispDate = ThisDate And FAItemRec.DsplFlag = 1 Then
'      '                     0                   1                      2                     3
'      Print #RptHandle, Employer; dlm; FAItemRec.ItemTag; dlm; FAItemRec.IDESC1; dlm; FAItemRec.IDEPT; dlm;
'      '                         4                       5                       6
'      Print #RptHandle, FAItemRec.ORGCOST; dlm; FAItemRec.CURRVAL; dlm; fpDateDisp.Text; dlm;
'      '                                    7                                             8
'      Print #RptHandle, "Disposal Amount __________________ "; dlm; "Method: AUCTION __ SALVAGE __ SOLD__ OTHER __"
'    End If
'
'    frmFAShowPctComp.ShowPctComp x, TagCnt
'    If frmFAShowPctComp.Out = True Then
'      Close
'      frmFAShowPctComp.Out = False
'      EnableCloseButton Me.hwnd, True
'      Me.cmdExit.Enabled = True
'      Me.cmdClear.Enabled = True
'      Me.cmdSave.Enabled = True
'      Me.cmdPrint.Enabled = True
'      Unload frmFAShowPctComp
'      Exit Sub
'    End If
'  Next x
'
'  EnableCloseButton Me.hwnd, True
'  Me.cmdExit.Enabled = True
'  Me.cmdClear.Enabled = True
'  Me.cmdSave.Enabled = True
'  Me.cmdPrint.Enabled = True
'  Unload frmFAShowPctComp
'
'  Close RptHandle
'
'  arFAItemsForDsplList.Show
'  frmFALoadReport.Show
'
'  Exit Sub
'
'End Sub

Private Sub FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim cnt As Integer
  '-1 means all rows or all columns....0 means headers
'    GoTo SkipAdjust
    Select Case ScreenW
      Case 1280
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 5
          coladj = 10
          vaSpread1.FontSize = 18
          vaSpread1.RowHeight(-1) = 22
          vaSpread1.RowHeight(0) = 22
        Else
          COne = 13
          coladj = 4.5
          vaSpread1.RowHeight(-1) = 18
          vaSpread1.RowHeight(0) = 18
        End If
      Case 1152
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 14
          coladj = 7
          vaSpread1.FontSize = 14
          vaSpread1.RowHeight(0) = 18.5
          vaSpread1.RowHeight(-1) = 18.5
        Else
          COne = 6.65
          coladj = 2.25
          vaSpread1.RowHeight(0) = 16
          vaSpread1.RowHeight(-1) = 17
        End If
      Case 1024
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 13.49
          coladj = 5.65
          vaSpread1.RowHeight(0) = 14
          vaSpread1.RowHeight(-1) = 14
        Else
          COne = 1.2
          coladj = 0 '.35
        End If
      Case 800
        COne = 0
        coladj = -0.6
        vaSpread1.Font.Size = 10
        vaSpread1.RowHeight(-1) = 14
      Case Else
    End Select
SkipAdjust:
    vaSpread1.ColWidth(1) = vaSpread1.ColWidth(1)
    vaSpread1.ColWidth(2) = vaSpread1.ColWidth(2) + coladj
    vaSpread1.ColWidth(3) = vaSpread1.ColWidth(3) + coladj
    vaSpread1.ColWidth(4) = vaSpread1.ColWidth(4)
    vaSpread1.ColWidth(5) = vaSpread1.ColWidth(5) + coladj
    vaSpread1.ColWidth(6) = vaSpread1.ColWidth(6) + coladj

End Sub

Private Function Check4ValidDates() As Boolean
  Dim x As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TotalAccts As Long
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim FAItemCnt As Long
  Dim Nextx As Integer
  Dim ThisDate As Integer
  Dim DateCnt As Integer
  Dim ActiveDate As Integer
  Dim DepFile As Integer
  Dim NumOfDprRecs As Integer
  Dim FADep(1) As FADepFileType
  
  On Error GoTo ERRORSTUFF
  ActiveDate = Date2Num(fpDateActive)
  
  Check4ValidDates = True
  
  DepFile = FreeFile
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  NumOfDprRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  If NumOfDprRecs > 0 Then
    Close
    'if there are temporary depreciation files then there is the potential for
    'inaccurate data if fixed asset is set for disposal...then depreciated (tax implications)...then disposed after depreciation...
    'in effect the disposed of fixed asset was depreciated when it was no longer owned
    If Date2Num(fpDateActive.Text) <= FADep(1).DprDay Then
      DoEvents
      frmFADsplMess.Label1.Height = 1200
      frmFADsplMess.Label1.Top = 500
      frmFADsplMess.Label1.Caption = "A pending depreciation date, " + MakeRegDate(FADep(1).DprDay) + ", is scheduled for a date later than the disposal date entered, " + fpDateActive.Text + ":"
      DoEvents
      frmFADsplMess.Label2.Height = 1500
      frmFADsplMess.Label2.Top = 1600
      frmFADsplMess.Label2.Caption = "Assets depreciated after " + fpDateActive.Text + " should not be disposed of on " + fpDateActive.Text + ". Please make the disposal date later than " + MakeRegDate(FADep(1).DprDay) + " or delete the pending depreciation."
      DoEvents
      frmFADsplMess.Label3.Height = 1500
      frmFADsplMess.Label3.Top = 2900
      frmFADsplMess.Label3.Caption = "Disposing of a fixed asset on a date that comes before a posted (or pending) depreciation date for that item indicates that the depreciation tax benefit was taken on that item which was no longer in inventory."
      frmFADsplMess.Show vbModal
      If frmFADsplMess.fptxtChoice.Text = "abort" Then
        Check4ValidDates = False
        UserBailsFlag = True
        Exit Function
      Else
        MainLog ("User warned that this disposal list should not be scheduled for a disposal on " + fpDateActive.Text + " when there is a depreciation date pending for " + MakeRegDate(FADep(1).DprDay) + ". They continued with the save procedure anyway.")
      End If
    End If
  End If
  Close
  
  OpenTagIdxFile TagIdxHandle
  TotalAccts = LOF(TagIdxHandle) \ Len(TagIdx)
  
  If TotalAccts = 0 Then
    Close
    Exit Function 'no tag data on file
  End If
  
  ReDim TagIdxRecs(1 To TotalAccts) As Integer
  
  For x = 1 To TotalAccts 'get sorted list of tag numbers
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load up array with record pointers
  Next x
  Close TagIdxHandle
  
  OpenFAItemFile FAHandle
  For x = 1 To TotalAccts 'if this disposal date comes before the last time depreciation
  'was processed then the town's depreciation will have been based on an incorrect value...
  'in essence the town could have received the benefit of this item's depreciation value
  'when they didn't own it
     Get FAHandle, x, FAItemRec
     If FAItemRec.CDEPDATE > ActiveDate Then
       Exit For
     End If
  Next x
  
  If x <= TotalAccts Then 'for loop above was stopped before it got to the end
  'of the records
    Check4ValidDates = False
    frmFADsplMess.Label1.Top = 700
    frmFADsplMess.Label1.Caption = "A depreciation processing date (" + MakeRegDate(FAItemRec.CDEPDATE) + ") comes after the disposal date entered (" + fpDateActive.Text + "):"
    DoEvents
    frmFADsplMess.Label2.Top = 1900
    frmFADsplMess.Label2.Height = 1500
    'caveat - a fixed asset that was recorded as being purchased before the last depreciation date
    'and then disposed of before the last depreciation date would not be affected ...chances of this happening - almost 0%
    frmFADsplMess.Label2.Caption = "1. Under most circumstances a fixed asset should not be disposed of on any date that comes before a depreciation date."
    frmFADsplMess.Label3.Top = 3300
    frmFADsplMess.Label3.Caption = "2. Continuing at this point would record assets that are no longer owned as being depreciated."
    frmFADsplMess.Show vbModal
    If frmFADsplMess.fptxtChoice.Text = "abort" Then
      Close
      Unload frmFADsplMess
      UserBailsFlag = True
      Exit Function
    Else
      Unload frmFADsplMess
      MainLog ("User warned that the disposal date (" + MakeRegDate(ActiveDate) + ") is before a later depreciation date (" + MakeRegDate(FAItemRec.CDEPDATE) + ") and continued building depreciation list anyway in frmFADispItemList.")
    End If
  End If
  Close

  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADisplItemList", "Check4ValidDates", Erl)
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
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
    Unload Me
  
End Function

