VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFABuildYrEndDep 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year End Processing"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFABuildYrEndDep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText fptxtLastDprDate 
      Height          =   396
      Left            =   5184
      TabIndex        =   3
      Top             =   5796
      Width           =   1356
      _Version        =   196608
      _ExtentX        =   2392
      _ExtentY        =   698
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
      ForeColor       =   8421504
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
   Begin EditLib.fpDateTime fpDateLastPurch 
      Height          =   444
      Left            =   4944
      TabIndex        =   2
      ToolTipText     =   $"frmFABuildYrEndDep.frx":08CA
      Top             =   4596
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   783
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
      Text            =   "1/24/2003"
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
   Begin EditLib.fpText fpDateLastYear 
      Height          =   396
      Left            =   7020
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Read only field showing the last valid year depreciation took place. N/A means no prior depreciation has been saved."
      Top             =   2580
      Width           =   1116
      _Version        =   196608
      _ExtentX        =   1968
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      InvalidColor    =   8454143
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
      CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
      MaxLength       =   4
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fpDateCurrYear 
      Height          =   396
      Left            =   7008
      TabIndex        =   1
      ToolTipText     =   "Key in the year desired for this depreciation processing."
      Top             =   3204
      Width           =   1116
      _Version        =   196608
      _ExtentX        =   1968
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      InvalidColor    =   8454143
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
      Text            =   ""
      CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
      MaxLength       =   4
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   3489
      TabIndex        =   10
      Top             =   6924
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmFABuildYrEndDep.frx":0960
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   690
      Left            =   6279
      TabIndex        =   9
      ToolTipText     =   "Click this button to begin the depreciation process."
      Top             =   6924
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmFABuildYrEndDep.frx":0B3C
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   1632
      X2              =   10032
      Y1              =   5316
      Y2              =   5316
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asset Purchase Date Entered for Most Recent Posted Depreciation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2304
      TabIndex        =   8
      Top             =   5460
      Width           =   7164
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Year Depreciation Posted:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   3564
      TabIndex        =   7
      Top             =   2676
      Width           =   3324
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Latest Date of Asset Purchase to INCLUDE in the Processing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   588
      Left            =   3708
      TabIndex        =   6
      Top             =   3828
      Width           =   4236
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Current Depreciation Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   3300
      TabIndex        =   5
      Top             =   3300
      Width           =   3564
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUILDING YEAR END DEPRECIATION FILE"
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
      Left            =   2292
      TabIndex        =   4
      Top             =   1080
      Width           =   7068
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   924
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   4044
      Left            =   1632
      Top             =   2340
      Width           =   8412
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   876
      Width           =   8652
   End
End
Attribute VB_Name = "frmFABuildYrEndDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim SoSoftFlag As Boolean

Private Sub cmdExit_Click()
  frmFAYearEndMenu.Show
  Close
  DoEvents
  Unload frmFABuildYrEndDep
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
      SendKeys "%P"
      Call cmdProcess_Click
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
      ClearInUse PWcnt 'resets password data to zero
      MainLog ("FixedAssets.exe terminated via menu bar on frmFABuildYrEndDep.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub LoadMe()
  Dim YearHandle As Integer
  Dim FAYear As FAYearEndType
  Dim YearSize As Integer
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Long
  Dim x As Long
  Dim BigDate As Integer
  Dim NewYear As Integer
  
  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  If NumOfFARecs = 0 Then GoTo NoFARecs
  BigDate = 0
  For x = 1 To NumOfFARecs
    Get FAHandle, x, FAItemRec
    If FAItemRec.CDEPDATE > BigDate Then
      BigDate = FAItemRec.CDEPDATE
    End If
  Next x
  
  fptxtLastDprDate = MakeRegDate(BigDate)
  
NoFARecs:
  NewYear = Val(Mid(fptxtLastDprDate.Text, 7, 4))
  NewYear = NewYear + 1
  SoSoftFlag = False
  fpDateCurrYear = NewYear
  fpDateLastPurch = Mid(fptxtLastDprDate.Text, 1, 6) + "/" + CStr(NewYear)
  OpenYearFile YearHandle 'this file maintains the current depreciation
  'years
  YearSize = LOF(YearHandle) \ Len(FAYear)
  If YearSize = 0 Then
    fpDateLastYear.Text = "N/A" 'the first depreciation begins with
    'N/A in the last depreciation year's field
  Else
    Get YearHandle, 1, FAYear
    fpDateLastYear.Text = FAYear.LastYear 'the last depreciation year
    'value
  End If
  Close YearHandle
  
End Sub

Private Sub cmdProcess_Click()
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim NumOfDepts As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim FASetup As FASetupRecType
  Dim SetupHandle As Integer
  Dim SetUpSize As Integer
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim YearHandle As Integer
  Dim FAYear As FAYearEndType
  Dim YearSize As Integer
  Dim CurYear$
  Dim ProcessThru$, x As Integer
  Dim cnt&
  Dim PSDate As Integer
  Dim PEDate As Integer
  Dim DepFile As Integer
  Dim Nextx As Integer
  Dim IAqurDate As Integer
  Dim Day2Dep As Integer
  Dim DepPerDay#
  Dim MaxDep#, CurDep#
  Dim NextEditRecord!
  Dim FADep(1) As FADepFileType
  Dim DeprType$
  Dim UseThisPct As Double
  Dim ThisType$
  Dim HistCnt As Long
  Dim DprHistRec As DprHistType
  Dim HHandle As Integer
  Dim DoWhatFlag As NotSubsequentOption
  Dim BigDate As Integer
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Integer
  Dim TextLen As Integer
  Dim LastYear As Integer
  Dim FromHandle As Integer
  Dim One As Integer
  
  On Error GoTo ERRORSTUFF
  TextLen = Len(QPTrim$(fpDateCurrYear.Text))
  If TextLen <> 4 Or Mid(fpDateCurrYear.Text, 1, 2) <> "20" Then
    MsgBox "Please enter a valid four digit current year."
    fpDateCurrYear.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpDateCurrYear.Text) = "" Then
    MsgBox "Please enter a date for the current depreciation year."
    fpDateCurrYear.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpDateLastPurch.Text) = "" Then
    MsgBox "Please enter a date for the latest asset purchase."
    fpDateLastPurch.SetFocus
  End If
  
  OpenTempDisposedDate GHandle
  DateCnt = LOF(GHandle) / Len(DateRec)
  If DateCnt = 0 Then GoTo NoDisposal
  For x = 1 To DateCnt
    Get GHandle, x, DateRec
      If DateRec.DsplDate <> 0 And DateRec.DsplDate <= Date2Num(fpDateLastPurch.Text) Then
        DoEvents
        frmFADprMess.Label1.Height = 1200
        frmFADprMess.Label2.Top = 600
        frmFADprMess.Label1.Caption = "A fixed asset disposal date (" + MakeRegDate(DateRec.DsplDate) + ") is scheduled on or before this depreciation date (" + fpDateLastPurch.Text + "). This disposal activity should be finalized before continuing:"
        DoEvents
        frmFADprMess.Label2.Height = 1500
        frmFADprMess.Label2.Top = 1600
        frmFADprMess.Label2.Caption = "Assets scheduled for disposal before " + fpDateLastPurch.Text + " should not be depreciated at this a time. Please finalize the items scheduled for disposal on " + MakeRegDate(DateRec.DsplDate) + " before continuing here."
        DoEvents
        frmFADprMess.Label3.Height = 1500
        frmFADprMess.Label3.Top = 3000
        frmFADprMess.Label3.Caption = "If a fixed asset is disposed of on a date that comes before it's last depreciation date then the depreciation benefit was received on a fixed asset that was not in inventory."
        DoEvents
        frmFADprMess.Show vbModal
        If frmFADprMess.fptxtChoice.Text = "abort" Then
          Unload frmFADprMess
          Close
          Exit Sub
        Else
          Unload frmFADprMess
          MainLog ("User warned that a pending disposal date (" + MakeRegDate(DateRec.DsplDate) + ") is scheduled on or before this depreciation date (" + fpDateLastPurch.Text + ") and elected to continue with the depreciation build anyway in frmFABuildYrEndDep.")
          Exit For
        End If
      End If
  Next x
NoDisposal:

  If Exist(DprHistFileName) Then 'look for the file holding all
  'depreciation records
    If Check4SubsequentYear = False Then 'CheckForSubsequentYear examines
    'the depreciation record and determines if the year entered for depreciation
    'now is the next valid year
      If SoSoftFlag = True Then 'user wanted to abort processing because of
      'warning issued in Check4SubsequentYear
        Close
        Exit Sub
      End If
      DoEvents
      frmFADprMess.Label1.Caption = "The year entered for depreciation is not the subsequent year from the last year depreciated:"
      DoEvents
      frmFADprMess.Label2.Top = 1500
      frmFADprMess.Label2.Caption = "1. Skipping a depreciation year will invalidate the year skipped from being depreciated in the future."
      DoEvents
      frmFADprMess.Label3.Top = 3000
      frmFADprMess.Label3.Caption = "2. Skipping a depreciation year eliminates the tax benefit that would be realized if assets were depreciated for that year. "
      frmFADprMess.Show vbModal
      If frmFADprMess.fptxtChoice.Text = "abort" Then
        Unload frmFADprMess
        fpDateCurrYear.SetFocus
        Close
        Exit Sub
      Else
        Unload frmFADprMess
        MainLog ("User issued warning that the year to be depreciated was not subsequent to the last depreciation year. User elected to run the depreciation anyway in frmFABuildYrEndDep.")
      End If
    End If
  End If
  
  If Val(fpDateCurrYear.Text) < 1950 Or Val(fpDateCurrYear.Text) > 2100 Then 'designed to catch
  'an inadvertent entry
    MsgBox "Please enter a year between 1949 and 2101 for Current Depreciation Year"
    fpDateCurrYear.SetFocus
    Close
    Exit Sub
  End If
  
  OpenFASetUpFile SetupHandle
  
  SetUpSize = LOF(SetupHandle) \ Len(FASetup)
  
  If SetUpSize > 0 Then 'if setup has been saved then look for what type
  'of depreciation to use...here a file exists but there is nothing in it
    Get SetupHandle, 1, FASetup
    ThisType = QPTrim$(FASetup.DeprType)
    If ThisType = "Prorate 1st year" Then
      DeprType = 1
    ElseIf ThisType = "Fixed 1st year percentage" Then
      DeprType = 2
      UseThisPct = FASetup.Pct1St * 0.01
    ElseIf ThisType = "Whole year" Then
      DeprType = 3
    Else 'a depreciation type is required so if nothing has been saved then
    'the user is either given the opportunity to jump to the setup screen
    'or this depreciation defaults to whole year
      If MsgBox("A depreciation method has not been saved. Would you like to jump to the fixed asset depreciation set up screen?", vbYesNo) = vbYes Then
        FromHandle = FreeFile
        Open "fromBuildDep.dat" For Output As FromHandle Len = 2
        One = 1
        Print #FromHandle, One
        Close
        frmFASystemSetup.Show
        DoEvents
'        Unload Me
        Exit Sub
      Else
        MsgBox "Depreciation type is defaulting to Whole Year."
        DeprType = 3
      End If
    End If
  Else 'here no file exists
    If MsgBox("A depreciation method has not been saved. Would you like to jump to the fixed asset depreciation set up screen?", vbYesNo) = vbYes Then
      FromHandle = FreeFile
      Open "fromBuildDep.dat" For Output As FromHandle Len = 2
      One = 1
      Print #FromHandle, One
      Close
      frmFASystemSetup.Show
      DoEvents
      frmFABuildYrEndDep.Hide
      Exit Sub
    Else
      MsgBox "Depreciation type is defaulting to Whole Year."
      DeprType = 3
    End If
  End If
  
  Close SetupHandle

  If Len(QPTrim$(fpDateCurrYear.Text)) = 4 Then 'check first to make sure the year
  'entered for depreciation is a valid for digit year
    CurYear$ = QPTrim$(fpDateCurrYear.Text)
    OpenDprHistFile HHandle
    HistCnt = LOF(HHandle) / Len(DprHistRec)
    If HistCnt > 0 Then
      For x = 1 To HistCnt
      Get HHandle, x, DprHistRec 'look to see if this year has been processed already...
      'if it has then stop the process now...if it has and the SoSoft flag is set to true then
      'we know this depreciation might be a redo so we overlook this screen
        If CurYear$ = QPTrim$(DprHistRec.DprYear) And DprHistRec.SoSoftFlag = False Then
          MsgBox "Depreciation for this year has already taken place."
          fpDateCurrYear.SetFocus
          Close
          Exit Sub
        End If
      Next x
      Close HHandle
    End If
    
    ProcessThru$ = QPTrim$(fpDateLastPurch.Text) 'important assignment
    GoSub ProcessDepreciation 'all clear so go ahead and depreciate
  Else
    MsgBox "Please enter a valid date for the current year."
    fpDateLastPurch.SetFocus
    Close
    Exit Sub
  End If
    
  Close
  
  If NextEditRecord! <> 0 Then
    MsgBox "Depreciation file building completed for " + CurYear$ + "."
    MainLog ("Depreciation built for " + CurYear$ + " in frmFABuildYrEndDep.")
  Else
    MsgBox "No assets qualify for depreciation for this period."
    MainLog ("Depreciation for year " + CurYear$ + " attempted but failed because no fixed assets qualify for depreciation for this year in frmFABuildYrEndDep.")
  End If
  
  frmFAYearEndMenu.Show
  DoEvents
  Unload frmFABuildYrEndDep
  Exit Sub
  
ProcessDepreciation:
  If PWUser = "Sosoft Support" And SoSoftFlag = True Then GoTo SoSoftSignIn
  If Right$(ProcessThru$, 4) <> CurYear Then 'the correct way to process depreciation
  'is to make the process thru date in the same year as the depreciation year.
    DoEvents
    frmFADprMess.Label1.Caption = "If the last purchase date entered does not fall into the same year as the year being depreciated then:"
    DoEvents
    frmFADprMess.Label2.Caption = "1. Assets purchased after the last purchase date will not be depreciated."
    DoEvents
    frmFADprMess.Label3.Caption = "2. Items purchased in the 365 days prior to the purchase date entered will have their life left calculation reset to the same as if they were purchased the year being depreciated. "
    DoEvents
    If ThisType = "Prorate 1st year" Then
      frmFADprMess.Label4.Caption = "3. The Prorate 1st year depreciation method will cause the depreciation amount to be less than the amount intended for any item purchased within 365 days prior to the last purchase date."
      DoEvents
    ElseIf ThisType = "Fixed 1st year percentage" Then
      frmFADprMess.Label4.Caption = "3. The Fixed 1st year percentage depreciation method will cause the depreciation amount to be equal to the percent intended for use for any item purchased within 365 days prior to the last purchase date."
      DoEvents
    End If
    frmFADprMess.Show vbModal
    If frmFADprMess.fptxtChoice.Text = "abort" Then
      Unload frmFADprMess
      Close
      fpDateLastPurch.SetFocus
      Exit Sub
    Else
      Unload frmFADprMess
      MainLog ("User warned in building year end depreciation that the process thru date (" + ProcessThru$ + ") is in a different year than the depreciation date entered by the user (" + CurYear$ + "). The user elected to continue anyway in frmFABuildYrEndDep.")
    End If
  End If
  
SoSoftSignIn:
  PSDate = Date2Num(ProcessThru$)
  PSDate = PSDate - 365 'used to determine if an item was
  'purchased within the year prior to the depreciation date
  PEDate = Date2Num(ProcessThru$) '
  'Build Depreciation File
  'Build Work File
  DepFile = FreeFile
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  Close DepFile
  KillFile ("FADPREDT.DAT") 'destroy any existing temporary depreciation file before
  'building a new one...only one allowed
  'Open Deprec Edit File
  DepFile = FreeFile
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  'Open Item File
    
  OpenDeptIdxFile DIdxHandle

  DIdxRecNums = LOF(DIdxHandle) / Len(DeptIdx)
  
  If DIdxRecNums = 0 Then
    Close
    MsgBox "No departments have been indexed. Please save at least one department code."
    Exit Sub
  End If
  
  ReDim DIdx(1 To DIdxRecNums) As Integer
  For x = 1 To DIdxRecNums
    Get DIdxHandle, x, DeptIdx
    DIdx(x) = CInt(QPTrim$(DeptIdx.DeptNumb)) 'load array with pointers to records
  Next x
  Close DIdxHandle
  
  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  If NumOfFARecs = 0 Then
    Close
    MsgBox "No fixed assets have been saved. Depreciation cannot take place."
    Exit Sub
  End If
  BigDate = 0 'if the user attempts to run a depreciation before one full year
  'since the last depreciation he will get a warning flag
  For x = 1 To NumOfFARecs
    Get FAHandle, x, FAItemRec
    If FAItemRec.CDEPDATE > BigDate Then
      BigDate = FAItemRec.CDEPDATE
    End If
  Next x
  
  If SoSoftFlag = True Then GoTo SoSoft
  
  If BigDate > -11001 Then '-11001 is the value saved when a new asset is saved...
  'it flags the asset as a depreciation that has not yet been saved...so if
  'after all asset records have been scoured and the last date found is the same as
  'what a new fixed asset gets then we know depreciation has not taken place so we
  'don't need to check any further
    If Date2Num(ProcessThru$) - BigDate < 365 Then 'determines if the requested last purchase date for this depreciation
    'is less than a year from the last depreciation date...366 accommodates leap year
        DoEvents
        frmFADprMess.Label1.Height = 1000
        frmFADprMess.Label1.Caption = "The latest asset purchase date, " + QPTrim$(fpDateLastPurch.Text) + ", and the last depreciation date, " + MakeRegDate(BigDate) + " are not ordered one year apart. This could cause unexpected depreciation results:"
        DoEvents
        frmFADprMess.Label2.Height = 1000
        frmFADprMess.Label2.Top = 1700
        frmFADprMess.Label2.Caption = "1. It is best to maintain a consistent annual depreciation date so that all depreciation amounts are calculated in an equal time frame."
        DoEvents
        If ThisType = "Prorate 1st year" Then
          frmFADprMess.Label3.Top = 3000
          frmFADprMess.Label3.Caption = "2. Assets purchased during the 365 days prior to the last purchase date will not have the same depreciation amount as other assets purchased under the exact same conditions in prior years."
          DoEvents
        End If
        frmFADprMess.Show vbModal
        If frmFADprMess.fptxtChoice.Text = "abort" Then
        Unload frmFADprMess
        Close
        Exit Sub
      Else
        Unload frmFADprMess
        MainLog ("User warned that the date entered for the last purchase date for this depreciation, " + QPTrim$(fpDateLastPurch.Text) + ", is less than one full year since the last depreciation, " + MakeRegDate(BigDate) + " and elected to continue in frmFABuildYrEndDep.")
      End If
    End If
  End If
  
SoSoft:
  frmFAShowPctComp.Label1 = "Building Year End Depreciation"
  frmFAShowPctComp.Show
  frmFAShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  Nextx = 1
  Do
    For cnt& = 1 To NumOfFARecs 'move through fixed assets records
      Get FAHandle, cnt&, FAItemRec
'      If QPTrim$(FAItemRec.ItemTag) = "10-3" Then Stop
      If FAItemRec.IDEPT <> DIdx(Nextx) Or FAItemRec.AQURDATE > PEDate Or FAItemRec.CURRVAL = 0 Then 'any asset that
      'does not match the current department we are wanting (this report is ordered by tag number then by dept num),
      'or it was purchased after the last purchase date or the asset has been depreciated down to zero (or disposed of),
      'will not be depreciated.
        GoTo SkipThisAsset
      End If
      'Calc Depreciation for This Period
      If FAItemRec.ILIFE > 0 And FAItemRec.DEPYN <> "N" Then
      'if this fixed asset has an original life greater than zero and has
      'not been flagged as "do not depreciate" then go ahead and
      'depreciate this asset
        CurDep# = OldRound#(FAItemRec.ORGCOST / FAItemRec.ILIFE) 'CurDep#
        'is assigned what would be a normal whole year's depreciation for this asset
        If DeprType = 2 Then 'this type is the first year percentage type
          IAqurDate = FAItemRec.AQURDATE
          If IAqurDate >= PSDate And IAqurDate <= PEDate Then 'screen out any assets that
          'were not purchased within the year prior to the last purchase date
            CurDep# = OldRound#(FAItemRec.ORGCOST * UseThisPct) 'since this asset
            'will not be depreciated at the whole year rate we need to establish the
            'depreciation amount for a partial year...this asset simply is depreciated by
            'the percentage entered on the set up screen times the original purchase price
            'to arrive at the first year depreciation
            FADep(1).PctFlag = True
          Else
            FADep(1).PctFlag = False
          End If
        ElseIf DeprType = 1 Then 'this type is based on when during the year this asset
        'was purchased and the percentage of the year left after its purchase date is
        'applied to the normal whole year depreciation amount
          IAqurDate = FAItemRec.AQURDATE
          If IAqurDate >= PSDate And IAqurDate <= PEDate Then 'first check to see if
          'this asset was purchase within the year prior to the last purchase date
            Day2Dep = PEDate - IAqurDate 'number of days of the last year this asset has been
            'considered a valid fixed asset
            DepPerDay# = OldRound(CurDep# / 365) 'figure one day's worth of depreciation at
            'whole year rate
            CurDep# = OldRound#(DepPerDay# * Day2Dep) 'now figure the depreciation for this asset
            FADep(1).PctFlag = True
          Else
            FADep(1).PctFlag = False
          End If
        End If
  
        MaxDep# = OldRound#(FAItemRec.ORGCOST - FAItemRec.DEP2DATE) 'figure how much value
        'is left in this asset...a value over which it cannot be depreciated
        If OldRound#(MaxDep# + FAItemRec.DEP2DATE) > FAItemRec.ORGCOST Then 'this depreciation amount
        'exceeds the max allowable
          MaxDep# = OldRound#(MaxDep# - (OldRound#(MaxDep# + FAItemRec.DEP2DATE) - FAItemRec.ORGCOST))
        End If
        If MaxDep# < 0 Then MaxDep# = 0 'can't have a negative depreciation
        If CurDep# > MaxDep# Then CurDep# = MaxDep# 'can't over depreciate
        If Abs(CurDep# - MaxDep#) <= 0.04 Then 'filter out rounding shortcomings
          FADep(1).AssetRecord = cnt&
          FADep(1).CurYrDep = OldRound#(MaxDep#)
          FADep(1).CurrYear = fpDateCurrYear.Text
          FADep(1).DprDay = Date2Num(fpDateLastPurch.Text)
          NextEditRecord! = NextEditRecord! + 1
          Put DepFile, NextEditRecord!, FADep(1) 'save this depreciation
          'in this temporary depreciation file
        ElseIf OldRound#(CurDep#) >= 0.01 Then 'anything less than a penny is
        'not valid
          FADep(1).AssetRecord = cnt&
          FADep(1).CurYrDep = OldRound#(CurDep#)
          FADep(1).CurrYear = fpDateCurrYear.Text
          FADep(1).DprDay = Date2Num(fpDateLastPurch.Text)
          NextEditRecord! = NextEditRecord! + 1
          Put DepFile, NextEditRecord!, FADep(1)
        End If
      Else
        GoTo SkipThisAsset
      End If
      GoTo MoveOn
SkipThisAsset:
      'added on 9/22/2004 to allow all assets to appear on the Pre-
      'Posting report
      If FAItemRec.IDEPT <> DIdx(Nextx) Then GoTo MoveOn
      If FAItemRec.DispDate > 0 And FAItemRec.DsplFlag <> 1 Then GoTo MoveOn
        FADep(1).PctFlag = False
        FADep(1).AssetRecord = cnt&
        FADep(1).CurYrDep = 0
        FADep(1).CurrYear = fpDateCurrYear.Text
        FADep(1).DprDay = Date2Num(fpDateLastPurch.Text)
        NextEditRecord! = NextEditRecord! + 1
        Put DepFile, NextEditRecord!, FADep(1) 'save this depreciation
'      End If
MoveOn:
    Next cnt&
    If Nextx = DIdxRecNums Then Exit Do
    frmFAShowPctComp.ShowPctComp Nextx, DIdxRecNums
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    Nextx = Nextx + 1
  Loop
  
ExitRpt:

  frmFAShowPctComp.Out = False
  Unload frmFAShowPctComp
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True

  Close
  
  Return

ERRORSTUFF:
   Unload frmFAShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFABuildYrEndDep", gstrcProgName, Erl)
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
End Sub

Private Function Check4SubsequentYear() As Boolean
  Dim x As Long
  Dim DprRec As DprHistType
  Dim DprCnt As Long
  Dim DHHandle As Integer
  Dim LastYear As Integer
  Dim ThisYear As Integer
  
  On Error GoTo ERRORSTUFF
  Check4SubsequentYear = True
  
  ThisYear = CInt(fpDateCurrYear)
  
  OpenDprHistFile DHHandle
  DprCnt = LOF(DHHandle) / Len(DprRec)
  
  If DprCnt = 0 Then
    Close
    Exit Function 'exit with function = true
  End If
  
  For x = 1 To DprCnt
    Get DHHandle, x, DprRec
    If DprRec.SoSoftFlag = True Then
      SoSoftFlag = True
      If PWUser = "Sosoft Support" Then
        fpDateCurrYear.BackColor = &H80FFFF
        fpDateLastPurch.BackColor = &H80FFFF
        If MsgBox("Message to Southern Software Support: Make sure your dates (move this screen and look for 2 yellow fields) are set correctly (the year should be the same year as the year for which the reversal took place) and go ahead and depreciate this 'after reversal' procedure. OK to continue?", vbYesNo) = vbNo Then
          Check4SubsequentYear = False
        Else
          Check4SubsequentYear = True
        End If
        fpDateCurrYear.BackColor = &HFFFFFF
        fpDateLastPurch.BackColor = &HFFFFFF
        Close
        Exit Function
      End If
        
      frmFADprMess.Label2.Caption = "The depreciation procedure for " + DprRec.DprYear + " was a depreciation reversal. It is recommended that you call Southern Software at 1-800-842-8190 for assistance before continuing. Press ESC to abort this depreciation."
      DoEvents
      frmFADprMess.Label2.Height = 1500
      frmFADprMess.Label2.Top = 1500
      frmFADprMess.Show vbModal
      If frmFADprMess.fptxtChoice = "abort" Then
        Close
        Check4SubsequentYear = False 'when Check4SubsequentYear is false and
        'SoSoftFlag is true then the cmd_Process will see it and abort the process
        Exit Function
      Else 'Ok we're going to continue
        MainLog ("The user warned that there was a pending depreciation reversal for year " + DprRec.DprYear + ". The user elected to continue anyway in frmFABuildYrEndDep.")
        If DprRec.DprYear <> ThisYear Then 'year entered is not the same one as the
        'year currently oending after a reversal
          frmFADprMess.Label2.Caption = "The year selected for depreciation (" + CStr(ThisYear) + ") is not the same year as the pending depreciation year reversed by Southern Software (" + DprRec.DprYear + "). This is not recommended. Press ESC to abort this depreciation."
          DoEvents
          frmFADprMess.Label2.Height = 1500
          frmFADprMess.Label2.Top = 1500
          frmFADprMess.Show vbModal
          If frmFADprMess.fptxtChoice = "abort" Then
            Close
            Check4SubsequentYear = False
            Exit Function
          End If
          MainLog ("With SoSoftFlag = True the user warned that the year entered for depreciation (" + CStr(ThisYear) + ") was not the same year as the pending reversal year (" + DprRec.DprYear + "). The user elected to continue anyway in frmFABuildYrEndDep.")
        End If
        Exit For 'else continue with processing depreciation
      End If
    End If
  Next x
  
  'if SoSoftFlag = True then the next for loop will be true because SoSoftFlag
  'is only set with the most recent year which means that the previous year will
  'have been saved and this year will be subsequent to it...
  'the function will exit with Check4SubsequentYear = true
  If SoSoftFlag = True Then Exit Function
  For x = 1 To DprCnt
    Get DHHandle, x, DprRec
    LastYear = CInt(DprRec.DprYear)
    If ThisYear - LastYear = 1 Then 'if any depreciation file exists that makes this true then
    'we know this year is the next subsequent year so exit with function = true unless SoSoftFlag = True
      Close
      Exit Function
    End If
  Next x
  
  Check4SubsequentYear = False 'if Check4SubsequentYear is false then cmd_Process will give the user
  'a warning and options as to what they can do next
  Close DHHandle
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFABuildYrEndDep", "Check4SubsequentYear", Erl)
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

Private Sub fpDateLastPurch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpDateLastYear.SetFocus
  End If
End Sub
