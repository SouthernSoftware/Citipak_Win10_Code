VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmEmpORBITProcessing 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retirement/ORBIT Processing"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   Icon            =   "frmEmpORBITProcessing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   23810.71
   ScaleMode       =   0  'User
   ScaleWidth      =   28905.92
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText fptxtFileDest 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   5520
      Width           =   4815
      _Version        =   196608
      _ExtentX        =   8493
      _ExtentY        =   661
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   6405
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   6360
      Width           =   2175
      _Version        =   131072
      _ExtentX        =   3836
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
      ButtonDesigner  =   "frmEmpORBITProcessing.frx":08CA
   End
   Begin EditLib.fpDateTime fpdtRptPeriod 
      Height          =   375
      Left            =   6300
      TabIndex        =   0
      ToolTipText     =   "This is the time period during which pay checks were written."
      Top             =   3000
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
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
      AlignTextH      =   1
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
      Text            =   "05/2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19200101"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   3045
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   6360
      Width           =   2175
      _Version        =   131072
      _ExtentX        =   3836
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
      ButtonDesigner  =   "frmEmpORBITProcessing.frx":0AA6
   End
   Begin EditLib.fpDateTime fpdtPayPdBegin 
      Height          =   375
      Left            =   6300
      TabIndex        =   1
      Top             =   3840
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   0   'False
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
      AlignTextH      =   1
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
      Text            =   "10/01/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19200101"
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
   Begin EditLib.fpDateTime fpdtPayPdEnd 
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Top             =   4680
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   0   'False
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
      AlignTextH      =   1
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
      Text            =   "10/01/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19200101"
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "NOTE: If you are RE-Processing then any edited data will be removed."
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
      Height          =   375
      Left            =   1358
      TabIndex        =   12
      Top             =   7920
      Width           =   8895
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Pay Period Begin Date:"
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
      Height          =   375
      Left            =   3540
      TabIndex        =   11
      Top             =   3915
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Pay Period End Date:"
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
      Height          =   255
      Left            =   3540
      TabIndex        =   10
      Top             =   4785
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4935
      Left            =   1200
      Top             =   2400
      Width           =   9255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "File Name and Location:"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Report Period:"
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
      Height          =   375
      Left            =   3540
      TabIndex        =   6
      Top             =   3090
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4058
      TabIndex        =   5
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   855
      Index           =   1
      Left            =   1478
      Top             =   360
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retirement/ORBIT Processing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2925
      TabIndex        =   3
      Top             =   600
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1478
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmEmpORBITProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim AgencyNum As String
Public MaxSalary As Double

Private Sub cmdProcess_Click()
  Dim EndDate As Integer
  Dim BegDate As Integer
  Dim OHRec As OrbitHeader
  Dim OHHandle As Integer
  Dim Answer As VbMsgBoxResult
  
  If Look4MaxSalary = True Then
    Answer = MsgBox("Employees are approaching the maximum salary level for NC ORBIT. Press 'Yes' to review these employees.", vbYesNo)
    If Answer = vbYes Then
      frmEmpORBITMaxSalaryList.Show vbModal
      Exit Sub
    End If
  End If
  
  If AgencyNum = "" Then
    MsgBox ("An agency number is a requirement for this report. Please enter and save the agency number on the Employer Setup screen.")
    Exit Sub
  End If
  
  If fpdtPayPdBegin.Text = "" Then
    MsgBox ("Please enter a valid pay period begin date")
    fpdtPayPdBegin.SetFocus
    Exit Sub
  End If
  
  If fpdtPayPdEnd.Text = "" Then
    MsgBox ("Please enter a valid pay period end date")
    fpdtPayPdEnd.SetFocus
    Exit Sub
  End If
  
  BegDate = Date2Num(fpdtPayPdBegin.Text)
  EndDate = Date2Num(fpdtPayPdEnd.Text)
  If Abs(BegDate - EndDate) > 31 Then
    MsgBox ("The pay period cannot exceed 31 days.")
    fpdtPayPdBegin.SetFocus
    Exit Sub
  End If
    
  If EndDate < BegDate Then
    MsgBox ("The pay period end date comes before the pay period begin date. Please correct this situation.")
    fpdtPayPdBegin.SetFocus
    Exit Sub
  End If
  
  If MakeFile4Submission = True Then
    MsgBox ("Processing has completed successfully.")
  Else
    KillFile OrbitDetail
    MsgBox ("No employees qualify for ORBIT processing in the time period entered.")
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpORBITProcessing.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmORBITMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub LoadMe()
  Dim OHeaderRec As OrbitHeader
  Dim OHHandle As Integer
  Dim ThisYr As String
  Dim ThisMn As String
  Dim ThisDate As String
  Dim RptDate As String
  
  ThisDate = Date
  Call MakeRptDate(ThisDate, ThisMn, ThisYr)
  If Exist(OrbitHeader) Then
    OpenOrbHeader OHHandle
    Get OHHandle, 1, OHeaderRec
    Label3.Caption = "AGENCY NUMBER: " & QPTrim$(OHeaderRec.AgencyNum)
    AgencyNum = QPTrim$(OHeaderRec.AgencyNum)
    MaxSalary = OHeaderRec.MaxSalary
    If Exist(OrbitDetail) Then
      If OHeaderRec.RptPeriod <> "" Then
        RptDate = OHeaderRec.RptPeriod
        fpdtRptPeriod.Text = Mid(RptDate, 5, 2) & "/" & Mid(RptDate, 1, 4)
      End If
      If OHeaderRec.PayPrdBeginDate <> 0 Then
        fpdtPayPdBegin.Text = MakeRegDate(OHeaderRec.PayPrdBeginDate)
      Else
        fpdtPayPdBegin.Text = Date
      End If
      If OHeaderRec.PayPrdEndDate <> 0 Then
        fpdtPayPdEnd.Text = MakeRegDate(OHeaderRec.PayPrdEndDate)
      Else
        fpdtPayPdEnd.Text = Date
      End If
    Else
      fpdtRptPeriod.Text = ThisDate
      fpdtPayPdBegin.Text = Date
      fpdtPayPdEnd.Text = Date
    End If
  Else
    fpdtRptPeriod.Text = ThisDate
    fpdtPayPdBegin.Text = Date
    fpdtPayPdEnd.Text = Date
  End If
'  fptxtFileDest.Text = "NCORBIT\" & ReplaceString(ThisDate, "/", "") & AgencyNum & ".CNT"
  fptxtFileDest.Text = "NCORBIT\" & RptDate & AgencyNum & ".CNT" 'added 10/23/07
  Close OHHandle
  
End Sub

Private Function MakeRptDate(ByRef ThisDate As String, ByRef ThisMn As String, ByRef ThisYr As String) As String
 Dim x As Integer
 Dim y As Integer
 Dim ch As String
 
 For x = 1 To Len(ThisDate)
   ch = Mid(ThisDate, x, 1)
   If ch <> "/" Then
     ThisMn = ThisMn + ch
     Exit For
   End If
 Next x
  
 For y = Len(ThisDate) - 3 To Len(ThisDate)
   ch = Mid(ThisDate, y, 1)
   ThisYr = ThisYr & ch
 Next y
  
 If Len(ThisMn) = 1 Then
   ThisMn = "0" & ThisMn
 End If
 
 ThisDate = ThisMn & "/" & ThisYr

End Function

Private Sub fpdtRptPeriod_LostFocus()
  Dim Month As String
  Dim Year As String
  Dim RptDate As String
  If InStr(fpdtRptPeriod, "/") = 2 Then
    fpdtRptPeriod.Text = "0" & fpdtRptPeriod.Text
  End If
  RptDate = Mid(fpdtRptPeriod.Text, 4, 7) 'added 10/23/07
  RptDate = RptDate + Mid(fpdtRptPeriod.Text, 1, 2) 'added 10/23/07
'  fptxtFileDest.Text = "NCORBIT\" & ReplaceString(fpdtRptPeriod.Text, "/", "") & AgencyNum & ".CNT"
  fptxtFileDest.Text = "NCORBIT\" & RptDate & AgencyNum & ".CNT" 'added 10/23/07
  
  Month = Mid(fpdtRptPeriod.Text, 1, 2)
  Year = Mid(fpdtRptPeriod.Text, 4, 7)
  fpdtPayPdBegin.Text = Month & "/" & "01/" & Year
  Select Case CInt(Month)
    Case 1, 3, 5, 7, 8, 10, 12
      fpdtPayPdEnd.Text = Month & "/" & "31/" & Year
    Case 4, 6, 9, 11
      fpdtPayPdEnd.Text = Month & "/" & "30/" & Year
    Case 2
      Select Case CInt(Year)
        Case 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036
          fpdtPayPdEnd.Text = Month & "/" & "29/" & Year
        Case Else
          fpdtPayPdEnd.Text = Month & "/" & "28/" & Year
      End Select
    Case Else
  End Select

End Sub

Function RoundDbl#(DblNum#)
  RoundDbl# = (Int((DblNum# * 100) + 0.5) / 100)
End Function

Private Function MakeFile4Submission() As Boolean
  Dim HiDate As Integer
  Dim LowDate As Integer
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfERecs As Integer
  Dim x As Integer
  Dim IdxRec As NumbSortIdxType
  Dim IdxHandle As Integer
  Dim TransRec As TransRecType
  Dim THandle As Integer
  Dim RetRec As RetRecType
  Dim RetRecNew As RetRecType
  Dim RHandle As Integer
  Dim RptName$, TransRecNum&
  Dim RecNo As Long, NumOfRecs As Integer
  Dim OHRec As OrbitHeader
  Dim ODRec As OrbitDetail
  Dim OTRec As OrbitTrailer
  Dim OERec As OrbitEmpData
  Dim OHHandle As Integer
  Dim ODHandle As Integer
  Dim OTHandle As Integer
  Dim OEHandle As Integer
  Dim NumOfOERecs As Integer
  Dim NumOfODRecs As Integer
  Dim NextRec As Long
  Dim ThisODRec As Long
  Dim y As Long, TotSalary#, TotEmpyMatch#, TotEmprMatch#
  Dim PayPdBegDate As String
  Dim PayPdEndDate As String
  Dim RptPeriod As String
  Dim EndDate As Integer
  Dim BegDate As Integer
  Dim ODCnt As Integer
  'from here down to UseOT are variables added on 10/1/07 to allow for
  'a separate transaction for OT Pay
  Dim Dif As Double
  Dim OTPct As Double
  Dim RegPct As Double
  Dim OTAndReg As Double
  Dim OTRetGrossPay As Double
  Dim OTRetireAmt As Double
  Dim OTMatchRetAmt As Double
  Dim RegRetGrossPay As Double
  Dim RegRetireAmt As Double
  Dim RegMatchRetAmt As Double
  Dim TotRetGrossPay As Double
  Dim TotRetireAmt As Double
  Dim TotMatchRetAmt As Double
  Dim RetHandle As Integer
  Dim RetireRec As RetireRecType
  Dim NumOfRetRecs As Integer
  Dim UseOT As Boolean
  
  MakeFile4Submission = False
  
  RptPeriod = FormatThisPayPd(fpdtRptPeriod.Text, 1)
  
  EndDate = Date2Num(fpdtPayPdEnd.Text)
  BegDate = Date2Num(fpdtPayPdBegin.Text)
  OpenOrbHeader OHHandle
  OHRec.FrmtVersion = "001"
  OHRec.PayPrdBeginDate = Date2Num(fpdtPayPdBegin.Text)
  OHRec.PayPrdEndDate = Date2Num(fpdtPayPdEnd.Text)
  OHRec.RecType = "H"
  OHRec.RptPeriod = RptPeriod
  OHRec.FileCreateDate = Date2Num(Date)
  OHRec.AgencyNum = AgencyNum
  OHRec.MaxSalary = MaxSalary
  Put OHHandle, 1, OHRec
  Close OHHandle
  
  OpenEmpIdxNNameFile IdxHandle
  NumOfRecs = LOF(IdxHandle) \ 2
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxHandle, x, IdxBuff(x)
  Next x
  Close IdxHandle
  
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.cmdCancel.Visible = False
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  Close
  If Exist(OrbitDetail) Then
    KillFile OrbitDetail
  End If
  
  OpenRetFile RetHandle
  NumOfRetRecs = LOF(RetHandle) / Len(RetireRec)
  OpenEmpData2File EHandle
  OpenTransHistFile THandle
  ODCnt = 0
  OpenOrbDetail ODHandle, NumOfODRecs
  OpenOrbEmpData OEHandle, NumOfOERecs
  For x = 1 To NumOfOERecs
    Get OEHandle, x, OERec
'    If QPTrim$(OERec.LastName) = "KIMBLE" Then Stop
    If OERec.Deleted = True Then GoTo SkipIt
    UseOT = False
    If QPTrim$(OERec.PayType) = "LEAVEPAY" Then
      ODRec.Salary = 0
      ODRec.IncDecSalary = "+"
      ODRec.IncDecEmpleeCntrb = "+"
      ODRec.EmployeeCntrb = 0
      ODRec.EmployerCntrb = 0
      ODRec.OTPay = 0
      ODRec.RegPay = 0
      ODRec.CheckDate = BegDate
      ODRec.CheckNum = 0
      ODRec.PayPrdBeginDate = FormatThisPayPd(MakeRegDate(OHRec.PayPrdBeginDate), 2)
      ODRec.PayPrdEndDate = FormatThisPayPd(MakeRegDate(OHRec.PayPrdEndDate), 2)
      NextRec = 0
      GoTo LeavePay
    End If

    If OERec.EmpRecNum = 0 Then GoTo SkipIt
   
    Get EHandle, OERec.EmpRecNum, EmpRec
    If NumOfRetRecs > 0 Then
      For y = 1 To NumOfRetRecs
        Get RetHandle, y, RetireRec
        If QPTrim$(RetireRec.TYPEDES1) = QPTrim$(EmpRec.EMPRETTP) Then
          If RetireRec.TYPEOT1 = "Y" Then
            UseOT = True
          End If
        End If
      Next y
    End If
    NextRec = EmpRec.LastTransRec
    Do While NextRec > 0
      Get THandle, NextRec, TransRec
'      If TransRec.CheckNum = 23 Then Stop
      If TransRec.CheckDate >= BegDate And TransRec.CheckDate <= EndDate Then
        If QPTrim$(OERec.EligibleDate) <> "0" Then
          If ORBITDate2Num(OERec.EligibleDate) > TransRec.PayPdStart Then 'TransRec.PayPdStart
            ODRec.Salary = 0
            ODRec.IncDecSalary = "+"
            ODRec.IncDecEmpleeCntrb = "+"
            ODRec.EmployeeCntrb = 0
            ODRec.EmployerCntrb = 0
            ODRec.OTPay = 0
            ODRec.RegPay = 0
            ODRec.CheckDate = BegDate
            ODRec.CheckNum = 0
            ODRec.PayPrdBeginDate = FormatThisPayPd(MakeRegDate(TransRec.PayPdStart), 2) 'OHRec.PayPrdBeginDate), 2)
            ODRec.PayPrdEndDate = FormatThisPayPd(MakeRegDate(TransRec.PayPdEnd), 2) 'OHRec.PayPrdEndDate), 2)
            GoTo LeavePay
          End If
        End If
        PayPdBegDate = FormatThisPayPd(MakeRegDate(TransRec.PayPdStart), 2)
        PayPdEndDate = FormatThisPayPd(MakeRegDate(TransRec.PayPdEnd), 2)
        ODRec.OTPay = CStr(TransRec.TotOTWage)
        If QPTrim$(OERec.PlanCode) = "LOCRS" Then 'added 5/5/08 because we changed the
        'transrec.retgrosspay to only get a value if the employee is on a retirement
        'plan...in this case the employee is retired and getting a wage the state
        'requires to be submitted but is not on a retirement plan anymore and his
        'retgrosspay = 0.
          If TransRec.RetGrossPay = 0 Then
            TransRec.RetGrossPay = TransRec.GrossPay
          End If
        End If
        ODRec.RegPay = OldRound(TransRec.RetGrossPay - TransRec.TotOTWage) 'added 10/22/07
        If ODRec.RegPay > 0 And ODRec.OTPay > 0 Then
          GoSub CreateOTPayTrans
          ODRec.RegPay = TransRec.RetGrossPay
          ODRec.OTPay = 0
        End If
        ODRec.Salary = TransRec.RetGrossPay
        TotSalary# = TotSalary# + TransRec.RetGrossPay
        If TransRec.RetGrossPay >= 0 Then
          ODRec.IncDecSalary = "+"
        ElseIf TransRec.RetGrossPay < 0 Then
          ODRec.IncDecSalary = "-"
        End If
        ODRec.EmployeeCntrb = TransRec.RetireAmt
        TotEmpyMatch# = TotEmpyMatch# + TransRec.RetireAmt
        If TransRec.RetireAmt >= 0 Then
          ODRec.IncDecEmpleeCntrb = "+"
        ElseIf TransRec.RetireAmt < 0 Then
          ODRec.IncDecEmpleeCntrb = "-"
        End If
        ODRec.EmployerCntrb = TransRec.MatchRetAmt
        TotEmprMatch# = OldRound(TotEmprMatch# + TransRec.MatchRetAmt)
        ODRec.CheckDate = TransRec.CheckDate
        ODRec.CheckNum = TransRec.CheckNum
        ODRec.PayPrdBeginDate = PayPdBegDate
        ODRec.PayPrdEndDate = PayPdEndDate
LeavePay:
        ODRec.AddLine1 = OERec.AddLine1
        ODRec.AddLine2 = OERec.AddLine2
        ODRec.Adjustment = OERec.Adjustment
        ODRec.AgencyNum = OERec.AgencyNum
        ODRec.City = OERec.City
        ODRec.ContrPdEmpBegDate = OERec.ContrPdEmpBegDate
        ODRec.ContrPdEmpEndDate = OERec.ContrPdEmpEndDate
        ODRec.ContrPdEmpPrd = OERec.ContrPdEmpPrd
        ODRec.DateOfBirth = OERec.DateOfBirth
        ODRec.DeptNum = OERec.DeptNum
        ODRec.EligibleDate = OERec.EligibleDate
        ODRec.EmpNum = OERec.EmpNum
        ODRec.EmpRecNum = OERec.EmpRecNum
        ODRec.EmployDate = OERec.EmployDate
        ODRec.FirstName = OERec.FirstName
        ODRec.Gender = OERec.Gender
        ODRec.JobClass = OERec.JobClass
        ODRec.LastName = OERec.LastName
        ODRec.MemberID = OERec.MemberID
        ODRec.MiddleName = OERec.MiddleName
        ODRec.OutOfCntryAdd = OERec.OutOfCntryAdd
        If CDbl(ODRec.OTPay) > 0 And CDbl(ODRec.RegPay) = 0 Then
          ODRec.PayType = "OVERTIME"
        Else
          ODRec.PayType = OERec.PayType
        End If
        ODRec.PlanCode = OERec.PlanCode
        ODRec.RecType = OERec.RecType
        ODRec.SharedPosition = OERec.SharedPosition
        ODRec.SSN = OERec.SSN
        ODRec.State = OERec.State
        ODRec.Suffix = OERec.Suffix
        ODRec.TerminationDate = OERec.TerminationDate
        ODRec.TermType = OERec.TermType
        If OERec.TerminationDate > 0 Then
          ODRec.VacHours = TransRec.VacUsed
        Else
          ODRec.VacHours = 0
        End If
        ODRec.Zip = QPTrim$(OERec.Zip)
        ODRec.Deleted = OERec.Deleted
        ODRec.PostedYN = "N"
        ODCnt = ODCnt + 1
        Put ODHandle, ODCnt, ODRec
      End If
      NextRec = TransRec.PrevTransRec
    Loop
SkipIt:
    FrmShowPctComp.ShowPctComp x, NumOfOERecs '12/31/2009 changed from NumOfRecs
  Next x
  
  OpenOrbTrailer OTHandle
  OTRec.RecType = "F"
  OTRec.AgencyNum = AgencyNum
  OTRec.RptPeriod = RptPeriod
  OTRec.RecCount = ODCnt 'figured at file creation time
  OTRec.IncDecSalary = " " 'figured at file creation time
  OTRec.TotalSalary = 0 'figured at file creation time
  OTRec.IncDecTtlEmpContrb = " " 'figured at file creation time
  OTRec.TotalEmpCntrb = 0 'figured at file creation time
  Put OTHandle, 1, OTRec
  Close
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  If ODCnt > 0 Then
    MakeFile4Submission = True
  End If
  
  Exit Function
  
CreateOTPayTrans:
  TotRetGrossPay = TransRec.RetGrossPay
  TotRetireAmt = TransRec.RetireAmt
  TotMatchRetAmt = TransRec.MatchRetAmt
  OTAndReg = OldRound(CDbl(ODRec.OTPay) + CDbl(ODRec.RegPay))
'  If ODCnt = 15 Then Stop
  OTPct = CDbl(ODRec.OTPay) / OTAndReg
  RegPct = CDbl(ODRec.RegPay) / OTAndReg
  TransRec.RetGrossPay = OldRound(TransRec.RetGrossPay * OTPct)
  OTRetGrossPay = TransRec.RetGrossPay
  If UseOT = False Then
    TransRec.RetireAmt = 0
    TransRec.MatchRetAmt = 0
  Else
    TransRec.RetireAmt = OldRound(TransRec.RetireAmt * OTPct)
    OTRetireAmt = TransRec.RetireAmt
    TransRec.MatchRetAmt = OldRound(TransRec.MatchRetAmt * OTPct)
    OTMatchRetAmt = TransRec.MatchRetAmt
  End If
  ODRec.Salary = TransRec.RetGrossPay
  TotSalary# = TotSalary# + TransRec.RetGrossPay
  If TransRec.RetGrossPay >= 0 Then
    ODRec.IncDecSalary = "+"
  ElseIf TransRec.RetGrossPay < 0 Then
    ODRec.IncDecSalary = "-"
  End If
  ODRec.EmployeeCntrb = TransRec.RetireAmt
  TotEmpyMatch# = TotEmpyMatch# + TransRec.RetireAmt
  If TransRec.RetireAmt >= 0 Then
    ODRec.IncDecEmpleeCntrb = "+"
  ElseIf TransRec.RetireAmt < 0 Then
    ODRec.IncDecEmpleeCntrb = "-"
  End If
  ODRec.EmployerCntrb = TransRec.MatchRetAmt
  TotEmprMatch# = OldRound(TotEmprMatch# + TransRec.MatchRetAmt)
  ODRec.RegPay = 0
  ODRec.CheckDate = TransRec.CheckDate
  ODRec.CheckNum = TransRec.CheckNum
  ODRec.PayPrdBeginDate = PayPdBegDate
  ODRec.PayPrdEndDate = PayPdEndDate
  ODRec.AddLine1 = OERec.AddLine1
  ODRec.AddLine2 = OERec.AddLine2
  ODRec.Adjustment = OERec.Adjustment
  ODRec.AgencyNum = OERec.AgencyNum
  ODRec.City = OERec.City
  ODRec.ContrPdEmpBegDate = OERec.ContrPdEmpBegDate
  ODRec.ContrPdEmpEndDate = OERec.ContrPdEmpEndDate
  ODRec.ContrPdEmpPrd = OERec.ContrPdEmpPrd
  ODRec.DateOfBirth = OERec.DateOfBirth
  ODRec.DeptNum = OERec.DeptNum
  ODRec.EligibleDate = OERec.EligibleDate
  ODRec.EmpNum = OERec.EmpNum
  ODRec.EmpRecNum = OERec.EmpRecNum
  ODRec.EmployDate = OERec.EmployDate
  ODRec.FirstName = OERec.FirstName
  ODRec.Gender = OERec.Gender
  ODRec.JobClass = OERec.JobClass
  ODRec.LastName = OERec.LastName
  ODRec.MemberID = OERec.MemberID
  ODRec.MiddleName = OERec.MiddleName
  ODRec.OutOfCntryAdd = OERec.OutOfCntryAdd
  ODRec.PayType = "OVERTIME"
  ODRec.PlanCode = OERec.PlanCode
  ODRec.RecType = OERec.RecType
  ODRec.SharedPosition = OERec.SharedPosition
  ODRec.SSN = OERec.SSN
  ODRec.State = OERec.State
  ODRec.Suffix = OERec.Suffix
  ODRec.TerminationDate = OERec.TerminationDate
  ODRec.TermType = OERec.TermType
  If OERec.TerminationDate > 0 Then
    ODRec.VacHours = TransRec.VacUsed
  Else
    ODRec.VacHours = 0
  End If
  ODRec.Zip = QPTrim$(OERec.Zip)
  ODRec.Deleted = OERec.Deleted
  ODRec.PostedYN = "N"
  ODCnt = ODCnt + 1
  Put ODHandle, ODCnt, ODRec
  RegRetGrossPay = OldRound(TotRetGrossPay * RegPct)
  
  Dif = OldRound(TotRetGrossPay - (RegRetGrossPay + OTRetGrossPay))
  TransRec.RetGrossPay = OldRound(RegRetGrossPay + Dif)
  If UseOT = False Then
    TransRec.RetireAmt = TotRetireAmt
    TransRec.MatchRetAmt = TotMatchRetAmt
  Else
    RegRetireAmt = OldRound(TotRetireAmt * RegPct)
    Dif = OldRound(TotRetireAmt - (RegRetireAmt + OTRetireAmt))
    TransRec.RetireAmt = OldRound(RegRetireAmt + Dif)
  
    RegMatchRetAmt = OldRound(TotMatchRetAmt * RegPct)
    Dif = OldRound(TotMatchRetAmt - (RegMatchRetAmt + OTMatchRetAmt))
    TransRec.MatchRetAmt = OldRound(RegMatchRetAmt + Dif)
  End If
  Return
  
End Function

Public Sub ZeroFill(ByRef thisNum$, ThisLen As Integer)
  Dim x As Integer
  Dim thischar$
  Dim BCnt As Integer
  Dim ThisTemp$
  
  For x = 1 To ThisLen
    thischar = Mid(thisNum, x, 1)
    If thischar = " " Then
      BCnt = BCnt + 1
    End If
  Next x
  
  For x = 1 To BCnt
    ThisTemp = ThisTemp + "0"
  Next x
  
  thisNum$ = ThisTemp + QPTrim$(thisNum)
  
End Sub

Private Function FormatThisPayPd(ByRef ThisDate As String, ByVal Vers As Integer) As String
  Dim ch As String
  Dim DateLen As Integer
  Dim FSPstn As Integer
  Dim x As Integer
  Dim ThisDay As String
  Dim ThisMonth As String
  Dim ThisYear As String
  
  FSPstn = 0
  DateLen = Len(ThisDate)
  For x = 1 To DateLen
    ch = Mid(ThisDate, x, 1)
    If ch = "/" Then
      FSPstn = x
      Exit For
    End If
  Next x
  
  ThisMonth = Mid(ThisDate, 1, FSPstn - 1)
  If Len(ThisMonth) = 1 Then ThisMonth = "0" & ThisMonth
  ThisDay = Mid(ThisDate, FSPstn + 1, 2)
  If Len(ThisDay) = 1 Then ThisDay = "0" + ThisDay
  ThisYear = Mid(ThisDate, DateLen - 3, DateLen)
  If Vers = 2 Then
    ThisDate = ThisYear & ThisMonth & ThisDay
  ElseIf Vers = 1 Then
    ThisDate = ThisYear & ThisMonth
  End If
  FormatThisPayPd = ThisDate
  
End Function

Private Function Look4MaxSalary() As Boolean
  Dim Emp2Rec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfERecs As Integer
  Dim AvePay As Double
  Dim PayCnt As Integer
  Dim x As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim TransRec As TransRecType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim ORec As OrbitEmpData
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim NextRec As Long
  Dim ThisSalary As Double
  Dim Name As String * 50
  Dim AvePayS As String * 13
  Dim Gross As String * 13
  
  Look4MaxSalary = False
  If MaxSalary = 0 Then Exit Function
  BegDate = Date2Num(("01/01/") & Mid(fpdtPayPdEnd.Text, 7, 4))
  EndDate = Date2Num(fpdtPayPdEnd.Text)
  OpenEmpData2File EHandle
  NumOfERecs = LOF(EHandle) / Len(Emp2Rec)
  OpenTransHistFile THandle
  NumOfTRecs = LOF(THandle) / Len(TransRec)
  OpenOrbEmpData OHandle, NumOfORecs
  
  ReDim MaxSalaryPL(1 To 1) As String
  MaxCnt = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo Deleted
    Get EHandle, ORec.EmpRecNum, Emp2Rec
    NextRec = Emp2Rec.LastTransRec
    ThisSalary = 0
    PayCnt = 0
    Do While NextRec > 0
      Get THandle, NextRec, TransRec
      If TransRec.PayPdStart >= BegDate And TransRec.PayPdEnd <= EndDate Then
        ThisSalary = ThisSalary + TransRec.RetGrossPay
        PayCnt = PayCnt + 1
      End If
      NextRec = TransRec.PrevTransRec
    Loop
    If PayCnt > 0 Then
      AvePay = ThisSalary / PayCnt
      If ThisSalary + AvePay >= MaxSalary Then
        MaxCnt = MaxCnt + 1
        ReDim Preserve MaxSalaryPL(1 To MaxCnt) As String
        LSet Name = QPTrim$(Emp2Rec.EmpLName) & ", " & QPTrim$(Emp2Rec.EmpFName)
        RSet AvePayS = Using$("$###,###.##", AvePay)
        RSet Gross = Using$("###,###.##", ThisSalary)
        MaxSalaryPL(MaxCnt) = Name & "  " & AvePayS & "  " & Gross
      End If
    End If
Deleted:
  Next x
  Close EHandle
  Close OHandle
  Close THandle
  
  If MaxCnt > 1 Then
    Look4MaxSalary = True
  End If
  
End Function

Private Function ORBITDate2Num(ByVal ThisDate As String) As Integer
  Dim ThisDay As String
  Dim ThisMonth As String
  Dim ThisYear As String
  Dim x As Integer, ch As String, cnt As Integer
  
  ThisMonth = Mid(ThisDate, 5, 2)
  If Len(ThisMonth) = 1 Then ThisMonth = "0" & ThisMonth
  ThisDay = Mid(ThisDate, 7, 2)
  If Len(ThisDay) = 1 Then ThisDay = "0" + ThisDay
  ThisYear = Mid(ThisDate, 1, 4)
  ThisDate = ThisMonth & "/" & ThisDay & "/" & ThisYear
  If Mid(ThisDate, 3, 1) <> "/" Or Mid(ThisDate, 6, 1) <> "/" Then ThisDate = "12/31/1979"
  cnt = 0
  For x = 1 To Len(ThisDate)
    If Mid(ThisDate, x, 1) = "/" Then cnt = cnt + 1
  Next x
  If cnt > 2 Then ThisDate = "12/31/1979"
  ORBITDate2Num = Date2Num(ThisDate)
  
End Function
