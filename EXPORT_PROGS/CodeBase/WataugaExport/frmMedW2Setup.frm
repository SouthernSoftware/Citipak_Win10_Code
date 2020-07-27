VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmMedW2Setup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Med Only W-2 Extract"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11640
   Icon            =   "frmMedW2Setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread vaSpreadW2 
      Height          =   4830
      Left            =   1395
      TabIndex        =   0
      Top             =   2280
      Width           =   8850
      _Version        =   196613
      _ExtentX        =   15610
      _ExtentY        =   8520
      _StockProps     =   64
      ColsFrozen      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   13684944
      MaxCols         =   4
      MaxRows         =   52
      RetainSelBlock  =   0   'False
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmMedW2Setup.frx":08CA
      VisibleCols     =   4
      TextTip         =   2
   End
   Begin EditLib.fpDateTime fptxtYear 
      Height          =   375
      Left            =   5910
      TabIndex        =   3
      ToolTipText     =   "Enter the Year to extract W2 information here."
      Top             =   1560
      Width           =   1260
      _Version        =   196608
      _ExtentX        =   2222
      _ExtentY        =   656
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
      Text            =   "2005"
      DateCalcMethod  =   1
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
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
      Appearance      =   0
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   480
      Left            =   8850
      TabIndex        =   4
      Top             =   7485
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmMedW2Setup.frx":0ED7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   480
      Left            =   7290
      TabIndex        =   5
      Top             =   7485
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmMedW2Setup.frx":10B3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear2 
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmMedW2Setup.frx":128F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear3 
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   7320
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmMedW2Setup.frx":1479
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear4 
      Height          =   615
      Left            =   4920
      TabIndex        =   8
      Top             =   7320
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmMedW2Setup.frx":1666
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Social Security Exempt Employees Only"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   840
      Width           =   3135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   885
      Left            =   1440
      Top             =   7200
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Extract Year"
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
      Left            =   4185
      TabIndex        =   2
      Top             =   1635
      Width           =   1410
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   972
      Index           =   1
      Left            =   1464
      Top             =   288
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W-2 Extraction Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   405
      Width           =   6015
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   585
      Left            =   3990
      Top             =   1470
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1092
      Left            =   1464
      Top             =   168
      Width           =   8652
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
Attribute VB_Name = "frmMedW2Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private Sub cmdClear2_Click()
  Dim x As Integer '8/20/04
  
  For x = 1 To 52
    vaSpreadW2.Row = x
    vaSpreadW2.Col = 2
    vaSpreadW2.Text = ""
  Next x

End Sub

Private Sub cmdClear3_Click()
  Dim x As Integer '8/20/04
  
  For x = 1 To 52
    vaSpreadW2.Row = x
    vaSpreadW2.Col = 3
    vaSpreadW2.Text = ""
  Next x
  
End Sub

Private Sub cmdClear4_Click()
  Dim x As Integer '8/20/04
  
  For x = 1 To 52
    vaSpreadW2.Row = x
    vaSpreadW2.Col = 4
    vaSpreadW2.Text = ""
  Next x

End Sub

Private Sub cmdPrintForm_Click()
  PrintForm
End Sub

Private Sub cmdSave_Click()
  Dim W2Type As Integer
  Dim x As Integer
  Dim TempVal As String
  
  vaSpreadW2.Col = 3
  For x = 1 To 52
    vaSpreadW2.Row = x
    TempVal = QPTrim$(vaSpreadW2.Text)
    If Len(QPTrim$(vaSpreadW2.Text)) > 0 Then
      If TempVal <> "12a" And TempVal <> "12b" And _
      TempVal <> "12c" And TempVal <> "12d" And _
      TempVal <> "14a" And TempVal <> "14b" Then
        MsgBox "Please make a valid selection from the list provided."
        vaSpreadW2.SetFocus
        vaSpreadW2.SetActiveCell 3, x
        Exit Sub
      End If
    End If
  Next x
  
  W2Type = 2
  Call ExtractW2Info(W2Type%, Me)
  KillFile "PRDATA\W2RPNIDX.DAT" 'deletes the current reprint saves
'  MsgBox "Your information has been saved."
  frmW2Message.Label1.Caption = "W2 extraction has completed successfully. If you intend to file your Forms W2 electronically please take the time to test your data using ACCUWAGE before editing any further. Then make your editing changes and run ACCUWAGE again to make sure new data has been entered correctly."
  frmW2Message.Label1.Top = 450
  frmW2Message.Label1.Height = 1300
  frmW2Message.Show vbModal
  frmW2Processing.Show
  DoEvents
  Unload frmMedW2Setup
  MainLog ("Second W2 extraction saved.")
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
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF2:
      Call cmdClear2_Click
      KeyCode = 0
    Case vbKeyF3:
      Call cmdClear3_Click
      KeyCode = 0
    Case vbKeyF4:
      Call cmdClear4_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call FixSpread
  Call LoadThisForm
  Me.HelpContextID = hlpExtractMedW2
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub cmdExit_Click()
  Dim DoWhatFlag As SaveChangeOptions1
  Dim changeFlag As Boolean
  Dim W2Handle As Integer
  Dim W2SetUpRec As W2SetUpType
  Dim x As Integer
  
  changeFlag = False
  OpenW2SetUp W2Handle
  Get W2Handle, 1, W2SetUpRec
  Close W2Handle
  
  If Val(fptxtYear.Text) <> W2SetUpRec.ExtrYear Then
    changeFlag = True
    fptxtYear.SetFocus
    GoTo changeFound
  End If
  
  For x = 1 To 51
    vaSpreadW2.Col = 2
    vaSpreadW2.Row = x
    If QPTrim$(vaSpreadW2.Text) = "Deferred Compensation" And QPTrim$(W2SetUpRec.Deds(x - 1).CHKDED) = "Deferred Compensatio" Then GoTo ThisIsOK
    If QPTrim$(vaSpreadW2.Text) <> QPTrim$(W2SetUpRec.Deds(x - 1).CHKDED) Then
      changeFlag = True
      vaSpreadW2.SetActiveCell 2, x
      vaSpreadW2.SetFocus
      GoTo changeFound
    End If
ThisIsOK:
    vaSpreadW2.Col = 3
    vaSpreadW2.Row = x
    If QPTrim$(vaSpreadW2.Text) <> QPTrim$(W2SetUpRec.Deds(x - 1).AMTBOX) Then
      changeFlag = True
      vaSpreadW2.SetActiveCell 3, x
      vaSpreadW2.SetFocus
      GoTo changeFound
    End If
    vaSpreadW2.Col = 4
    vaSpreadW2.Row = x
    If QPTrim$(vaSpreadW2.Text) <> QPTrim$(W2SetUpRec.Deds(x - 1).DedCode) Then
      changeFlag = True
      vaSpreadW2.SetActiveCell 4, x
      vaSpreadW2.SetFocus
      GoTo changeFound
    End If
  Next x
changeFound:
  If changeFlag = True Then
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges 'save changes
      Call cmdSave_Click
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      frmW2Processing.Show
      DoEvents
      Unload frmMedW2Setup
    Case Else:
       'Do nothing because we don't know about any options except
       'save, review or abandon...used as a placeholder for adding
       'other options at a later date
    End Select
  Else
    changeFlag = False
    frmW2Processing.Show
    DoEvents
    Unload frmMedW2Setup
  End If
End Sub

Private Sub LoadThisForm()
   Dim EmpData2FileHandle As Integer
   Dim EmpData2FileRec As EmpData2Type
   Dim EmpRecLen As Long, Today As String * 11
   Dim DedCodeFileRec As DedCodeRecType
   Dim DedCodeFileHandle As Integer
   Dim DecRecLen As Integer, x As Integer
   Dim W2Handle As Integer
   Dim W2SetUpRec As W2SetUpType
   Dim RetType(0 To 51) As String
   
   OpenDedCodeFile DedCodeFileHandle
   DecRecLen = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
   For x = 1 To 52
      If x = 1 Then
        vaSpreadW2.Col = 1
        vaSpreadW2.Row = 1
        vaSpreadW2.Text = "Retirement"
        GoTo Zero
      ElseIf x = 2 Then
        vaSpreadW2.Col = 1
        vaSpreadW2.Row = 2
        vaSpreadW2.Text = "Tax Fringe"
        GoTo Zero
      Else
        Get DedCodeFileHandle, x - 2, DedCodeFileRec
        vaSpreadW2.Col = 1
        vaSpreadW2.Row = x
        vaSpreadW2.Text = QPTrim$(DedCodeFileRec.DCDESC1)
      End If
Zero:
   Next x
   
   Close DedCodeFileHandle
   OpenW2SetUp W2Handle
   Get W2Handle, 1, W2SetUpRec
   Close W2Handle
   For x = 1 To 52
      RetType(x - 1) = QPTrim$(W2SetUpRec.Deds(x - 1).CHKDED)
      If RetType(x - 1) = "Deferred Compensatio" Then RetType(x - 1) = "Deferred Compensation"
   Next x
   For x = 1 To 52
     vaSpreadW2.Col = 2
     vaSpreadW2.Row = x
     vaSpreadW2.Text = RetType(x - 1)
     vaSpreadW2.Col = 3
     vaSpreadW2.Row = x
     vaSpreadW2.Text = QPTrim$(W2SetUpRec.Deds(x - 1).AMTBOX)
     vaSpreadW2.Col = 4
     vaSpreadW2.Row = x
     vaSpreadW2.Text = QPTrim$(W2SetUpRec.Deds(x - 1).DedCode)
   Next x
'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   fptxtYear.Text = W2SetUpRec.ExtrYear
End Sub

Private Function FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  '-1 means all rows or all columns....0 means headers
    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 14
        coladj = 11
        vaSpreadW2.FontSize = 18
        vaSpreadW2.RowHeight(-1) = 22
        vaSpreadW2.RowHeight(0) = 22
      Else
        COne = 6
        coladj = 4.5
        vaSpreadW2.RowHeight(-1) = 18
        vaSpreadW2.RowHeight(0) = 18
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 12
        coladj = 7.5
        vaSpreadW2.FontSize = 14
        vaSpreadW2.RowHeight(0) = 18.5
        vaSpreadW2.RowHeight(-1) = 18.5
      Else
        COne = 4
        coladj = 2
        vaSpreadW2.RowHeight(0) = 15
        vaSpreadW2.RowHeight(-1) = 15
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 10.5
        coladj = 3.1
        vaSpreadW2.FontBold = True
        vaSpreadW2.RowHeight(0) = 17.5
        vaSpreadW2.RowHeight(-1) = 17.5
      Else
        COne = -0.5
        coladj = 0
      End If
      Case 800
        COne = -4.5
        coladj = 0.75
        vaSpreadW2.Font.Size = 10
        vaSpreadW2.RowHeight(-1) = 12.2
      Case Else
       
    End Select
    vaSpreadW2.ColWidth(1) = vaSpreadW2.ColWidth(1) + COne
    vaSpreadW2.ColWidth(2) = vaSpreadW2.ColWidth(2) + coladj
    vaSpreadW2.ColWidth(3) = vaSpreadW2.ColWidth(3) + coladj
    vaSpreadW2.ColWidth(4) = vaSpreadW2.ColWidth(4) + coladj

End Function

Private Sub fptxtYear_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    vaSpreadW2.Col = 2
    vaSpreadW2.Row = 1
    vaSpreadW2.SetActiveCell 2, 1
    vaSpreadW2.SetFocus
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn(RegExit)
      MainLog ("Payroll.exe terminated via menu bar on frmMedW2Setup.")
      End
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
  MainLog ("Second W2 extraction screen printed.")
End Sub

Private Sub vaSpreadW2_Change(ByVal Col As Long, ByVal Row As Long)
  Dim box As String
  If Col = 2 Then
    vaSpreadW2.Col = 2
    vaSpreadW2.Row = Row
    box = QPTrim$(vaSpreadW2.Text)
    If box <> "Pension Plan" And box <> "Deferred Compensation" _
    And box <> "" And box <> "P" And box <> "D" And box <> "p" And box <> "d" Then
      MsgBox "ERROR: Entry is not valid. Select from the Pick List or leave blank."
        vaSpreadW2.Text = ""
        vaSpreadW2.SetActiveCell 2, Row
        Exit Sub
    End If
  ElseIf Col = 3 Then
    vaSpreadW2.Col = 3
    vaSpreadW2.Row = Row
    box = QPTrim$(vaSpreadW2.Text)
    If box <> "" And box <> "12a" And box <> "12b" And box <> "12c" And box <> "12d" And box <> "14a" And box <> "14b" Then
      MsgBox "ERROR: Entry is not valid. Select from the Pick List or leave blank."
      vaSpreadW2.Text = ""
      vaSpreadW2.SetActiveCell 3, Row
      Exit Sub
    End If
  End If

End Sub

Private Sub vaSpreadW2_KeyPress(KeyAscii As Integer)
   If vaSpreadW2.ActiveCol = 4 Then
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
   End If

End Sub

Private Sub vaSpreadW2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  If Col = 2 Then
    vaSpreadW2.Col = 2
    vaSpreadW2.Row = Row
    If QPTrim$(vaSpreadW2.Text) = "D" Or QPTrim$(vaSpreadW2.Text) = "d" Then vaSpreadW2.Text = "Deferred Compensation"
    If QPTrim$(vaSpreadW2.Text) = "P" Or QPTrim$(vaSpreadW2.Text) = "p" Then vaSpreadW2.Text = "Pension Plan"
  End If

End Sub



