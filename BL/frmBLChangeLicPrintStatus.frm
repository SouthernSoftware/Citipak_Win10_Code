VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBLChangeLicPrintStatus 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bueinss License Change License Print Status"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLChangeLicPrintStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5892
      Left            =   1920
      TabIndex        =   3
      Top             =   1500
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10398
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLChangeLicPrintStatus.frx":08CA
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   3105
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'License Processing' menu."
         Top             =   4650
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmBLChangeLicPrintStatus.frx":08E6
      End
      Begin EditLib.fpDateTime fpDateXDate 
         Height          =   465
         Left            =   3315
         TabIndex        =   0
         Tag             =   $"frmBLChangeLicPrintStatus.frx":0AC4
         Top             =   3930
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3228
         _ExtentY        =   820
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
         InvalidColor    =   12648447
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
         Text            =   "03/19/2003"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
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
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   5280
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   $"frmBLChangeLicPrintStatus.frx":0C6B
         Top             =   4650
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmBLChangeLicPrintStatus.frx":0E6B
      End
      Begin fpBtnAtlLibCtl.fpBtn fpcmdXList 
         Height          =   360
         Left            =   5190
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmBLChangeLicPrintStatus.frx":104A
         Top             =   4005
         Width           =   1920
         _Version        =   131072
         _ExtentX        =   3387
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmBLChangeLicPrintStatus.frx":1138
      End
      Begin fpBtnAtlLibCtl.fpBtn fpcmdHelp 
         Height          =   645
         Left            =   630
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmBLChangeLicPrintStatus.frx":131E
         Top             =   4650
         Width           =   2160
         _Version        =   131072
         _ExtentX        =   3810
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmBLChangeLicPrintStatus.frx":13EE
      End
      Begin VB.Label lblBalloon 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "HELP BALLOONS ON"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Left            =   672
         TabIndex        =   9
         Top             =   5328
         Width           =   2100
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Expiration Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   750
         TabIndex        =   6
         Top             =   4110
         Width           =   2400
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2172
         Left            =   576
         Top             =   1440
         Width           =   6732
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLChangeLicPrintStatus.frx":15D1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1452
         Left            =   960
         TabIndex        =   5
         Top             =   1824
         Width           =   6012
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Set Licenses To Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2190
         TabIndex        =   4
         Top             =   570
         Width           =   3735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1536
         Top             =   432
         Width           =   4908
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   2400
      TabIndex        =   10
      Top             =   7584
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   3000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6150
      Left            =   1801
      Top             =   1359
      Width           =   8050
   End
End
Attribute VB_Name = "frmBLChangeLicPrintStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  KillFile "setstatus.dat"
  frmBLPrintLicMenu.Show
  DoEvents
  Unload frmBLChangeLicPrintStatus
End Sub

Private Sub cmdProcess_Click()
  Dim TrHandle As Integer
  Dim CustRec As ARCustRecType
  Dim TRNumRecs As Integer
  Dim cnt As Integer
  Dim TempChrg As TempChargesType
  Dim TempHandle As Integer
  Dim NextTempRec As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer
  Dim ThisCode$
  Dim ThisRev As Double
  Dim ThisFee As Double
  Dim TotalFees As Double
  Dim ThisTempRec As Integer
  Dim LaserRec1 As LaserLetterType1
  Dim LaserRec2 As LaserLetterType2
  Dim LHandle As Integer
  Dim YCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Exist("artmppst.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "The business license fee register has already been processed. Please process the business license register again after this operation has completed successfully. Also be sure to evaluate any business license forms already printed."
    frmBLMessageBoxJr.Label1.Top = 500
    frmBLMessageBoxJr.Label1.Height = 1300
    frmBLMessageBoxJr.Show vbModal
  End If
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  ThisTempRec = 1
  OpenCustFile TrHandle
  TRNumRecs = LOF(TrHandle) \ Len(CustRec)
  If TRNumRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customer records saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenCatCodeFile CodeHandle
  NumOfCodeRecs = LOF(CodeHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "There are no category records saved. Do you want to continue anyway?"
    frmBLMessageBoxJrWOpts.Label1.Top = 800
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  ReDim OmitList(1 To 1) As Long
  OmitCnt = 0
  frmBLShowPctComp.Label1 = "Setting Business License Fee Customer Flags"
  frmBLShowPctComp.cmdCancel.Visible = False
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  For cnt = 1 To TRNumRecs
    Get TrHandle, cnt, CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo SkipThis
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
      If MakeRegDate(CustRec.VALID) = QPTrim$(fpDateXDate.Text) Then
        CustRec.IssueLicense = "Y"
        YCnt = YCnt + 1
        Put TrHandle, cnt, CustRec
      End If
    End If
SkipThis:
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      fpcmdHelp.Enabled = True
      
      Exit Sub
    End If
  Next cnt
  Close         'Close all open files now
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  If YCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are zero customers whose expiration date is " + fpDateXDate.Text + "."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  frmBLMessageBoxJr.Label1.Caption = CStr(YCnt) + " Business Licenses set for customers whose licenses expire on " + fpDateXDate.Text + "."
  frmBLMessageBoxJr.Label1.Top = 800
  frmBLMessageBoxJr.Show vbModal
  
  MainLog ("Set Licenses to Print processed for " + fpDateXDate.Text + ".")
  frmBLPrintLicMenu.Show
  DoEvents
  Unload frmBLChangeLicPrintStatus
  Close
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLChangeLicPrintStatus", "cmdProcess_Click", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%E"
      Call fpcmdXList_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call fpcmdHelp_Click
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
      KillFile "setstatus.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLChangeLicPrintStatus.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim DHandle As Integer
  
  lblBalloon.Visible = False
'  fpDateXDate.ToolTipText = "Enter a business license expiration date to create a file of businesses whose licenses will expire on this date."
'  fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'  cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'  cmdProcess.ToolTipText = "Press the 'Process' button to create a file containing all customers whose licenses expire on the date entered and set them up for license renewal."
'  cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
  One = 1
  DHandle = FreeFile
  Open "setstatus.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  fpDateXDate = Date
End Sub

Private Sub fpcmdHelp_Click()
  If InStr(fpcmdHelp.Text, "On") Then
    fpcmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fpDateXDate.ToolTipText = ""
    fpcmdXList.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
    fpcmdHelp.ToolTipText = ""
  ElseIf InStr(fpcmdHelp.Text, "Off") Then
    fpcmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpDateXDate.ToolTipText = "Enter a business license expiration date to create a file of businesses whose licenses will expire on this date."
'    fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'    cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'    cmdProcess.ToolTipText = "Press the 'Process' button to create a file containing all customers whose licenses expire on the date entered and set them up for license renewal."
'    fpcmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
  End If
End Sub

Private Sub fpcmdXList_Click()
  frmBLXDateList.Show vbModal
End Sub
