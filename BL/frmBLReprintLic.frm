VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBLReprintLic 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint Business Licenses"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLReprintLic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5292
      Left            =   1872
      TabIndex        =   2
      Top             =   1776
      Width           =   7932
      _Version        =   196609
      _ExtentX        =   13991
      _ExtentY        =   9334
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDShadowColor=   -2147483633
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLReprintLic.frx":08CA
      Begin EditLib.fpText fptxtFirstNum 
         Height          =   396
         Left            =   3168
         TabIndex        =   0
         Tag             =   $"frmBLReprintLic.frx":08E6
         Top             =   2016
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
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
         ControlType     =   0
         Text            =   ""
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ,"
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
      Begin EditLib.fpText fptxtLastNum 
         Height          =   396
         Left            =   3168
         TabIndex        =   1
         Tag             =   $"frmBLReprintLic.frx":0994
         Top             =   3120
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
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
         ControlType     =   0
         Text            =   ""
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ,"
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
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   636
         Left            =   480
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmBLReprintLic.frx":0A53
         Top             =   4032
         Width           =   2172
         _Version        =   131072
         _ExtentX        =   3831
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmBLReprintLic.frx":0B23
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdReprint 
         Height          =   630
         Left            =   5130
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   $"frmBLReprintLic.frx":0D06
         Top             =   4035
         Width           =   2325
         _Version        =   131072
         _ExtentX        =   4101
         _ExtentY        =   1111
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
         ButtonDesigner  =   "frmBLReprintLic.frx":0DC5
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   2928
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'License Processing' menu."
         Top             =   4032
         Width           =   1932
         _Version        =   131072
         _ExtentX        =   3408
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmBLReprintLic.frx":0FAA
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
         Left            =   528
         TabIndex        =   7
         Top             =   4704
         Width           =   2100
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "First Sequence Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2832
         TabIndex        =   5
         Top             =   1584
         Width           =   2556
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Sequence Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         TabIndex        =   4
         Top             =   2688
         Width           =   2700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Reprint Business Licenses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   396
         Left            =   1776
         TabIndex        =   3
         Top             =   624
         Width           =   4572
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   828
         Left            =   1536
         Top             =   384
         Width           =   4956
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   540
      Left            =   1872
      TabIndex        =   8
      Top             =   7440
      Width           =   876
      _Version        =   131072
      _ExtentX        =   1545
      _ExtentY        =   952
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
      MaxWidth        =   5000
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5616
      Left            =   1716
      Top             =   1626
      Width           =   8220
   End
End
Attribute VB_Name = "frmBLReprintLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim SmallNum As Double
  Dim LargeNum As Double
  
Private Sub cmdExit_Click()
  frmBLPrintLicMenu.Show
  DoEvents
  Unload frmBLReprintLic
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fptxtFirstNum.ToolTipText = ""
    fptxtLastNum.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdReprint.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtFirstNum.ToolTipText = "Each business license printed on a tractor fed printer includes a tracking number. Enter the beginning tractor number in this field."
'    fptxtLastNum.ToolTipText = "Each business license printed on a tractor fed printer includes a tracking number. Enter the ending tractor number in this field."
'    cmdExit.ToolTipText = "Press 'Cancel' to exit this screen."
'    cmdReprint.ToolTipText = "Press to begin reprinting all business forms from the first to the last sequence number."
  End If
End Sub

Private Sub cmdReprint_Click()
  If Not Exist("artmplic.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "License printing has not yet taken place, reprints are not possible."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  Call PrintText
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
    DoEvents
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
      SendKeys "%R"
      Call cmdReprint_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLReprintLic.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TPHandle As Integer
  Dim TempPrint As TempLicPrintType
  Dim x As Integer
  Dim NumOfTempRecs As Integer
  Dim BigNum As Double
  
  On Error Resume Next
  
  lblBalloon.Visible = False
'  fptxtFirstNum.ToolTipText = "Each business license printed on a tractor fed printer includes a tracking number. Enter the beginning tractor number in this field."
'  fptxtLastNum.ToolTipText = "Each business license printed on a tractor fed printer includes a tracking number. Enter the ending tractor number in this field."
'  cmdExit.ToolTipText = "Press 'Cancel' to exit this screen."
'  cmdReprint.ToolTipText = "Press to begin reprinting all business forms from the first to the last sequence number."
  If Exist("artmplic.dat") Then
    BigNum = 0
    OpenTempLicPrint TPHandle
    NumOfTempRecs = LOF(TPHandle) / Len(TempPrint)
    Get TPHandle, 1, TempPrint
    fptxtFirstNum.Text = TempPrint.SeqNum
    SmallNum = TempPrint.SeqNum
    Get TPHandle, NumOfTempRecs, TempPrint
    fptxtLastNum = TempPrint.SeqNum
    LargeNum = TempPrint.SeqNum
    
    Close TPHandle
  End If
 
End Sub
Private Sub PrintText()
  Dim ReportFile$
  Dim FF$, x As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim TCat$, ZCnt&, cnt&
  Dim NumOfTransRecs As Double
  Dim NextTransRec As Double
  Dim CategoryRecord1 As Integer
  Dim CategoryRecord2 As Integer
  Dim CategoryRecord3 As Integer
  Dim CategoryRecord4 As Integer
  Dim CategoryRecord5 As Integer
  Dim TotalBillAmt#
  Dim PostDate$
  Dim CustomerNumber As Integer
  Dim Prev As Long
  Dim CategoryDesc$
  Dim CategoryDesc1$
  Dim CategoryDesc2$
  Dim CategoryDesc3$
  Dim CategoryDesc4$
  Dim CategoryDesc5$, DidCnt As Integer
  Dim LICENSE#, ll As Integer
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim Heading1 As Integer
  Dim Heading2 As Integer
  Dim Heading3 As Integer
  Dim Heading4 As Integer
  Dim tab1 As Integer
  Dim tab2 As Integer
  Dim Tab3 As Integer
  Dim Tab4 As Integer
  Dim SHeading1$
  Dim SHeading2$
  Dim SHeading3$
  Dim SHeading4$
  Dim IssueDate$
  Dim SCnt As Integer, LCnt As Integer
  Dim TempPrint As TempLicPrintType
  Dim TPHandle As Integer
  Dim NumOfTempPrintRecs As Integer
  Dim Year$, ExpireDate$
  Dim TempPostRec As TempTransPostType
  Dim TempPostHandle As Integer
  Dim NumOfTempRecs As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim CustNameIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim Nextx As Double, y As Double
  
  On Error GoTo ERRORSTUFF
  
  If Exist("artownsu.dat") Then
    OpenTownFile TownHandle
    Get TownHandle, 1, TownRec
    Close TownHandle
  End If
  
  If QPTrim$(fptxtFirstNum.Text) = "" Then
    fptxtFirstNum.BackColor = 65535
    frmBLMessageBoxJr.Label1.Caption = "Please enter a first license number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtFirstNum.BackColor = &H80000005
    Exit Sub
  End If
  
  If QPTrim$(fptxtLastNum.Text) = "" Then
    fptxtLastNum.BackColor = 65535
    frmBLMessageBoxJr.Label1.Caption = "Please enter a last license number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtLastNum.BackColor = &H80000005
    Exit Sub
  End If
  
  If Val(fptxtFirstNum.Text) > Val(fptxtLastNum.Text) Then
    fptxtFirstNum.BackColor = 65535
    fptxtLastNum.BackColor = 65535
    frmBLMessageBoxJr.Label1.Caption = "Error: The first number must be smaller than the last number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtFirstNum.BackColor = &H80000005
    fptxtLastNum.BackColor = &H80000005
    fptxtFirstNum.SetFocus
    Exit Sub
  End If
  
  If Val(fptxtFirstNum.Text) < SmallNum Then
    fptxtFirstNum.BackColor = 65535
    frmBLMessageBoxJr.Label1.Caption = "Error: Invalid first sequence number. The first number cannot be smaller than " + Str(SmallNum) + "."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtFirstNum.BackColor = &H80000005
    fptxtFirstNum.Text = SmallNum
    fptxtFirstNum.SetFocus
    Exit Sub
  End If
  
  If Val(fptxtLastNum.Text) > LargeNum Then
    fptxtLastNum.BackColor = 65535
    frmBLMessageBoxJr.Label1.Caption = "Error: Invalid last sequence number. The last number cannot be greater than " + Str(LargeNum) + "."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtLastNum.BackColor = &H80000005
    fptxtLastNum.Text = LargeNum
    fptxtLastNum.SetFocus
    Exit Sub
  End If
  
  ReportFile$ = "REPRTLIC.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenTransFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  Close THandle
  
  OpenTempLicPrint TPHandle
  NumOfTempPrintRecs = LOF(TPHandle) / Len(TempPrint)
'  LICENSE# = Val(fptxtFirstNum.Text)
  
  If NumOfTempPrintRecs > 0 Then
    Get TPHandle, 1, TempPrint
    If Len(TempPrint.Head1) > 0 Then tab1 = Len(QPTrim$(TempPrint.Head1)) / 2 Else tab1 = 0
    If Len(TempPrint.Head2) > 0 Then tab2 = Len(QPTrim$(TempPrint.Head2)) / 2 Else tab2 = 0
    If Len(TempPrint.Head3) > 0 Then Tab3 = Len(QPTrim$(TempPrint.Head3)) / 2 Else Tab3 = 0
    If Len(TempPrint.Head4) > 0 Then Tab4 = Len(QPTrim$(TempPrint.Head4)) / 2 Else Tab4 = 0
  End If
  
  If TempPrint.Order = "A" Then 'name order
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
    ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
    ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  
  OpenTempPostFile TempPostHandle
  NumOfTempRecs = LOF(TempPostHandle) / Len(TempPostRec)
  
  If NumOfTempRecs <> NumOfTempPrintRecs Then
    frmBLMessageBoxJr.Label1.Caption = "Please print business license forms to the screen again. Two necessary files (temporary post and temporary license print) are not matching up as expected."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  Nextx = 0
  ReDim PrintIdx(1 To 1) As Double
  
  For x = 1 To NumOfCustIdxRecs
    For y = 1 To NumOfTempRecs
      Get TempPostHandle, y, TempPostRec
        If CDbl(TempPostRec.CustomerNumber) = IdxRecs(x) Then
          Nextx = Nextx + 1
          ReDim Preserve PrintIdx(1 To Nextx) As Double
          PrintIdx(Nextx) = y 'Val(TempRec.CustomerNumber)
          Exit For
        End If
    Next y
  Next x
    
  OpenCustFile CustHandle
  
  frmBLShowPctComp.Label1 = "Reprinting Customer Business Licenses"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdReprint.Enabled = False
  
  For x = 1 To NumOfTempPrintRecs
    Get TPHandle, x, TempPrint
    If TempPrint.SeqNum < Val(fptxtFirstNum.Text) Or TempPrint.SeqNum > Val(fptxtLastNum.Text) Then GoTo Skip
'    Get TempPostHandle, x, TempPostRec
    Get TempPostHandle, PrintIdx(x), TempPostRec
'    Get CustHandle, TempPrint.RecNum, CustRec
    Get CustHandle, TempPostRec.CustomerNumber, CustRec
    If (QPTrim$(CustRec.IssueLicense) = "Y") And (QPTrim$(CustRec.Inactive) <> "Y") Then
      CustomerNumber = TempPostRec.CustomerNumber 'TempPrint.RecNum
      For ll = 1 To 5
        Print #RptHandle,
      Next ll
      DidCnt = DidCnt + 1
      Print #RptHandle, Tab(37 - tab1); QPTrim$(TempPrint.Head1)
      Print #RptHandle, Tab(37 - tab2); QPTrim$(TempPrint.Head2)
      Print #RptHandle, Tab(37 - Tab3); QPTrim$(TempPrint.Head3)
      Print #RptHandle, Tab(37 - Tab4); QPTrim$(TempPrint.Head4)
      Print #RptHandle, Tab(66); TempPrint.ThisYear
      If CustRec.Prorate < 100 Then
        Print #RptHandle, Tab(11); "Cust #"; Tab(19); QPTrim$(Using("####0", CustomerNumber)); Tab(26); "Fee prorated at " + CStr(CustRec.Prorate) + "%"
      Else
        Print #RptHandle, Tab(11); "Cust #"; Tab(19); QPTrim$(Using("####0", CustomerNumber))
      End If
      Print #RptHandle, Tab(11); QPTrim$(CustRec.BillName)
      Print #RptHandle, Tab(11); QPTrim$(CustRec.ADDRESS1); Tab(58); Using("#######0", TempPrint.LicNum) 'LICENSE#)
      Print #RptHandle, Tab(11); CustRec.ADDRESS2
      Print #RptHandle, Tab(11); RTrim$(CustRec.City); "  "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
      Print #RptHandle, Tab(55); TempPrint.Issue;
      Print #RptHandle, Tab(64); TempPrint.Expire
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle, Tab(11); QPTrim$(CustRec.CustName)
      Print #RptHandle,
      Print #RptHandle,
      SCnt = 23
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT1)) = 0 Then GoTo To2
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT1);
      If TempPrint.FeeYN = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC1);
        Print #RptHandle, Tab(62); Using("####0.00", TempPostRec.CatFee1)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC1)
      End If
      SCnt = SCnt + 1
To2:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT2)) = 0 Then GoTo To3
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT2);
      If TempPrint.FeeYN = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC2);
        Print #RptHandle, Tab(62); Using("####0.00", TempPostRec.CatFee2)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC2)
      End If
      SCnt = SCnt + 1
To3:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT3)) = 0 Then GoTo To4
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT3);
      If TempPrint.FeeYN = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC3);
        Print #RptHandle, Tab(62); Using("####0.00", TempPostRec.CatFee3)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC3)
      End If
      SCnt = SCnt + 1
To4:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT4)) = 0 Then GoTo To5
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT4);
      If TempPrint.FeeYN = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC4);
        Print #RptHandle, Tab(62); Using("####0.00", TempPostRec.CatFee4)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC4)
      End If
      SCnt = SCnt + 1
To5:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT5)) = 0 Then GoTo ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT5);
      If TempPrint.FeeYN = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC5);
        Print #RptHandle, Tab(62); Using("####0.00", TempPostRec.CatFee5)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC5)
      End If
      SCnt = SCnt + 1

ExitFormPrint1:
      If OldRound(TownRec.IssFee) > 0 And TempPrint.FeeYN = True Then
        Print #RptHandle, Tab(15); "ISSUE FEE"; Tab(62); Using("####0.00", OldRound(TownRec.IssFee))
        SCnt = SCnt + 1
      End If

      For LCnt = SCnt To 31
        Print #RptHandle,
      Next
      Print #RptHandle, ""

      For LCnt = 33 To 35
        Print #RptHandle, ""
      Next LCnt
      'Calc Total License Amount Here
      TotalBillAmt# = OldRound(TempPostRec.CatFee1 + TempPostRec.CatFee2 + TempPostRec.CatFee3 + TempPostRec.CatFee4 + TempPostRec.CatFee5)
      TotalBillAmt# = OldRound(TotalBillAmt# + TownRec.IssFee)

      If TempPrint.FeeYN = True Then
        Print #RptHandle, Tab(62); Using("####0.00", TotalBillAmt#) ' - OldRound(CustRec.AcctBal))
      Else
        Print #RptHandle, ""
      End If
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      If TempPrint.TBalYN = False Then
        Print #RptHandle, Tab(62); Using("####0.00", TotalBillAmt#)
      Else
        Print #RptHandle, Tab(62); Using("####0.00", TempPostRec.AcctBal)
      End If
      Print #RptHandle,
      Print #RptHandle, "~"
    End If
Skip:
    frmBLShowPctComp.ShowPctComp x, NumOfTempPrintRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdReprint.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdReprint.Enabled = True
  
  Print #RptHandle, Chr$(12);
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now

  ViewPrint ReportFile$, "Business License Reprinting", True
  KillFile ReportFile$
  
  MainLog ("Business license tractor fed forms reprinted.")
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLReprintLic", "PrintText", Erl)
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

