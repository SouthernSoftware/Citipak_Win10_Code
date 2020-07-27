VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBLReprintAppsRnwls 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Reprint Applications Renewals"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLReprintAppsRnwls.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5292
      Left            =   1848
      TabIndex        =   2
      Top             =   1770
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
      Picture         =   "frmBLReprintAppsRnwls.frx":08CA
      Begin EditLib.fpText fptxtFirstNum 
         Height          =   396
         Left            =   3024
         TabIndex        =   0
         Tag             =   $"frmBLReprintAppsRnwls.frx":08E6
         Top             =   2064
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
         Left            =   3072
         TabIndex        =   1
         Tag             =   $"frmBLReprintAppsRnwls.frx":09D7
         Top             =   3264
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
         Tag             =   $"frmBLReprintAppsRnwls.frx":0B25
         Top             =   4080
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
         ButtonDesigner  =   "frmBLReprintAppsRnwls.frx":0BF5
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   630
         Left            =   2880
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Applications' menu."
         Top             =   4080
         Width           =   2025
         _Version        =   131072
         _ExtentX        =   3572
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
         ButtonDesigner  =   "frmBLReprintAppsRnwls.frx":0DD8
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdReprint 
         Height          =   630
         Left            =   5130
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   $"frmBLReprintAppsRnwls.frx":0FB6
         Top             =   4080
         Width           =   2025
         _Version        =   131072
         _ExtentX        =   3572
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
         ButtonDesigner  =   "frmBLReprintAppsRnwls.frx":1075
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
         Top             =   4752
         Width           =   2100
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "First Application/Renewal Number"
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
         Left            =   2064
         TabIndex        =   5
         Top             =   1680
         Width           =   3804
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Reprint Applications/Renewals"
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
         Height          =   396
         Left            =   1776
         TabIndex        =   4
         Top             =   624
         Width           =   4572
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Last Application/Renewal Number"
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
         Left            =   2112
         TabIndex        =   3
         Top             =   2880
         Width           =   3804
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   2256
      TabIndex        =   8
      Top             =   7296
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5616
      Left            =   1716
      Top             =   1626
      Width           =   8220
   End
End
Attribute VB_Name = "frmBLReprintAppsRnwls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmBLIssueAppsLics.Show
  DoEvents
  Unload frmBLReprintAppsRnwls
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
'    fptxtFirstNum.ToolTipText = "Enter the first reprint reference number here."
'    fptxtLastNum.ToolTipText = "Enter the last reprint reference number here."
'    cmdExit.ToolTipText = "Press 'Cancel' to exit this screen."
'    cmdReprint.ToolTipText = "Press 'Start Reprint' to generate application reprints."
  End If
End Sub

Private Sub cmdReprint_Click()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim Code$, ll As Integer
  Dim Year$, FF$
  Dim AppFormat$
  Dim ReturnAdd$
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustSearchNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer, LCnt As Integer
  Dim cnt As Integer
  Dim ThisCode$, SCnt As Integer
  Dim TotalCust As Integer
  Dim ReportFile$, RptHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Lp As Integer
  Dim LicTotal#
  Dim CatCode$
  Dim ZZCnt As Integer
  Dim Snt&, Amt#
  Dim CODEDESC$
  Dim CodeType$
  Dim DESC1$
  Dim BaseAmt1#, BaseAmt2#, BaseAmt3#, BaseAmt4#, BaseAmt5#, BaseAmt6#
  Dim Revenue1#, Revenue2#, Revenue3#, Revenue4#, Revenue5#, Revenue6#
  Dim Percent1#, Percent2#, Percent3#, Percent4#, Percent5#, Percent6#
  Dim Maximum1#, Maximum2#, Maximum3#, Maximum4#, Maximum5#, Maximum6#
  Dim TempCustRec As TempCustRecType
  Dim TempHandle As Integer
  Dim NumOfRecs As Integer
  Dim FirstRec As Integer
  Dim LastRec As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AppType As Integer
  Dim TownLen As Integer
  Dim ThisTab As Integer
  Dim AddLen As Integer
  Dim CityLen As Integer
  Dim tab2 As Integer
  Dim Tab3 As Integer
  Dim Tab4 As Integer
  Dim LessBase$
  Dim Dash$
  Dim TLen As Integer
  Dim TT$
  Dim BaseFee$
  Dim MultiBY$
  Dim YrUpDown$(1 To 10)
  Dim IssFee#
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  IssFee# = TownRec.IssFee
  AppType = TownRec.AppForm
  
  For x = 1 To 10
    YrUpDown$(x) = "0000"
  Next x
  
  If QPTrim$(fptxtFirstNum.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid number in the 'First Application/Renewal Number' field."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtFirstNum.SetFocus
    Close
    Exit Sub
  End If
  
  If QPTrim$(fptxtLastNum.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid number in the 'Last Application/Renewal Number' field."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtLastNum.SetFocus
    Close
    Exit Sub
  End If
  
  FirstRec = Val(fptxtFirstNum.Text)
  LastRec = Val(fptxtLastNum.Text)
  If FirstRec > LastRec Then
    fptxtFirstNum.BackColor = 65535
    frmBLMessageBoxJr.Label1.Caption = "Please make sure the first number is less than the last number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtFirstNum.BackColor = -2147483643
    Exit Sub
  End If

  OpenTempCustRec TempHandle
  Get TempHandle, 1, TempCustRec
  AppType = TempCustRec.AppType
  NumOfRecs = LOF(TempHandle) / Len(TempCustRec)
  OpenCustFile CHandle
  ReportFile$ = "RPRTAPPS.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  If AppType > 1 Then
    OpenCatCodeFile CodeHandle
    NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
    If AppType = 2 Then
      GoSub PrintCustom2
    ElseIf AppType = 3 Then
      GoSub PrintCustom3
    ElseIf AppType = 4 Then
      GoSub PrintCustom4
    ElseIf AppType = 5 Then
      GoSub PrintCustom5
    ElseIf AppType = 6 Then
      GoSub PrintCustom6
    ElseIf AppType = 7 Then
      GoSub PrintCustom7
    ElseIf AppType = 8 Then
      GoSub PrintCustom8
    ElseIf AppType = 9 Then
      GoSub PrintCustom9
    End If
  Else
    GoSub PrintStandard
  End If
  
PrintCustom2:
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  ThisTab = TownLen / 2
  ThisTab = Abs(38 - ThisTab)
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
  
    Print #RptHandle, ""
    Print #RptHandle, Tab(30); "LICENSE APPLICATION"
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppTownOf); Tab(58); "ACCOUNT NO.    " + Using("####0", TempCustRec.CustRecNum);
    Print #RptHandle, Tab(55); "START DATE: "; Tab(67); QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", " + QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "APPLICANT'S NAME:    "; CustRec.BillName
    Print #RptHandle, Tab(5); "APPLICANT'S ADDRESS: "; QPTrim$(CustRec.ADDRESS1)
    Print #RptHandle, Tab(5); "                     "; QPTrim$(CustRec.ADDRESS2)
    Print #RptHandle, Tab(5); "                     "; QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + "  " + QPTrim$(CustRec.ZipCode)
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "TAX MAP______  BLOCK______ LOT______     ZONING DISTRICT_________________"
    Print #RptHandle, Tab(5); "FEDERAL ID/SS NUMBER__________________ " + QPTrim$(TownRec.AppState) + " TAX ID NUMBER__________________"
    Print #RptHandle, Tab(5); "TYPE OF BUSINESS:________________________________________________________"
    Print #RptHandle, Tab(5); "APPLICATION FOR:___ NEW___ RENEWAL___ GOING OUT OF BUSINESS(DATE)________"
    Print #RptHandle, Tab(5); "OWNERSHIP:___ CORPORATION___ PARTNERSHIP____ INDIVIDUAL-NO EMPLOYEES_____"
    Print #RptHandle, Tab(5); "NAME OF OWNER, PARTNER OR PRINCIPAL______________________________________"
    Print #RptHandle, Tab(5); "TELEPHONE NO. LOCAL:_____________ HOME:____________ EMERGENCY:___________"
    Print #RptHandle, Tab(5); "FAX NO._____________  E-MAIL:____________________________________________"
    Print #RptHandle,
    Print #RptHandle, Tab(5); "IS HAZARDOUS WASTE INVOLVED IN OPERATION? ____NO ____YES (ATTACH DETAILS)"
    Print #RptHandle, Tab(5); "CODE CLEARANCE: __ZONING ___INSPECTION __FIRE __HEALTH ___LAW ENFORCEMENT"
    Print #RptHandle,
    Print #RptHandle, Tab(28); "COMPUTATION OF LICENSE TAX"
    Print #RptHandle, Tab(5); "COMPUTE LICENSE TAX ACCORDING TO THE FOLLOWING SCHEDULE AND MAKE CHECKS"
    Print #RptHandle, Tab(5); "PAYABLE TO: "; QPTrim$(TownRec.AppTownOf) + "  DELIVER BY DUE DATE: "; Tab(61); QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay) + ", " + QPTrim$(YrUpDown(2))
    Print #RptHandle,
    Print #RptHandle, Tab(5); "GROSS INCOME FOR PRECEDING CALENDAR OR FISCAL YEAR....$_________________"
    Print #RptHandle, Tab(5); "LESS INCOME ON WHICH A LICENSE TAX WAS PAID TO ANOTHER"
    Print #RptHandle, Tab(5); "CITY OR COUNTY FOR OPERATIONS OUTSIDE CITY/COUNTY.....$_________________"
    Print #RptHandle, Tab(5); "BALANCE OF GROSS INCOME SUBJECT TO LICENSE TAX........$_________________"
    Print #RptHandle, Tab(5); "TAX:   RATE CLASS MINIMUM ON FIRST " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(1))) + ": " + QPTrim$(Using("$#,###,##0.00", TownRec.AppBaseFee(1))) + " PLUS"
    Print #RptHandle, Tab(5); QPTrim$(Using("$#,###,##0.00", TownRec.AppBaseFee(2))) + " PER " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(2))) + " FOR INCOME OVER " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(3)))
    Print #RptHandle, Tab(5); "[See declining rate schedule for over $1 million]       [OFFICE USE ONLY]"
    Print #RptHandle, Tab(5); "                           TOTAL LICENSE TAX $_________ [PAYMENT RECORD]"
    
    If QPTrim$(TempCustRec.AmtPct) = "Pct" Then
      Print #RptHandle, Tab(5); "PENALTY AFTER DUE DATE IS " + CStr(TownRec.AppPct) + "% PER MONTH $_________ [CHECK NO. ____________]"
    Else
      Print #RptHandle, Tab(5); "PENALTY AFTER DUE DATE IS " + QPTrim$(Using("$##,##0.00", TownRec.AppPct)) + " PER MONTH $_________ [CHECK NO. ____________]"
    End If
    
    Print #RptHandle, Tab(5); "TOTAL LICENSE TAX AND PENALTY $_________     [DATE RECEIVED____________]"
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(35); "CERTIFICATION"
    Print #RptHandle, Tab(5); "I (WE) DO CERTIFY THAT THE ABOVE INFORMATION AND AMOUNT RETURNED AS GROSS"
    Print #RptHandle, Tab(5); "INCOME FROM MY BUSINESS IS TRUE AND CORRECT. AND I HAVE MADE NO DEDUCTIONS"
    Print #RptHandle, Tab(5); "EXCEPT INCOME ON WHICH I HAVE PAID BUSINESS LICENSE TAX TO ANOTHER CITY OR"
    Print #RptHandle, Tab(5); "COUNTY, FOR WHICH I HAVE PROOF OF PAYMENT. I AM FAMILIAR WITH THE PENALTY"
    Print #RptHandle, Tab(5); "PROVISIONS OF THE ORDINANCE AND GROUNDS FOR LICENSE REVOCATION, INCLUDING"
    Print #RptHandle, Tab(5); "MAKING FALSE OR FRAUDULENT STATEMENTS IN THIS APPLICATION. I CERTIFY THAT"
    Print #RptHandle, Tab(5); "ALL BUSINESS PERSONAL PROPERTY TAXES AND PAYABLES DUE TO THE CITY/COUNTY"
    Print #RptHandle, Tab(5); "HAVE BEEN PAID, AND THAT THE ABOVE BUSINESS NAME IS THE SAME AS REPORTED"
    Print #RptHandle, Tab(5); "ON DOCUMENTS FILED WITH THE STATE AND FEDERAL GOVERNMENTS. I UNDERSTAND MY"
    Print #RptHandle, Tab(5); "BUSINESS INCOME TAX RETURNS AND OTHER DOCUMENTS MAY BE INSPECTED TO VERIFY"
    Print #RptHandle, Tab(5); "GROSS INCOME OR OTHER BUSINESS DATA."
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(5); "___________________________________________________________________________"
    Print #RptHandle, Tab(5); "SIGNATURE                          TITLE                           DATE"
    Print #RptHandle, Chr$(12);

  Next cnt

  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  MainLog ("Application #2 reprinted.")
  
  Exit Sub

'-----------------------------------------------------
PrintCustom3:
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  ThisTab = TownLen / 2
  ThisTab = Abs(38 - ThisTab)
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
    Print #RptHandle, ""
    Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.AppTownOf) '"TOWN OF RIVERSIDE"
    Print #RptHandle, Tab(24); "BUSINESS LICENSE APPLICATION"
    Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Business Name: "; CustRec.CustName; CustRec.CustNumb
    Print #RptHandle, Tab(5); "              -----------------------------------------------------------"
    Print #RptHandle, Tab(5); "Street Address of Business: "
    Print #RptHandle, Tab(5); "                           ----------------------------------------------"
    Print #RptHandle, Tab(5); "Zoning of Business Location: "
    Print #RptHandle, Tab(5); "                            ---------------------------------------------"
    Print #RptHandle, Tab(5); "Telephone Number: "
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, Tab(5); "Applicant's Name: "; CustRec.BillName
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "Applicant's Address: "; CustRec.ADDRESS1
    Print #RptHandle, Tab(5); "                    -----------------------------------------------------"
    If QPTrim$(CustRec.WPHONE) = "(" Then CustRec.WPHONE = ""
    Print #RptHandle, Tab(23); QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); Tab(57); "Phone: "; QPTrim$(CustRec.WPHONE)
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Rem 22 lines printed here
    Print #RptHandle, Tab(5); "TYPE OF BUSINESS LICENSE APPLYING FOR:"
    Print #RptHandle, Tab(5); ""
    If TownRec.IssFee > 0 Then
      Print #RptHandle, Tab(5); "_______ Contracting or Construction " + QPTrim(Using("$#,##0.00", TownRec.AppBaseFee(1))) + " plus " + QPTrim$(Using("$#,##0.00", TownRec.IssFee)) + " Issuance Fee."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Retail Sales " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + " plus " + QPTrim$(Using("##0", TownRec.AppNumer)) + "/" + QPTrim$(Using("##0", TownRec.AppDenom)) + " of " + QPTrim$(Using("##0%", (TownRec.AppGrsPct / 100))) + " of gross receipts"
      Print #RptHandle, Tab(5); "        over " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1))) + " plus " + QPTrim$(Using("$#, ##0.00", TownRec.IssFee)) + " Issuance Fee."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Financial, Real Estate or Professional Service " + QPTrim(Using("$#,##0.00", TownRec.AppBaseFee(3))) + " plus " + QPTrim$(Using("$#,##0.00", TownRec.IssFee))
      Print #RptHandle, Tab(5); "        Issuance Fee."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Repair, Personal, Business or Delivery Service " + QPTrim(Using("$#,##0.00", TownRec.AppBaseFee(4))) + " plus " + QPTrim$(Using("$#,##0.00", TownRec.IssFee))
      Print #RptHandle, Tab(5); "        Issuance Fee."
    Else
      Print #RptHandle, Tab(5); "_______ Contracting or Construction: " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + "."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Retail Sales " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + " plus " + QPTrim$(Using("##0", TownRec.AppNumer)) + "/" + QPTrim$(Using("##0", TownRec.AppDenom)) + " of " + QPTrim$(Using("##0%", (TownRec.AppGrsPct / 100))) + " of gross receipts"
      Print #RptHandle, Tab(5); "        over " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1))) + "."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Financial, Real Estate or Professional Service: " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + "."
      Print #RptHandle, Tab(5);
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Repair, Personal, Business or Delivery Service: " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + "."
      Print #RptHandle, Tab(5);
    End If
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Other (Specify) ______________________________________________"
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "Estimate of ______________ gross receipts or preceding year's gross "
    Print #RptHandle, Tab(5); "receipts ______________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "AMOUNT OF LICENSE TAX FOR " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", THROUGH " + QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay) + ", " + QPTrim$(YrUpDown(2)) + " IS:$_______"
    Print #RptHandle, Tab(5); "ANY SPECIAL CONDITIONS OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED "
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "ACTIVITY SHALL BE CONDUCTED: ____________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "I certify that the statements and figures set forth on this application"
    Print #RptHandle, Tab(5); "are true to the best of my knowledge."
    Print #RptHandle, Tab(5); "                                      ___________________________________"
    Print #RptHandle, Tab(5); "                                            Signature of Applicant"
    Print #RptHandle, Tab(5); ""
    If TempCustRec.AmtPct = "Pct" Then
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("##0%", (TownRec.AppPct / 100))) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    Else
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("$##,##0.00", TownRec.AppPct)) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    End If
    Print #RptHandle, Tab(5);
    Print #RptHandle, Tab(5); "Return Application and Fee to:"
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppAdd1)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)
    Print #RptHandle, Chr$(12);
'    End If
  Next cnt

  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  
  KillFile ReportFile$
  
  MainLog ("Application #3 reprinted.")
  
  Exit Sub
'-----------------------------------------------------

PrintCustom4:
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  
  ThisTab = TownLen / 2 'start centering process
  ThisTab = Abs(39 - ThisTab) '63 = end of line 2...63 - 16 = 47 length of line from beginning tab
  '...47/2 = 23.5 ...+ 16 = middle point of line...round down to 39
  
  AddLen = Len(QPTrim$(TownRec.TownAdd1))
  CityLen = Len(QPTrim$(TownRec.City) + ", " + QPTrim$(TownRec.State) + "  " + QPTrim$(TownRec.ZipCode))
  
  tab2 = TownLen / 2
  tab2 = Abs(38 - tab2) '38 = mid point of line 3
  Tab3 = AddLen / 2
  Tab3 = Abs(38 - Tab3)
  Tab4 = CityLen / 2
  Tab4 = Abs(38 - Tab4)
  
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(4)) = "Curr" Then
    YrUpDown(4) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(4)) = "+1" Then
    YrUpDown(4) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(4)) = "-1" Then
    YrUpDown(4) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(5)) = "Curr" Then
    YrUpDown(5) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(5)) = "+1" Then
    YrUpDown(5) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(5)) = "-1" Then
    YrUpDown(5) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(6)) = "Curr" Then
    YrUpDown(6) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(6)) = "+1" Then
    YrUpDown(6) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(6)) = "-1" Then
    YrUpDown(6) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(7)) = "Curr" Then
    YrUpDown(7) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(7)) = "+1" Then
    YrUpDown(7) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(7)) = "-1" Then
    YrUpDown(7) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(8)) = "Curr" Then
    YrUpDown(8) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(8)) = "+1" Then
    YrUpDown(8) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(8)) = "-1" Then
    YrUpDown(8) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(9)) = "Curr" Then
    YrUpDown(9) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(9)) = "+1" Then
    YrUpDown(9) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(9)) = "-1" Then
    YrUpDown(9) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(10)) = "Curr" Then
    YrUpDown(10) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(10)) = "+1" Then
    YrUpDown(10) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(10)) = "-1" Then
    YrUpDown(10) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
      Print #RptHandle, "" '33
      Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppTownOf)
      Print #RptHandle, Tab(16); "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE" 'line 2
      Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1)); Tab(70); "PAGE 2"
      Print #RptHandle, ""
      Print #RptHandle, Tab(2); "Dear Business Owner:"
      Print #RptHandle,
      Print #RptHandle, Tab(2); "     For the purpose of computing Business, Professional and Occupational"
      Print #RptHandle, Tab(2); "License (BPOL) Tax promulgated by Virginia Code Section 58.1-3700 et seq."
      Print #RptHandle, Tab(2); "and " + QPTrim$(TownRec.AppCity) + " Town Ordinance #" + QPTrim$(TownRec.AppCityOrd) + " adopted "
      Print #RptHandle, Tab(2); MakeRegDate(TownRec.AppAdoptDate) + " please complete and return this form with the required"
      Print #RptHandle, Tab(2); "information no later than " + QPTrim$(TownRec.AppDiscMonth) + " " + CStr(TownRec.AppDiscDay) + ", " + QPTrim$(YrUpDown(2)) + "."
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle, Tab(2); "Respectfully,"
      Print #RptHandle, Tab(2); QPTrim$(TownRec.AppTownOf)
      Print #RptHandle, Tab(2); QPTrim$(TownRec.AppMayorCouncil)
      Print #RptHandle, Tab(2); String$(76, "-")
      Print #RptHandle, Tab(2);
      Print #RptHandle, Tab(2);
      Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppTownOf)
      Print #RptHandle, Tab(Tab3); QPTrim$(TownRec.TownAdd1)
      Print #RptHandle, Tab(Tab4); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)
      Print #RptHandle, Tab(24); "Application for Town Licenses" 'line 3
      Print #RptHandle,
      Print #RptHandle, Tab(2); "For period beginning " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", " + QPTrim$(YrUpDown(3)) + " (or start of business in " + QPTrim$(YrUpDown(4))
      Print #RptHandle, Tab(2); "and ending " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(5))
      Print #RptHandle, Tab(2);
      Print #RptHandle, Tab(2); "NAME OF APPLICANT: "; QPTrim$(CustRec.BillName)
      Print #RptHandle, Tab(2); "      TRADING AS : "; QPTrim$(CustRec.CustName)
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "BUSINESS ADDRESS:"; Tab(40); "HOME ADDRESS"
      Print #RptHandle, Tab(2); "MAIL: "; QPTrim$(CustRec.ADDRESS1); Tab(40); "MAIL: ________________________________"
      Print #RptHandle, Tab(8); QPTrim$(CustRec.ADDRESS2)
      Print #RptHandle, Tab(8); RTrim$(CustRec.City); " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); Tab(40); "      ________________________________"
      Print #RptHandle, Tab(2); "911:  ______________________________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(8); "______________________________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "PHONE: _________________________"; Tab(40); "PHONE: ______________________________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "A SEPARATE LICENSE WILL BE ISSUED FOR EACH TYPE OF BUSINESS"
      Print #RptHandle, Tab(2); "PERFORMED, AS REQUIRED PER THE " + UCase(QPTrim$(TownRec.AppCityOrd)) + ".  THIS WILL NOT"
      Print #RptHandle, Tab(2); "RESULT IN ANY ADDITONAL COST TO BUSINESSES.  PLEASE REPORT GROSS"
      Print #RptHandle, Tab(2); "RECEIPTS FOR EACH CLASSIFICATION THAT APPLIES TO YOUR BUSINESS."
      Print #RptHandle, Chr$(12);
      Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppTownOf)
      Print #RptHandle, Tab(16); "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE"
      Print #RptHandle, Tab(31); "For Year: " + QPTrim$(YrUpDown(1)); Tab(70); "PAGE 1"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "WHOLESALE MERCHANT:"
      Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppWholeMonth) + "-" + CStr(TownRec.AppWholeDay) + "-" + QPTrim(YrUpDown(6)) + " as shown by applicants records"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); Tab(60); "$_______________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "RETAIL MERCHANT:"
      Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppRetailMonth) + "-" + CStr(TownRec.AppRetailDay) + "-" + QPTrim(YrUpDown(7)) + " as shown by applicants records"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); Tab(60); "$_______________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "FINANCIAL, REAL ESTATE AND PROFESSIONAL:"
      Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppFinMonth) + "-" + CStr(TownRec.AppFinDay) + "-" + QPTrim(YrUpDown(8)) + " as shown by applicants records"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); Tab(60); "$_______________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "CONTRACTING:"
      Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppContMonth) + "-" + CStr(TownRec.AppContDay) + "-" + QPTrim(YrUpDown(9)) + " as shown by applicants records"
      Print #RptHandle, Tab(2); "Subject to Virginia Code Sec 58.1-3715)"; Tab(60); "$_______________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "REPAIR, PERSONAL or BUSINESS SERVICES:"
      Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppRepairMonth) + "-" + CStr(TownRec.AppRepairDay) + "-" + QPTrim(YrUpDown(10)) + " as shown by applicants records"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); Tab(60); "$_______________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "If uncertain of your business classification(s), please call the Town Office at"
      Print #RptHandle, Tab(2); QPTrim$(TownRec.AppPhone) + " for assistance."
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "I do affirm that the foregoing figures are true, complete and accurate"
      Print #RptHandle, Tab(2); "to the best of my knowledge."
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(35); "Signature ___________________________________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(34); "Print Name ___________________________________"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(24); "*** IMPORTANT ***"
      Print #RptHandle, Tab(2); ""
      Print #RptHandle, Tab(2); "APPLICATION MUST BE RETURNED PRIOR TO " + UCase(QPTrim$(TownRec.AppFiscMonth)) + " " + CStr(TownRec.AppFiscDay) + " OF EACH YEAR"
      Print #RptHandle, Tab(2); "TO AVOID PENALTY. LICENSE FEES ARE DUE PRIOR TO " + UCase(QPTrim$(TownRec.AppLicRetMonth)) + " " + CStr(TownRec.AppLicRetDay)
      Print #RptHandle, Tab(2); "OF EACH YEAR TO AVOID PENALTY AND INTEREST. INTENTIONALLY PROVIDING"
      Print #RptHandle, Tab(2); "INSUFFICIENT OR INACCURATE INFORMATION MAY RESULT IN LEGAL RECOURSE"
      Print #RptHandle, Tab(2); "BY THE TOWN OF " + UCase(QPTrim$(TownRec.AppCity)) + " AS SET FORTH BY VIRGINIA CODE."
      Print #RptHandle, Tab(2); Chr$(12);
'    End If
  Next cnt
  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  
  MainLog ("Application #4 reprinted.")
  
  Exit Sub
'-----------------------------------------------------

PrintCustom5:
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))

  ThisTab = TownLen / 2
  ThisTab = Abs(38 - ThisTab)
  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
    Print #RptHandle, ""
    Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(24); "BUSINESS LICENSE APPLICATION"
    Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Business Name: "; QPTrim$(CustRec.CustName)
    Print #RptHandle, Tab(5); "              -----------------------------------------------------------"
    Print #RptHandle, Tab(5); "Street Address of Business: "
    Print #RptHandle, Tab(5); "                           ----------------------------------------------"
    Print #RptHandle, Tab(5); "Zoning of Business Location: "
    Print #RptHandle, Tab(5); "                            ---------------------------------------------"
    Print #RptHandle, Tab(5); "Telephone Number: "
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, Tab(5); "Applicant's Name: "; QPTrim$(CustRec.BillName)
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "Applicant's Address: "; QPTrim$(CustRec.ADDRESS1)
    Print #RptHandle, Tab(5); "                    -----------------------------------------------------"
    Print #RptHandle, Tab(5); "Telephone Number: "; QPTrim$(CustRec.WPHONE)
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Rem 22 lines printed here
    Print #RptHandle, Tab(5); "TYPE OF BUSINESS LICENSE APPLYING FOR:"
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Contracting or Construction " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + " or " + QPTrim$(Using("#.###", TownRec.AppCentsPer(1))) + " cents per " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1)))
    Print #RptHandle, Tab(5); "           gross receipts whichever is greater."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_______ Retail Sales " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + " or " + QPTrim$(Using("#.###", TownRec.AppCentsPer(2))) + " cents per " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(2))) + " whichever"
    Print #RptHandle, Tab(5); "           is greater."
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Financial, Real Estate or Professional Service " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + " or "
    Print #RptHandle, Tab(5); "           " + QPTrim$(Using("#.###", TownRec.AppCentsPer(3))) + " cents per " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(3))) + " whichever is greater."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_______ Repair, Personal or Business Service " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + " or " + QPTrim$(Using("#.###", TownRec.AppCentsPer(4))) + " cents per"
    Print #RptHandle, Tab(5); "           " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(4))) + " whichever is greater."
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Other (Specify) ______________________________________________"
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "Estimate of ______________ gross receipts or preceding year's gross "
    Print #RptHandle, Tab(5); "receipts ______________________. Enclose copy of most recent schedule C"
    Print #RptHandle, Tab(5); "or other comparable federal document."
    Print #RptHandle, Tab(5); "AMOUNT OF LICENSE TAX FOR " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", THROUGH " + QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay) + ", " + QPTrim(YrUpDown(2)) + " IS:$_______"
    Print #RptHandle, Tab(5); "ANY SPECIAL CONDITIONS OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED "
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "ACTIVITY SHALL BE CONDUCTED: ____________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "I certify that the statements and figures set forth on this application"
    Print #RptHandle, Tab(5); "are true to the best of my knowledge."
    Print #RptHandle, Tab(5); "                                      ___________________________________"
    Print #RptHandle, Tab(5); "                                            Signature of Applicant"
    Print #RptHandle, Tab(5); ""
    If TempCustRec.AmtPct = "Pct" Then
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("##0%", (TownRec.AppPct / 100))) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    Else
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("$##,##0.00", TownRec.AppPct)) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    End If
    Print #RptHandle, Tab(5);
    Print #RptHandle, Tab(5); "Return Application and Fee to:"
    Print #RptHandle, Tab(5); QPTrim$(TownRec.TownName)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.TownAdd1)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.City) + ", " + QPTrim$(TownRec.State) + "  " + QPTrim$(TownRec.ZipCode)
    Print #RptHandle, Chr$(12);
'    End If
  Next cnt
  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  
  MainLog ("Application #5 reprinted.")
  
  Exit Sub

'-----------------------------------------------------
Return

PrintCustom6:
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  Dash$ = String$(30, "_")
  MultiBY$ = CStr(TownRec.AppPct)
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  ThisTab = TownLen / 2
  ThisTab = Abs(39 - ThisTab)
  AddLen = Len(QPTrim$(TownRec.AppAdd1))
  CityLen = Len(QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip))
  
  tab2 = TownLen / 2
  tab2 = Abs(39 - tab2)
  Tab3 = AddLen / 2
  Tab3 = Abs(39 - Tab3)
  Tab4 = CityLen / 2
  Tab4 = Abs(39 - Tab4)
  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
    Print #RptHandle, ""
    Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.AppTownOf) '"TOWN OF ELLOREE"
    Print #RptHandle, Tab(Tab3); QPTrim$(TownRec.AppAdd1) '"P.O. BOX 28"
    Print #RptHandle, Tab(Tab4); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip) '"ELLOREE, S.C. 29047"
    Print #RptHandle, Tab(20); "APPLICATION FOR BUSINESS LICENSE FOR YEAR " + QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); CustRec.BillName
    Print #RptHandle, Tab(5); CustRec.ADDRESS1
    Print #RptHandle, Tab(5); CustRec.ADDRESS2
    Print #RptHandle, Tab(5); QPTrim$(CustRec.City); ", "; CustRec.State; " "; CustRec.ZipCode
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "To engage in business or profession, make a separate application"
    Print #RptHandle, Tab(5); "for each business and each location.  Send fee with application to"
    Print #RptHandle, Tab(5); "The " + QPTrim$(TownRec.AppTownOf) + ":" 'Town of Elloree:"
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "         Owners Name:______________________________________________"
    Print #RptHandle, Tab(5); "Business Description:______________________________________________"
    Print #RptHandle, Tab(5); "      Business Phone:______________________________________________"
    Print #RptHandle, Tab(5); "   Federal ID Number:______________________________________________"
    Print #RptHandle, Tab(5); "     State ID Number:______________________________________________"
    Print #RptHandle, Tab(5); "___________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "To calculate your " + QPTrim$(TownRec.AppTownOf) + " Business License Fee, Use the"
    Print #RptHandle, Tab(5); "formula below."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "1.  Gross Sales"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "2.  Less Base Amount"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "3.  Excess Gross"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "4.  Base Rate Fee"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "5.  If No. 3 is Greater than"
    Print #RptHandle, Tab(5); "    Zero, divide No. 3 by 1,000"
    Print #RptHandle, Tab(5); "    and round UP"; Tab(40); Dash$
    Print #RptHandle,
    Print #RptHandle, Tab(5); "6.  Multiply #5 by "; MultiBY$; Tab(40); Dash$
    Print #RptHandle,
    Print #RptHandle, Tab(5); "7.  Total License Fee # 4 + # 6"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "8.  Add  penalty (" + QPTrim$(Using("$##0.00", TownRec.AppColFee)) + " Collector's"
    If TempCustRec.AmtPct = "Pct" Then
      Print #RptHandle, Tab(5); "    Fee and " + QPTrim$(Using("#0.00%", TownRec.AppGrsPct / 100)) + " per month after"
    Else
      Print #RptHandle, Tab(5); "    Fee and " + QPTrim$(Using("$##,##0.00", TownRec.AppGrsPct)) + " per month after"
    End If
    Print #RptHandle, Tab(9); QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay); Tab(40); Dash$
    Print #RptHandle, Tab(5); "9.  TOTAL DUE (# 7 + # 8)"; Tab(40); Dash$
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(5); "This is to certify that the amount of total gross for the business"
    Print #RptHandle, Tab(5); "transacted at or through the above location for the calendar year"
    Print #RptHandle, Tab(5); "ending " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", or the last complete fiscal year is true and"
    Print #RptHandle, Tab(5); "correct, and that this report corresponds with the amount that was"
    Print #RptHandle, Tab(5); "reported to the SC Tax Commission or Insurance Commission and with"
    Print #RptHandle, Tab(5); "the Internal Revenue Service."
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); Dash$; Tab(40); Dash$
    Print #RptHandle, Tab(5); "Firm Name/ Individual Signature"; Tab(40); "By:"
    Print #RptHandle, Chr$(12);
  Next cnt
    
  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  
  MainLog ("Application #6 reprinted.")
  
  Exit Sub

'-----------------------------------------------------
PrintCustom7:
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.TownName))
  AddLen = Len(QPTrim$(TownRec.TownAdd1))
  CityLen = Len(QPTrim$(TownRec.City) + ", " + QPTrim$(TownRec.State) + "  " + QPTrim$(TownRec.ZipCode))
  
  tab2 = TownLen / 2
  tab2 = Abs(38 - tab2)
  Tab3 = AddLen / 2
  Tab3 = Abs(38 - Tab3)
  Tab4 = CityLen / 2
  Tab4 = Abs(38 - Tab4)

  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
      Print #RptHandle, ""
      Print #RptHandle, Tab(tab2); QPTrim$(TownRec.TownName) '"TOWN OF STEPHENS CITY"
      Print #RptHandle, Tab(24); "BUSINESS LICENSE APPLICATION"
      Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1))
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Please print or type:"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Applicant Name: "; CustRec.BillName; Tab(58); "Phone:"
      Print #RptHandle, Tab(5); "               ----------------------------------------------------------"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Trade Name: "; Tab(54); "FEIN or SS#"
      Print #RptHandle, Tab(5); "           --------------------------------------------------------------"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Mailing Address:                            Physical Address:"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "------------------------------------------  -----------------------------"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "------------------------------------------  -----------------------------"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "------------------------------------------  -----------------------------"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Phone:                                      Phone:"
      Print #RptHandle, Tab(5); "      ------------------------------------        -----------------------"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Nature Of Business:"
      Print #RptHandle, Tab(5); "                   ------------------------------------------------------"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Gross receipts                 Estimated                         Actual"
      Print #RptHandle, Tab(5); "for year ending"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppFiscMonth) + " " + CStr(TownRec.AppFiscDay) + ", " + QPTrim$(YrUpDown(2)); Tab(28); "       -----------                     -----------"
      Print #RptHandle, Tab(5); "(Wholesalers Only...Enter Purchases)"
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "CONTRACTORS ONLY"
      Print #RptHandle, Tab(5); "Please Note: All contractors must have valid Workmans Compensation coverage"
      Print #RptHandle, Tab(5); "in effect for the time period covered by this license. Failure to have"
      Print #RptHandle, Tab(5); "proper coverage will cause your license to be revoked."
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "____ I certify that I am in compliance with the provisions of the Virginia"
      Print #RptHandle, Tab(5); "Workmans Compensation Act, and I will notify the " + QPTrim$(TownRec.AppTownOf)
      Print #RptHandle, Tab(5); "if this coverage lapses during the period that this license is in effect."
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "I hereby swear (or affirm) that the statements are true, full and correct to"
      Print #RptHandle, Tab(5); "the best of my knowledge."
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "___________________________________________              ________________"
      Print #RptHandle, "                    Signature                                      Date      "
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "*************************************************************************"
      Print #RptHandle, Tab(5); "FOR OFFICE USE ONLY"
      Print #RptHandle, Tab(5); "Zoning classification approved for this type of business"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Approved by ______________________________              ________________"
      Print #RptHandle, "                          Signature                               Date      "
  
      Print #RptHandle, Chr$(12);
  Next cnt
    
  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  
  MainLog ("Application #2 reprinted.")
 
  Exit Sub

'-----------------------------------------------------
PrintCustom8:
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  AddLen = Len(QPTrim$(TownRec.AppAdd1))
  CityLen = Len(QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip))
  
  ThisTab = TownLen / 2
  ThisTab = Abs(41 - ThisTab)
  tab2 = Len(QPTrim$(TownRec.AppMayorCouncil))
  tab2 = tab2 / 2
  tab2 = Abs(41 - tab2)

  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
      Print #RptHandle, ""
      Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.AppTownOf)  '"CITY OF ATMORE"
      Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppMayorCouncil)
      Print #RptHandle, ""
      Print #RptHandle, Tab(60); "Date: "; MakeRegDate(TempCustRec.MiscNum)
      Print #RptHandle, ""
      Print #RptHandle, Tab(2); "NOTICE FOR RENEWAL OF BUSINESS LICENSE FOR PERIOD ENDING: " + QPTrim$(TownRec.AppFiscMonth) + ", " + QPTrim$(YrUpDown(1))
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Business Account # "; TempCustRec.CustRecNum
      Print #RptHandle, Tab(5); CustRec.BillName
      Print #RptHandle, Tab(5); CustRec.ADDRESS1
      Print #RptHandle, Tab(5); CustRec.ADDRESS2
      Print #RptHandle, Tab(5); RTrim$(CustRec.City); " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode)
      Print #RptHandle, ""
      Print #RptHandle, String$(79, "-")
      Print #RptHandle, Tab(2); "Code"; Tab(9); "Type of License"
      Print #RptHandle, String$(79, "-")
      Lp = 17
'-----------------------------------------------------------
      CatCode$ = QPTrim$(CustRec.BILLCAT1)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT1;
      Print #RptHandle, Tab(9); CustRec.DESC1; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1

'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT2)) = 0 Then GoTo EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT2)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT2;
      Print #RptHandle, Tab(9); CustRec.DESC2; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT3)) = 0 Then GoTo EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT3)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT3;
      Print #RptHandle, Tab(9); CustRec.DESC3; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT4)) = 0 Then GoTo EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT4)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT4;
      Print #RptHandle, Tab(9); CustRec.DESC4; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
      If Lp >= 54 Then
        GoSub PrintHeader8
      End If
      
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT5)) = 0 Then GoTo EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT5)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT5;
      Print #RptHandle, Tab(9); CustRec.DESC5; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$#,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$#,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$#,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$#,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$#,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$#,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
EndAtmore1:
      If Lp >= 36 Then
        GoSub PrintHeader8
      End If
  
      Print #RptHandle,
      Print #RptHandle, Tab(5); "Make Checks Payable To:"; Tab(45); "License Total: _________________"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppTownOf); Tab(45); "Penalty:       _________________"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppAdd1); Tab(45); "Interest:      _________________"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppCity); Tab(45); "Issue Fee: "; Tab(67); Using("$##0.00", TownRec.IssFee) + " " + QPTrim$(TownRec.SpareSpace)
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip); Tab(45); "               -----------------"
      Print #RptHandle, Tab(5); ""; Tab(45); "Total Due:     _________________"
      Print #RptHandle,
      Print #RptHandle, Tab(5); "License renewals are due " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + " and delinquent after ";
      Print #RptHandle, QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay)
      
      If QPTrim$(TempCustRec.AmtPct) = "Pct" Then
        If TownRec.AppGrsPct = 8 Or TownRec.AppGrsPct = 11 Then
          Print #RptHandle, Tab(5); "at which time an ";
        Else
          Print #RptHandle, Tab(5); "at which time a ";
        End If
        
        Print #RptHandle, CStr(TownRec.AppGrsPct) + "% penalty will be charged. Renewals after " + QPTrim(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay)
        
        If TownRec.AppDiscPct = 8 Or TownRec.AppDiscPct = 11 Then
          Print #RptHandle, Tab(5); "will be charged an " + CStr(TownRec.AppDiscPct) + "% penalty. If you have any questions regarding this"
        Else
          Print #RptHandle, Tab(5); "will be charged a " + CStr(TownRec.AppDiscPct) + "% penalty. If you have any questions regarding this"
        End If
      Else 'amount
        Print #RptHandle, Tab(5); "at which time a ";
        Print #RptHandle, QPTrim$(Using$("$##,##0.00", TownRec.AppGrsPct)) + " penalty will be charged. Renewals after " + QPTrim(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay)
        Print #RptHandle, Tab(5); "will be charged a " + QPTrim$(Using("$##,##0.00", TownRec.AppDiscPct)) + " penalty. If you have any questions regarding this"
      End If
      
      
      Print #RptHandle, Tab(5); "notice, please call " + QPTrim$(TownRec.AppPhone) + "."
      Print #RptHandle,
      Print #RptHandle, Tab(10); "RENEWALS THAT DO NOT CONTAIN SIGNATURE AND GROSS RECEIPTS"
      Print #RptHandle, Tab(10); "(WHERE REQUIRED) WILL NOT BE PROCESSED."
      Print #RptHandle,
      Print #RptHandle, Tab(5); "I CERTIFY THAT THE ABOVE INFORMATION IS CORRECT"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "NAME ________________________________ TITLE ________________________"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "SUBSCRIBED AND SWORN TO BEFORE ME THIS ______ DAY OF ______, ______."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "NOTARY PUBLIC ____________________________________________"
      Print #RptHandle,
      Print #RptHandle,
  
      Print #RptHandle, Chr$(12);
NotNow8:
  Next cnt
    
  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  
  MainLog ("Application #8 reprinted.")
  
  Exit Sub
  
PrintHeader8:
  Print #RptHandle, Chr$(12)
  Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.AppTownOf)  '"CITY OF ATMORE"
  Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppMayorCouncil)
  Print #RptHandle,
  Lp = 3
  
  Return
'-----------------------------------------------------
PrintCustom9:
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = TempCustRec.ThisYear
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(TempCustRec.ThisYear) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  
  tab2 = TownLen / 2
  tab2 = Abs(42 - tab2)

  For cnt = FirstRec To LastRec
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppTownOf) ' "TOWN OF HEMINGWAY, SOUTH CAROLINA"
      Print #RptHandle, Tab(26); "   BUSINESS LICENSE APPLICATION"
      Print #RptHandle, Tab(26); "         For Year: "; QPTrim$(YrUpDown(1))
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(3); "Business Name: "; QPTrim$(CustRec.CustName)
      Print #RptHandle, Tab(3); "              -------------------------------------------------------------"
      Print #RptHandle, Tab(3); "Mailing Address: "; QPTrim$(CustRec.ADDRESS1)
      Print #RptHandle, Tab(3); "                 "; QPTrim$(CustRec.City); " "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
      Print #RptHandle, Tab(3); "                -----------------------------------------------------------"
      Print #RptHandle, Tab(3); "Business Address: "
      Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
      Print #RptHandle, Tab(3); "Telephone Number: "
      Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
      Print #RptHandle, Tab(3); "Type of Business: "
      Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
      Print #RptHandle, Tab(3); "Social Security Number:"
      Print #RptHandle, Tab(3); "                       ----------------------------------------------------"
      Print #RptHandle, Tab(3); "Federal Identification Number: "
      Print #RptHandle, Tab(3); "                              ---------------------------------------------"
      Print #RptHandle, Tab(3); "Gross Income Previous Year:"
      Print #RptHandle, Tab(3); "                           ------------------------------------------------"
      Print #RptHandle, Tab(3); "License as Calculated:"
      Print #RptHandle, Tab(3); "                      -----------------------------------------------------"
      If QPTrim$(TempCustRec.AmtPct) = "Pct" Then
        Print #RptHandle, Tab(3); QPTrim$(Using("##0", TownRec.AppDiscPct)) + "% Discount, If Paid by " + QPTrim$(TownRec.AppDiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppDiscDay)) + ":"
        Print #RptHandle, Tab(3); "                                   ----------------------------------------"
        Print #RptHandle, Tab(3); QPTrim$(Using("##0", TownRec.AppPct)) + "% Penalty Per Month After " + QPTrim$(TownRec.AppPenMonth) + " " + QPTrim$(Using("#0", TownRec.AppPenDay)) + ":"
      Else
        Print #RptHandle, Tab(3); QPTrim$(Using("$##,##0.00", TownRec.AppDiscPct)) + " Discount, If Paid by " + QPTrim$(TownRec.AppDiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppDiscDay)) + ":"
        Print #RptHandle, Tab(3); "                                   ----------------------------------------"
        Print #RptHandle, Tab(3); QPTrim$(Using("$##,##0.00", TownRec.AppPct)) + " Penalty Per Month After " + QPTrim$(TownRec.AppPenMonth) + " " + QPTrim$(Using("#0", TownRec.AppPenDay)) + ":"
      End If
      Print #RptHandle, Tab(3); "                                 ------------------------------------------"
      Print #RptHandle, Tab(3); "TOTAL AMOUNT DUE: "
      Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
      Print #RptHandle, Tab(3); ""
      Print #RptHandle, Tab(3); ""
      Print #RptHandle, Tab(3); "   This is to certify that the above is a true statement of the business"
      Print #RptHandle, Tab(3); "transacted at or through the above location for the calender year ending"
      Print #RptHandle, Tab(3); QPTrim$(TownRec.AppFiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppFiscDay)) + ", " + QPTrim$(YrUpDown(2)); ", and that the report corresponds with the records with"
      Print #RptHandle, Tab(3); "the S.C. Tax Commission of Insurance Commissioner and with the Collector of"
      Print #RptHandle, Tab(3); "Internal Revenue of the United States. I understand that the Town Ordinance"
      Print #RptHandle, Tab(3); "provides for penalties of making false or fraudulent statements in this"
      Print #RptHandle, Tab(3); "application. All licenses are subject to being audited. Failure to provide"
      Print #RptHandle, Tab(3); "all information requested will result in an audit form all required sources."
      Print #RptHandle, Tab(3); ""
      Print #RptHandle, Tab(3); "___________________________________________________________________________"
      Print #RptHandle, Tab(3); "Signature                       Title                              Date"
      Print #RptHandle, Tab(3); ""
      Print #RptHandle, Tab(3); ""
      Print #RptHandle, Tab(3); "FOR OFFICE USE ONLY                        PLEASE REMIT TO:"
      Print #RptHandle, Tab(3); "SIC CODE___________________ "; Tab(46); QPTrim$(TownRec.AppTownOf)               'TOWN OF HEMINGWAY"
      Print #RptHandle, Tab(3); "RATE CLASS_________________ "; Tab(46); QPTrim$(TownRec.AppAdd1)               'P.O. BOX 968"
      Print #RptHandle, Tab(3); "LICENSE NUMBER_____________ "; Tab(46); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)              'HEMINGWAY S.C. 29554"
      Print #RptHandle, Chr$(12);
    End If
  Next cnt
    
  Close         'Close all open files now

  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  
  MainLog ("Application #9 reprinted.")
  
  Exit Sub

'-----------------------------------------------------


PrintStandard:
  
  Get TempHandle, 1, TempCustRec 'it is possible (but not probable) for the user to run
  'renewal applications then change the renewal form and then try to run reprints...
  'this could cause problems if all the data needed to run the reprints wasn't available
  'with the change
  If TempCustRec.AppType <> TownRec.AppForm Then
    frmBLMessageBoxJr.Label1.Caption = "The last renewal applications printed were not the same form as that currently saved. Please rerun the application renewals."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdReprint.Enabled = False
  For cnt = FirstRec To LastRec 'NumOfIdxRecs 'IdxTrNumRecs
    Get TempHandle, cnt, TempCustRec
    Get CHandle, TempCustRec.CustRecNum, CustRec
    GoSub PrintSTDForm
    frmBLShowPctComp.ShowPctComp cnt, LastRec
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdReprint.Enabled = True
      Exit Sub
    End If
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdReprint.Enabled = True
  
  Close         'Close all open files now
  ViewPrint ReportFile$, "Applications", True
  KillFile ReportFile$
  
  MainLog ("Standard application reprinted.")
  
  Exit Sub
  
PrintSTDForm:
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  For ll = 1 To 5
    Print #RptHandle, ""
  Next
  Print #RptHandle, 'Tab(37 - tab1); 'Heading1$
  Print #RptHandle, 'Tab(37 - tab2); 'Heading2$
  Print #RptHandle, 'Tab(37 - Tab3); 'Heading3$
  Print #RptHandle, 'Tab(37 - Tab4); 'Heading4$
  Print #RptHandle, 'Tab(66); Year$ ' Form$(2, 0)
  Print #RptHandle,
  Print #RptHandle, Tab(11); CustRec.BillName
  Print #RptHandle, Tab(11); CustRec.ADDRESS1
  Print #RptHandle, Tab(11); CustRec.ADDRESS2
  Print #RptHandle, Tab(11); RTrim$(CustRec.City); "  "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(11); QPTrim$(CustRec.CustName)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); QPTrim$(TempCustRec.CatCode(1));
  Print #RptHandle, Tab(15); QPTrim$(TempCustRec.CatDesc(1));
  Print #RptHandle, Tab(62); Using("##,##0.00", TempCustRec.Fee(1))
  SCnt = 24
  If Val(CustRec.BILLCAT2) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(TempCustRec.CatCode(2));
  Print #RptHandle, Tab(15); QPTrim$(TempCustRec.CatDesc(2));
  Print #RptHandle, Tab(62); Using("##,##0.00", TempCustRec.Fee(2))
  SCnt = 25
  If Val(CustRec.BILLCAT3) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(TempCustRec.CatCode(3));
  Print #RptHandle, Tab(15); QPTrim$(TempCustRec.CatDesc(3));
  Print #RptHandle, Tab(62); Using("##,##0.00", TempCustRec.Fee(3))
  SCnt = 26
  If Val(CustRec.BILLCAT4) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(TempCustRec.CatCode(4));
  Print #RptHandle, Tab(15); QPTrim$(TempCustRec.CatDesc(4));
  Print #RptHandle, Tab(62); Using("##,##0.00", TempCustRec.Fee(4))
  SCnt = 27
  If Val(CustRec.BILLCAT5) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(TempCustRec.CatCode(5));
  Print #RptHandle, Tab(15); QPTrim$(TempCustRec.CatDesc(5));
  Print #RptHandle, Tab(62); Using("##,##0.00", TempCustRec.Fee(5))
  SCnt = 28
  
ExitFormPrint:
  
  If TempCustRec.IssFee > 0 Then
    Print #RptHandle, Tab(15); "Issuance Fee"; Tab(62); Using$("##,##0.00", TempCustRec.IssFee)
  End If
  For LCnt = SCnt To 35
    Print #RptHandle, ""
  Next
  Print #RptHandle, Tab(62); Using("##,##0.00", TempCustRec.MiscNum)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(62); Using("##,##0.00", TempCustRec.MiscNum)
  Print #RptHandle,
  Print #RptHandle,
  TotalCust = TotalCust + 1
Return
  

GetCode:
  For Snt& = 1 To NumOfARCatRecs
    Get CodeHandle, Snt&, CodeRec
    If QPTrim$(CodeRec.CatCode) = CatCode$ Then
      CODEDESC$ = QPTrim$(CodeRec.CODEDESC)
      Select Case CodeRec.CodeType
      Case "F"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case "M"
        DESC1$ = "Per Each"
        Amt# = CodeRec.RateStep
        CodeType$ = CodeRec.CodeType
      Case Is = "S"
        BaseAmt1# = CodeRec.BaseAmt1
        Revenue1# = CodeRec.Recpt1
        Percent1# = CodeRec.Percent1
        Maximum1# = CodeRec.Maximum1
        BaseAmt2# = CodeRec.BaseAmt2
        Revenue2# = CodeRec.Recpt2
        Percent2# = CodeRec.Percent2
        Maximum2# = CodeRec.Maximum2
        BaseAmt3# = CodeRec.BaseAmt3
        Revenue3# = CodeRec.Recpt3
        Percent3# = CodeRec.Percent3
        Maximum3# = CodeRec.Maximum3
        BaseAmt4# = CodeRec.BaseAmt4
        Revenue4# = CodeRec.Recpt4
        Percent4# = CodeRec.Percent4
        Maximum4# = CodeRec.Maximum4
        BaseAmt5# = CodeRec.BaseAmt5
        Revenue5# = CodeRec.Recpt5
        Percent5# = CodeRec.Percent5
        Maximum5# = CodeRec.Maximum5
        BaseAmt6# = CodeRec.BaseAmt6
        Revenue6# = CodeRec.Recpt6
        Percent6# = CodeRec.Percent6
        Maximum6# = CodeRec.Maximum6
        CodeType$ = CodeRec.CodeType
      Case Else
        CodeType$ = "N"
      End Select
      Exit For
    End If
  Next Snt&

GotCode:
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLReprintAppsRnwls", "cmdReprint_Click", Erl)
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
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLReprintAppsRnwls.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TempCustRec As TempCustRecType
  Dim TempHandle As Integer
  Dim NumOfRecs As Integer
  Dim x As Integer
  
  On Error Resume Next
  
  lblBalloon.Visible = False
  
  fptxtFirstNum.ToolTipText = "Enter the first reprint reference number here."
  fptxtLastNum.ToolTipText = "Enter the last reprint reference number here."
  cmdExit.ToolTipText = "Press 'Cancel' to exit this screen."
  cmdReprint.ToolTipText = "Press 'Start Reprint' to generate application reprints."
  
  If Exist("artmpcus.dat") Then
    OpenTempCustRec TempHandle
    NumOfRecs = LOF(TempHandle) / Len(TempCustRec)
    If NumOfRecs = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no application renewal records saved that can be reprinted. Please print application renewal forms first."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Close TempHandle
      Exit Sub
    End If
  End If
  
  fptxtFirstNum.Text = "1"
  fptxtLastNum.Text = NumOfRecs
  Close TempHandle
    
End Sub
