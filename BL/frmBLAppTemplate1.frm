VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmBLAppTemplate1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Application Renewal Template #1"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   Icon            =   "frmBLAppTemplate1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   696
      Left            =   9468
      TabIndex        =   0
      Tag             =   "Press the 'Cancel' button to close this screen and return to the Town Setup screen."
      Top             =   6420
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1228
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
      ButtonDesigner  =   "frmBLAppTemplate1.frx":08CA
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   8652
      Left            =   1980
      TabIndex        =   1
      Top             =   48
      Width           =   7116
      _Version        =   196609
      _ExtentX        =   12552
      _ExtentY        =   15261
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483639
      Caption         =   ""
      Picture         =   "frmBLAppTemplate1.frx":0AA8
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
         Caption         =   "Business City, State  Zip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   20
         Top             =   7056
         Width           =   1740
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Caption         =   "Business Address 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   19
         Top             =   6864
         Width           =   1404
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         Caption         =   "Business Address 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   18
         Top             =   6672
         Width           =   1404
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         Caption         =   "Business Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   17
         Top             =   6480
         Width           =   1308
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         Caption         =   "Business Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   16
         Top             =   7536
         Width           =   1308
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000009&
         Caption         =   "Total Fees"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   5808
         TabIndex        =   15
         Top             =   4752
         Width           =   780
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         Caption         =   "------------------------------------------------------------------------------------------------------------------------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   48
         TabIndex        =   14
         Top             =   5232
         Width           =   7020
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         Caption         =   "Total Fees"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   5808
         TabIndex        =   13
         Top             =   4224
         Width           =   780
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "Code        Category Description                                                                       Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   720
         TabIndex        =   12
         Top             =   3600
         Width           =   6204
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "Code        Category Description                                                                       Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   720
         TabIndex        =   11
         Top             =   3408
         Width           =   6204
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "Code        Category Description                                                                       Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   720
         TabIndex        =   10
         Top             =   3216
         Width           =   6204
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Code        Category Description                                                                       Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   720
         TabIndex        =   9
         Top             =   3024
         Width           =   6204
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Business Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   8
         Top             =   2448
         Width           =   1308
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Business Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   6
         Top             =   1392
         Width           =   1308
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Business Address 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   5
         Top             =   1584
         Width           =   1404
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Business Address 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   4
         Top             =   1776
         Width           =   1404
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Business City, State  Zip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1200
         TabIndex        =   3
         Top             =   1968
         Width           =   1740
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000009&
         Caption         =   "Code        Category Description                                                                       Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   720
         TabIndex        =   2
         Top             =   2832
         Width           =   6204
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   684
      Left            =   9468
      TabIndex        =   21
      Tag             =   "Press this 'Next App' button to close this application screen and open up the screen for application #2."
      Top             =   4536
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1206
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
      ButtonDesigner  =   "frmBLAppTemplate1.frx":0AC4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   684
      Left            =   9468
      TabIndex        =   22
      Tag             =   "Press 'Save' to save the currently active application as application #1. Otherwise there are no fields on this screen to save."
      Top             =   7368
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1206
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
      ButtonDesigner  =   "frmBLAppTemplate1.frx":0CA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLast 
      Height          =   672
      Left            =   9468
      TabIndex        =   23
      Tag             =   "Press this 'Last App' to close this screen and open the screen for application #9."
      Top             =   5496
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1185
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
      ButtonDesigner  =   "frmBLAppTemplate1.frx":0E7F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   9456
      TabIndex        =   24
      Tag             =   $"frmBLAppTemplate1.frx":105E
      ToolTipText     =   "Press to bring up a brief help screen."
      Top             =   3312
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmBLAppTemplate1.frx":1128
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   9984
      TabIndex        =   26
      Top             =   1104
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
      Left            =   9360
      TabIndex        =   25
      Top             =   4080
      Width           =   2052
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   876
      Left            =   9264
      Top             =   3108
      Width           =   2268
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Renewal Application #1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   9492
      TabIndex        =   7
      Top             =   1920
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   972
      Left            =   9408
      Top             =   1764
      Width           =   1980
   End
   Begin VB.Menu mnuOPtions 
      Caption         =   "Options"
      Begin VB.Menu mnuPntScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBLAppTemplate1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload frmBLAppTemplate1
  frmBLTownSetup.fpcmbAppType.SetFocus
End Sub

Private Sub cmdHelp_Click()
  
  If InStr(cmdHelp.Text, "On") Then
    lblBalloon.Visible = True
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Label1.Caption = "This form is designed so that it can be printed on a pre-supplied business license form."
    frmBLMessageBoxJr.Show vbModal
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    cmdHelp.ToolTipText = ""
    cmdNext.ToolTipText = ""
    cmdLast.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSave.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdHelp.ToolTipText = "Press this button to activate/deactivate instructional balloons."
'    cmdNext.ToolTipText = "Press to move to application template #2."
'    cmdLast.ToolTipText = "Press to move to business application #9."
'    cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'    cmdSave.ToolTipText = "Press to save the data on this screen."
  End If
End Sub

Private Sub cmdLast_Click()
  frmBLFreeFormatApp1.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim x As Integer
  Dim TempCustRec As TempCustRecType
  Dim TempHandle As Integer
  Dim TempCnt As Integer
  
  On Error GoTo ERRORSTUFF
  'added as a precaution to prevent the user from running application
  'renewal form #1 then coming here to save different data and then
  'trying to run application renewal reprints which will use this
  'latest saved data while the originals have the old data...now the
  'user will have to print applications over
  If Exist("artmpcus.dat") Then
    OpenTempCustRec TempHandle
    TempCnt = LOF(TempHandle) / Len(TempCustRec)
    If TempCnt > 0 Then
      Get TempHandle, 1, TempCustRec
      If TempCustRec.AppType = 1 Then KillFile "artmpcus.dat"
    End If
    Close TempHandle
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
      TownRec.AppForm = 1
    Put THandle, 1, TownRec
  Else
    TownRec.TownName = ""
    TownRec.Contact = ""
    TownRec.TownAdd1 = ""
    TownRec.TownAdd2 = ""
    TownRec.City = ""
    TownRec.State = ""
    TownRec.ZipCode = ""
    TownRec.TownPhone = ""
    TownRec.SpareSpace = ""
    TownRec.AppForm = 1
    TownRec.DLQNotice = 0
    TownRec.AppAdd1 = ""
    TownRec.AppBaseFee(1) = 0
    TownRec.AppBaseFee(2) = 0
    TownRec.AppBaseFee(3) = 0
    TownRec.AppBaseFee(4) = 0
    TownRec.AppCentsPer(1) = 0
    TownRec.AppCentsPer(2) = 0
    TownRec.AppCentsPer(3) = 0
    TownRec.AppCentsPer(4) = 0
    TownRec.AppFirstDay = ""
    TownRec.AppLastDay = ""
    TownRec.AppGrsRcpts(1) = 0
    TownRec.AppGrsRcpts(2) = 0
    TownRec.AppGrsRcpts(3) = 0
    TownRec.AppGrsRcpts(4) = 0
    TownRec.AppColFee = 0
    TownRec.AppGrsPct = 0
    TownRec.AppDenom = 0
    TownRec.AppNumer = 0
    TownRec.AppState = ""
    TownRec.AppCity = ""
    TownRec.AppTownOf = ""
    TownRec.AppZip = ""
    TownRec.AppPct = 0
    TownRec.AppPayBy = 0
    TownRec.AppPhone = ""
    TownRec.AppAdminName = ""
    TownRec.AppAdminTitle = ""
    TownRec.AppDiscPct = 0
    TownRec.AppDiscMonth = ""
    TownRec.AppDiscDay = 0
    TownRec.AppPenMonth = ""
    TownRec.AppPenDay = 0
    TownRec.AppFiscMonth = ""
    TownRec.AppFiscDay = 0
    TownRec.AppMayorCouncil = ""
    TownRec.AppWholeMonth = 0
    TownRec.AppWholeDay = 0
    TownRec.AppRetailMonth = 0
    TownRec.AppRetailDay = 0
    TownRec.AppFinMonth = 0
    TownRec.AppFinDay = 0
    TownRec.AppContMonth = 0
    TownRec.AppContDay = 0
    TownRec.AppRepairMonth = 0
    TownRec.AppRepairDay = 0
    TownRec.AppStartMonth = ""
    TownRec.AppStartDay = 0
    TownRec.AppLicRetMonth = ""
    TownRec.AppLicRetDay = 0
    TownRec.AppAdoptDate = 0
    TownRec.AppCityOrd = ""
    For x = 1 To 10
      TownRec.AppYrUpDown(x) = "0"
    Next x
    TownRec.DlqAdd1 = ""
    TownRec.DlqAdminName = ""
    TownRec.DlqAdminTitle = ""
    TownRec.DlqCity = ""
    TownRec.DlqPhone = ""
    TownRec.DlqPhone2 = ""
    TownRec.DlqFax = ""
    TownRec.DlqState = ""
    TownRec.DlqTownName = ""
    TownRec.DlqZip = ""
    TownRec.DlqFirstDay = ""
    TownRec.DlqLastDay = ""
    TownRec.DlqFirstHour = ""
    TownRec.DlqLastHour = ""
    TownRec.DlqClerkName = ""
    TownRec.DlqMayorCouncil = ""
    TownRec.LicNumPermYN = "No"
    TownRec.UseAmtPctYN = "Pct"
    TownRec.PENCASHACCT = 0
    TownRec.PENRECGLNUM = 0
    TownRec.PENREVGLNUM = 0
    TownRec.IssFee = 0
    TownRec.AcctMeth = ""
    TownRec.LaserLtr = "N"
    TownRec.GL2Cats = "N"
    OpenTownFile THandle
    Put THandle, 1, TownRec
  End If
  Close THandle
  
  frmBLSucSave.Label1.Caption = "Your renewal application notice #1 data has been saved successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  Call cmdExit_Click
  frmBLTownSetup.fpcmbAppType.Text = "1. APP STANDARD"
  frmBLTownSetup.fpcmdApps.Text = "F3 S&how App Type 1"
  
  MainLog ("Application #1 saved.")
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAppTemplate1", "cmdSave_Click", Erl)
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
  lblBalloon.Visible = False
'  cmdHelp.ToolTipText = "Press this button to activate/deactivate instructional balloons."
'  cmdNext.ToolTipText = "Press to move to application template #2."
'  cmdLast.ToolTipText = "Press to move to business application #9."
'  cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'  cmdSave.ToolTipText = "Press to save the data on this screen."
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
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%N"
      Call cmdNext_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%L"
      Call cmdLast_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLAppTemplate1.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub cmdNext_Click()
  frmBLAppTemplate2.Show
  DoEvents
  Unload Me
End Sub
