VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLIssueAppsLics 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Issue Applications and Licenses Menu"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "frmIssueAppsLics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   5424
      TabIndex        =   1
      Top             =   7008
      Width           =   684
      _Version        =   131072
      _ExtentX        =   1206
      _ExtentY        =   529
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
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
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
   Begin fpBtnAtlLibCtl.fpBtn cmdAppList 
      Height          =   492
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Press to call up a report screen for printing current active business license customers."
      Top             =   2715
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmIssueAppsLics.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintAppsLics 
      Height          =   492
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Press to bring up an application processing screen."
      Top             =   3329
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmIssueAppsLics.frx":0AB1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprints 
      Height          =   480
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Press to make reprints of existing applications."
      Top             =   3943
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmIssueAppsLics.frx":0CA0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAdvLtr 
      Height          =   492
      Left            =   3960
      TabIndex        =   5
      Tag             =   $"frmIssueAppsLics.frx":0E91
      Top             =   4545
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmIssueAppsLics.frx":0F22
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMailingLabels 
      Height          =   492
      Left            =   3960
      TabIndex        =   6
      Tag             =   "Click this button to open a screen that begins the mailing label printing process."
      Top             =   5159
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmIssueAppsLics.frx":1112
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   3960
      TabIndex        =   7
      Top             =   5773
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmIssueAppsLics.frx":1303
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   480
      Left            =   3960
      TabIndex        =   8
      Tag             =   "Click this button to return to the main Business License menu."
      Top             =   6390
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmIssueAppsLics.frx":14E8
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   3
      Left            =   8550
      Top             =   1995
      Width           =   985
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   1970
      Top             =   2000
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "APPLICATIONS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2775
      TabIndex        =   0
      Top             =   1170
      Width           =   6012
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8666
      X2              =   8666
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2086
      Y1              =   2133
      Y2              =   8005
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2795
      Y1              =   8025
      Y2              =   8025
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8655
      X2              =   9369
      Y1              =   8010
      Y2              =   8010
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1455
      Top             =   820
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1455
      Top             =   690
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   1966
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2086
      Top             =   2130
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8550
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8655
      Top             =   2130
      Width           =   732
   End
End
Attribute VB_Name = "frmBLIssueAppsLics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "Turn Menu &Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "Turn Menu &Help On"
    btnHelp.AutoScan = fpAutoScanOff
  End If
End Sub

Private Sub cmdAdvLtr_Click()
  Dim DHandle As Integer
  Dim One As Integer
  Dim LaserRec1 As LaserLetterType1
  Dim LHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile TownHandle
    Get TownHandle, 1, TownRec
    Close TownHandle
  End If
  
  If QPTrim$(TownRec.LaserLtr) = "1" Then
    If Not Exist("arlaser1.dat") Then
      One = 1
      DHandle = FreeFile
      Open "issueappslics.dat" For Output As DHandle Len = 2
      Print #DHandle, One
      Close DHandle
      frmBLMessageBoxJrWOpts.Label1.Caption = "Before using the advance renewal letter feature the letter must be created. This is done on the 'Business License Advance Letter Build' screen reached by way of the Town Setup screen. Would you like to jump to that screen now?"
      frmBLMessageBoxJrWOpts.Label1.Top = 600
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        frmBLAdvanceLetter.Show
        DoEvents
        Unload frmBLIssueAppsLics
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      End If
    End If
  ElseIf QPTrim$(TownRec.LaserLtr) = "2" Then
    If Not Exist("arlaser2.dat") Then
      One = 1
      DHandle = FreeFile
      Open "issueappslics.dat" For Output As DHandle Len = 2
      Print #DHandle, One
      Close DHandle
      frmBLMessageBoxJrWOpts.Label1.Caption = "Before using the advance renewal letter feature the letter must be created. This is done on the 'Business License Advance Letter Build' screen reached by way of the Town Setup screen. Would you like to jump to that screen now?"
      frmBLMessageBoxJrWOpts.Label1.Top = 600
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        frmBLAdvLetter2.Show
        DoEvents
        Unload frmBLIssueAppsLics
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      End If
    End If
  ElseIf QPTrim$(TownRec.LaserLtr) = "3" Then
    If Not Exist("arlaser3.dat") Then
      One = 1
      DHandle = FreeFile
      Open "issueappslics.dat" For Output As DHandle Len = 2
      Print #DHandle, One
      Close DHandle
      frmBLMessageBoxJrWOpts.Label1.Caption = "Before using the advance renewal letter feature the letter must be created. This is done on the 'Business License Advance Letter Build' screen reached by way of the Town Setup screen. Would you like to jump to that screen now?"
      frmBLMessageBoxJrWOpts.Label1.Top = 600
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        frmBLAdvanceLtr3.Show
        DoEvents
        Unload frmBLIssueAppsLics
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      End If
    End If
  Else
    frmBLMessageBoxJrWOpts.Label1.Caption = "Advance laser letters are created via the Town Setup screen. The advance letter option on that screen is currently set to '0' indicating that advance letters are not being used. Would you like to jump to the Town Setup screen now?"
    frmBLMessageBoxJrWOpts.Label1.Top = 600
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLTownSetup.Show
      DoEvents
      Unload frmBLIssueAppsLics
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
   
  frmBLPrintAdvanceLetter.Show
  DoEvents
  Unload frmBLIssueAppsLics
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLIssueAppsLics", "cmdAdvLtr_Click", Erl)
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

Private Sub cmdAppList_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLAppListIssue.Show
  DoEvents
  Unload frmBLIssueAppsLics
End Sub

Private Sub cmdExit_Click()
  KillFile "issueappslics.dat"
  frmBLMainMenu.Show
  DoEvents
  Unload frmBLIssueAppsLics
End Sub

Private Sub cmdMailingLabels_Click()
  Dim DHandle As Integer
  Dim One As Integer
  
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  One = 1
  DHandle = FreeFile
  Open "issueappslics.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  frmBLMailLbls.Show
  DoEvents
  Unload frmBLIssueAppsLics
End Sub

Private Sub cmdPrintAppsLics_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenTownFile THandle
  Get THandle, 1, TownRec
  Close THandle
  
  If Exist("artownsu.dat") Then
    If TownRec.AppForm = 11 Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "No application renewal form has been selected. Would you like to jump to the Town Setup screen to select an application renewal form?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        frmBLTownSetup.Show
        frmBLTownSetup.fpcmbAppType.SetFocus
        DoEvents
        Unload frmBLIssueAppsLics
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      End If
    End If
  Else
    frmBLMessageBoxJrWOpts.Label1.Caption = "Town setup records have not been saved. Would you like to jump to the Town Setup screen now?"
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLTownSetup.Show
      frmBLTownSetup.fptxtTownName.SetFocus
      DoEvents
      Unload frmBLIssueAppsLics
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
  
  frmBLPrintAppsRenwls.Show
  DoEvents
  Unload frmBLIssueAppsLics
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLIssueAppsLics", "cmdPrintAppsLics_Click", Erl)
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

Private Sub cmdReprints_Click()
  If Not Exist("artmpcus.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "No renewal applications have been created in the text version so no reprints are possible."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  frmBLReprintAppsRnwls.Show
  DoEvents
  Unload frmBLIssueAppsLics

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
      SendKeys "%X"
      Call cmdExit_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLIssueAppsLics.")
      Call Terminate
      End
    End If
  End If
End Sub


