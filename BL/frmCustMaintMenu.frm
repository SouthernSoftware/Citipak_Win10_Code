VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLCustMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Maintenance Menu"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "frmCustMaintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   5338
      TabIndex        =   1
      Top             =   7322
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
   Begin fpBtnAtlLibCtl.fpBtn cmdAddNewCust 
      Height          =   492
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Click this button to add a brand new customer."
      Top             =   2525
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditCust 
      Height          =   492
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Click this button to make changes to existing customers."
      Top             =   3126
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":0AB0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustListRpt 
      Height          =   492
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Click this button to display and print a report of customer listings."
      Top             =   3727
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":0C9D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMailingLabels 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Click this button to create customer mailing labels."
      Top             =   4328
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":0E84
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAdjust 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Tag             =   "Click this button to make changes to a customer's balance."
      Top             =   4932
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":1075
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReindex 
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Tag             =   $"frmCustMaintMenu.frx":1262
      Top             =   5536
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":1323
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   6140
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":150F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Tag             =   "Click this button to return to the main Business License menu."
      Top             =   6746
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmCustMaintMenu.frx":16F4
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   3
      Left            =   8550
      Top             =   1995
      Width           =   985
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2085
      X2              =   2795
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER MAINTENANCE"
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
      Left            =   2686
      TabIndex        =   0
      Top             =   1178
      Width           =   6012
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   1966
      X2              =   2926
      Y1              =   1893
      Y2              =   1893
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8446
      X2              =   9406
      Y1              =   1893
      Y2              =   1893
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
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8666
      X2              =   9369
      Y1              =   8009
      Y2              =   8009
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1455
      Top             =   820
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1455
      Top             =   696
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1965
      Top             =   1890
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2086
      Top             =   2126
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8548
      Top             =   1896
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8666
      Top             =   2124
      Width           =   732
   End
End
Attribute VB_Name = "frmBLCustMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdAddNewCust_Click()
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim NumOfTownRecs As Integer
  
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No category codes have been saved. Please save data for at least one category code. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenTownFile TownHandle
  NumOfTownRecs = LOF(TownHandle) / Len(TownRec)
  If NumOfTownRecs = 0 Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "No town control files have been saved. Town control files are edited on the Town Setup screen. Would you like to jump there now?"
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLTownSetup.Show
      DoEvents
      Unload frmBLCustMaintMenu
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  Close TownHandle
  
  frmBLCustEdit.Show
  DoEvents
  Unload frmBLCustMaintMenu
End Sub

Private Sub cmdExit_Click()
  frmBLMainMenu.Show
  DoEvents
  Unload frmBLCustMaintMenu
End Sub

Private Sub cmdCustListRpt_Click()
  Dim PrintType$
  
  On Error Resume Next
  If Not Exist("arcustnameidx.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "No customer name index saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Call PrintText
    Case "Exit"
  End Select
  cmdHelp.Text = "Turn Menu &Help On"
  btnHelp.AutoScan = fpAutoScanOff

End Sub

Private Sub cmdEditCust_Click()
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim NumOfTownRecs As Integer
  
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No category codes have been saved. Please save data for at least one category code. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenTownFile TownHandle
  NumOfTownRecs = LOF(TownHandle) / Len(TownRec)
  If NumOfTownRecs = 0 Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "No town control files have been saved. Town control files are edited on the Town Setup screen. Would you like to jump there now?"
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLTownSetup.Show
      DoEvents
      Unload frmBLCustMaintMenu
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  Close TownHandle
  
  frmBLCustomerLookup.Show
  DoEvents
  Unload frmBLCustMaintMenu
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "Turn Menu &Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "Turn Menu &Help On"
    btnHelp.AutoScan = fpAutoScanOff
  End If

End Sub

Private Sub cmdMailingLabels_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLMailLbls.Show
  DoEvents
  Unload frmBLCustMaintMenu
End Sub
Private Sub cmdAdjust_Click()
  On Error Resume Next
  
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLAdjustBal.Show
  DoEvents
  Unload frmBLCustMaintMenu
End Sub

Private Sub cmdReindex_Click()
  
  If Exist("arcust.dat") Then
    Call CreateCustNameIdx
    Call CreateCustNumIdx
    Call CreateLicNumIdx
    Call CreateCustSearchNameIdx
  Else
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Re-indexing aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  cmdHelp.Text = "Turn Menu &Help On"
  btnHelp.AutoScan = fpAutoScanOff
  
  frmBLMessageBoxJr.Label1.Caption = "Customer names, customer search names, customer license numbers and customer numbers have been reindexed."
  frmBLMessageBoxJr.Label1.Top = 800
  frmBLMessageBoxJr.Show vbModal
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  GCustNum = 0
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustMaintMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim TrHandle As Integer
  Dim TRNumRecs As Integer
  Dim IdxTrHandle As Integer
  Dim CustIdxRec As CustNameIdxType
  Dim IdxTrNumRecs As Integer
  Dim CustRec As ARCustRecType
  Dim x As Integer
  Dim cnt As Integer
  Dim TotalCust As Integer
  Dim Page As Integer
  
  On Error GoTo ERRORSTUFF
  
  ReportFile$ = "ARCUST.PRN"    'Report File Name
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  OpenCustFile TrHandle
  
  OpenCustNameIdxFile IdxTrHandle
  IdxTrNumRecs = LOF(IdxTrHandle) \ Len(CustIdxRec)
  
  ReDim CustIdxs(1 To IdxTrNumRecs) As Integer
  For x = 1 To IdxTrNumRecs
    Get IdxTrHandle, x, CustIdxRec
    CustIdxs(x) = CustIdxRec.CustRec
  Next x
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintRptHeader
  frmBLShowPctComp.Label1 = "Loading Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  For cnt = 1 To IdxTrNumRecs
    Get TrHandle, CustIdxs(cnt), CustRec
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintRptHeader
      End If
      Print #RptHandle, QPTrim$(CustRec.CustNumb); Tab(14); QPTrim$(CustRec.BillName);
      Print #RptHandle, Tab(67); Using("##0.000", CustRec.Prorate); "%";

      Print #RptHandle, Tab(14); QPTrim$(CustRec.CustName); Tab(50); QPTrim$(CustRec.LICENSE); Tab(65); MakeRegDate(CustRec.VALID)
      Print #RptHandle, "Categories: "; QPTrim(CustRec.BILLCAT1$); " / "; QPTrim$(CustRec.BILLCAT2$); " / "; QPTrim$(CustRec.BILLCAT3$); " / "; QPTrim$(CustRec.BILLCAT4$); " / "; QPTrim$(CustRec.BILLCAT5$)
      Print #RptHandle, String$(79, "-")
      TotalCust = TotalCust + 1
      LineCnt = LineCnt + 4
    End If
    frmBLShowPctComp.ShowPctComp cnt, IdxTrNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next cnt

  GoSub PrintRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  
  ViewPrint ReportFile$, "Customer Listing", True
  
  Kill ReportFile$
  
  Exit Sub

PrintRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(18); "Business License : Customer 'Quick' Listing"
  Print #RptHandle, Tab(21); "      Report Date: "; Date$; Tab(68); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, "Cust #"; Tab(10); "Billing Name"; Tab(65); "ProRate"
  Print #RptHandle, Tab(10); "Customer Name"; Tab(48); "License #"; Tab(65); "Valid To"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
Return
  
PrintRptEnding:
  Print #RptHandle, "Number of Customers .. "; Using("###,##0", TotalCust)
  Print #RptHandle, FF$
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustMaintMenu", "PrintText", Erl)
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

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TrHandle As Integer
  Dim TRNumRecs As Integer
  Dim IdxTrHandle As Integer
  Dim CustIdxRec As CustNameIdxType
  Dim IdxTrNumRecs As Integer
  Dim CustRec As ARCustRecType
  Dim x As Integer
  Dim cnt As Integer
  Dim TotalCust As Integer
  Dim dlm$
  Dim TownName$
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName$ = QPTrim$(TownRec.TownName)
  
  ReportFile$ = "BLRPTS\ARQKCUST.RPT"    'Report File Name
  OpenCustFile TrHandle
  
  OpenCustNameIdxFile IdxTrHandle
  IdxTrNumRecs = LOF(IdxTrHandle) \ Len(CustIdxRec)
  
  ReDim CustIdxs(1 To IdxTrNumRecs) As Integer
  For x = 1 To IdxTrNumRecs
    Get IdxTrHandle, x, CustIdxRec
    CustIdxs(x) = CustIdxRec.CustRec
  Next x
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  frmBLShowPctComp.Label1 = "Loading Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  For cnt = 1 To IdxTrNumRecs
    Get TrHandle, CustIdxs(cnt), CustRec
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
      Print #RptHandle, TownName$; dlm;
      Print #RptHandle, QPTrim$(CustRec.CustNumb); dlm; QPTrim$(CustRec.BillName); dlm;
      Print #RptHandle, CustRec.Prorate / 100; dlm;
      Print #RptHandle, QPTrim$(CustRec.CustName); dlm; QPTrim$(CustRec.LICENSE); dlm; MakeRegDate(CustRec.VALID); dlm;
      Print #RptHandle, QPTrim(CustRec.BILLCAT1$); dlm; QPTrim$(CustRec.BILLCAT2$); dlm; QPTrim$(CustRec.BILLCAT3$); dlm; QPTrim$(CustRec.BILLCAT4$); dlm; QPTrim$(CustRec.BILLCAT5$)
      TotalCust = TotalCust + 1
    End If
    frmBLShowPctComp.ShowPctComp cnt, IdxTrNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next cnt

  Close         'Close all open files now
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  
  arBLQuickList.Show
  frmBLLoadReport.Show
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustMaintMenu", "PrintGraphics", Erl)
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

