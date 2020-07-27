VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFADisposalMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disposal Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFADisposalMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   495
      Left            =   4035
      TabIndex        =   3
      ToolTipText     =   "Click this button to print a report listing the assets earmarked for disposal by disposal date."
      Top             =   4155
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADisposalMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBuildList 
      Height          =   495
      Left            =   4035
      TabIndex        =   2
      ToolTipText     =   "Click this button to select multiple items to dispose of."
      Top             =   3495
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADisposalMenu.frx":0AB1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSingle 
      Height          =   495
      Left            =   4032
      TabIndex        =   1
      ToolTipText     =   "Click this button to dispose of and post this disposal of a single item."
      Top             =   2832
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADisposalMenu.frx":0C9D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditList 
      Height          =   495
      Left            =   4032
      TabIndex        =   4
      ToolTipText     =   "Click this button to edit any list of items designated for disposal by date of disposal"
      Top             =   4824
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADisposalMenu.frx":0E85
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   495
      Left            =   4032
      TabIndex        =   5
      ToolTipText     =   $"frmFADisposalMenu.frx":1070
      Top             =   5488
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADisposalMenu.frx":1102
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdUNDispose 
      Height          =   495
      Left            =   4032
      TabIndex        =   6
      ToolTipText     =   "Click this button to reverse any single item that has been posted as disposed of."
      Top             =   6152
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADisposalMenu.frx":12E8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4032
      TabIndex        =   7
      ToolTipText     =   "Click this button to return to the main menu."
      Top             =   6816
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADisposalMenu.frx":14D8
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   135
      Index           =   4
      Left            =   2110
      Top             =   2090
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   135
      Index           =   3
      Left            =   8610
      Top             =   2100
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIXED ASSET DISPOSAL MENU"
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
      Left            =   2820
      TabIndex        =   0
      Top             =   1246
      Width           =   6012
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8700
      X2              =   8700
      Y1              =   2206
      Y2              =   8078
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2203
      Y2              =   8075
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2219
      X2              =   2929
      Y1              =   8090
      Y2              =   8090
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8703
      X2              =   9403
      Y1              =   8090
      Y2              =   8090
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1500
      Top             =   895
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   766
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8700
      Top             =   2194
      Width           =   732
   End
End
Attribute VB_Name = "frmFADisposalMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdBuildList_Click()
  Dim DepFile As Integer
  Dim Nextx As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDprRecs As Long
  Dim DprHistRec As DprHistType
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  
  If TagIdxCnt = 0 Then
    MsgBox "There are no fixed assets saved."
    Exit Sub
  End If

  frmFADispItemList.Show
  DoEvents
  Unload frmFADisposalMenu
End Sub

Private Sub cmdEditList_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Long
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  
  If TagIdxCnt = 0 Then
    MsgBox "There are no fixed assets saved."
    Exit Sub
  End If
  
  OpenTempDisposedDate GHandle
  DateCnt = LOF(GHandle) / Len(DateRec) 'if there
  Close GHandle
  
  If DateCnt = 0 Then
    MsgBox "There are no disposal dates saved"
    Exit Sub
  End If
  
  frmFAEditDisposedOf.Show
  DoEvents
  Unload frmFADisposalMenu
End Sub

Private Sub cmdExit_Click()
  frmFAMainMenu.Show
  Close
  DoEvents
  Unload frmFADisposalMenu
End Sub

Private Sub cmdPost_Click()
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Long
  
  OpenTempDisposedDate GHandle
  DateCnt = LOF(GHandle) / Len(DateRec)
  Close GHandle
  If DateCnt = 0 Then
    MsgBox "There are no pending disposal dates to be posted."
    Exit Sub
  End If
  
  frmFAPostDisposal.Show
  DoEvents
  Unload frmFADisposalMenu
End Sub

Private Sub cmdPrint_Click()
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Long
  
  OpenTempDisposedDate GHandle
  DateCnt = LOF(GHandle) / Len(DateRec) 'if there
  Close GHandle
  
  If DateCnt = 0 Then
    MsgBox "There are no disposal dates saved"
    Exit Sub
  End If
  
  frmFAPrintDsplList.Show
  DoEvents
  Unload frmFADisposalMenu
End Sub

Private Sub cmdSingle_Click()
  Dim DepFile As Integer
  Dim Nextx As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDprRecs As Long
  Dim DprHistRec As DprHistType
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  
  If TagIdxCnt = 0 Then
    MsgBox "There are no fixed assets saved."
    Exit Sub
  End If
  
  frmFADsplSingle.Show
  DoEvents
  Unload frmFADisposalMenu
End Sub

Private Sub cmdUNDispose_Click()
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDprRecs As Long
  
  DepFile = FreeFile
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  NumOfDprRecs = LOF(DepFile) / Len(FADep(1))
  If NumOfDprRecs > 0 Then
    Close
    If MsgBox("There are pending depreciation files that have not been posted. Please post these files before continuing. Would you like to jump to the depreciation post screen?", vbYesNo) = vbYes Then
      frmFAYearEndPost.Show
      DoEvents
      Unload frmFADisposalMenu
    End If
    Exit Sub
  End If
  
  Close
  frmFAReverseDspl.Show
  DoEvents
  Unload frmFADisposalMenu
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFADisposalMenu.")
      Call Terminate
      End
    End If
  End If
End Sub



