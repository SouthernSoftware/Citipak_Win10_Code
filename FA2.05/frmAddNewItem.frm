VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFAItemMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Maintenance Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmAddNewItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdAddNewItem 
      Height          =   492
      Left            =   4032
      TabIndex        =   1
      ToolTipText     =   "Click this button to add a brand new item."
      Top             =   3456
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmAddNewItem.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditExisting 
      Height          =   495
      Left            =   4032
      TabIndex        =   2
      ToolTipText     =   "Click this button to bring up a list of items that can be edited."
      Top             =   4128
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
      ButtonDesigner  =   "frmAddNewItem.frx":0AAC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMasterItemList 
      Height          =   495
      Left            =   4032
      TabIndex        =   3
      ToolTipText     =   "Click this button to begin retrieving data to create a detailed report of fixed asset attributes."
      Top             =   4800
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
      ButtonDesigner  =   "frmAddNewItem.frx":0C95
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4032
      TabIndex        =   4
      ToolTipText     =   "Click this button to return to the Maintenance Menu."
      Top             =   5472
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
      ButtonDesigner  =   "frmAddNewItem.frx":0E7D
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   120
      Index           =   4
      Left            =   8602
      Top             =   2100
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   120
      Index           =   3
      Left            =   2101
      Top             =   2093
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   975
      Index           =   1
      Left            =   1500
      Top             =   1021
      Width           =   8655
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8682.765
      X2              =   9402.579
      Y1              =   7861.279
      Y2              =   7861.279
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2200.433
      X2              =   2929.246
      Y1              =   7881.747
      Y2              =   7881.747
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2146.112
      Y2              =   7884.67
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8699.76
      X2              =   8699.76
      Y1              =   2149.036
      Y2              =   7884.67
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM MAINTENANCE MENU"
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
      Left            =   2832
      TabIndex        =   0
      Top             =   1248
      Width           =   6012
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
Attribute VB_Name = "frmFAItemMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAddNewItem_Click()
  Dim One As Integer
  Dim FileHandle As Integer
  
  On Error Resume Next
  If Exist("FADPREDT.DAT") Then
    If MsgBox("A depreciation process has been executed but has not been posted. Adding a new item under these conditions would exclude the added item from the current unposted depreciation. Would you like to jump to the depreciation posting screen?", vbYesNo) = vbYes Then
      One = 1
      FileHandle = FreeFile
      Open "fromItemMaintMenu.dat" For Output As FileHandle Len = 2
      Print #FileHandle, One
      Close FileHandle
      frmFAYearEndPost.Show
      DoEvents
      Unload frmFAItemMaintMenu
      Exit Sub
    Else
      Exit Sub
    End If
  End If
  GRecNum = 0
  AddItemFlag = True
  frmFAEditItemWTabs.Caption = "Fixed Assets Add New Item"
  frmFAEditItemWTabs.Label2 = "Fixed Assets Add New Item"
  frmFAEditItemWTabs.Show
  DoEvents
  Unload frmFAItemMaintMenu
End Sub

Private Sub cmdEditExisting_Click()
  Dim One As Integer
  Dim FileHandle As Integer
  
  On Error Resume Next
  If Exist("FADPREDT.DAT") Then
    If MsgBox("A depreciation process has been executed but has not been posted. Editing an item under these conditions would exclude any change from the current unposted depreciation. Would you like to jump to the depreciation posting screen?", vbYesNo) = vbYes Then
      One = 1
      FileHandle = FreeFile
      Open "fromItemMaintMenu.dat" For Output As FileHandle Len = 2
      Print #FileHandle, One
      Close FileHandle
      frmFAYearEndPost.Show
      DoEvents
      Unload frmFAItemMaintMenu
      Exit Sub
    Else
      Exit Sub
    End If
  End If
  
  frmFAItemLookUp.Show
  DoEvents
  Unload frmFAItemMaintMenu
End Sub

Private Sub cmdExit_Click()
  KillFile "itemmaintmenu.dat"
  frmFAMainMenu.Show
  Close
  DoEvents
  Unload frmFAItemMaintMenu
  
End Sub

Private Sub cmdMasterItemList_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmFAMasterItemListing.Show
  DoEvents
  Unload frmFAReportMenu

End Sub

Private Sub Form_Load()
  Dim One As Integer
  Dim FileHandle As Integer
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  AddItemFlag = False
  
  One = 1
  FileHandle = FreeFile
  Open "itemmaintmenu.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAItemMaintMenu.")
      Call Terminate
      End
    End If
  End If
End Sub




