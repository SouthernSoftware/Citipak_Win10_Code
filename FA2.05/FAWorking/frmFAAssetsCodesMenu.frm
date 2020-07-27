VERSION 5.00
Begin VB.Form frmFAAssetsCodesMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assets Codes Menu"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Assets Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4005
      TabIndex        =   3
      Top             =   5616
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintCodeListing 
      BackColor       =   &H008F8265&
      Caption         =   "&Print Code Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4005
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   4848
      Width           =   3612
   End
   Begin VB.CommandButton cmdEditExistingCode 
      BackColor       =   &H008F8265&
      Caption         =   "&Edit Existing Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4005
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   4080
      Width           =   3612
   End
   Begin VB.CommandButton cmdAddNewAssetCode 
      BackColor       =   &H008F8265&
      Caption         =   "&Add New Asset Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4005
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   3312
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1500
      Top             =   886
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8590
      X2              =   9540
      Y1              =   2029.445
      Y2              =   2029.445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2029.445
      Y2              =   2029.445
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8703
      X2              =   9403
      Y1              =   7882.003
      Y2              =   7882.003
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2219
      X2              =   2929
      Y1              =   7882.003
      Y2              =   7882.003
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2146.36
      Y2              =   7867.388
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2146.36
      Y2              =   2146.36
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8592
      X2              =   9542
      Y1              =   2146.36
      Y2              =   2146.36
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8700
      X2              =   8700
      Y1              =   2149.283
      Y2              =   7870.311
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   9550
      X2              =   9550
      Y1              =   2140.514
      Y2              =   2028.471
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   8590
      X2              =   8590
      Y1              =   2046.008
      Y2              =   2148.309
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2100
      X2              =   2100
      Y1              =   2029.445
      Y2              =   2146.36
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   2146.36
      Y2              =   2029.445
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8580
      X2              =   9540
      Y1              =   1912.53
      Y2              =   1912.53
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   2100
      X2              =   3060
      Y1              =   1912.53
      Y2              =   1912.53
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ASSETS CODES MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   4
      Top             =   1246
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   766
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8700
      Top             =   2194
      Width           =   732
   End
End
Attribute VB_Name = "frmFAAssetsCodesmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAddNewAssetCode_Click()
  frmFAEditAssetCode.Show
  DoEvents
  Unload frmFAAssetsCodesmenu
End Sub

Private Sub cmdEditExistingCode_Click()
  frmFACodeLookUp.Show
  DoEvents
  Unload frmFAAssetsCodesmenu
End Sub

Private Sub cmdExit_Click()
  frmFAMainMenu.Show
  DoEvents
  Unload frmFAAssetsCodesmenu
End Sub

Private Sub cmdPrintCodeListing_Click()

  ReDim Arr(1 To 1) As Struct 'Template for the sort array
  Dim CodeRec As FAAssetCodeRecType
  Dim CodeRecLen As Integer
  Dim Dash80$, FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ItemCnt As Integer
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim FAFile As Integer
  Dim NumOfFARecs As Integer
  Dim Cnt&
  Dim CodeRecNo As Integer
  Dim Page As Integer
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisX As Integer
  Dim SmallNum As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempArr As Struct
  Dim OddRecNums As Integer
  
  CodeRecLen = Len(CodeRec)

  ReportFile$ = "FACODE.PRN"   'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)

  MaxLines = 50
  LineCnt = 0
  ItemCnt = 0

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle

  GoSub PrintHeader

  OpenFACodeNameFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(CodeRec)

  GoSub GetIndex

  For Cnt = 1 To NumOfFARecs
    CodeRecNo = Arr(Cnt).RecNum
    Get FAFile, CodeRecNo, CodeRec
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    'Check For Disposed Of
    Print #RptHandle, CodeRec.ASSETCODE;
    Print #RptHandle, Tab(25); CodeRec.AssetDesc;
    Print #RptHandle, Tab(70); CodeRec.AssetStatus
    LineCnt = LineCnt + 1
SkipEm:
  Next Cnt

  GoSub PrintEnding
  Close         'Close all open files now

  ViewPrint ReportFile$, "Code Listing", False

  Kill ReportFile$
  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(29); "Master Code Listing"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle,
  Print #RptHandle, "Asset Catagory Code"; Tab(25); "Description"; Tab(70); "Status"

  Print #RptHandle, Dash80$
  LineCnt = 6
  Return

PrintEnding:
  Print #RptHandle, FF$
  Return

GetIndex:
  ReDim Arr(1 To NumOfFARecs) As Struct
  For Cnt = 1 To NumOfFARecs
    Get FAFile, Cnt, CodeRec
    Arr(Cnt).who = LTrim$(CodeRec.ASSETCODE)
    Arr(Cnt).RecNum = Cnt
  Next

  Call SortAssetCodes(Arr(), NumOfFARecs, 1, 1, False)
  
  Return

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  CodeNum = 0
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
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
      Call UnloadAllFormsAndOpn
'      ClearInUse PWcnt
      MainLog ("Payroll.exe terminated via menu bar on frmPayrollMainMenu.")
      End
    End If
  End If
End Sub

