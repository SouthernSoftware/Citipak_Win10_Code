VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFACodeLookUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Code Listing"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   3792
      Left            =   2892
      TabIndex        =   1
      Top             =   2352
      Width           =   5820
      _Version        =   196608
      _ExtentX        =   10266
      _ExtentY        =   6689
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   3
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmFACodeLookUp.frx":0000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   4608
      TabIndex        =   3
      Top             =   7296
      Width           =   1884
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "F10 &Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   7188
      TabIndex        =   2
      Top             =   7296
      Width           =   1884
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master Fixed Asset Code Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   0
      Top             =   1068
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   924
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   4764
      Left            =   2340
      Top             =   1932
      Width           =   6972
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   876
      Width           =   8652
   End
End
Attribute VB_Name = "frmFACodeLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFAAssetsCodesmenu.Show
  DoEvents
  Unload frmFACodeLookUp
End Sub

Private Sub cmdProcess_Click()
  Call fpList1_DblClick
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
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
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%T"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%L"
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmPayrollMainMenu.")
      End
    End If
  End If
End Sub

Private Sub LoadMe()
   Dim CodeRec As FAAssetCodeRecType
   Dim ACHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim DoWhatFlag As BadFACodeNumOption
   Dim n As Integer
   Dim Nextx As Integer
   Dim Y As Integer, Cnt As Integer
   Dim ThisText$, CodeRecNo As Integer
   
   On Error GoTo ERRORSTUFF
   
   If Not Exist("FACODES.DAT") Then
     MsgBox "Path to FACODES.DAT could not be found"
     Exit Sub
   End If

   OpenFACodeNameFile ACHandle
   TotalAccts = LOF(ACHandle) \ Len(CodeRec)
   
   If TotalAccts = 0 Then Exit Sub
   ReDim Arr(1 To TotalAccts) As Struct
   GoSub GetIndex
   
   For x = 1 To TotalAccts
     CodeRecNo = Arr(x).RecNum
     Get ACHandle, CodeRecNo, CodeRec
     If Len(QPTrim(CodeRec.ASSETCODE)) = 0 Then GoTo BadCode
     fpList1.InsertRow = QPTrim$(CodeRec.ASSETCODE) & " " & Chr$(9) & QPTrim$(CodeRec.AssetDesc) & " " & Chr$(9) & QPTrim$(CodeRec.AssetStatus)
BadCode:
   Next x
   Close ACHandle
   fpList1.Row = 0
   fpList1.Selected = True
ZeroText:
   Exit Sub
   
GetIndex:
  For Cnt = 1 To TotalAccts
    Get ACHandle, Cnt, CodeRec
    Arr(Cnt).who = LTrim$(CodeRec.ASSETCODE)
    Arr(Cnt).RecNum = Cnt
  Next

  Call SortAssetCodes(Arr(), TotalAccts, 1, 1, False)
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmGLPickList", "Form Load", Erl)
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
    Unload Me
End Sub

Private Sub fpList1_DblClick()
   Dim CodeRec As FAAssetCodeRecType
   Dim ACHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim Desc$
   Dim Code$
   Dim Status$
   Dim Found As Boolean
   
   
   fpList1.Col = 0
   Code$ = QPTrim$(fpList1.ColText)
   fpList1.Col = 1
   Desc$ = QPTrim$(fpList1.ColText)
   fpList1.Col = 2
   Status$ = QPTrim$(fpList1.ColText)
   
   OpenFACodeNameFile ACHandle
   TotalAccts = LOF(ACHandle) \ Len(CodeRec)
   
   If TotalAccts = 0 Then Exit Sub
   
   For x = 1 To TotalAccts
     Get ACHandle, x, CodeRec
     If Code$ = QPTrim$(CodeRec.ASSETCODE) And Desc$ = QPTrim$(CodeRec.AssetDesc) And Status$ = QPTrim$(CodeRec.AssetStatus) Then
       Found = True
       fpList1.Row = -1
       CodeNum = x
       Exit For
     Else
       Found = False
       GoTo NotAMatch
     End If
      
NotAMatch:
   Next x
  Close ACHandle
  
  If Found = True Then
    frmFAEditAssetCode.Show
    DoEvents
    Unload frmFACodeLookUp
  Else
    MsgBox "No match found. Please try again."
    Exit Sub
  End If
End Sub
