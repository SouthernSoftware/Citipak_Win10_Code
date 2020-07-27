VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFAItemLookUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Asset Item LookUp"
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
   Begin LpLib.fpList fpListSearch 
      Height          =   2004
      Left            =   1284
      TabIndex        =   9
      Top             =   6000
      Width           =   9096
      _Version        =   196608
      _ExtentX        =   16044
      _ExtentY        =   3535
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
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   1
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
      BorderStyle     =   1
      BorderColor     =   8454143
      BorderWidth     =   2
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   3
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
      ColumnHeaderShow=   -1  'True
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
      ColDesigner     =   "frmItemLookUp.frx":0000
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "F10 &Search"
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
      Left            =   6204
      TabIndex        =   2
      Top             =   4560
      Width           =   1884
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
      Left            =   3564
      TabIndex        =   1
      Top             =   4560
      Width           =   1884
   End
   Begin EditLib.fpText fptxtTagNumber 
      Height          =   396
      Left            =   4368
      TabIndex        =   0
      ToolTipText     =   $"frmItemLookUp.frx":032F
      Top             =   2304
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   ""
      MaxLength       =   150
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Left            =   4368
      TabIndex        =   5
      ToolTipText     =   $"frmItemLookUp.frx":03D6
      Top             =   2928
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   ""
      MaxLength       =   150
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fpTxtSerialNum 
      Height          =   396
      Left            =   4368
      TabIndex        =   7
      ToolTipText     =   $"frmItemLookUp.frx":047D
      Top             =   3552
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   ""
      MaxLength       =   150
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   1824
      TabIndex        =   8
      Top             =   3600
      Width           =   2124
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   1824
      TabIndex        =   6
      Top             =   2976
      Width           =   2124
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2400
      TabIndex        =   4
      Top             =   2352
      Width           =   1548
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3900
      Left            =   1008
      Top             =   1776
      Width           =   9804
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   624
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets Item LookUp"
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
      TabIndex        =   3
      Top             =   768
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   576
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAItemLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFAMainMenu.Show
  DoEvents
  Unload frmFAItemLookUp
End Sub

Private Sub cmdSearch_Click()
  Dim FAHandle As Integer
  Dim NumOfRecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim x As Long
  Dim Found As Boolean
  Dim TagFlag As Boolean
  Dim DescFlag As Boolean
  Dim SerialFlag As Boolean
  Dim TempTag$
  Dim TempDesc$
  Dim TempSerial$
  Dim FoundCnt As Integer
  Dim MatchCnt As Integer
  Dim PrintDesc$
  Dim OnlyOneFound$
  
  fpListSearch.Clear
  
  TagFlag = False
  DescFlag = False
  SerialFlag = False
  
  If QPTrim$(fptxtTagNumber.Text) <> "" Then
    TagFlag = True
    TempTag = QPTrim$(fptxtTagNumber)
  ElseIf QPTrim$(fptxtDesc.Text) <> "" Then
    DescFlag = True
    TempDesc = QPTrim$(fptxtDesc)
  ElseIf QPTrim$(fpTxtSerialNum.Text) <> "" Then
    SerialFlag = True
    TempSerial = QPTrim$(fpTxtSerialNum)
  End If
  
  OpenFAItemFile FAHandle
  
  NumOfRecs = LOF(FAHandle) \ Len(FAItemRec)
  If NumOfRecs = 0 Then
    MsgBox "No records on file."
    Close FAHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfRecs
    Get FAHandle, x, FAItemRec
    Found = True
    If TagFlag = True Then
      If InStr(UCase$(FAItemRec.ITEMTAG), TempTag) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    ElseIf DescFlag = True Then
      If InStr(UCase$(FAItemRec.IDESC1), TempDesc) > 0 Or InStr(UCase$(FAItemRec.IDESC2), TempDesc) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    ElseIf SerialFlag = True Then
      If InStr(UCase$(FAItemRec.SERIALNO), TempSerial) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    If Found Then
      FoundCnt = FoundCnt + 1
      fpListSearch.Row = -1
      MatchCnt = MatchCnt + 1
      RecNum = x
      If QPTrim$(FAItemRec.IDESC1) <> "" Then
        PrintDesc$ = QPTrim$(FAItemRec.IDESC1)
      Else
        PrintDesc$ = QPTrim$(FAItemRec.IDESC2)
      End If
      fpListSearch.InsertRow = "   " & QPTrim$(FAItemRec.ITEMTAG) & Chr$(9) & "    " & PrintDesc$ & "  " & Chr$(9) & "  " & QPTrim$(FAItemRec.SERIALNO)
      DoEvents
      'only used if no more than one found
      OnlyOneFound = QPTrim$(FAItemRec.ITEMTAG)
    End If
NotAMatch:
  Next x
  
  If MatchCnt <= 0 Then
    MsgBox "No match found"
    Exit Sub
    Close
  End If
    
  If FoundCnt = 1 Then
'    frmLoadingEmpEdit.Show
'    DoEvents
    For x = 1 To NumOfRecs
      Get FAHandle, x, FAItemRec
        If OnlyOneFound = QPTrim$(FAItemRec.ITEMTAG) Then
          RecNum = x
          Exit For
        Else
          Found = False
          GoTo NotThisTime
        End If
NotThisTime:
    Next x
    
    fptxtTagNumber.Text = ""
    fptxtDesc.Text = ""
    fpTxtSerialNum.Text = ""
    fpListSearch.Clear
    FoundCnt = 0
    frmFAEditItem.Caption = "Fixed Asset Edit Item"
    frmFAEditItem.Label2 = "Fixed Asset Edit Item"
    frmFAEditItem.Show
    DoEvents
    Unload frmFAItemLookUp
  End If
  Close FAHandle
    
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
'    Case vbKeyF7:
'      Call cmdChoice_Click
'      SendKeys "%H"
'      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
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


Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub


Private Sub fpListSearch_DblClick()
  Dim FAHandle As Integer
  Dim NumOfRecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim x As Long
  Dim TagNum$
  Dim Desc$
  Dim SerialNum$
  Dim PrintDesc$
  Dim Found As Boolean
  
  fpListSearch.Col = 0
  'trap for double clicking on nothing
  If QPTrim$(fpListSearch.ColText) = "" Then
    MsgBox "No item has been selected"
    Exit Sub
  End If
  TagNum$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 1
  Desc$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 2
  SerialNum$ = QPTrim$(fpListSearch.ColText)
  
  OpenFAItemFile FAHandle
  NumOfRecs = LOF(FAHandle) \ Len(FAItemRec)
  For x = 1 To NumOfRecs
    Get FAHandle, x, FAItemRec
    If QPTrim$(FAItemRec.IDESC1) <> "" Then
      PrintDesc$ = QPTrim$(FAItemRec.IDESC1)
    Else
      PrintDesc$ = QPTrim$(FAItemRec.IDESC2)
    End If
  
    If InStr(UCase$(FAItemRec.ITEMTAG), TagNum$) > 0 And InStr(UCase$(PrintDesc$), Desc$) > 0 And InStr(FAItemRec.SERIALNO, SerialNum$) >= 0 _
    And Len(QPTrim$(FAItemRec.ITEMTAG)) = Len(QPTrim$(TagNum$)) Then '8/7 added Len = Len because
    'if two people had the same name and the emp number of one had a number that
    'included the other's (ie. 123 vs 1234) then then smaller number would not be accessed ever
      Found = True
      fpListSearch.Row = -1
      RecNum = x
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close FAHandle
  
  frmFAEditItem.Caption = "Fixed Asset Edit Item"
  frmFAEditItem.Label2 = "Fixed Asset Edit Item"
  frmFAEditItem.Show
  DoEvents
  Unload frmFAItemLookUp
    

End Sub
