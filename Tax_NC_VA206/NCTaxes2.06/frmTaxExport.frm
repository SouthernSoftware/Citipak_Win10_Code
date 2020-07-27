VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxExport 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Export Function"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxExport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbExOrder 
      Height          =   405
      Left            =   5273
      TabIndex        =   0
      Top             =   3758
      Width           =   3015
      _Version        =   196608
      _ExtentX        =   5318
      _ExtentY        =   714
      Text            =   ""
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
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
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
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
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmTaxExport.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3473
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6998
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmTaxExport.frx":0CA5
   End
   Begin EditLib.fpText fptxtRFile 
      Height          =   495
      Left            =   5693
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4358
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   873
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
      AutoAdvance     =   0   'False
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
      Text            =   "RTAXDATA1"
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   6113
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6998
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmTaxExport.frx":0E81
   End
   Begin EditLib.fpText fptxtPFile 
      Height          =   495
      Left            =   5693
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4958
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   873
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
      AutoAdvance     =   0   'False
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
      Text            =   "PTAXDATA1"
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export Personal To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3173
      TabIndex        =   10
      Top             =   5078
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".TXT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7853
      TabIndex        =   9
      Top             =   5003
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".TXT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7853
      TabIndex        =   7
      Top             =   4403
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export Real To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3533
      TabIndex        =   5
      Top             =   4478
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export Orders:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3353
      TabIndex        =   2
      Top             =   3878
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2655
      Left            =   2513
      Top             =   3278
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1493
      Top             =   1298
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Export Function"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2933
      TabIndex        =   1
      Top             =   1433
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1493
      Top             =   1238
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmTaxCustMaintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Call PrintIt
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
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpExportCustomer
  Call LoadMe

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxExport.")
      Call Terminate
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If

End Sub

Private Sub LoadMe()
  fpcmbExOrder.Text = "Name Order"
  fpcmbExOrder.AddItem "Name Order"
  fpcmbExOrder.AddItem "Account Number"
  fpcmbExOrder.AddItem "Search Name"
End Sub

Private Sub fpcmbExOrder_Change()
  If QPTrim$(fpcmbExOrder.Text) = "" Then
    fpcmbExOrder.Text = "Name Order"
  End If
End Sub

Private Sub PrintIt()
  Dim TaxSetupRec As TaxMasterType
  Dim TaxCustRec As TaxCustType
  Dim PropertyRec As PropertyRecType
  Dim PersRec As PersonalRecType
  Dim ReportFile$
  Dim x As Long
  Dim CustCnt As Long
  Dim RptHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim DetailFlag As Boolean
  Dim RptOutR As Integer
  Dim RptOutRFile$
  Dim RptOutP As Integer
  Dim RptOutPFile$
  Dim NumOfPersRecs As Long
  Dim PersHandle As Integer
  Dim RealRec As PropertyRecType
  Dim NumOfRealRecs As Long
  Dim RealHandle As Integer
  Dim TaxRec As TaxCustType
  Dim NumOfTaxRecs As Long
  Dim TaxHandle As Integer
  Dim UseNameIdx As Integer
  Dim PropertyRecord As Long
  Dim FileNameR$, CustRecNo As Long
  Dim FileNameP$
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim dlm$
  
  'on error goto ERRORSTUFF
  dlm$ = "~"
  UseNameIdx = 0
'  ReportFile$ = "TAXRPTS\TaxExport.RPT"   'Report File Name
  DetailFlag = False
  CustCnt = 0
  
  OpenTaxCustFile TaxHandle, NumOfTaxRecs
  If QPTrim$(fpcmbExOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbExOrder.SetFocus
      Close
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    UseNameIdx = 1
    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    NumOfTaxRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbExOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfSrchRecs
    If NumOfSrchRecs = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbExOrder.SetFocus
      Close
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfSrchRecs) As Long
    UseNameIdx = 2
    For x = 1 To NumOfSrchRecs
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    NumOfTaxRecs = NumOfSrchRecs
  End If
  
  FileNameR$ = QPTrim$(fptxtRFile.Text) + ".TXT"
  FileNameP$ = QPTrim$(fptxtPFile.Text) + ".TXT"
  If QPTrim$(FileNameR$) = QPTrim$(FileNameP) Then
    MsgBox ("The two file names must each be unique.")
    fptxtRFile.SetFocus
    Exit Sub
  End If
  
  RptOutR = FreeFile
  Open FileNameR$ For Output As #RptOutR
  RptOutP = FreeFile
  Open FileNameP$ For Output As #RptOutP

  OpenPersPropFile PersHandle, NumOfPersRecs
  OpenRealPropFile RealHandle, NumOfRealRecs
  
  frmTaxShowPctComp.Label1 = "Gathering Export Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  Print #RptOutR, "Type"; dlm; "Cust #"; dlm; "Cust Name"; dlm; "Add 1"; dlm; "Add 2"; dlm;
  Print #RptOutR, "City"; dlm; "State"; dlm; "Zip"; dlm; "Soc Sec #"; dlm; "Real Pin#"; dlm;
  Print #RptOutR, "Prop Note 1"; dlm; "Prop Note 2"; dlm; "Prop Note 3"; dlm; "Prop Value"; dlm;
  Print #RptOutR, "Senior Exemption"; dlm; "Other Exemption"; dlm; "Mort Code"; dlm; "Map"; dlm; "Block"; dlm; "Lot Number"; dlm;
  Print #RptOutR, "Late List"; dlm; "Opt'l Search"
  
  Print #RptOutP, "Type"; dlm; "Cust #"; dlm; "Cust Name"; dlm; "Add 1"; dlm; "Add 2"; dlm;
  Print #RptOutP, "City"; dlm; "State"; dlm; "Zip"; dlm; "Soc Sec #"; dlm; "Pers Pin #"; dlm;
  Print #RptOutP, "Pers Value"; dlm; "Senior Exemption"; dlm; "Other Exemption"; dlm; "Desc 1"; dlm;
  Print #RptOutP, "Desc 2"; dlm; "Desc 3"; dlm; "Desc 4"; dlm; "Desc 5"; dlm; "Late List"; dlm;
  Print #RptOutP, "Opt'l Search"; dlm; "Merch Cap Value"; dlm; "Farm Value"; dlm; "Mobile Home Value"; dlm;
  Print #RptOutP, "Mach Tools Value"; dlm; "Pers Value"
  
  For x = 1 To NumOfTaxRecs
    If UseNameIdx > 0 Then
      CustRecNo = IdxArray(x)
    Else
      CustRecNo = x
    End If

    Get TaxHandle, CustRecNo, TaxRec

    If Not TaxRec.Deleted Then
      If TaxRec.FirstPropRec > 0 Then 'closed
        PropertyRecord = TaxRec.FirstPropRec
        Do While PropertyRecord <> 0
          Get #RealHandle, PropertyRecord, RealRec
          Print #RptOutR, "R"; dlm;
          Print #RptOutR, Using("#####", CustRecNo);
          Print #RptOutR, dlm;
          Print #RptOutR, QPTrim$(TaxRec.CustName); dlm;
          Print #RptOutR, QPTrim$(TaxRec.Addr1); dlm;
          Print #RptOutR, QPTrim$(TaxRec.Addr2); dlm;
          Print #RptOutR, QPTrim$(TaxRec.City); dlm;
          Print #RptOutR, TaxRec.State; dlm;
          Print #RptOutR, TaxRec.Zip; dlm;
          Print #RptOutR, TaxRec.CSSN; dlm;
          Print #RptOutR, QPTrim$(RealRec.RealPin); dlm;
          Print #RptOutR, RealRec.PROPNOT1; dlm;
          Print #RptOutR, RealRec.PROPNOT2; dlm;
          Print #RptOutR, RealRec.PROPNOT3; dlm;
          Print #RptOutR, Using("##########.##", RealRec.PROPVALU);
          Print #RptOutR, dlm;
          Print #RptOutR, Using("##########.##", RealRec.EXMPSENI);
          Print #RptOutR, dlm;
          Print #RptOutR, Using("##########.##", RealRec.EXMPOTHR); dlm;
          Print #RptOutR, QPTrim$(RealRec.MORTCODE); dlm;
          Print #RptOutR, RealRec.Map; dlm;
          Print #RptOutR, RealRec.BLOCK; dlm;
          Print #RptOutR, QPTrim$(RealRec.LOTNUMB); dlm;
          Print #RptOutR, RealRec.LateList; dlm;
          Print #RptOutR, QPTrim$(RealRec.OptSearch)
          PropertyRecord = RealRec.NextRec
          Loop
      End If

        'NOW CHECK PERSONAL PROPERTY
      If TaxRec.FirstPersRec > 0 Then 'closed
        PropertyRecord = TaxRec.FirstPersRec
        Do While PropertyRecord <> 0
          Get #PersHandle, PropertyRecord, PersRec
          Print #RptOutP, "P"; dlm;
          Print #RptOutP, Using("#####", CustRecNo);
          Print #RptOutP, dlm;
          Print #RptOutP, TaxRec.CustName; dlm;
          Print #RptOutP, TaxRec.Addr1; dlm;
          Print #RptOutP, TaxRec.Addr2; dlm;
          Print #RptOutP, TaxRec.City; dlm;
          Print #RptOutP, TaxRec.State; dlm;
          Print #RptOutP, TaxRec.Zip; dlm;
          Print #RptOutP, TaxRec.CSSN; dlm;
          Print #RptOutP, PersRec.PropPin; dlm;
          Print #RptOutP, Using("########.##", PersRec.PersVal);
          Print #RptOutP, dlm;
          Print #RptOutP, Using("########.##", PersRec.EXMPSENI);
          Print #RptOutP, dlm;
          Print #RptOutP, Using("########.##", PersRec.EXMPOTHR);
          Print #RptOutP, dlm;
          Print #RptOutP, PersRec.DESC1; dlm;
          Print #RptOutP, PersRec.DESC2; dlm;
          Print #RptOutP, PersRec.DESC3; dlm;
          Print #RptOutP, PersRec.Desc4; dlm;
          Print #RptOutP, PersRec.Desc5; dlm;
          Print #RptOutP, PersRec.LateList; dlm;
          Print #RptOutP, PersRec.OptSearch; dlm;
          Print #RptOutP, Using("########.##", PersRec.MCVALUE); dlm;
          Print #RptOutP, Using("########.##", PersRec.CVALUE); dlm;
          Print #RptOutP, Using("########.##", PersRec.MHVALUE); dlm;
          Print #RptOutP, Using("########.##", PersRec.MTVALUE); dlm;
          Print #RptOutP, Using("########.##", PersRec.PersVal)
          PropertyRecord = PersRec.NextRec
         Loop
       End If
     End If
SkipEm:
     CustCnt = CustCnt + 1
     frmTaxShowPctComp.ShowPctComp x, NumOfTaxRecs
     If frmTaxShowPctComp.Out = True Then
       Close
       frmTaxShowPctComp.Out = False
       Unload frmTaxShowPctComp
       EnableCloseButton Me.hwnd, True
       cmdProcess.Enabled = True
       cmdExit.Enabled = True
       Exit Sub
     End If
  Next
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If CustCnt = 0 Then
    frmTaxMsg.Label1.Caption = "There are no export files to print."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  Close
  MsgBox ("The files have been saved in the Citipak folder sucessfully.")
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxExport", "PrintIt", Erl)
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

Private Sub fptxtPFile_Change()
  If InStr(fptxtPFile.Text, ".") Then
    Call TaxMsg(900, "This file will be saved as a .txt file by default. Please do not use a '.' in the name.")
    fptxtPFile.Text = ReplaceString(fptxtPFile.Text, ".", "")
    Exit Sub
  End If

End Sub

Private Sub fptxtRFile_Change()
  If InStr(fptxtRFile.Text, ".") Then
    Call TaxMsg(900, "This file will be saved as a .txt file by default. Please do not use a '.' in the name.")
    fptxtRFile.Text = ReplaceString(fptxtRFile.Text, ".", "")
    Exit Sub
  End If
End Sub
