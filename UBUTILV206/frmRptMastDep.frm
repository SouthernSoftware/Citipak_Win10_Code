VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptMastDep 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Deposit Listing"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptMastDep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   375
      Left            =   5610
      TabIndex        =   2
      Top             =   3840
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      Text            =   ""
      Columns         =   1
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
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
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptMastDep.frx":08CA
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Credits in revenues from dep trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1170
      TabIndex        =   14
      Top             =   7380
      Width           =   4005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print those with both but do not match"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1170
      TabIndex        =   13
      Top             =   6876
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Dep that do not match"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1170
      TabIndex        =   12
      Top             =   6374
      Width           =   4035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Cust Dep No/Distributions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1170
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Only Dep w/ Distributions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1170
      TabIndex        =   10
      Top             =   5370
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   9360
      TabIndex        =   4
      Top             =   6690
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F10 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7680
      TabIndex        =   3
      Top             =   6690
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "11:00 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/4/2012"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   5628
      TabIndex        =   1
      Top             =   3312
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5628
      TabIndex        =   0
      Top             =   2808
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   2085
      Left            =   2610
      Top             =   2475
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Route:"
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
      Height          =   324
      Index           =   2
      Left            =   4080
      TabIndex        =   9
      Top             =   3372
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Route:"
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
      Height          =   324
      Index           =   0
      Left            =   3984
      TabIndex        =   8
      Top             =   2868
      Width           =   1476
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Order:"
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
      Height          =   324
      Index           =   7
      Left            =   3264
      TabIndex        =   7
      Top             =   3864
      Width           =   2220
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   855
      Left            =   2092
      Top             =   810
      Width           =   7965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Calculated vs Stored Master Deposit Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2092
      TabIndex        =   6
      Top             =   1080
      Width           =   8010
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   975
      Left            =   2092
      Top             =   720
      Width           =   7995
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmRptMastDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim PrnOptions As Integer ' 1 is for only with Distributions, 2 is for only deposits on cust and no dist
' 3 is for any that do not match, 4 is for only those that do not match but both have amt,
'  5 is for all with either

Private Sub cmdExit_Click()
 Load frmUBEditMenu
  frmUBEditMenu.Show
  Unload Me
  
End Sub


Private Sub Command1_Click()
PrnOptions = 1
cmdPrint_Click
'If MsgBox("Are you sure you wish to recalculate deposit amounts and change the balance shown on customer?", vbYesNo) = vbYes Then
'  CalcDepBal
'End If
End Sub

Private Sub Command2_Click()
PrnOptions = 2
cmdPrint_Click
End Sub

Private Sub Command3_Click()
PrnOptions = 4
cmdPrint_Click
End Sub

Private Sub Command4_Click()
PrnOptions = 4
cmdPrint_Click
End Sub

Private Sub Command5_Click()
PrnOptions = 5
cmdPrint_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptMastDep by " + "Util OPer"
        
      End If
    End If
  End If
End Sub

Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
     
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub
Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboPrintOrder.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Route Selection, The Beginning Route Should Be Less or Equal to Ending Route.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
    End If
  Else
    MsgBox "Route Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function

Private Sub cmdPrint_Click()
  If ValidRoutes Then
    DeActivateControls Me, True
      SpecialBalRpt
    ActivateControls Me, True
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fptxtRoute1 = "00"
  fptxtRoute2 = "99"
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub SpecialBalRpt()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer, UBTransRecLen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, TCnt As Integer, Detail As String
  Dim AcctNumber As Long, UBCust As Integer, UsingAcct As Boolean, Depoff As Integer
  Dim IndexName As String, UBRpt As Integer, SEQNUMB As String, GTDep As Double
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer, ToPrintD2 As String
  Dim cnt As Long, TDeposit As Double, ToPrint As String, RevCnt As Integer, RCnt As Integer
  Dim Book As String, CustCnt As Long, ReportFile As String, MaxRevenue As Integer
  Dim GTDeptot As Double, UBTrans As Integer, Trans As Long
  Dim First As Integer, Last As Integer, Rev As String, AndPos As String
  Dim TabStop As Integer, Det As Boolean, Order As String
  Dim ToPrintD As String, ToPrintH1 As String, TRevName As String
  Dim ToPrintH2 As String, UBRpt2 As Integer, ToPrintS As String
  Dim DetFlag As Integer
  Dim Report2 As String
  Dim IndPRevAmt As Double, IndCRevAmt As Double
  Det = True
  UsingBook = False
  UsingAcct = False
  UsingName = False
ReDim RevAmts(1 To 15) As Double
ReDim Deptot(1 To 15) As Double


  Select Case fpcboPrintOrder.ListIndex
    Case 0
      IndexName$ = NameIndexFile
      UsingName = True
    Case 1
      IndexName$ = ""
      UsingAcct = True
    Case 2
      IndexName$ = BookIndexFile
      UsingBook = True
   End Select
  MaxRevenue = 15
  '***************
  MaxLines = 52
  FrmShowPctComp.Label1 = "Creating Customer Deposit Listing"
  FrmShowPctComp.Show , Me
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  TOWNNAME$ = QPTrim$(UBSetUp(1).UTILNAME)
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim RevenueName(1 To 15) As String * 10
  For RCnt = 1 To 15
    TRevName$ = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    If Len(TRevName$) > 0 Then
      AndPos = InStr(TRevName$, "&")
      If AndPos Then
        Mid$(TRevName$, AndPos) = " "
      End If
      RevenueName(RCnt) = TRevName$
     Else
      RevenueName(RCnt) = "Rev - " + Str(RCnt)
    End If
  Next
  ToPrint$ = ""
  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
   ReDim RevTotals(1 To MaxRevenue) As Double

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBDPLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  
  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Then
      AcctNumber = IdxBuff(cnt).RecNum
    Else
      AcctNumber = cnt
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitDepositListing
    End If

    Get UBCust, AcctNumber, UBCustRec(1)
        GTDep = 0
        ReDim RevAmts(1 To 15) As Double
        ReDim Deptot(1 To 15) As Double
        Trans& = UBCustRec(1).LastTrans
        Do While Trans& <> 0
          Get UBTrans, Trans&, UBTransRec(1)
'              Select Case UBTransRec(1).TransType
'
'                Case 7    'deposit payment
'                  Deptot = Round#(Deptot + UBTransRec(1).Transamt)
'                Case 5 ' apply
'                  Deptot = Round#(Deptot + UBTransRec(1).Transamt)
'                Case 9 ' refund
'                  Deptot = Round#(Deptot - UBTransRec(1).Transamt)
'                Case 37 ' depcrdrem
'
'                Case 39 ' deppayvoid
'                  Deptot = Round#(Deptot - UBTransRec(1).Transamt)
'                Case Else

'              End Select
             'If AcctNumber = 3060 Then Stop
              If UBTransRec(1).TransType = TranDepositPayment Then
                  For RevCnt = 1 To 15
                    If UBTransRec(1).RevAmt(RevCnt) <> 0 Then
                      
                      RevAmts(RevCnt) = Round#(RevAmts(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
                      Deptot(RevCnt) = Round#(Deptot(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
                      If Deptot(RevCnt) < 0 Then Depoff = 1
                      GTDep = Round#(GTDep + UBTransRec(1).RevAmt(RevCnt))
                    End If
                  Next
                ElseIf (UBTransRec(1).TransType = TranAppliedDeposit) Or (UBTransRec(1).TransType = TranRefundDeposit) Or (UBTransRec(1).TransType = TranDepPaymentVoid) Then
                 For RevCnt = 1 To 15
                    If UBTransRec(1).RevAmt(RevCnt) <> 0 Then
                      
                      RevAmts(RevCnt) = Round#(RevAmts(RevCnt) - UBTransRec(1).RevAmt(RevCnt))
                      Deptot(RevCnt) = Round#(Deptot(RevCnt) - UBTransRec(1).RevAmt(RevCnt))
                     If Deptot(RevCnt) < 0 Then Depoff = 1
                     GTDep = Round#(GTDep - UBTransRec(1).RevAmt(RevCnt))
                    End If
                  Next
                End If
              


          Trans& = UBTransRec(1).PrevTrans
          Loop
          'If Deptot < 0 Then Deptot = 0
          ''UBCustRec(1).DepositAmt = Deptot
    'IF UBCustRec(1).Status <> "A" AND UBCustRec(1).Status <> "I" THEN
    If PrnOptions = 1 Then 'this is with dist only
'        If Round#(UBCustRec(1).DepositAmt) = 0 And Deptot > 0 Then
'           GoSub Printoneout
'        End If
     ElseIf PrnOptions = 2 Then 'this is for only custdep no dist
'         If Round#(UBCustRec(1).DepositAmt) > 0 And Deptot = 0 Then
'           GoSub Printoneout
'        End If
     ElseIf PrnOptions = 3 Then 'ANY THAT do not match
'        If Round#(UBCustRec(1).DepositAmt) <> Deptot Then
'           GoSub Printoneout
'        End If
     ElseIf PrnOptions = 4 Then
        If (UBCustRec(1).DepositAmt > 0) And Round#(UBCustRec(1).DepositAmt) <> GTDep Then
          GoSub Printoneout2
        End If
     ElseIf PrnOptions = 5 Then 'either have amt
'       If Round#(UBCustRec(1).DepositAmt) > 0 Or Deptot > 0 Then
'           GoSub Printoneout
'        End If

    For RevCnt = 1 To 15

      If (UBCustRec(1).DepositAmt > 0) Then 'And ((Round#(UBCustRec(1).DepositAmt) <> GTDep)) Then
       If Deptot(RevCnt) < 0 Then
          GoSub Printoneout
          Depoff = 0
        End If
      End If
    Next
    End If
  Next
  GoSub DepSkipEm
Printoneout:
 
  
    ToPrint$ = Str$(AcctNumber) + "~"
    ToPrint$ = ToPrint$ + Str(Deptot(RevCnt))
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 38)
    ToPrint$ = ToPrint$ + "~" + Str$(UBCustRec(1).DepositAmt)
    ToPrint$ = ToPrint$ + "~" + UBCustRec(1).Status
    
    TDeposit# = Round#(TDeposit# + UBCustRec(1).DepositAmt)
    'GTDeptot# = Round#(GTDeptot# + Deptot#)
    CustCnt = CustCnt + 1
    Print #UBRpt, ToPrint$
   ' GoSub PrintDetail
Return
Printoneout2:
 
  
    ToPrint$ = Str$(AcctNumber) + "~"
    ToPrint$ = ToPrint$ + Str(GTDep)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 38)
    ToPrint$ = ToPrint$ + "~" + Str$(UBCustRec(1).DepositAmt)
    ToPrint$ = ToPrint$ + "~" + UBCustRec(1).Status
    
    TDeposit# = Round#(TDeposit# + UBCustRec(1).DepositAmt)
    'GTDeptot# = Round#(GTDeptot# + Deptot#)
    CustCnt = CustCnt + 1
    Print #UBRpt, ToPrint$
   ' GoSub PrintDetail
Return
PrintDetail:

  TCnt = 0
  Detail$ = Space$(18)
  First = 1
  ToPrintD$ = ""
  ToPrintD2$ = ""
  Last = 15

  For RCnt = 1 To 15
   
 
      LSet Detail$ = RevenueName(RCnt)
      ToPrintD2$ = Str$(RevAmts(RCnt)) + "~"
  
    
    If Det Then
      ToPrintD$ = ToPrintD$ + QPTrim(Detail$) + "~" + ToPrintD2$
    Else
      ToPrintD$ = ToPrintD$ + "~~~"
    End If
  
  Next

  If Det Then
    'Print #UBRpt,
    Print #UBRpt, ToPrint$ + "~" + ToPrintD$
   Else
    'Linecnt = DLineCnt
    Print #UBRpt, ToPrint$ + "~" + ToPrintD$
  End If
  ToPrint$ = ""
  ToPrintD$ = ""
  Return
DepSkipEm:
'
'          ToPrint$ = "Total Trans Calc~"
'          ToPrint$ = ToPrint$ + Str(GTDeptot)
'          ToPrint$ = ToPrint$ + "~ "
'          ToPrint$ = ToPrint$ + "~ "
'          ToPrint$ = ToPrint$ + "~ "
 '         Print #UBRpt, ToPrint$
          ToPrint$ = ""
  Close UBCust, UBRpt

  Erase IdxBuff, UBCustRec
  If CustCnt > 0 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptMastDep
    ARptMastDepList.txtDate = Now
    ARptMastDepList.txtTown = TOWNNAME$
    ARptMastDepList.Title = "Calculated v Stored Customer Deposit List"
    ARptMastDepList.txtTotCust = CustCnt
    ARptMastDepList.GetName ReportFile$
    ARptMastDepList.startrpt

  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If

ExitDepositListing:
  Exit Sub


DoDepositRptFooter:
'  Print #UBRpt, Dash80$
'  Print #UBRpt, "Totals:"; Tab(10); "Customers: "; Using("#####,#", CustCnt);
'  Print #UBRpt, Tab(60); Using("#####,#.##", TDeposit#)
'
'  Print #UBRpt, FF$
    
  Return

End Sub
Private Sub CalcDepBal()
  Dim UBCustRecLen As Integer, UBTransRecLen As Integer
  Dim AcctNumber As Long, UBCust As Integer
  Dim NumOfRecs As Long, Handle As Integer
  Dim cnt As Long
  Dim CustCnt As Long
  Dim Deptot As Double, GTDeptot As Double, UBTrans As Integer, Trans As Long



  '***************
  FrmShowPctComp.Label1 = "Calculating Customer Deposits"
  FrmShowPctComp.Show , Me
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs = LOF(UBCust) \ UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  
  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
    End If
    AcctNumber = cnt
    Get UBCust, AcctNumber, UBCustRec(1)
        Deptot = 0
        Trans& = UBCustRec(1).LastTrans
        Do While Trans& <> 0
          Get UBTrans, Trans&, UBTransRec(1)
              Select Case UBTransRec(1).TransType
           
                Case 7    'deposit payment
                  Deptot = Round#(Deptot + UBTransRec(1).Transamt)
                Case 5 ' apply
                  Deptot = Round#(Deptot - Abs(UBTransRec(1).Transamt))
                Case 9 ' refund
                  Deptot = Round#(Deptot - Abs(UBTransRec(1).Transamt))
                Case 37 ' depcrdrem
                  
                Case 39 ' deppayvoid
                  Deptot = Round#(Deptot - Abs(UBTransRec(1).Transamt))
                Case Else
              End Select
          Trans& = UBTransRec(1).PrevTrans
          Loop
        If Deptot <> UBCustRec(1).DepositAmt Then
          If Deptot <= 0 Then Deptot = 0
          UBCustRec(1).DepositAmt = Deptot
          Put UBCust, AcctNumber, UBCustRec(1)
          GTDeptot# = Round#(GTDeptot# + Deptot#)
          CustCnt = CustCnt + 1
        End If
 
  Next

  Close UBCust
  MsgBox "Changed - " + Str(CustCnt), vbOKOnly
  UBLog "Recalculated deposits for - " + Str(CustCnt) + " customers by Util OPer"
End Sub


