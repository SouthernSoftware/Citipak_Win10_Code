VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmRptMastCust 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Customer Report"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptMastCust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboCustStatus 
      Height          =   375
      Left            =   5100
      TabIndex        =   1
      Top             =   4035
      Width           =   2010
      _Version        =   196608
      _ExtentX        =   3545
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
      ColDesigner     =   "frmRptMastCust.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   5085
      TabIndex        =   2
      Top             =   4605
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptMastCust.frx":0BED
   End
   Begin LpLib.fpCombo fpcboRptOrd 
      Height          =   375
      Left            =   5085
      TabIndex        =   0
      Top             =   3405
      Width           =   4290
      _Version        =   196608
      _ExtentX        =   7567
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
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptMastCust.frx":0F1B
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
      Left            =   9816
      TabIndex        =   4
      Top             =   7200
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
      Left            =   8136
      TabIndex        =   3
      Top             =   7200
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
            TextSave        =   "8:19 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "7/31/2018"
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Status:"
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
      Index           =   1
      Left            =   2808
      TabIndex        =   9
      Top             =   4068
      Width           =   2076
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
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
      Height          =   372
      Left            =   2568
      TabIndex        =   8
      Top             =   4632
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Report Order:"
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
      Left            =   3144
      TabIndex        =   7
      Top             =   3456
      Width           =   1740
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   1128
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Master Customer Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3768
      TabIndex        =   6
      Top             =   1368
      Width           =   4668
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2412
      Left            =   2052
      Top             =   2952
      Width           =   8100
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   1008
      Width           =   5772
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
Attribute VB_Name = "frmRptMastCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptMastCust
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboCustStatus.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboCustStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCustStatus.ListDown = True
  End If
  If fpcboCustStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboRptOrd.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptMastCust by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub


Private Sub cmdPrint_Click()
  DeActivateControls Me, True
  If fpcboRptType.ListIndex = 1 Then
    MasterCustomerListing1
    ActivateControls Me, True
  ElseIf fpcboRptType.ListIndex = 0 Then
    MasterCustomerListing2
  Else
    ActivateControls Me, True
  End If
End Sub
Private Sub fpcboRptOrd_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptOrd.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboRptOrd.ListIndex = -1
    fpcboRptOrd.Action = ActionClearSearchBuffer
  End If
  If fpcboRptOrd.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboCustStatus.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
        KeyCode = 0
      End If
    End If
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
  fpcboRptOrd.AddItem "Customer Name Order"
  fpcboRptOrd.AddItem "Account Number Order"
  fpcboRptOrd.AddItem "Location Number Order"
  fpcboRptOrd.AddItem "Service Address Order"
  fpcboRptOrd.ListIndex = 0
  fpcboCustStatus.AddItem "ALL"
  fpcboCustStatus.AddItem "Active"
  fpcboCustStatus.AddItem "Inactive"
  fpcboCustStatus.AddItem "Balance"
  fpcboCustStatus.AddItem "Pending"
  fpcboCustStatus.AddItem "Delinquent"
  fpcboCustStatus.AddItem "Final"
  fpcboCustStatus.AddItem "DELETED"
  fpcboCustStatus.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpMasterCustomer
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub MasterCustomerListing1()
  Dim UsingBook As Boolean, UsingAcct As Boolean, UsingName As Boolean
  Dim PageNo As Integer, UsingAddr As Boolean, MaxLines As Integer
  Dim Dash80 As String, IndexName As String, UBCustRecLen As Integer
  Dim UBSetupreclen As Integer, IdxRecLen As Integer, IdxFileSize As Long
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCust As Integer, UBRpt As Integer, UBSetUp1 As Integer
  Dim AcctNumber As Long, LocationNumber As String, Zip As String
  Dim CustCnt As Integer, ReportFile As String, CustStat As String
  Dim UsingStat As Boolean, CStatus As String
  Dim blnDeleted As Boolean
  blnDeleted = False
  UsingBook = False
  UsingAcct = False
  UsingName = False
  UsingStat = False
 ' DeActivateControls Me
  FrmShowPctComp.Label1 = "Creating Master Customer Listing"
  FrmShowPctComp.Show , Me
  PageNo = 0

    Select Case fpcboRptOrd.ListIndex
    Case 0
      IndexName$ = UBPath$ + "UBCUSTNM.IDX"
      UsingName = True
    Case 1
      IndexName$ = ""
      UsingAcct = True
    Case 2
      IndexName$ = UBPath$ + "UBCUSTBK.IDX"
      UsingBook = True
    Case 3
      IndexName$ = TempIndexName
      UsingAddr = True
    End Select
  Select Case fpcboCustStatus.ListIndex
    Case 0:
      UsingStat = False
      CustStat = ""
    Case 1:
      UsingStat = True
      CustStat = "A"
    Case 2:
      UsingStat = True
      CustStat = "I"
    Case 3:
      UsingStat = True
      CustStat = "B"
    Case 4:
      UsingStat = True
      CustStat = "P"
    Case 5:
      UsingStat = True
      CustStat = "D"
    Case 6:
      UsingStat = True
      CustStat = "F"
    Case 7
      IndexName$ = ""
      UsingAcct = True
      blnDeleted = True
      UsingStat = False
      CustStat = ""
      UsingAddr = False
      UsingName = False
      UsingBook = False
    Case Else
      UsingStat = False
      CustStat = ""
      
    End Select
    

  '***************
  MaxLines = 52
  PageNo = 0
  Dash80$ = String$(80, "-")

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
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
    
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptMastCust
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
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

  UBCust = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  '  UBLoca = FREEFILE
  '  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBLoca LEN = UBLocaRecLen
  ReportFile$ = UBPath$ + "UBMCLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  UBSetUp1 = FreeFile
  Open UBPath$ + "UBSETUP.DAT" For Random Access Read Write Shared As UBSetUp1 Len = UBSetupreclen
  If LOF(UBSetUp1) / UBSetupreclen = 0 Then
    TOWNNAME$ = "Undefined"
  Else
    Get UBSetUp1, 1, UBSetUp(1)
    TOWNNAME$ = QPTrim$(UBSetUp(1).UTILNAME)
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
'    TownName$ = Space$(TabStop) + TownName$
  End If
  Close UBSetUp1
'  BlockClear
 ' ShowProcessingScrn "Master Customer Report."

  GoSub DoMastCustRptHeader

  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Or UsingAddr Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
      AcctNumber = IdxBuff(cnt).RecNum
    Else
      Get UBCust, cnt, UBCustRec(1)
      AcctNumber = cnt
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitMastCustListing
    End If
    LocationNumber$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB

    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoMastCustRptHeader
    End If
    
    If blnDeleted Then
      If UBCustRec(1).DelFlag <> -1 Then
        GoTo bskipem:
      Else
        GoTo PrintEmOK
      End If
    End If
    
    If UBCustRec(1).DelFlag <> 0 Then
       GoTo bskipem:
    End If

PrintEmOK:
      If UsingStat Then           'if they care about the cust status, or want all.
        CStatus$ = Left$(QPTrim$(UBCustRec(1).Status), 1)
        If CStatus$ <> CustStat$ Then
          GoTo bskipem
        End If
      End If
      Print #UBRpt, "Acct# :"; Using("#####", AcctNumber);
      Print #UBRpt, Tab(15); RTrim$(UBCustRec(1).CustName);
      Print #UBRpt, Tab(48); "Location: "; LocationNumber$
      Print #UBRpt, "Status: "; UBCustRec(1).Status; Tab(15); QPTrim$(UBCustRec(1).ADDR1);
      Print #UBRpt, Tab(48); "H Phone#: "; QPTrim$(UBCustRec(1).HPHONE)
      Print #UBRpt, "Type  : "; QPTrim$(UBCustRec(1).CUSTTYPE); Tab(15); QPTrim$(UBCustRec(1).ADDR2);
      Print #UBRpt, Tab(48); "W Phone#: "; QPTrim$(UBCustRec(1).WPHONE)
      Print #UBRpt, "Draft : "; QPTrim$(UBCustRec(1).USEDRAFT);
      Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
'      If Len(Zip$) > 5 Then
'        Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
'      End If
      Print #UBRpt, Tab(15); QPTrim$(UBCustRec(1).CITY); ", "; QPTrim$(UBCustRec(1).STATE); " "; Zip$;
      Print #UBRpt, Tab(48); "Ser Addr: "; QPTrim$(UBCustRec(1).ServAddr)
      Print #UBRpt, Dash80$
      CustCnt = CustCnt + 1
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
    
'   End If
bskipem:
   ' ShowPctComp cnt, NumOfRecs
  
  
  Next
  
  GoSub DoMastCustRptFooter:

  Close

  Erase IdxBuff, UBCustRec

 ' If Not AbortFlag Then
 '   PrintRptFile "Master Customer Report.", "UBMCLIST.RPT", 1, RetCode, EntryPoint
 ' End If
  ViewPrint ReportFile$, "Master Customer Report"
'ActivateControls Me
ExitMastCustListing:

  Exit Sub

DoMastCustRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(24); "Master Customer Listing Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, Dash80$
  LineCnt = 4
  Return

DoMastCustRptFooter:
  Print #UBRpt, "Totals:"; Tab(10); "Customers: "; Using("#####,#", CustCnt)
  Return


End Sub
Private Sub MasterCustomerListing2()
  Dim UsingBook As Boolean, UsingAcct As Boolean, UsingName As Boolean
  Dim PageNo As Integer, UsingAddr As Boolean, MaxLines As Integer
  Dim Dash80 As String, IndexName As String, UBCustRecLen As Integer
  Dim UBSetupreclen As Integer, IdxRecLen As Integer, IdxFileSize As Long
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCust As Integer, UBRpt As Integer, UBSetUp1 As Integer
  Dim AcctNumber As Long, LocationNumber As String, Zip As String
  Dim CustCnt As Integer, ToPrint As String, ReportFile As String
  Dim UsingStat As Boolean, CStatus As String, CustStat As String
  Dim blnDeleted As Boolean
  blnDeleted = False
  UsingStat = False
  UsingBook = False
  UsingAcct = False
  UsingName = False
  'DeActivateControls Me
  FrmShowPctComp.Label1 = "Creating Master Customer Listing"
  FrmShowPctComp.Show , Me
  PageNo = 0

    Select Case fpcboRptOrd.ListIndex
    Case 0
      IndexName$ = UBPath$ + "UBCUSTNM.IDX"
      UsingName = True
    Case 1
      IndexName$ = ""
      UsingAcct = True
    Case 2
      IndexName$ = UBPath$ + "UBCUSTBK.IDX"
      UsingBook = True
    Case 3
      IndexName$ = TempIndexName
      UsingAddr = True
    End Select
  Select Case fpcboCustStatus.ListIndex
    Case 0:
      UsingStat = False
      CustStat = ""
    Case 1:
      UsingStat = True
      CustStat = "A"
    Case 2:
      UsingStat = True
      CustStat = "I"
    Case 3:
      UsingStat = True
      CustStat = "B"
    Case 4:
      UsingStat = True
      CustStat = "P"
    Case 5:
      UsingStat = True
      CustStat = "D"
    Case 6:
      UsingStat = True
      CustStat = "F"
    
    Case 7
      IndexName$ = ""
      UsingAcct = True
      blnDeleted = True
      UsingStat = False
      CustStat = ""
      UsingAddr = False
      UsingName = False
      UsingBook = False

    Case Else
      UsingStat = False
      CustStat = ""
    End Select

  '***************
  MaxLines = 52
  PageNo = 0
  Dash80$ = String$(80, "-")

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
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
    
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptMastCust
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
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

  UBCust = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  '  UBLoca = FREEFILE
  '  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBLoca LEN = UBLocaRecLen
  ReportFile$ = UBPath$ + "UBMCLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  UBSetUp1 = FreeFile
  Open UBPath$ + "UBSETUP.DAT" For Random Access Read Write Shared As UBSetUp1 Len = UBSetupreclen
  If LOF(UBSetUp1) / UBSetupreclen = 0 Then
    TOWNNAME$ = "Undefined"
  Else
    Get UBSetUp1, 1, UBSetUp(1)
    TOWNNAME$ = QPTrim$(UBSetUp(1).UTILNAME)
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
'    TownName$ = Space$(TabStop) + TownName$
  End If
  Close UBSetUp1
'  BlockClear
 ' ShowProcessingScrn "Master Customer Report."

  'GoSub DoMastCustRptHeader

  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Or UsingAddr Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
      AcctNumber = IdxBuff(cnt).RecNum
    Else
      Get UBCust, cnt, UBCustRec(1)
      AcctNumber = cnt
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitMastCustListing
    End If
    LocationNumber$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB

    ToPrint$ = ""
'-----------------------
    If blnDeleted Then
      If UBCustRec(1).DelFlag <> -1 Then
        GoTo bskipem:
      Else
        GoTo PrintEmOK2
      End If
    End If
    
    If UBCustRec(1).DelFlag <> 0 Then
       GoTo bskipem:
    End If

PrintEmOK2:

'    If UBCustRec(1).DelFlag <> -1 Then
      If UsingStat Then           'if they care about the cust status, or want all.
        CStatus$ = Left$(QPTrim$(UBCustRec(1).Status), 1)
        If CStatus$ <> CustStat$ Then
          GoTo bskipem
        End If
      End If

      ToPrint$ = Str$(AcctNumber) + "~"
      ToPrint$ = ToPrint$ + RTrim$(UBCustRec(1).CustName)
      ToPrint$ = ToPrint$ + "~" + LocationNumber$
      ToPrint$ = ToPrint$ + "~" + UBCustRec(1).Status + "~" + QPTrim$(UBCustRec(1).ADDR1)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).HPHONE)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).CUSTTYPE) + "~" + QPTrim$(UBCustRec(1).ADDR2)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).WPHONE)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).USEDRAFT)
      Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
'      If Len(Zip$) > 5 Then
'        Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
'      End If
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).CITY) + "~" + QPTrim$(UBCustRec(1).STATE) + "~" + Zip$
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ServAddr)
      'Print #UBRpt, Dash80$
      'LineCnt = LineCnt + 5
      CustCnt = CustCnt + 1
      Print #UBRpt, ToPrint$
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
    'End If
   ' ShowPctComp cnt, NumOfRecs
bskipem:
  Next

  'GoSub DoMastCustRptFooter:

  Close

  Erase IdxBuff, UBCustRec

 ' If Not AbortFlag Then
 '   PrintRptFile "Master Customer Report.", "UBMCLIST.RPT", 1, RetCode, EntryPoint
 ' End If
'  ViewPrint "UBMCLIST.RPT", "Master Customer Report"
  Load frmLoadingRpt
  frmLoadingRpt.setwherefrom frmRptMastCust
  ARptMastCustList.txtDate = Now
  ARptMastCustList.txtTown = TOWNNAME$
  ARptMastCustList.Title = "Master Customer Listing Report"
  ARptMastCustList.totCust = CustCnt
  ARptMastCustList.GetName ReportFile$
  ARptMastCustList.startrpt
' ActivateControls Me
ExitMastCustListing:

  Exit Sub

End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
