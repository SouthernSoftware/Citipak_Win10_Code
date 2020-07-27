VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnVendLabl 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Labels"
   ClientHeight    =   8610
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnVendLabl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboVend1 
      Height          =   405
      Left            =   5130
      TabIndex        =   0
      Top             =   2850
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
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
      ScrollBarH      =   3
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnVendLabl.frx":08CA
   End
   Begin LpLib.fpCombo fpcboVend2 
      Height          =   405
      Left            =   5130
      TabIndex        =   1
      Top             =   3480
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
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
      ScrollBarH      =   3
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnVendLabl.frx":0C7D
   End
   Begin VB.CommandButton cmd3Column 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Three Column"
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
      Left            =   4776
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   1092
   End
   Begin VB.CommandButton cmdSingle 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Single Column"
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
      Left            =   6528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1092
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Include Inactives:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   3096
      TabIndex        =   2
      Top             =   4056
      Width           =   2244
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
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
      Left            =   10032
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8256
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
            TextSave        =   "2:11 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "7/7/2010"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Do You Wish To Print 3 Column Sheet or Single Column Labels ? "
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
      Height          =   828
      Left            =   3936
      TabIndex        =   10
      Top             =   4752
      Width           =   4788
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Vendor:"
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
      Height          =   420
      Left            =   2976
      TabIndex        =   7
      Top             =   2928
      Width           =   2004
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4428
      Left            =   1920
      Top             =   2352
      Width           =   8316
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Vendor Labels"
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
      Left            =   3984
      TabIndex        =   6
      Top             =   1176
      Width           =   4332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   936
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2490
      Picture         =   "frmPrnVendLabl.frx":1030
      Top             =   2730
      Width           =   360
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Vendor:"
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
      Height          =   420
      Left            =   3168
      TabIndex        =   5
      Top             =   3540
      Width           =   1812
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   816
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
Attribute VB_Name = "frmPrnVendLabl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType

Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Sub cmdExit_Click()
  frmAPVendMaintMenu.Show
  Unload Me
End Sub

Private Sub cmdsingle_Click()
  'Unload frmAPVendLablOPt
  PrintVendorLabels
End Sub
Private Sub cmd3Column_Click()
 ' Unload frmAPVendLablOPt
  PrnVendLabelsLaser
  
End Sub


Private Sub fpcboVend1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVend1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVend1.ListIndex = -1
    fpcboVend1.Action = ActionClearSearchBuffer
  End If
  If fpcboVend1.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcboVend1_LostFocus()
  fpcboVend1.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboVend2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVend2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVend2.ListIndex = -1
    fpcboVend2.Action = ActionClearSearchBuffer
  End If
  If fpcboVend2.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcboVend2_LostFocus()
  fpcboVend2.Action = ActionClearSearchBuffer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpVendLab
  VendCodeNameIA fpcboVend1
  VendCodeNameIA fpcboVend2
  fpcboVend1.ListIndex = 0
  fpcboVend2.ListIndex = fpcboVend2.ListCount - 1
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Public Sub PrintVendorLabels()
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim PRNFile As Integer, cnt As Integer, HowMany As Integer
  Dim ReportFile As String, ToPrint As String, Header As String
  Dim MaskFile As Integer, MaskReportFile As String, VCode As String
  Dim ToPrintM As String, HeaderM As String, Vend1St As String
  Dim Vactive As String, doone As Boolean, VendLst As String
  FrmShowPctComp.Label1 = "Creating Vendor Labels"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnVendLabl
  fpcboVend1.col = 0
  fpcboVend2.col = 0
  Vend1St$ = QPTrim$(fpcboVend1.ColText)
  VendLst$ = QPTrim$(fpcboVend2.ColText)

  ToPrint$ = Space$(30)
  'PrintHelp "Print Vendor Labels"

  PRNFile = FreeFile
  ReportFile$ = "vndrlbl.PRN"
  Open ReportFile$ For Output As #PRNFile

  'PrintHelp "Processing report. Please wait."

  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs

  '--fix for del not to show
  For cnt = 1 To NumActiveVendors
    FrmShowPctComp.ShowPctComp cnt, NumActiveVendors
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnVendLabl
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get VendorIdxFile, cnt, VendorIdx
    VCode$ = QPTrim$(VendorIdx.VendorCode)
    If VCode$ >= Vend1St$ And VCode$ <= VendLst$ Then
      Get VendorFile, VendorIdx.RecNum, Vendor
      If Vendor.DelFlag = 0 Then
        If Check1.Value = 1 Then
          doone = True
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
          Else
            Vactive$ = "Inactive"
          End If
        End If
        If Check1.Value = 0 Then
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
            doone = True
          Else
            doone = False
          End If
        End If
        If doone = True Then

    HowMany = HowMany + 1

    LSet ToPrint$ = Vendor.vnum
    Print #PRNFile, ToPrint$

    LSet ToPrint$ = Vendor.VNAME
    Print #PRNFile, ToPrint$

    LSet ToPrint$ = Vendor.Addr1
    Print #PRNFile, ToPrint$

    LSet ToPrint$ = Vendor.Addr2
    Print #PRNFile, ToPrint$

    LSet ToPrint$ = QPTrim$(Vendor.City) + " " + Vendor.STATE + " " + Vendor.Zip
    Print #PRNFile, ToPrint$
    LSet ToPrint$ = ""
    Print #PRNFile, ToPrint$
  End If
  End If
  End If
  Next

  Close
  GoSub PrintLblMask

  Header$ = "Vendor Labels"
  ActivateControls frmPrnVendLabl
  ViewPrint ReportFile$, Header$, False, , True, MaskReportFile$
  'KILL ReportFile$
Exit Sub

PrintLblMask:

  MaskFile = FreeFile
  MaskReportFile$ = "LBLMASK.PRN"
  Open MaskReportFile$ For Output As #MaskFile
  ToPrintM$ = Space$(30)

  LSet ToPrintM$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  Print #MaskFile, ToPrintM$

  LSet ToPrintM$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  Print #MaskFile, ToPrintM$

  LSet ToPrintM$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  Print #MaskFile, ToPrintM$

  LSet ToPrintM$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  Print #MaskFile, ToPrintM$

  LSet ToPrintM$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  Print #MaskFile, ToPrintM$

  LSet ToPrintM$ = ""
  Print #MaskFile, ToPrintM$

  Close MaskFile

  HeaderM$ = "Label Alignent Test"
 ' PrintRptFile Header$, MaskReportFile$, LPTNo, RetCode%, EntryPoint

Return
CancelExit:
  Exit Sub
End Sub
Public Sub PrnVendLabelsLaser()
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim PRNFile As Integer, cnt As Integer, HowMany As Integer
  Dim ReportFile As String, ToPrint As String, Header As String
  Dim MaskFile As Integer, MaskReportFile As String, VCode As String
  Dim ToPrintM As String, HeaderM As String, Vend1St As String
  Dim ToPrint1 As String, ToPrint2 As String, ToPrint3 As String
  Dim Vactive As String, doone As Boolean, VendLst As String, lblnum As Integer
  FrmShowPctComp.Label1 = "Creating Vendor Labels"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnVendLabl
  fpcboVend1.col = 0
  fpcboVend2.col = 0
  Vend1St$ = QPTrim$(fpcboVend1.ColText)
  VendLst$ = QPTrim$(fpcboVend2.ColText)
  lblnum = 0
  ToPrint$ = Space$(30)
  'PrintHelp "Print Vendor Labels"

  PRNFile = FreeFile
  ReportFile$ = "vndrlbl.PRN"
  Open ReportFile$ For Output As #PRNFile

  'PrintHelp "Processing report. Please wait."

  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs

  '--fix for del not to show
  For cnt = 1 To NumActiveVendors
    FrmShowPctComp.ShowPctComp cnt, NumActiveVendors
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnVendLabl
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    Get VendorIdxFile, cnt, VendorIdx
    VCode$ = QPTrim$(VendorIdx.VendorCode)
    If VCode$ >= Vend1St$ And VCode$ <= VendLst$ Then
      Get VendorFile, VendorIdx.RecNum, Vendor
      If Vendor.DelFlag = 0 Then
        If Check1.Value = 1 Then
          doone = True
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
          Else
            Vactive$ = "Inactive"
          End If
        End If
        If Check1.Value = 0 Then
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
            doone = True
          Else
            doone = False
          End If
        End If
        If doone = True Then
    HowMany = HowMany + 1
    ToPrint1$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME) + "~"
    ToPrint1$ = ToPrint1$ + QPTrim(Vendor.Addr1) + "~" + QPTrim(Vendor.Addr2) + "~"
    ToPrint1$ = ToPrint1$ + QPTrim$(Vendor.City) + " " + Vendor.STATE + " " + Vendor.Zip
    lblnum = lblnum + 1
'    cnt = cnt + 1
'    End If
'  If Not cnt > NumActiveVendors Then
'    Get VendorIdxFile, cnt, VendorIdx
'    VCode$ = QPTrim$(VendorIdx.VendorCode)
'    If VCode$ >= Vend1St$ And VCode$ <= VendLst$ Then
'      Get VendorFile, VendorIdx.RecNum, Vendor
'      If Vendor.DelFlag = 0 Then
'        If Check1.Value = 1 Then
'          doone = True
'          If Vendor.ActiveFlag = 0 Then
'            Vactive$ = "Active"
'          Else
'            Vactive$ = "Inactive"
'          End If
'        End If
'        If Check1.Value = 0 Then
'          If Vendor.ActiveFlag = 0 Then
'            Vactive$ = "Active"
'            doone = True
'          Else
'            doone = False
'          End If
'        End If
'        If doone = True Then
'
'      HowMany = HowMany + 1
'      ToPrint2$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME) + "~"
'      ToPrint2$ = ToPrint2$ + QPTrim(Vendor.Addr1) + "~" + QPTrim(Vendor.Addr2) + "~"
'      ToPrint2$ = ToPrint2$ + QPTrim$(Vendor.City) + " " + Vendor.STATE + " " + Vendor.Zip
'    Else
'      ToPrint2$ = "~ ~ ~ ~ ~ ~"
'    End If
'    cnt = cnt + 1
'    End If
'    End If
'    If Not cnt > NumActiveVendors Then
'      Get VendorIdxFile, cnt, VendorIdx
'    VCode$ = QPTrim$(VendorIdx.VendorCode)
'    If VCode$ >= Vend1St$ And VCode$ <= VendLst$ Then
'      Get VendorFile, VendorIdx.RecNum, Vendor
'      If Vendor.DelFlag = 0 Then
'        If Check1.Value = 1 Then
'          doone = True
'          If Vendor.ActiveFlag = 0 Then
'            Vactive$ = "Active"
'          Else
'            Vactive$ = "Inactive"
'          End If
'        End If
'        If Check1.Value = 0 Then
'          If Vendor.ActiveFlag = 0 Then
'            Vactive$ = "Active"
'            doone = True
'          Else
'            doone = False
'          End If
'        End If
'        If doone = True Then
'          HowMany = HowMany + 1
'          ToPrint3$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME) + "~"
'          ToPrint3$ = ToPrint3$ + QPTrim(Vendor.Addr1) + "~" + QPTrim(Vendor.Addr2) + "~"
'          ToPrint3$ = ToPrint3$ + QPTrim$(Vendor.City) + " " + Vendor.STATE + " " + Vendor.Zip
'        Else
'          ToPrint3$ = "~ ~ ~ ~ ~ ~"
'        End If
'        cnt = cnt + 1
'      End If
'    End If
    If lblnum = 1 Then
      ToPrint3$ = ToPrint1
    ElseIf lblnum = 2 Then
      ToPrint3$ = ToPrint3$ + "~" + ToPrint1$
    ElseIf lblnum = 3 Then
      ToPrint$ = ToPrint3$ + "~" + ToPrint1$
      ToPrint3$ = ""
      Print #PRNFile, ToPrint$
      lblnum = 0
    End If
    ToPrint$ = ""
    ToPrint1$ = ""
   
    End If
    End If
    End If
  Next

  Close
  FrmShowPctComp.ShowPctComp 1, 1
  Load frmLoadingRpt
  ActivateControls frmPrnVendLabl
  ARptAPVendLabLas.GetName ReportFile$
  ARptAPVendLabLas.startrpt

Exit Sub


Return
CancelExit:
  Exit Sub
End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
