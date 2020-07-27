VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnVendFile 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor File Report"
   ClientHeight    =   8610
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnVendFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   6150
      TabIndex        =   0
      Top             =   3030
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ColDesigner     =   "frmPrnVendFile.frx":08CA
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
      Left            =   4104
      TabIndex        =   1
      Top             =   3624
      Width           =   2244
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Go"
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
      Left            =   8256
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   1332
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
            TextSave        =   "10:40 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "12/1/2006"
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
      Left            =   3696
      TabIndex        =   6
      Top             =   3048
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2100
      Left            =   1920
      Top             =   2352
      Width           =   8316
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Vendor File Report"
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
      TabIndex        =   5
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
      Picture         =   "frmPrnVendFile.frx":0C30
      Top             =   2730
      Width           =   360
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
Attribute VB_Name = "frmPrnVendFile"
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
  Unload frmPrnVendFile
End Sub

Private Sub cmdGo_Click()
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  Else
    rptopt = 2
  End If
  If rptopt = 1 Then
    PrintVendorFile
  ElseIf rptopt = 2 Then
    PrintVendorFile2
  End If
 
End Sub

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      Check1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdGo.SetFocus
        KeyCode = 0
      End If
    End If
  End If
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
    Case vbKeyF10:
      SendKeys "%G"
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
  Me.HelpContextID = hlpVendFile
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PrintVendorFile()
  Dim MaxLines As Integer, VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, HowMany As Integer, ToPrint5 As String
  Dim ReportFile As String, ToPrint As String, ToPrint2 As String
  Dim ToPrint4 As String, Header As String, ToPrint3 As String
  Dim A1Len As Integer, a2Len As Integer, PA1Len As Integer, PA2Len As Integer
  Dim AcctIdx As GLAcctIndexType, Memo As String
  Dim GLAcct As GLAcctRecType
  Dim FundCode As String, Vactive As String, doone As Boolean
  Dim ChkFund As Boolean
  FrmShowPctComp.Label1 = "Creating Vendor File Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnVendFile

  'FF$ = Chr$(12)
  'MaxLines = 60
  'Linecnt = 0
  'Page = 0
  Header$ = "Vendor File"

  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs
  PRNFile = FreeFile
  ReportFile$ = "vndrfile.PRN"
  Open ReportFile$ For Output As #PRNFile

  For cnt = 1 To NumActiveVendors
    FrmShowPctComp.ShowPctComp cnt, NumActiveVendors
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnVendFile
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
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
    'ToPrint$ = Space$(80)
    'IF Vendor.DELFLAG <> 0 THEN STOP

'    If Linecnt > MaxLines - 8 Then
'      Print #PRNFile, FF$
'      GoSub PrintVRptHeader
'    End If
'
    '--Line 1
    ToPrint$ = Space$(80)
    ToPrint$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME) + "~"
    ToPrint$ = ToPrint$ + QPTrim(Vendor.PaytoName) + "~"
    

    '--Line 2
    A1Len = Len(QPTrim$(Vendor.Addr1))
    PA1Len = Len(QPTrim$(Vendor.PaytoAddr))
    If A1Len > 0 Or PA1Len > 0 Then
      ToPrint2$ = Space$(80)
      ToPrint2$ = QPTrim(Vendor.Addr1) + "~" + QPTrim(Vendor.PaytoAddr) + "~"
    Else
      ToPrint2$ = "~~"
    End If
    '--Line 3
    a2Len = Len(QPTrim$(Vendor.Addr2))
    PA2Len = Len(QPTrim$(Vendor.PaytoAddr2))
    If a2Len > 0 Or PA2Len > 0 Then
      ToPrint3$ = Space$(80)
      ToPrint3$ = QPTrim(Vendor.Addr2) + "~" + QPTrim(Vendor.PaytoAddr2) + "~"
    Else
      ToPrint3$ = "~~"
    End If

    '--Line 4
    ToPrint4$ = Space$(80)
    ToPrint4$ = QPTrim$(Vendor.City) + " " + QPTrim(Vendor.State) + " " + QPTrim(Vendor.Zip)
    ToPrint4$ = ToPrint4$ + "~" + QPTrim$(Vendor.PayToCity) + " " + QPTrim(Vendor.PaytoState) + " " + QPTrim(Vendor.PaytoZip) + "~"
    

    '--Line 5
    ToPrint5$ = Space$(80)
    ToPrint5$ = QPTrim(Vendor.Fedid) + "~" + QPTrim$(Str$(Vendor.VTerms)) + "~"
    ToPrint5$ = ToPrint5$ + QPTrim(Vendor.StCode) + "~" + QPTrim(Vendor.CoCode) + "~"
    ToPrint5$ = ToPrint5$ + QPTrim(Vendor.Get1099) + "~"
    'IF Vendor.Get1099 = "Y" THEN STOP

    '--Line 6
    If Len(QPTrim$(Vendor.Memo)) > 0 Then
      Memo$ = "      MEMO: " + QPTrim$(Vendor.Memo)
    Else
      Memo$ = ""
    End If
    ToPrint5$ = ToPrint5$ + QPTrim(Vendor.Contact) + "~" + QPTrim(Vendor.Phone) + "~"
    ToPrint5$ = ToPrint5$ + QPTrim(Vendor.Fax) + "~" + QPTrim(Vendor.DBA) + Memo$ + "~" + Vactive$
    Print #PRNFile, ToPrint$ + ToPrint2$ + ToPrint3$ + ToPrint4$ + ToPrint5$
  End If
  Next
  
  Close
  Load frmLoadingRpt
  ActivateControls frmPrnVendFile
  ARptVendFile.totvends = HowMany
  ARptVendFile.txtTown = GLUserName$
  ARptVendFile.Label1.Caption = Header$
  ARptVendFile.txtDate.Caption = Now
  ARptVendFile.GetName ReportFile$
  ARptVendFile.startrpt

 ' ViewPrint ReportFile$, Header$, True
 'KillFile ReportFile$
  
Exit Sub

CancelExit:
  Exit Sub
End Sub

Private Sub PrintVendorFile2()
  Dim MaxLines As Integer, VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, HowMany As Integer
  Dim ReportFile As String, ToPrint As String, Page As Integer
  Dim FF As String, Header As String, DivLine As String
  Dim A1Len As Integer, a2Len As Integer, PA1Len As Integer, PA2Len As Integer
  Dim AcctIdx As GLAcctIndexType
  Dim GLAcct As GLAcctRecType
  Dim FundCode As String, Vactive As String, doone As Boolean
  Dim ChkFund As Boolean
  FrmShowPctComp.Label1 = "Creating Vendor File Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnVendFile

  FF$ = Chr$(12)
  MaxLines = 60
  Linecnt = 0
  Page = 0
  Header$ = "Vendor File"

  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs
  PRNFile = FreeFile
  ReportFile$ = "vndrfile.PRN"
  Open ReportFile$ For Output As #PRNFile
  ToPrint$ = Space$(80)
  DivLine$ = String$(80, "-")

  GoSub PrintVRptHeader

  For cnt = 1 To NumActiveVendors
    FrmShowPctComp.ShowPctComp cnt, NumActiveVendors
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnVendFile
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
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
    ToPrint$ = Space$(80)
    'IF Vendor.DELFLAG <> 0 THEN STOP

    If Linecnt > MaxLines - 8 Then
      Print #PRNFile, FF$
      GoSub PrintVRptHeader
    End If

    '--Line 1
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 2) = QPTrim(Vendor.vnum)
    Mid$(ToPrint$, 14) = QPTrim(Vendor.VNAME)
   ''' Mid$(ToPrint$, 46) = Str(VendorIdx.RecNum)
    Mid$(ToPrint$, 46) = QPTrim(Vendor.PaytoName)
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1

    '--Line 2
    A1Len = Len(QPTrim$(Vendor.Addr1))
    PA1Len = Len(QPTrim$(Vendor.PaytoAddr))
    If A1Len > 0 Or PA1Len > 0 Then
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 4) = Vactive$
      Mid$(ToPrint$, 14) = QPTrim(Vendor.Addr1)
      Mid$(ToPrint$, 46) = QPTrim(Vendor.PaytoAddr)
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1
    End If
    '--Line 3
    a2Len = Len(QPTrim$(Vendor.Addr2))
    PA2Len = Len(QPTrim$(Vendor.PaytoAddr2))
    If a2Len > 0 Or PA2Len > 0 Then
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 14) = QPTrim(Vendor.Addr2)
      Mid$(ToPrint$, 46) = QPTrim(Vendor.PaytoAddr2)
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1
    End If

    '--Line 4
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 14) = QPTrim$(Vendor.City) + " " + QPTrim(Vendor.State) + " " + QPTrim(Vendor.Zip)
    Mid$(ToPrint$, 46) = QPTrim$(Vendor.PayToCity) + " " + QPTrim(Vendor.PaytoState) + " " + QPTrim(Vendor.PaytoZip)
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1

    '--Line 5
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 14) = "ID: " + QPTrim(Vendor.Fedid)
    Mid$(ToPrint$, 30) = "Terms: " + QPTrim$(Str$(Vendor.VTerms))
    Mid$(ToPrint$, 41) = "State: " + QPTrim(Vendor.StCode)
    Mid$(ToPrint$, 53) = "County: " + QPTrim(Vendor.CoCode)
    Mid$(ToPrint$, 68) = "1099? " + QPTrim(Vendor.Get1099)
    'IF Vendor.Get1099 = "Y" THEN STOP
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1

    '--Line 6
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 14) = "Contact: " + QPTrim(Vendor.Contact)
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1

    '--Line 7
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 14) = "Phone: " + QPTrim(Vendor.Phone)
    Mid$(ToPrint$, 46) = "Fax: " + QPTrim(Vendor.Fax)
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1

    '--Line 8 and 9
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 12) = "  DBA: " + QPTrim(Vendor.DBA)
    Mid$(ToPrint$, 45) = " MEMO: " + QPTrim(Vendor.Memo)
    Print #PRNFile, ToPrint$
    Print #PRNFile, DivLine$
    Linecnt = Linecnt + 2
  End If
  Next

  Print #PRNFile,
  Print #PRNFile, HowMany; "Vendors listed."
  Print #PRNFile, FF$

  Close

  ActivateControls frmPrnVendFile
  ViewPrint ReportFile$, Header$, True
  KillFile ReportFile$
  
Exit Sub

PrintVRptHeader:
Page = Page + 1
  Print #PRNFile, GLUserName$
  Print #PRNFile, Header$
  Print #PRNFile,
  Print #PRNFile, "Report Date: "; Date$; Tab(67); "Page #"; Page
  Print #PRNFile, "Vendor Code  Vendor Address                   Check Address          Default Dist"
  Print #PRNFile, String$(80, "-")

  Linecnt = 5
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
