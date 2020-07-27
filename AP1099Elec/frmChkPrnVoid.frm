VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChkPrnCancel 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancel Printed A/P Check"
   ClientHeight    =   8844
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmChkPrnVoid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8844
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList lstChecks 
      Height          =   2688
      Left            =   2064
      TabIndex        =   8
      Top             =   3120
      Width           =   8100
      _Version        =   196608
      _ExtentX        =   14287
      _ExtentY        =   4741
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmChkPrnVoid.frx":08CA
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Alt-C &Clear All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   8772
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6264
      Width           =   1500
   End
   Begin VB.CommandButton cmdMark 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Alt-M &Mark All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6876
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6264
      Width           =   1500
   End
   Begin VB.CommandButton cmdReverse 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F7 &Reinstate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   5328
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7176
      Width           =   1524
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 Cance&l"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   7260
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7176
      Width           =   1524
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
      Height          =   468
      Left            =   9120
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7176
      Width           =   1524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8484
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "9:50 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "11/18/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel Printed A/P Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3684
      TabIndex        =   7
      Top             =   1056
      Width           =   4836
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   816
      Width           =   7020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4356
      Left            =   1662
      Top             =   2520
      Width           =   8868
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2940
      Left            =   1932
      Top             =   3000
      Width           =   8340
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   3
      Left            =   9120
      TabIndex        =   6
      Top             =   2784
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   1
      Left            =   5640
      TabIndex        =   5
      Top             =   2784
      Width           =   828
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   7
      Left            =   2328
      TabIndex        =   4
      Top             =   2784
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Select Check, Click or use SpaceBar to Highlight."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   636
      Left            =   1740
      TabIndex        =   3
      Top             =   6192
      Width           =   5028
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   696
      Width           =   7020
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
Attribute VB_Name = "frmChkPrnCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Dim Vendor As VendorRecType
Dim APChkInfo As CheckInfoType3
Dim CHKinfo() As CheckInfoType3
Dim TempRec As Long
Dim APAcct As String
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Call MainLog("Close via APChkPrnCancel.")
        KillFile ("APCHK.opn")
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmAPChkProcessMenu.Show
  Unload frmChkPrnCancel
End Sub

Private Sub cmdReverse_Click()
  If lstChecks.SelCount <> 0 Then
    If MsgBox("Are You Sure You Wish To Continue With Reinstatement?", vbYesNo, "Reinstate Check") = vbNo Then
      Exit Sub
    Else
      UnCancelPrintedCheck
      cmdExit_Click
    End If
  Else
    MsgBox "You Must Make A Selection First", vbOKOnly, "No Selection"
    lstChecks.SetFocus
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
      SendKeys "%V"
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%R"
      KeyCode = 0
    Case Else:
  End Select

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdClear_Click()
  lstChecks.Action = ActionDeselectAll
End Sub

Private Sub cmdMark_Click()
  lstChecks.Action = ActionSelectAll
End Sub

Private Sub cmdCancel_Click()
  If lstChecks.SelCount <> 0 Then
  
    If MsgBox("Are You Sure You Wish To Continue With Cancel?", vbYesNo, "Cancel Check") = vbNo Then
      Exit Sub
    Else
      CancelPrintedCheck
      cmdExit_Click
    End If
  Else
    MsgBox "You Must Make A Selection First", vbOKOnly, "No Selection"
    lstChecks.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpCancelChks
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetAPAcct APAcct
  StatusBar1.Panels.Item(1).Text = GLUserName
  ListEntries
End Sub


Private Sub ListEntries()
'this fills the listbox with iNVOICE Entries for User to Select
  Dim Dash As String, FF As String, ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer
  Dim VendorFile As Integer, NumVRecs As Long, PrintFile As Integer
  Dim Cnt2 As Long, TCheckAmt As Double, Title As String, Temp As Integer
  Dim low As Long, High As Long, tempstr As String, TempAmt As String
  Dim tempstrchk As String, Void As String
  ChkinfoFile = FreeFile
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  ChkInfoRecLen = Len(CHKinfo(1))
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  ReDim CHKinfo(1 To VCnt) As CheckInfoType3
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For Temp = 1 To VCnt
    Get ChkinfoFile, Temp, CHKinfo(Temp)
  Next
  OpenVendorFile VendorFile, NumVRecs
  lstChecks.SortState = SortStateSuspend
  For cnt = 1 To VCnt
    tempstr = Space$(96)
    TempAmt$ = Space$(20)
    Void$ = Space$(15)
    Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
    For Cnt2& = CHKinfo(cnt).StartChk To CHKinfo(cnt).LastChk
      tempstrchk = Using("#######", Str$(Cnt2&))
      If CHKinfo(cnt).VoidFlag Then
        Mid$(tempstr, 18, 50) = QPTrim(Vendor.VNAME)
        TempAmt$ = "  CANCELED BY USER"
        Void$ = "X"
      ElseIf Cnt2& < CHKinfo(cnt).LastChk Then
        Mid$(tempstr, 18, 50) = QPTrim(Vendor.VNAME)
        TempAmt$ = "              VOID"
      Else
        TCheckAmt# = Round(TCheckAmt# + CHKinfo(cnt).ChkAmt)
        Mid$(tempstr, 18, 50) = QPTrim(Vendor.VNAME)
        TempAmt$ = Using("$##,###,###.##", Str$(CHKinfo(cnt).ChkAmt))
      End If
      lstChecks.AddItem tempstrchk & Chr$(9) & tempstr & Chr$(9) & TempAmt$ & Chr$(9) & cnt & Chr$(9) & Void$
      lstChecks.ListApplyTo = ListApplyToIndividual
      lstChecks.col = 2
      lstChecks.AlignH = AlignHRight
    Next
  Next
  lstChecks.SortState = 1
  Close VendorFile
End Sub

Private Sub CancelPrintedCheck()
  Dim ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer
  Dim Pcnt As Integer, nocnt As Integer
  ChkinfoFile = FreeFile
  ChkInfoRecLen = Len(APChkInfo)
  nocnt = 0
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
    For Pcnt = 0 To lstChecks.ListCount - 1
      If lstChecks.Selected(Pcnt) Then
        lstChecks.col = 4
        lstChecks.Row = Pcnt
        If QPTrim(lstChecks.ColList) <> "X" Then
          lstChecks.col = 3
          lstChecks.Row = Pcnt
          cnt = QPTrim(lstChecks.ColList)
          Get ChkinfoFile, cnt, APChkInfo
          APChkInfo.VoidFlag = True
          Put ChkinfoFile, cnt, APChkInfo
        Else
          nocnt = nocnt + 1
        End If
      End If
    Next
  Close
  Call MainLog("APChk Cancel Printed Chk.")
  If nocnt > 0 Then
  If nocnt = lstChecks.SelCount Then
    MsgBox "Selected Check(s) Already Canceled.", vbOKOnly, "Canceled Check"
  Else
    MsgBox "One or More of the Checks Selected Were Already Canceled.", vbOKOnly, "Canceled Checks"
  End If
  End If
End Sub
Private Sub UnCancelPrintedCheck()
  Dim ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer
  Dim Pcnt As Integer, nocnt As Integer
  nocnt = 0
  ChkinfoFile = FreeFile
  ChkInfoRecLen = Len(APChkInfo)
  VCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
    For Pcnt = 0 To lstChecks.ListCount - 1
      If lstChecks.Selected(Pcnt) Then
        lstChecks.col = 4
        lstChecks.Row = Pcnt
        If QPTrim(lstChecks.ColList) = "X" Then
          lstChecks.col = 3
          lstChecks.Row = Pcnt
          cnt = QPTrim(lstChecks.ColList)
          Get ChkinfoFile, cnt, APChkInfo
          APChkInfo.VoidFlag = False
          Put ChkinfoFile, cnt, APChkInfo
        Else
          nocnt = nocnt + 1
        End If
      End If
    Next
  Close
  Call MainLog("APChk Undo Cancel Printed Chk.")
  If nocnt > 0 Then
  If nocnt = lstChecks.SelCount Then
    MsgBox "Selected Check(s) Active.", vbOKOnly, "Active Check"
  Else
    MsgBox "One or More of the Checks Selected Were Already Active.", vbOKOnly, "Active Check"
  End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

