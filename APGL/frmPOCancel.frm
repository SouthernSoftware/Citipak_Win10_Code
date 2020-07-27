VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOCancel 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancel/Clear Open Purchase Order"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPOCancel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboVendCode 
      Height          =   405
      Left            =   5115
      TabIndex        =   0
      Top             =   3285
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   1
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
      ColDesigner     =   "frmPOCancel.frx":08CA
   End
   Begin LpLib.fpList fplstPOs 
      Height          =   735
      Left            =   5100
      TabIndex        =   1
      Top             =   4875
      Width           =   4215
      _Version        =   196608
      _ExtentX        =   7435
      _ExtentY        =   1296
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
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
      Columns         =   0
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
      ColDesigner     =   "frmPOCancel.frx":0C89
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
      Height          =   468
      Left            =   6984
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7344
      Width           =   1236
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
      Left            =   8856
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1236
   End
   Begin EditLib.fpText fptxtVendName 
      Height          =   396
      Left            =   5112
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Width           =   3780
      _Version        =   196608
      _ExtentX        =   6667
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8484
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
            TextSave        =   "1:24 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/14/2018"
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
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1092
      Left            =   4944
      Top             =   4728
      Width           =   4524
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order:"
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
      Left            =   2712
      TabIndex        =   8
      Top             =   4776
      Width           =   2028
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel or Clear Open Purchase Order "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3306
      TabIndex        =   7
      Top             =   1416
      Width           =   5580
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   1176
      Width           =   7020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3276
      Left            =   2568
      Top             =   2880
      Width           =   7212
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Code:"
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
      Left            =   3024
      TabIndex        =   5
      Top             =   3360
      Width           =   1716
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   1056
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
Attribute VB_Name = "frmPOCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim EncAcct As String
Dim POControl As POControlRecType
Dim POEdit As POFORMRecType2
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmPOProcessMenu.Show
  Unload frmPOCancel
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog "Close AP"
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub cmdGo_Click()
  Dim Pcnt As Integer, cnt As Integer
  cnt = 0
  If MsgBox("Are You Sure Wish To Cancel This PO?", vbYesNo, "Continue?") = vbNo Then
     Exit Sub
  End If

  If fplstPOs.ListCount <> 0 Then
   For Pcnt = 0 To fplstPOs.ListCount - 1
    If fplstPOs.Selected(Pcnt) Then
      cnt = cnt + 1
      fplstPOs.Row = Pcnt
      CancelPO
      fpcboVendCode.ListIndex = -1
      fptxtVendName.Text = ""
      fplstPOs.Clear
      'fpcboVendCode.SetFocus
      Exit For
    End If
   Next
   If cnt = 0 Then
     MsgBox "You Must Select A Purchase Order First.", vbOKOnly, "Select PO"
     fplstPOs.SetFocus
   End If
 
  Else
    MsgBox "No Purchase Orders to Select, Try Another Vendor.", vbOKOnly, "No PO's"
    fpcboVendCode.SetFocus
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdGo_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpCancPO
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetEncAcct EncAcct
  StatusBar1.Panels.Item(1).Text = GLUserName
  VendCodeList fpcboVendCode
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
'    Me.SetFocus
  End If
End Sub
Private Sub fpcboVendCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendCode.ListDown = True
  End If
End Sub

Private Sub fpcboVendCode_LostFocus()
  fpcboVendCode.Action = ActionClearSearchBuffer
End Sub

Private Sub fplstPOs_GotFocus()
  If fplstPOs.ListCount = 0 Then
    MsgBox "No Open Purchase Orders For This Vendor, Please Select Another.", vbOKOnly, "No PO's"
    fpcboVendCode.SetFocus
  End If
End Sub

Private Sub fpcboVendCode_Click()
  fplstPOs.Clear
  If fpcboVendCode.ListIndex <> -1 Then
    LoadUp
  End If
End Sub
Private Sub LoadUp()
  Dim VendorFile As Integer, NumVRecs As Integer, VRecNum As Integer
  Dim Last As Integer, cnt As Integer, Dcnt As Integer, TmpAcct As Integer
  fpcboVendCode.col = 1
  VRecNum = fpcboVendCode.ColText
  fptxtVendName.Text = ""
  If VRecNum > 0 Then
    OpenVendorFile VendorFile, NumVRecs
    Get VendorFile, VRecNum, Vendor
    fptxtVendName.Text = Vendor.VNAME
    FindPO VRecNum
    End If
 Close
End Sub
Private Sub FindPO(vrec As Integer)
  Dim POCnt As Integer, NextTrans As Long, fmt As String
  Dim VendorFile As Integer, NumVRecs As Integer, tempstr As String
  Dim APLedgerFile As Integer, NumTrans As Long, LdRecLen As Integer
  ReDim APLedgerRec(1) As APLedger81RecType
  LdRecLen = Len(APLedgerRec(1))

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  fmt = "$ ###,###,###.##"
  Get VendorFile, vrec, Vendor
  NextTrans& = Vendor.FrstTran
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 4 Then
      POCnt = POCnt + 1
    End If
    NextTrans& = APLedgerRec(1).NextTrans

  Loop

  If POCnt <> 0 Then

  NextTrans& = Vendor.FrstTran
  fplstPOs.Clear
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 4 Then
      tempstr = Space$(50)
      Mid$(tempstr, 1) = QPTrim$(APLedgerRec(1).PONum)
      Mid$(tempstr, 10) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
      Mid$(tempstr, 20) = Using(fmt, Str$(APLedgerRec(1).Amt))
      Mid$(tempstr, 45) = NextTrans&
      fplstPOs.AddItem tempstr
    End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  End If
  Close
End Sub
Private Sub CancelPO()
  Dim VoidTransRecNum As Long
 
  VoidTransRecNum& = QPTrim(Val(Mid$(fplstPOs.Text, 45, 10)))
  VoidPOTrans VoidTransRecNum&
End Sub
Private Sub VoidPOTrans(VoidTransRecNum&)
  Dim LdRecLen As Integer, DistRecLEn As Integer, POLogFileName As String
  Dim APLedgerFile As Integer, NumTrans As Long, IFRec As Integer
  Dim POIFFile As String, GLIFRecLen As Integer, GLIFFile As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, NextDist As Long
  Dim AcctNum As String, BadAcct As Integer, ReportFile As String
  ReDim ApLedger(1) As APLedger81RecType
  ReDim DistRec(1) As APDistRecType
  LdRecLen = Len(ApLedger(1))
  DistRecLEn = Len(DistRec(1))


  POIFFile$ = "POVDIF.DAT"
  KillFile POIFFile$
  ReDim GLifRec(1) As GLTransRecType
  GLIFRecLen = Len(GLifRec(1))
  GLIFFile = FreeFile
  Open POIFFile$ For Random As GLIFFile Len = GLIFRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  Get APLedgerFile, VoidTransRecNum&, ApLedger(1)
  NextDist& = ApLedger(1).FrstDist
  Close APLedgerFile

  Do Until NextDist& = 0
    Get APDistFile, NextDist&, DistRec(1)
    IFRec = IFRec + 1
    'make sure distribution hasn't been liquidated
    If (QPTrim(DistRec(1).DistStat)) <> "L" Then
    '--Make Debit side of entry

      GLifRec(1).Src = "VD" + Format(Date$, "mmddyy")
      AcctNum$ = Left$(DistRec(1).DistAcctNum, GLFundLen) + EncAcct$
      GLifRec(1).AcctNum = AcctNum$
      GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", Date$)
      GLifRec(1).Desc = "CANCELLED PO"            'APLedger(1).PONum
      GLifRec(1).Ref = ApLedger(1).PONum
      GLifRec(1).CrAmt = 0
      GLifRec(1).DrAmt = DistRec(1).DistAmt
      Put GLIFFile, IFRec, GLifRec(1)
    
      IFRec = IFRec + 1
      'AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
      'GLIFRec(1).AcctNum = AcctNum$
      GLifRec(1).AcctNum = DistRec(1).DistAcctNum
      GLifRec(1).CrAmt = DistRec(1).DistAmt
      GLifRec(1).DrAmt = 0
      Put GLIFFile, IFRec, GLifRec(1)
    Else
      GLifRec(1).Src = "VD" + Format(Date$, "mmddyy")
      AcctNum$ = Left$(DistRec(1).DistAcctNum, GLFundLen) + EncAcct$
      GLifRec(1).AcctNum = AcctNum$
      GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", Date$)
      GLifRec(1).Desc = "CANCELLED PO"            'APLedger(1).PONum
      GLifRec(1).Ref = ApLedger(1).PONum
      GLifRec(1).CrAmt = 0
      GLifRec(1).DrAmt = 0
      Put GLIFFile, IFRec, GLifRec(1)
  
      IFRec = IFRec + 1
      'AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
      'GLIFRec(1).AcctNum = AcctNum$
      GLifRec(1).AcctNum = DistRec(1).DistAcctNum
      GLifRec(1).CrAmt = 0
      GLifRec(1).DrAmt = 0
      Put GLIFFile, IFRec, GLifRec(1)

    End If
    NextDist& = DistRec(1).NextDist

  Loop

  Close

  Post2PO POIFFile$, BadAcct, frmPOCancel, False
  If BadAcct <> 0 Then
    '--Couldn't find an account.
    '--Account was possibly deleted after entry made?
      MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
      ReportFile$ = "TempLog.PRN"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Error Log"
      End If
      frmCitiCancel.Show
      Unload frmPOCancel
      frmPOProcessMenu.Show
      Exit Sub

  End If
  Post2PO POIFFile$, BadAcct%, frmPOCancel, True
  If BadAcct <> 0 Then                  'posting problem
      MsgBox "Error, One or more transactions were not posted. Make sure the printer is ready and Press a Key to View Log.", vbOKOnly, "Posting Error"
      POLogFileName = "POlog.dat"
      ReportFile$ = "POlog.dat"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Posting Log"
      End If
   End If
  
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  Get APLedgerFile, VoidTransRecNum&, ApLedger(1)
  ApLedger(1).TRCode = -4
  Put APLedgerFile, VoidTransRecNum&, ApLedger(1)
  Close APLedgerFile

  KillFile POIFFile$
MsgBox "Cancel Purchase Order Complete.", vbOKOnly, "Completed"
End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
