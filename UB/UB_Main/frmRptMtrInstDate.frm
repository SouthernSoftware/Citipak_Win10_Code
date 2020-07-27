VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptMtrInstDate 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meter Installed Date Report"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptMtrInstDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   6330
      TabIndex        =   2
      Top             =   4650
      Width           =   1920
      _Version        =   196608
      _ExtentX        =   3387
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
      ColDesigner     =   "frmRptMtrInstDate.frx":08CA
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
      Left            =   10080
      TabIndex        =   4
      Top             =   7296
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
      Left            =   8400
      TabIndex        =   3
      Top             =   7296
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
            TextSave        =   "1:43 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "9/20/2006"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   6336
      TabIndex        =   1
      Top             =   4104
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      ButtonStyle     =   2
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   6336
      TabIndex        =   0
      Top             =   3576
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      ButtonStyle     =   2
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
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
      Left            =   3948
      TabIndex        =   9
      Top             =   4704
      Width           =   2340
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2100
      Left            =   2820
      Top             =   3192
      Width           =   6564
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date:"
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
      Left            =   4536
      TabIndex        =   8
      Top             =   3624
      Width           =   1668
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
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
      Left            =   4632
      TabIndex        =   7
      Top             =   4164
      Width           =   1572
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3186
      Top             =   1776
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Meter Installed Date Report"
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
      Left            =   3594
      TabIndex        =   6
      Top             =   2016
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3186
      Top             =   1656
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
Attribute VB_Name = "frmRptMtrInstDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptMtrInstDate
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
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
        UBLog "Closed via RptMtrInstDate by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtDate1) = False And CheckValDate(txtDate2) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    TempDate1 = DateDiff("d", "12/31/1979", txtDate1)
    TempDate2 = DateDiff("d", "12/31/1979", txtDate2)
    If TempDate1 > TempDate2 Then
      ValidDate = False
      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
    Else
      ValidDate = True
    End If
  End If
End Function

Private Sub cmdPrint_Click()
If ValidDate Then
  DeActivateControls Me, True
  If fpcboRptType.ListIndex = 0 Then
    MeterInstallDate2
  ElseIf fpcboRptType.ListIndex = 1 Then
    MeterInstallDate
    ActivateControls Me, True
  Else
    ActivateControls Me, True
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
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpMeterInstalledReport
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub MeterInstallDate()
  Dim FromDate As Integer, ThruDate As Integer
  Dim OKFlag As Boolean, IdxRecLen As Integer, IdxFileSize As Long
  Dim IdxNumRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim AcctNumb As Long, Status As String, CustPrinted As Boolean
  Dim MtrDate As Integer, UBCust As Integer, UBRpt As Integer
  Dim UBSetupreclen As Integer, MTCnt As Integer, cnt As Long, TempRev$
  Dim RptName As String, MaxLines As Integer, PageNo As Integer
  Dim Dash80 As String, IndexName As String, UBCustRecLen As Integer
  RptName$ = UBPath$ + "UBMTINDT.RPT"
  FrmShowPctComp.Label1 = "Creating Meter Installed Date Report"
  FrmShowPctComp.Show , Me

  MaxLines = 52
  PageNo = 0
  Dash80$ = String$(80, "-")
  IndexName$ = TempIndexName

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  FromDate = Date2Num(txtDate1)
  ThruDate = Date2Num(txtDate2)
  If FromDate < -29219 Then
    FromDate = -29219        '01-01-1900
  End If
  If FromDate <= ThruDate Then
    OKFlag = True
  End If
  If OKFlag <> True Then
    MsgBox "Invalid Date Range, Try Again.", vbOKOnly, "Invalid Dates"
    Exit Sub
  End If
  SortMtrDateIndex FromDate, ThruDate

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize&(IndexName$)
  IdxNumRecs = IdxFileSize& \ IdxRecLen
  If IdxNumRecs = 0 Then
     Close
     Unload FrmShowPctComp
     MsgBox "No Transactions", vbOKOnly, "No Trans"
     GoTo ExitMtrInsRpt
  End If
  ReDim IdxBuff(1 To IdxNumRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumRecs      'load it
  NumOfRecs& = IdxNumRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBRpt = FreeFile
  Open RptName$ For Output As UBRpt

'  BlockClear
'  ShowProcessingScrn "Meter Installed Report."

  GoSub DoMtrInsDateRptHeader

  For cnt = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitMtrInsRpt
    End If


    Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
    AcctNumb& = IdxBuff(cnt).RecNum

    Status$ = UBCustRec(1).Status
    'LocationNumber$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB

    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoMtrInsDateRptHeader
    End If
    CustPrinted = False
    For MTCnt = 1 To 7
    TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MTCnt).MTRType)
    If Len(TempRev$) <> 0 Then

      MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
      If MtrDate >= FromDate And MtrDate <= ThruDate Then
        If Not CustPrinted Then
          CustPrinted = True
          'PRINT #UBRpt, LocationNumber$;
          Print #UBRpt, Using("#####", AcctNumb&);
          Print #UBRpt, "  "; Status$;
          Print #UBRpt, Tab(17); QPTrim$(UBCustRec(1).CustName);
          Print #UBRpt, Tab(48); Left$(UBCustRec(1).ServAddr, 30)
                    LineCnt = LineCnt + 1
        End If
        Print #UBRpt, Num2Date(UBCustRec(1).LocMeters(MTCnt).InsDate);
        Print #UBRpt, Tab(17); QPTrim$(UBCustRec(1).LocMeters(MTCnt).MtrNum);
        Print #UBRpt, Tab(35); Using("##########", UBCustRec(1).LocMeters(MTCnt).CurRead)
        LineCnt = LineCnt + 1
      End If
    End If
    Next
    If CustPrinted Then
      Print #UBRpt, Dash80$
      LineCnt = LineCnt + 1
    End If
'MtrDateSkip:
   ' ShowPctComp cnt, NumOfRecs&
  Next
  Close
  Erase IdxBuff
  ViewPrint RptName$, "Meter Installed Date Report."
ExitMtrInsRpt:

Exit Sub

DoMtrInsDateRptHeader:

  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$; Tab(70); "Page #"; PageNo
  Print #UBRpt, Tab(28); "Meter Installed Date Report"; Tab(61); " Date: "; Date$
  Print #UBRpt, "Acct #  Status   Cust Name                    Service Addr"
  Print #UBRpt, "Ins Date         Mtr Numb              Current Reading"
  Print #UBRpt, Dash80$
  LineCnt = 5
  Return

Return
End Sub
Private Sub MeterInstallDate2()
  Dim FromDate As Integer, ThruDate As Integer
  Dim OKFlag As Boolean, IdxRecLen As Integer, IdxFileSize As Long
  Dim IdxNumRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim AcctNumb As Long, Status As String, CustPrinted As Boolean
  Dim MtrDate As Integer, UBCust As Integer, UBRpt As Integer
  Dim UBSetupreclen As Integer, MTCnt As Integer, cnt As Long, TempRev$
  Dim RptName As String, IndexName As String, UBCustRecLen As Integer
  Dim ToPrint As String, ToPrintM As String, CustCnt As Long, MtrCnt As Long
  RptName$ = UBPath$ + "UBMTINDT.RPT"
  FrmShowPctComp.Label1 = "Creating Meter Installed Date Report"
  FrmShowPctComp.Show , Me
  IndexName$ = TempIndexName

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  FromDate = Date2Num(txtDate1)
  ThruDate = Date2Num(txtDate2)
  If FromDate < -29219 Then
    FromDate = -29219        '01-01-1900
  End If
  If FromDate <= ThruDate Then
    OKFlag = True
  End If
  If OKFlag <> True Then
    MsgBox "Invalid Date Range, Try Again.", vbOKOnly, "Invalid Dates"
    ActivateControls Me, True
    Exit Sub
  End If
  SortMtrDateIndex FromDate, ThruDate

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize&(IndexName$)
  IdxNumRecs = IdxFileSize& \ IdxRecLen
  If IdxNumRecs = 0 Then
     Close
     Unload FrmShowPctComp
     MsgBox "No Transactions", vbOKOnly, "No Trans"
     ActivateControls Me, True
     GoTo ExitMtrInsRpt
  End If
  ReDim IdxBuff(1 To IdxNumRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumRecs      'load it
  NumOfRecs& = IdxNumRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBRpt = FreeFile
  Open RptName$ For Output As UBRpt

'  BlockClear
'  ShowProcessingScrn "Meter Installed Report."

'  GoSub DoMtrInsDateRptHeader

  For cnt = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitMtrInsRpt
    End If


    Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
    AcctNumb& = IdxBuff(cnt).RecNum

    Status$ = UBCustRec(1).Status
    'LocationNumber$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB

    CustPrinted = False
    For MTCnt = 1 To 7
     TempRev$ = QPTrim$(UBCustRec(1).LocMeters(MTCnt).MTRType)
     If Len(TempRev$) <> 0 Then

      MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
      If MtrDate >= FromDate And MtrDate <= ThruDate Then
        If Not CustPrinted Then
          CustPrinted = True
          'PRINT #UBRpt, LocationNumber$;
          ToPrint$ = Using("#####", AcctNumb&) + "~"
          ToPrint$ = ToPrint$ + Status$ + "~"
          ToPrint$ = ToPrint$ + QPTrim$(UBCustRec(1).CustName) + "~"
          ToPrint$ = ToPrint$ + Left$(UBCustRec(1).ServAddr, 30)
          CustCnt = CustCnt + 1
        End If
        ToPrintM$ = Num2Date(UBCustRec(1).LocMeters(MTCnt).InsDate) + "~"
        ToPrintM$ = ToPrintM$ + QPTrim$(UBCustRec(1).LocMeters(MTCnt).MtrNum) + "~"
        ToPrintM$ = ToPrintM$ + Using("##########", UBCustRec(1).LocMeters(MTCnt).CurRead)
        MtrCnt = MtrCnt + 1
      Print #UBRpt, ToPrint$ + "~" + ToPrintM$
      End If
     End If
    Next
   ' If CustPrinted Then
      'Print #UBRpt, Dash80$
      'Linecnt = Linecnt + 1
   ' End If
'MtrDateSkip:
   ' ShowPctComp cnt, NumOfRecs&
  Next
  Close
  Erase IdxBuff
  
  If CustCnt > 0 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptMtrInstDate
    ARptMtrInstDate.txtDate = Now
    ARptMtrInstDate.txtTown = TOWNNAME$
    ARptMtrInstDate.Title = "Meter Installed Date Report."
    ARptMtrInstDate.totMeters = MtrCnt
    ARptMtrInstDate.totCust = CustCnt
    ARptMtrInstDate.txtDateRange = "Report Date Range: " + txtDate1 + " - " + txtDate2
    ARptMtrInstDate.GetName RptName$
    ARptMtrInstDate.startrpt

  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If

 ' ViewPrint RptName$, "Meter Installed Date Report."
ExitMtrInsRpt:

Exit Sub

DoMtrInsDateRptHeader:
'
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$; Tab(70); "Page #"; PageNo
'  Print #UBRpt, Tab(28); "Meter Installed Date Report"; Tab(61); " Date: "; Date$
'  Print #UBRpt, "Acct #  Status   Cust Name                    Service Addr"
'  Print #UBRpt, "Ins Date         Mtr Numb              Current Reading"
'  Print #UBRpt, Dash80$
'  Linecnt = 5
  Return

Return
End Sub

