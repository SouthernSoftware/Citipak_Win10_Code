VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmRptCycleSum1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cycle Count Summary"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptCycleSum1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   6144
      TabIndex        =   4
      Top             =   4128
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmRptCycleSum1.frx":08CA
   End
   Begin VB.CheckBox chkDetail 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6168
      TabIndex        =   6
      Top             =   4632
      Width           =   2388
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
      Left            =   9504
      TabIndex        =   1
      Top             =   6864
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
      Left            =   7848
      TabIndex        =   0
      Top             =   6864
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8280
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
            TextSave        =   "10:53 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "5/13/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1860
      Left            =   2868
      Top             =   3600
      Width           =   6444
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
      Left            =   3624
      TabIndex        =   5
      Top             =   4176
      Width           =   2388
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   1416
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Cycle Count Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   3
      Top             =   1656
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   1296
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
Attribute VB_Name = "frmRptCycleSum1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptCycleSum
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
        cmdExit.SetFocus
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
        UBLog "Closed via RptCycleSum by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub


Private Sub cmdPrint_Click()
  DeActivateControls Me, True
  If fpcboRptType.ListIndex = 0 Then
    CycleCount2
  ElseIf fpcboRptType.ListIndex = 1 Then
    CycleCount
    ActivateControls Me, True
  Else
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
Private Sub CycleCount()
  Dim PageNo As Integer, Title As String, MaxLines As Integer
  Dim Dash80 As String, IndexName As String, UBCustRecLen As Integer
  Dim UBSetupreclen As Integer, IdxRecLen As Integer, IdxFileSize As Long
  Dim IdxNumOfRecs As Long, NumCust As Long, Handle As Integer
  Dim CCnt As Long, UBCust As Integer, UBRpt As Integer, UBFile As Integer
  Dim zz As Integer, CustCnt As Integer, ReportFile As String, Maxlength As Integer
  FrmShowPctComp.Label1 = "Creating Cycle Count Summary Report."
  FrmShowPctComp.Show , Me
  Maxlength = 50
  Title$ = "Cycle Count Summary Report"
  Dash80$ = String$(80, "-")

  ReDim Cycles(-1 To 99) As CycleType

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReportFile$ = UBPath$ + "UBCYCLE.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

'***************
  UBFile = FreeFile
  Open UBPath$ + "UBCUST.dat" For Random Shared As UBFile Len = UBCustRecLen
  NumCust& = LOF(UBFile) \ UBCustRecLen
  If chkDetail.Value = 1 Then GoSub DoLocaRptHeader
  For CCnt& = 1 To NumCust&
     FrmShowPctComp.ShowPctComp CCnt, NumCust
     If FrmShowPctComp.Out = True Then
       Close
       FrmShowPctComp.Out = False
       GoTo CycleAbortExit
     End If

    Get UBFile, CCnt&, UBCustRec(1)
    If UBCustRec(1).BILLCYCL >= 0 And (Not UBCustRec(1).DelFlag) Then
      Cycles(UBCustRec(1).BILLCYCL).CustCnt = Cycles(UBCustRec(1).BILLCYCL).CustCnt + 1
      If chkDetail.Value = 1 Then
        If LineCnt >= Maxlength Then
            Print #UBRpt, Chr$(12)
            GoSub DoLocaRptHeader
        End If
        Print #UBRpt, Tab(5); UBCustRec(1).BILLCYCL; Tab; (14); QPTrim$(UBCustRec(1).CustName); Tab(45); Using("######", CCnt&)
        LineCnt = LineCnt + 1
      End If
      If UBCustRec(1).Status = "P" Then
        Cycles(UBCustRec(1).BILLCYCL).PendCnt = Cycles(UBCustRec(1).BILLCYCL).PendCnt + 1
      End If
    Else
      If Not UBCustRec(1).DelFlag Then
        Cycles(-1).CustCnt = Cycles(-1).CustCnt + 1
        If chkDetail.Value = 1 Then
          If LineCnt >= Maxlength Then
              Print #UBRpt, Chr$(12)
              GoSub DoLocaRptHeader
          End If
          Print #UBRpt, Tab(5); "??"; Tab; (14); QPTrim$(UBCustRec(1).CustName); Tab(45); Using("######", CCnt&)
          LineCnt = LineCnt + 1
        End If
        If UBCustRec(1).Status = "P" Then
            Cycles(-1).PendCnt = Cycles(-1).PendCnt + 1
        End If
      End If
    End If
  Next
  Print #UBRpt, Chr$(12)
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$; Tab(55); "Page "; PageNo
  Print #UBRpt, "                           Cycle Count Summary Report"
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt,
  Print #UBRpt, "Cycle"; Tab(30); "Customer Count"; Tab(48); "Pending"
  For zz = 0 To 99
    If Cycles(zz).CustCnt > 0 Then
      Print #UBRpt, Using("######", zz); Tab(30); Using("######", Cycles(zz).CustCnt); Tab(48); Using("######", Cycles(zz).PendCnt)
    End If
  Next

  If Cycles(-1).CustCnt > 0 Then
    Print #UBRpt, "    ??";
    Print #UBRpt, Tab(30); Using("######", Cycles(-1).CustCnt); Tab(48); Using("######", Cycles(-1).PendCnt)
  End If
  Print #UBRpt, Dash80$
  Print #UBRpt, Chr$(12)


  Close

  Erase UBCustRec

'  If Not AbortFlag Then
'    PrintRptFile , , 1, RetCode, EntryPoint
'  End If
ViewPrint ReportFile$, Title$

Exit Sub

DoLocaRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$; Tab(55); "Page "; PageNo
  Print #UBRpt, "                           Cycle Count Summary Report"
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt,
  Print #UBRpt, "Bill Cycle            Customer                        Account"
  Print #UBRpt, Dash80$
  LineCnt = 6
Return
CycleAbortExit:
  Exit Sub
End Sub
Private Sub CycleCount2()
  Dim Title As String, UBCustRecLen As Integer, IdxFileSize As Long
  Dim IndexName As String, UBFile As Integer
  Dim UBSetupreclen As Integer, IdxRecLen As Integer
  Dim IdxNumOfRecs As Long, NumCust As Long, Handle As Integer
  Dim CCnt As Long, UBCust As Integer, UBRpt As Integer
  Dim zz As Integer, CustCnt As Integer, ReportFile As String
  Dim SubRpt As Integer, SubFile As String
  FrmShowPctComp.Label1 = "Creating Cycle Count Summary Report."
  FrmShowPctComp.Show , Me
  
  Title$ = "Cycle Count Summary Report"

  ReDim Cycles(-1 To 99) As CycleType

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReportFile$ = UBPath$ + "UBCYCLE.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  SubFile$ = UBPath$ + "UBSubCy.rpt"
  SubRpt = FreeFile
  Open SubFile$ For Output As SubRpt


'***************
  UBFile = FreeFile
  Open UBPath$ + "UBCUST.dat" For Random Shared As UBFile Len = UBCustRecLen
  NumCust& = LOF(UBFile) \ UBCustRecLen

  For CCnt& = 1 To NumCust&
     FrmShowPctComp.ShowPctComp CCnt, NumCust
     If FrmShowPctComp.Out = True Then
       Close
       FrmShowPctComp.Out = False
       ActivateControls Me, True
       GoTo CycleAbortExit
     End If

    Get UBFile, CCnt&, UBCustRec(1)
    If UBCustRec(1).BILLCYCL >= 0 And (Not UBCustRec(1).DelFlag) Then
      Cycles(UBCustRec(1).BILLCYCL).CustCnt = Cycles(UBCustRec(1).BILLCYCL).CustCnt + 1
      If chkDetail.Value = 1 Then
        Print #UBRpt, Str(UBCustRec(1).BILLCYCL) + "~" + QPTrim$(UBCustRec(1).CustName) + "~" + Using("######", CCnt&)
      End If
      If UBCustRec(1).Status = "P" Then
        Cycles(UBCustRec(1).BILLCYCL).PendCnt = Cycles(UBCustRec(1).BILLCYCL).PendCnt + 1
      End If
    Else
      If Not UBCustRec(1).DelFlag Then
        Cycles(-1).CustCnt = Cycles(-1).CustCnt + 1
        If chkDetail.Value = 1 Then
          Print #UBRpt, "??" + "~" + QPTrim$(UBCustRec(1).CustName) + "~" + Using("######", CCnt&)
        End If
        If UBCustRec(1).Status = "P" Then
            Cycles(-1).PendCnt = Cycles(-1).PendCnt + 1
        End If
      End If
    End If
  Next
  Print #SubRpt, "Cycle~Customer Count~Pending"
  For zz = 0 To 99
    If Cycles(zz).CustCnt > 0 Then
      Print #SubRpt, Using("######", zz) + "~" + Using("###############", Cycles(zz).CustCnt) + "~" + Using("######", Cycles(zz).PendCnt)
    End If
  Next

  If Cycles(-1).CustCnt > 0 Then
    Print #SubRpt, "    ??" + "~" + Using("######", Cycles(-1).CustCnt) + "~" + Using("######", Cycles(-1).PendCnt)
    
  End If


  Close

  Erase UBCustRec

'ViewPrint "UBCYCLE.RPT", Title$
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptCycleSum
    ARptCycleSum.txtDate = Now
    ARptCycleSum.txtTown = TOWNNAME$
    ARptCycleSum.Title = Title$
    ARptCycleSum.GetName ReportFile$, SubFile$
    ARptCycleSum.startrpt

Exit Sub

'DoLocaRptHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, "                           Cycle Count Summary Report"
'  Print #UBRpt, "Report Date: "; Date$
'  Print #UBRpt,
'  Print #UBRpt, "Bill Cycle            Customer Count            Pending"
'  Print #UBRpt, Dash80$
'Return
CycleAbortExit:
  Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
