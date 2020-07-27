VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmRptMtrNoSerial 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12216
   Icon            =   "frmMtrNoSerial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboNoLocation 
      Height          =   348
      Left            =   7392
      TabIndex        =   0
      Top             =   3768
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   614
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmMtrNoSerial.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   7392
      TabIndex        =   1
      Top             =   4440
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
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
      ColDesigner     =   "frmMtrNoSerial.frx":0C68
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
      Left            =   9234
      TabIndex        =   3
      Top             =   7224
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
      Left            =   7554
      TabIndex        =   2
      Top             =   7224
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4:34 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "3/31/2004"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Meters/No Serial Numbers"
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
      Left            =   3216
      TabIndex        =   7
      Top             =   1224
      Width           =   5700
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   984
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2172
      Left            =   2520
      Top             =   3192
      Width           =   7212
   End
   Begin VB.Label Label2 
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
      Left            =   4872
      TabIndex        =   6
      Top             =   4464
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Skip Accounts With No Location Number:"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   3816
      Width           =   4908
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   864
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
Attribute VB_Name = "frmRptMtrNoSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Grpt As Boolean
Private Sub cmdExit_Click()
  frmUBMeterMenu.Show
  Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptMtrNoSerial by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
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
        fpcboNoLocation.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboNoLocation_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboNoLocation.ListDown = True
  End If
  If fpcboNoLocation.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
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

Private Sub cmdPrint_Click()
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
       NoSerialNumRPT True
    ElseIf fpcboRptType.ListIndex = 1 Then
      NoSerialNumRPT False
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
  fpcboNoLocation.AddItem "Yes"
  fpcboNoLocation.AddItem "No"
  fpcboNoLocation.ListIndex = 1
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


Private Sub NoSerialNumRPT(graphicflag As Boolean)
  Dim UBCustRecLen As Integer, ReportFile As String
  Dim UBSetupLen As Integer, IdxName As String, lcnt As Long
  Dim NumOfRecs As Long, IdxNumOfRecs As Long, Handle As Integer
  Dim cnt As Long, VacFlag As Boolean, SAddr As String
  Dim Header As String, CustName As String, NumOfCust As Long
  Dim PrintedOne As Boolean, MeterType As String, Book As Integer
  Dim ToPrint As String, IdxRecLen As Long, RptHandle As Integer
  Dim MaxMeterAmt As Long, Page As Integer, IdxFileSize As Long
  Dim DidOne As Boolean, MtrCnt As Long, TempRev As String
  Dim NewFlag As Boolean, UBCust As Integer, BegRoute As Integer
  Dim EndRoute As Integer, NoLocFlag As Boolean
  FrmShowPctComp.Label1 = "Creating No Serial Number Report"
  FrmShowPctComp.Show , Me

  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  MaxLines = 52
  FF$ = Chr$(12)

  BegRoute = 1
  EndRoute = 99
  If fpcboNoLocation.ListIndex = 0 Then
    NoLocFlag = True
  Else
    NoLocFlag = False
  End If

  'Open Report File
  ReportFile$ = UBPath$ + "WBMTRLST.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  'Open the Utility Setup File to Grab Meter List Order (Seq or Loc)
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  CustName$ = Space$(30)

  ' Location Order ********************************************************
  'if UBSetupRec(1).
  IdxName$ = UBPath$ + "UBCUSTBK.IDX"
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&(IdxName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  Handle = FreeFile
  Open IdxName$ For Random Shared As Handle Len = IdxRecLen
  For cnt& = 1 To IdxNumOfRecs
    Get #Handle, cnt&, IdxBuff(cnt&)
  Next
  Close Handle

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  cnt& = 1
  If graphicflag = False Then
    GoSub PrnNOSerHeader
    'ShowProcessingScrn "Reading Meter Information"
    For lcnt& = 1 To IdxNumOfRecs
      FrmShowPctComp.ShowPctComp lcnt, IdxNumOfRecs&
      If FrmShowPctComp.Out = True Then
        FrmShowPctComp.Out = False
        GoTo ExitHere
      End If
  
      Get #UBCust, IdxBuff(lcnt&).RecNum, UBCustRec(1)
      If UBCustRec(1).DelFlag = 0 Then
        If UBCustRec(1).Book = "  " And UBCustRec(1).SEQNUMB = "      " Then
        '  'This if want to leave off the blank loc cust
          If NoLocFlag Then
            GoTo FlagSkip
          End If
        End If
        If InStr(UBCustRec(1).HHMSG1, "NOREAD") = 0 Then
          LSet CustName$ = QPTrim$(UBCustRec(1).CustName)
          If Len(UBCustRec(1).Status) > 0 Then
            For MtrCnt& = 1 To 7                'find last active meter
              If (UBCustRec(1).LocMeters(MtrCnt&).CurRead >= 0) Or (UBCustRec(1).LocMeters(MtrCnt&).PrevRead) > 0 Then
                If (Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrNum)) = 0) And Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)) > 0 Then
                  Print #RptHandle, QPTrim$(UBCustRec(1).Book); "-"; QPTrim$(UBCustRec(1).SEQNUMB);
                  Print #RptHandle, ; Tab(13); QPTrim$(UBCustRec(1).Status);
                  Print #RptHandle, ; Tab(18); CustName$;
                  Print #RptHandle, ; Tab(48); Left$(QPTrim$(UBCustRec(1).ServAddr), 30);
                  Print #RptHandle, ; Tab(79); QPTrim$(Str$(MtrCnt&))
                  LineCnt = LineCnt + 1
                  Exit For
                End If
              End If
            Next
          End If
        End If
      End If
FlagSkip:
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrnNOSerHeader
      End If
    Next
  
    Print #RptHandle, FF$
  Else
      For lcnt& = 1 To IdxNumOfRecs
      FrmShowPctComp.ShowPctComp lcnt, IdxNumOfRecs&
      If FrmShowPctComp.Out = True Then
        FrmShowPctComp.Out = False
        GoTo ExitHere
      End If
      Get #UBCust, IdxBuff(lcnt&).RecNum, UBCustRec(1)
      If UBCustRec(1).DelFlag = 0 Then
        If UBCustRec(1).Book = "  " And UBCustRec(1).SEQNUMB = "      " Then
        '  'This if want to leave off the blank loc cust
          If NoLocFlag Then
            GoTo GFlagSkip
          End If
        End If
        If InStr(UBCustRec(1).HHMSG1, "NOREAD") = 0 Then
          LSet CustName$ = QPTrim$(UBCustRec(1).CustName)
          If Len(UBCustRec(1).Status) > 0 Then
            For MtrCnt& = 1 To 7                'find last active meter
              If (UBCustRec(1).LocMeters(MtrCnt&).CurRead >= 0) Or (UBCustRec(1).LocMeters(MtrCnt&).PrevRead) > 0 Then
                If (Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrNum)) = 0) And Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt&).MtrType)) > 0 Then
                  ToPrint$ = QPTrim$(UBCustRec(1).Book) + "-" + QPTrim$(UBCustRec(1).SEQNUMB)
                  ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).Status)
                  ToPrint$ = ToPrint$ + "~" + CustName$
                  ToPrint$ = ToPrint$ + "~" + Left$(QPTrim$(UBCustRec(1).ServAddr), 30)
                  ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(MtrCnt&))
                  Print #RptHandle, ToPrint$
                  ToPrint$ = ""
                  Exit For
                End If
              End If
            Next
          End If
        End If
      End If
GFlagSkip:
    Next
  End If
  Close

  Header$ = "NO SERIAL NUMBER Report"

  Erase IdxBuff
  If graphicflag Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptMtrNoSerial
    ARptNOMtrNum.txtDate = Now
    ARptNOMtrNum.txtTown = TOWNNAME$
    ARptNOMtrNum.Title = Header$
    ARptNOMtrNum.GetName ReportFile$
    ARptNOMtrNum.startrpt
  Else
    'PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
    ViewPrint ReportFile$, Header$
  End If
  Exit Sub

PrnNOSerHeader:
  Page = Page + 1
  Print #RptHandle, "Date: "; Date$; Tab(30); "NO Meter Number Report";
  Print #RptHandle, Tab(70); "Page #"; Page
  Print #RptHandle, "Location Status Customer Name"; Tab(41); "Service Address"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
Return

ExitHere:
  Close
  Erase IdxBuff

End Sub

