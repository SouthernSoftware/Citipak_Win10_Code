VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBHHMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmUBHHMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHHImptExp 
      Caption         =   "Import/Export Handheld Readings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3851
      TabIndex        =   1
      Top             =   3168
      Width           =   4524
   End
   Begin VB.CommandButton cmdReadingNotes 
      Caption         =   "Print Meter Reading &Notes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3851
      TabIndex        =   2
      Top             =   3960
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitUBHHMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3851
      TabIndex        =   3
      Top             =   4752
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "3:13 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7/6/2018"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hand Held Meter Reading Menu"
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
      Left            =   3624
      TabIndex        =   4
      Top             =   1080
      Width           =   5148
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
End
Attribute VB_Name = "frmUBHHMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim HHType As String

Private Sub cmdHHImptExp_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer, ZZCnt As Integer
  
  If Len(HHType) > 0 Then

  If InStr(TOWNNAME$, "HARRISBURG") Then
    Load frmUBHarryImpExpHHRead
    frmUBHarryImpExpHHRead.lblWhatHH.Caption = HHType
    frmUBHarryImpExpHHRead.Timer1.Enabled = True
    frmUBHarryImpExpHHRead.Show
    DoEvents
    frmUBHarryImpExpHHRead.Show
    Unload Me
  Else
    Load frmUBImpExpHHRead
    frmUBImpExpHHRead.lblWhatHH.Caption = HHType
    frmUBImpExpHHRead.Timer1.Enabled = True
    frmUBImpExpHHRead.Show
    'frmUBImpExpHHRead.Timer1.Enabled = True
    DoEvents
    frmUBImpExpHHRead.Show
    Unload Me
  End If
  Else
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(1).FontSize
    For ZZCnt = 0 To 4
      frmMsgDialog.Label(ZZCnt).FontSize = FntSize + 2
    Next
    MsgText(0) = "ERROR:"
    MsgText(1) = "No Hand Held Device specified in your"
    MsgText(2) = "system configuration. Please go to"
    MsgText(3) = "Utility Billing System Setup and select"
    MsgText(4) = "the appropriate Hand Held device type."
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    'do error no hh defined
  End If
  
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  GetHHType
  Me.HelpContextID = hlpHandHeldMeter
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitUBHHMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo + vbCritical, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitUBHHMenu_Click
      KeyCode = 0
    Case vbKeyHome
      KeyCode = 0
      'cmdManualReadings.SetFocus
    Case vbKeyEnd
      KeyCode = 0
      cmdExitUBHHMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Function GetHHType$()
  Dim UBSetupLen As Integer
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  Select Case UBSetUpRec(1).HHDEVICE
  Case "X", "H", "S", "E", "U", "C", "D", "T", "L", "I", "Z", "B", "W", "J", "P", "G", "Y", "A"
    HHType = UBSetUpRec(1).HHDEVICE
  Case Else
    HHType = ""
  End Select
End Function

Private Sub cmdExitUBHHMenu_Click()
  frmUBMeterMenu.Show
  Unload frmUBHHMenu
End Sub

Private Sub cmdReadingNotes_Click()
  frmRptMeterNotes.HelpContextID = hlpPrintMeterReading
  Load frmRptMeterNotes
  frmRptMeterNotes.Exit2Flag = "2"
  DoEvents
  frmRptMeterNotes.Show
  Unload frmUBHHMenu
End Sub

'Dim CustAcct As Long
'Dim BeenDone As Boolean
'Dim fromform As Form, toform As Form, codeopt As Integer
'Dim uselook As Boolean, Answer As Integer, CredOKFlag As Boolean
'Dim CleveFlag As Boolean, TotalBalance As Double, TempDepAmt As Double
'Dim BtnFnt As Double
'
'Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
'  Set fromform = xfrm
'  Set toform = tfrm
'  If opt <> 0 Then
'    codeopt = opt
'  Else
'    codeopt = 0
'  End If
'  uselook = True
'End Sub

'Private Sub cmdExit_Click()
'  Chk4Change
'  If Answer = 1 Then
'    Exit Sub
'  ElseIf Answer = 2 Then
'    fpCmdSave_Click
'  End If
'  CustAcct = 0
'  fpCustRecNo = 0
'  BeenDone = False
'  If codeopt = 1 Then
'    ActivateControls frmCustEditLookUP
'  ElseIf codeopt = 2 Then
'    ActivateControls frmDisplayList
'  End If
'  If codeopt = 0 Then
'    Load frmUBDepositMenu
'    DoEvents
'    frmUBDepositMenu.Show
'  End If
'
'  UBLog "OUT: UTIL DepCredRem"
'  Unload Me
'  DoEvents
'End Sub
'Private Sub Form_Activate()
'  If Val(fpCustRecNo) > 0 And Not BeenDone Then
'    BeenDone = True
'    loadCustrec
'    DoEvents
'  End If
'
'End Sub
'
'Private Sub fpAmount_Change(Index As Integer)
'  CalcBALFlds
'End Sub
'Private Sub Chk4Change()
'  Answer = 0
'  If fpTotAdjust <> 0 Then
'    frmChangedWarning.Show vbModal, Me
'    Select Case SaveFlag
'    Case False
'      Answer = 3
'    Case True
'      Answer = 2
'    Case 1
'      Answer = 1
'    End Select
'  Else
'    Answer = 0
'  End If
'End Sub
'
''Private Sub fpAmount_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''  Dim x As Integer
''  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
''    If Index < MaxRevsCnt Then
''     For x = Index To (MaxRevsCnt - 1)
''      If fpAmount(x + 1).Enabled Then
''        fpAmount(x + 1).SetFocus
''        Exit For
''      Else
''        fpCmdSave.SetFocus
''        Exit For
''      End If
''     Next
''    End If
''  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
''    If Index > 0 Then
''     For x = Index To (MaxRevsCnt - 1)
''      If fpAmount(x - 1).Enabled Then
''        fpAmount(x - 1).SetFocus
''        Exit For
''      Else
''        fpCmdSave.SetFocus
''      End If
''     Next
''    End If
''  End If
''
''End Sub
'
''Private Sub fpCmdClear_Click()
''  Chk4Change
''  If Answer = 1 Then
''    Exit Sub
''  ElseIf Answer = 2 Then
''    fpCmdSave_Click
''  End If
''
''End Sub
'
''Private Sub fpCmdDist_Click()
''  Autodist
''End Sub
'
'Private Sub fpCmdMsg_Click()
'  If CustAcct& > 0 Then
'    frmCustMsgEdit.CustRec = CustAcct&
'    frmCustMsgEdit.Show vbModal
'    DoEvents
'    If CustHasMsg(CustAcct&) Then
'      MsgAlertTimer.Enabled = True
'    Else
'      MsgAlertTimer.Enabled = False
'      fpCmdMsg.ForeColor = &H80000012
'      'fpCmdMsg.FontSize = BtnFnt
'    End If
'  End If
'
'End Sub
'
'Private Sub fpCmdTranHist_Click()
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer
'  If Len(fptxtAccount) > 0 Then
'    If CustAcct& > 0 Then
'      DeActivateControls Me
'      DisplayCustTransList CustAcct&
'      ActivateControls Me
'    Else
'      frmMsgDialog.RetLabel = "-2"
'      FntSize = frmMsgDialog.Label(2).FontSize
'      frmMsgDialog.Label(2).FontSize = (FntSize + 2)
'      MsgText(0) = "ERROR:"
'      MsgText(1) = ""
'      MsgText(2) = ""
'      MsgText(3) = "There are NO transactions to display."
'      MsgText(4) = ""
'      MsgText(5) = ""
'      GetOKorNot MsgText(), True
'    End If
'  End If
'End Sub
'
'
'Private Sub fpCmdSave_Click()
'  CalcBALFlds
'  CheckApplyInfo
'  If CredOKFlag Then
'    If MsgBox("Are you sure you wish to save this transaction?", vbYesNo, "Save Transaction") = vbYes Then
'      SaveTransaction
'      CustAcct = 0
'      fpCustRecNo = 0
'      BeenDone = False
'      If codeopt = 1 Then
'        ActivateControls frmCustEditLookUP
'      ElseIf codeopt = 2 Then
'        ActivateControls frmDisplayList
'      End If
'      If codeopt = 0 Then
'        Load frmUBDepositMenu
'        DoEvents
'        frmUBDepositMenu.Show
'      End If
'
'      UBLog "OUT: UTIL DepCredRem"
'      Unload Me
'      DoEvents
'    End If
'  End If
'End Sub

'Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
'    KeyCode = 0
'    fpAmount(0).SetFocus
'  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
'    fpCmdSave.SetFocus
'  End If
'End Sub
'
'Private Sub mnuExit_Click()
'  cmdExit_Click
'End Sub
'
'Private Sub mnuPrnScn_Click()
'  PrintForm
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
'      Cancel = True
'    Else
'      UBLog "OUT: UTIL DepCredRem"
'    End If
'  End If
'End Sub
'
'Private Sub MsgAlertTimer_Timer()
'  Static tog As Double
'  Static TogState As Boolean
'  If Me.Visible Then
'    If BtnFnt# = 0 Then
'      BtnFnt# = fpCmdMsg.FontSize
'    End If
'    If TogState Then
'      tog = tog + 1
'    Else
'      tog = tog - 1
'    End If
'    Select Case tog
'    Case 1
'      fpCmdMsg.ForeColor = &H80000012
'      fpCmdMsg.FontSize = BtnFnt
'    Case 2
'      fpCmdMsg.ForeColor = &H80000011
'      fpCmdMsg.FontSize = BtnFnt - 0.7
'    Case 3
'      fpCmdMsg.ForeColor = &H80000011
'      fpCmdMsg.FontSize = BtnFnt - 1.4
'    Case 4
'      fpCmdMsg.ForeColor = &H80000010
'      fpCmdMsg.FontSize = BtnFnt - 2.1
'    Case 5
'      fpCmdMsg.ForeColor = &H80000010
'      fpCmdMsg.FontSize = BtnFnt - 2.8
'    Case 6
'      fpCmdMsg.ForeColor = &H8000000F
'      fpCmdMsg.FontSize = BtnFnt - 3.5
'    Case 7
'      fpCmdMsg.ForeColor = &H8000000F
'      fpCmdMsg.FontSize = BtnFnt - 4.2
'    Case 8
'      fpCmdMsg.ForeColor = &H8000000E
'      fpCmdMsg.FontSize = BtnFnt - 4.9
'    Case 9
'      fpCmdMsg.ForeColor = &H8000000E
'      fpCmdMsg.FontSize = BtnFnt - 5.6
'    End Select
'    Select Case tog
'    Case Is < 0, Is > 9
'      TogState = Not TogState
'    End Select
'  End If
''  DoEvents
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
''    Case vbKeyDown, vbKeyReturn:
''      SendKeys "{Tab}"
''      KeyCode = 0
''    Case vbKeyUp:
''      SendKeys "+{Tab}"
''      KeyCode = 0
'    Case vbKeyEscape:
'      KeyCode = 0
'      DoEvents
'      Call cmdExit_Click
'    Case vbKeyF4:
'      KeyCode = 0
'      DoEvents
'      Call fpCmdTranHist_Click
'    Case vbKeyF7:
'      KeyCode = 0
'      DoEvents
'      Call fpCmdMsg_Click
'    Case vbKeyF10:
'      KeyCode = 0
'      DoEvents
'      Call fpCmdSave_Click
'    Case Else:
'  End Select
'End Sub
'

'Private Sub LoadRevs()
'  Dim NumOfRevs As Integer, UBSetupLen As Integer, RevCnt As Integer
'  Dim InvRev As Integer
'  NumOfRevs = MaxRevsCnt
'
'  ReDim RevText$(1 To MaxRevsCnt)
'
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'
'  For RevCnt = 1 To MaxRevsCnt
'    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
'    If Len(RevText$(RevCnt)) = 0 Then
'      NumOfRevs = RevCnt - 1
'      Exit For
'    End If
'  Next
'
'  If NumOfRevs < MaxRevsCnt Then
'    ReDim Preserve RevText$(1 To NumOfRevs)
'  End If
'
'  For RevCnt = 1 To NumOfRevs
'    fpRevSource(RevCnt - 1) = RevText$(RevCnt)
'  Next
'  For InvRev = NumOfRevs To 14
'    fpRevSource(InvRev).Enabled = False
'    fpRevSource(InvRev).Visible = False
'    fpAmount(InvRev).Enabled = False
'    fpAmount(InvRev).Visible = False
'    fpCurrent(InvRev).Enabled = False
'    fpCurrent(InvRev).Visible = False
'    fpActual(InvRev).Enabled = False
'    fpActual(InvRev).Visible = False
'  Next
'
'End Sub
'
'Private Sub loadCustrec()
'  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
'  Dim CustFile As Integer, cnt As Integer
'  ReDim UBCustRec(1) As NewUBCustRecType
'  Dim NumOfRevs As Integer, RevCnt As Integer
'  NumOfRevs = MaxRevsCnt
'  UBCustRecLen = Len(UBCustRec(1))
'
''  If uselook = True Then
''    Unload frmCustEditLookUP
''    Unload frmDisplayList
''    uselook = False
''  End If
'  CustAcct = fpCustRecNo
'  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
''  If CustAcct& > NumOfCustRecs& Or CustAcct& <= 0 Then
''    CustAcct& = 0
''    LabelDel.Visible = True
''    GoTo SkiptoHere
''  End If
'
''  If IsDeleted(CustAcct&) Then
''    CustAcct& = 0
''    LabelDel.Caption = "Deleted Account!"
''    LabelDel.Visible = True
''    GoTo SkiptoHere
''  End If
'  CustFile = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
'  Get CustFile, CustAcct&, UBCustRec(1)
'  Close CustFile
'  fptxtAccount = Str$(CustAcct&)
'  fpCustName = UBCustRec(1).CustName
'  CustAcct& = Val(CustAcct)
'
'  For RevCnt = 0 To NumOfRevs - 1
'    If fpCurrent(RevCnt).Enabled = True Then
'      fpCurrent(RevCnt) = UBCustRec(1).CurrRevAmts(RevCnt + 1)
'    End If
'  Next
'  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'  fpBalance = TotalBalance#
'  If CustHasMsg(CustAcct) Then
'    MsgAlertTimer.Enabled = True
'  End If
'  Autodist
'  CalcBALFlds
'  Exit Sub
'SkiptoHere:
'
'End Sub

''Private Sub ClearScn()
''  Dim cnt As Integer
''  For cnt = 1 To 15
''    fpAmount(cnt - 1) = 0
''  Next
''  CalcCashFlds
''End Sub
