VERSION 5.00
Begin VB.Form frmfrmCMPaymentSource 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Payment Source"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmCMPaymentSource.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmfrmCMPaymentSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CustAcct As Long
Dim BeenDone As Boolean
Dim fromform As Form, toform As Form, codeopt As Integer
Dim uselook As Boolean, Answer As Integer, CredOKFlag As Boolean
Dim CleveFlag As Boolean, TotalBalance As Double, TempDepAmt As Double
Dim BtnFnt As Double

Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  uselook = True
End Sub

Private Sub cmdExit_Click()
  Chk4Change
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    fpCmdSave_Click
  End If
  CustAcct = 0
  fpCustRecNo = 0
  BeenDone = False
  If codeopt = 1 Then
    ActivateControls frmCustEditLookUP
  ElseIf codeopt = 2 Then
    ActivateControls frmDisplayList
  End If
  If codeopt = 0 Then
    Load frmUBDepositMenu
    DoEvents
    frmUBDepositMenu.Show
  End If

  UBLog "OUT: UTIL DepCredRem"
  Unload Me
  DoEvents
End Sub
Private Sub Form_Activate()
  If Val(fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    loadCustrec
    DoEvents
  End If
  
End Sub

Private Sub fpAmount_Change(Index As Integer)
  CalcBALFlds
End Sub
Private Sub Chk4Change()
  Answer = 0
  If fpTotAdjust <> 0 Then
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      Answer = 3
    Case True
      Answer = 2
    Case 1
      Answer = 1
    End Select
  Else
    Answer = 0
  End If
End Sub

'Private Sub fpAmount_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'  Dim x As Integer
'  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
'    If Index < MaxRevsCnt Then
'     For x = Index To (MaxRevsCnt - 1)
'      If fpAmount(x + 1).Enabled Then
'        fpAmount(x + 1).SetFocus
'        Exit For
'      Else
'        fpCmdSave.SetFocus
'        Exit For
'      End If
'     Next
'    End If
'  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
'    If Index > 0 Then
'     For x = Index To (MaxRevsCnt - 1)
'      If fpAmount(x - 1).Enabled Then
'        fpAmount(x - 1).SetFocus
'        Exit For
'      Else
'        fpCmdSave.SetFocus
'      End If
'     Next
'    End If
'  End If
'
'End Sub

'Private Sub fpCmdClear_Click()
'  Chk4Change
'  If Answer = 1 Then
'    Exit Sub
'  ElseIf Answer = 2 Then
'    fpCmdSave_Click
'  End If
'
'
'End Sub

'Private Sub fpCmdDist_Click()
'  Autodist
'End Sub

Private Sub fpCmdMsg_Click()
  If CustAcct& > 0 Then
    frmCustMsgEdit.CustRec = CustAcct&
    frmCustMsgEdit.Show vbModal
    DoEvents
    If CustHasMsg(CustAcct&) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      fpCmdMsg.ForeColor = &H80000012
      'fpCmdMsg.FontSize = BtnFnt
    End If
  End If

End Sub

Private Sub fpCmdTranHist_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  If Len(fptxtAccount) > 0 Then
    If CustAcct& > 0 Then
      DeActivateControls Me
      DisplayCustTransList CustAcct&
      ActivateControls Me
    Else
      frmMsgDialog.RetLabel = "-2"
      FntSize = frmMsgDialog.Label(2).FontSize
      frmMsgDialog.Label(2).FontSize = (FntSize + 2)
      MsgText(0) = "ERROR:"
      MsgText(1) = ""
      MsgText(2) = ""
      MsgText(3) = "There are NO transactions to display."
      MsgText(4) = ""
      MsgText(5) = ""
      GetOKorNot MsgText(), True
    End If
  End If
End Sub


Private Sub fpCmdSave_Click()
  CalcBALFlds
  CheckApplyInfo
  If CredOKFlag Then
    If MsgBox("Are you sure you wish to save this transaction?", vbYesNo, "Save Transaction") = vbYes Then
      SaveTransaction
      CustAcct = 0
      fpCustRecNo = 0
      BeenDone = False
      If codeopt = 1 Then
        ActivateControls frmCustEditLookUP
      ElseIf codeopt = 2 Then
        ActivateControls frmDisplayList
      End If
      If codeopt = 0 Then
        Load frmUBDepositMenu
        DoEvents
        frmUBDepositMenu.Show
      End If
    
      UBLog "OUT: UTIL DepCredRem"
      Unload Me
      DoEvents
    End If
  End If
End Sub
Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
    KeyCode = 0
    fpAmount(0).SetFocus
  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
    fpCmdSave.SetFocus
  End If
End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      UBLog "OUT: UTIL DepCredRem"
    End If
  End If
End Sub

Private Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = fpCmdMsg.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      fpCmdMsg.ForeColor = &H80000012
      fpCmdMsg.FontSize = BtnFnt
    Case 2
      fpCmdMsg.ForeColor = &H80000011
      fpCmdMsg.FontSize = BtnFnt - 0.7
    Case 3
      fpCmdMsg.ForeColor = &H80000011
      fpCmdMsg.FontSize = BtnFnt - 1.4
    Case 4
      fpCmdMsg.ForeColor = &H80000010
      fpCmdMsg.FontSize = BtnFnt - 2.1
    Case 5
      fpCmdMsg.ForeColor = &H80000010
      fpCmdMsg.FontSize = BtnFnt - 2.8
    Case 6
      fpCmdMsg.ForeColor = &H8000000F
      fpCmdMsg.FontSize = BtnFnt - 3.5
    Case 7
      fpCmdMsg.ForeColor = &H8000000F
      fpCmdMsg.FontSize = BtnFnt - 4.2
    Case 8
      fpCmdMsg.ForeColor = &H8000000E
      fpCmdMsg.FontSize = BtnFnt - 4.9
    Case 9
      fpCmdMsg.ForeColor = &H8000000E
      fpCmdMsg.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
'  DoEvents
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
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF4:
      KeyCode = 0
      DoEvents
      Call fpCmdTranHist_Click
    Case vbKeyF7:
      KeyCode = 0
      DoEvents
      Call fpCmdMsg_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call fpCmdSave_Click
    Case Else:
  End Select
End Sub


Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  UBLog " IN: UTIL DepCredRem"
'  If InStr(TownName$, "CLEVELAND") Then
'    CleveFlag = True
'  End If
'Change this so only true if really cleveland
'decide later on what to do.
 'CleveFlag = True

  LoadRevs
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Sub LoadRevs()
  Dim NumOfRevs As Integer, UBSetupLen As Integer, RevCnt As Integer
  Dim InvRev As Integer
  NumOfRevs = MaxRevsCnt

  ReDim RevText$(1 To MaxRevsCnt)

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    If Len(RevText$(RevCnt)) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    End If
  Next

  If NumOfRevs < MaxRevsCnt Then
    ReDim Preserve RevText$(1 To NumOfRevs)
  End If

  For RevCnt = 1 To NumOfRevs
    fpRevSource(RevCnt - 1) = RevText$(RevCnt)
  Next
  For InvRev = NumOfRevs To 14
    fpRevSource(InvRev).Enabled = False
    fpRevSource(InvRev).Visible = False
    fpAmount(InvRev).Enabled = False
    fpAmount(InvRev).Visible = False
    fpCurrent(InvRev).Enabled = False
    fpCurrent(InvRev).Visible = False
    fpActual(InvRev).Enabled = False
    fpActual(InvRev).Visible = False
  Next
  
End Sub

Private Sub loadCustrec()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile As Integer, cnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim NumOfRevs As Integer, RevCnt As Integer
  NumOfRevs = MaxRevsCnt
  UBCustRecLen = Len(UBCustRec(1))

'  If uselook = True Then
'    Unload frmCustEditLookUP
'    Unload frmDisplayList
'    uselook = False
'  End If
  CustAcct = fpCustRecNo
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
'  If CustAcct& > NumOfCustRecs& Or CustAcct& <= 0 Then
'    CustAcct& = 0
'    LabelDel.Visible = True
'    GoTo SkiptoHere
'  End If

'  If IsDeleted(CustAcct&) Then
'    CustAcct& = 0
'    LabelDel.Caption = "Deleted Account!"
'    LabelDel.Visible = True
'    GoTo SkiptoHere
'  End If
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  Close CustFile
  fptxtAccount = Str$(CustAcct&)
  fpCustName = UBCustRec(1).CustName
  CustAcct& = Val(CustAcct)
  
  For RevCnt = 0 To NumOfRevs - 1
    If fpCurrent(RevCnt).Enabled = True Then
      fpCurrent(RevCnt) = UBCustRec(1).CurrRevAmts(RevCnt + 1)
    End If
  Next
  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  fpBalance = TotalBalance#
  If CustHasMsg(CustAcct) Then
    MsgAlertTimer.Enabled = True
  End If
  Autodist
  CalcBALFlds
  Exit Sub
SkiptoHere:
  
End Sub
'Private Sub ClearScn()
'  Dim cnt As Integer
'  For cnt = 1 To 15
'    fpAmount(cnt - 1) = 0
'  Next
'  CalcCashFlds
'End Sub
Private Sub CalcBALFlds()
  Dim TAmt As Double, cnt As Integer, TAct As Double, TCur As Double
  TAct# = 0
  For cnt = 1 To MaxRevsCnt
    TAmt# = Round#(TAmt# + fpAmount(cnt - 1).DoubleValue)
    'TCur# = Round#(TCur# + fpCurrent(cnt - 1).DoubleValue)
    'fpActual(cnt - 1) = Round#(fpCurrent(cnt - 1).DoubleValue - fpAmount(cnt - 1).DoubleValue)
    fpActual(cnt - 1) = 0
    TAct# = Round#(TAct# + fpActual(cnt - 1).DoubleValue)
  Next
  fpTotActual = TAct#
  fpTotAdjust = TAmt#
  'fix other totals
End Sub
Private Sub Autodist()
  Dim cnt As Integer, ThisAmt As Double, UBTransRecLen As Integer
  Dim NumOfRevs As Integer, WhatRev As Integer, UBTran As Integer
  Dim CustFile As Integer, UBCustRecLen As Integer, ThisTran As Long
  Dim DZCnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType

  NumOfRevs = MaxRevsCnt
  
  'If Not CleveFlag Then
  For cnt = 1 To NumOfRevs
    WhatRev = cnt - 1
    ThisAmt# = Val(fpCurrent(WhatRev))
    If ThisAmt# <> 0 Then
      If ThisAmt# < 0 Then
        fpAmount(WhatRev) = Abs(ThisAmt#) 'Round(ThisAmt# - (ThisAmt# - ThisAmt#))
      Else
        fpAmount(WhatRev) = -ThisAmt# 'Round(ThisAmt# - (ThisAmt# - ThisAmt#))
      End If
    Else
      fpAmount(WhatRev) = 0
    End If

  Next

'End If
'If CleveFlag Then
'  UBCustRecLen = Len(UBCustRec(1))
'  CustFile = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
'  Get CustFile, CustAcct&, UBCustRec(1)
'  Close CustFile
'
'  ReDim DepRev(1 To 15) As Double
'  UBTran = FreeFile
'  ReDim UBTempDepTran(1) As UBTransRecType
'  UBTransRecLen = Len(UBTempDepTran(1))
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
'
'  ThisTran& = UBCustRec(1).LastTrans
'  Do While ThisTran& > 0
'    Get UBTran, ThisTran&, UBTempDepTran(1)
'    Select Case UBTempDepTran(1).TransType
'    Case TranDepositPayment, TranDepositPayment + 100
'      For DZCnt = 1 To 15
'        DepRev(DZCnt) = Round#(DepRev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
'      Next
'
'    Case TranAppliedDeposit, TranRefundDeposit
'      For DZCnt = 1 To 15
'        DepRev(DZCnt) = Round#(DepRev(DZCnt) - Abs(UBTempDepTran(1).RevAmt(DZCnt)))
'      Next
'    End Select
'    ThisTran& = UBTempDepTran(1).PrevTrans
'     Loop
'
'  Close UBTran
'
'   '   If CleveFlag Then 'And NotDone
'        'NotDone = False
'        For DZCnt = 1 To 15
'          'DepRev(DZCnt) = Round#(DepRev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
'          fpAmount(DZCnt - 1) = QPTrim$(Str$(DepRev(DZCnt)))
'         ' SaveField CurFlds(DZCnt) + 1, Form$(), Fld(), BadField
'        Next
'  End If
 
 
 End Sub
    
Private Sub SaveTransaction()
  Dim UBTransRecLen As Integer, NextTranRecs As Long
  Dim TransDate As Integer, Transamt As Double, CustChCnt As Integer
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile  As Integer, cnt As Integer, RevCnt As Integer
  Dim UBTran As Integer, NumOfTranRecs As Long, PrevLastTrans As Long
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTransRec(1) As UBTransRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBTransRecLen = Len(UBTransRec(1))

  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  Close CustFile

  TransDate = Date2Num(txtDate)
  Transamt# = fpTotAdjust.DoubleValue

  UBTransRec(1).TransDate = TransDate
   'UBTransRec(1)CustLocation = CustAcct&
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = CustAcct&
  UBTransRec(1).Transamt = Transamt#
  For cnt = 1 To 15
    If fpAmount(cnt - 1).Enabled = True Then
      If Len(fpAmount(cnt - 1)) > 0 Then
        UBTransRec(1).RevAmt(cnt) = fpAmount(cnt - 1).DoubleValue
      Else
        UBTransRec(1).RevAmt(cnt) = 0
      End If
    End If
  Next

  UBTransRec(1).TransDesc = "Deposit Credit Removal"
  For RevCnt = 1 To 15
    If fpActual(RevCnt - 1).Enabled = True Then
      UBCustRec(1).CurrRevAmts(RevCnt) = Round#(fpActual(RevCnt - 1).DoubleValue)
    End If
  Next

  UBCustRec(1).CurrBalance = 0 'Round#(fpTotActual.DoubleValue)
'  Select Case UBCustRec(1).PrevBalance
'  Case 0
'    'don't do anything
'  Case Is > 0
'    If UBCustRec(1).PrevBalance < fpTotAdjust.DoubleValue Then
      UBCustRec(1).PrevBalance = 0
'    Else
'      UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - fpTotAdjust.DoubleValue)
'      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - UBCustRec(1).PrevBalance)
'    End If
'  Case Is < 0
'    UBCustRec(1).PrevBalance = 0
'  End Select

  UBTransRec(1).TransType = TranDepCreditRemoval
  'Do not create trans to interface to gl -set as already interfaced....
  UBTransRec(1).Posted2GL = "Y"
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  UBCustRec(1).DepositAmt = 0

  CustFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&
  UBTransRec(1).BillMsg = QPTrim(fptxtNote.Text)
  If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 Then
    If UBCustRec(1).Status = "B" Then
      CustChCnt = CustChCnt + 1
      UBLog "DEP Credit Removal: SET CUST STATUS to I. Acct:" + Str$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).Status = "I"
    End If
  End If

  Put CustFile, UBTransRec(1).CustAcctNo, UBCustRec(1)
  Put UBTran, NextTranRecs&, UBTransRec(1)
  Close UBTran, CustFile
  MsgBox "Save procedure complete.", vbOKOnly, "Completed"

End Sub

Private Sub CheckApplyInfo()
  Dim TestDate As Integer
  CredOKFlag = True
  TestDate = Date2Num(txtDate)
  If TestDate < 0 Then
    CredOKFlag = False
    MsgBox "Invalid Date.", vbOKOnly, "Request Canceled."
    GoTo BadApp
  End If
  If fpTotActual.DoubleValue <> 0 Then
    CredOKFlag = False
    MsgBox "Invalid Amount. The Total Balance Should Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadApp
  End If
  Exit Sub
BadApp:
  Exit Sub
End Sub



