VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostBills 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Utility Bills"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   ControlBox      =   0   'False
   Icon            =   "frmPostBills.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer frmBlinkTimer 
      Interval        =   333
      Left            =   8970
      Top             =   5640
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
            TextSave        =   "3:44 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "6/10/2019"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   3840
      TabIndex        =   7
      Top             =   5064
      Width           =   1548
      _Version        =   131072
      _ExtentX        =   2730
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPostBills.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPost 
      Height          =   480
      Left            =   6930
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1560
      _Version        =   131072
      _ExtentX        =   2752
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   1
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPostBills.frx":0AA8
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press Esc to Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   324
      Index           =   6
      Left            =   3120
      TabIndex        =   6
      Top             =   4488
      Width           =   2976
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press F10 to Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   324
      Index           =   5
      Left            =   6192
      TabIndex        =   5
      Top             =   4488
      Width           =   3000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WITH THIS PROCEDURE!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   372
      Index           =   3
      Left            =   4344
      TabIndex        =   4
      Top             =   3648
      Width           =   3552
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UTILITY BILLING PROGRAM BEFORE CONTINUING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   348
      Index           =   2
      Left            =   3120
      TabIndex        =   3
      Top             =   3192
      Width           =   6000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALL UTILITY BILLING OPERATORS MUST EXIT THE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   348
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   2856
      Width           =   6000
   End
   Begin VB.Label LblWarn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING   WARNING  WARNING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   3816
      TabIndex        =   1
      Top             =   2184
      Width           =   4632
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   4116
      Left            =   2832
      Top             =   1920
      Width           =   6564
   End
End
Attribute VB_Name = "frmPostBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Fflag As Boolean
Dim DistArray() As DistArrayType

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF10:
      KeyCode = 0
      Call fpCmdPost_Click
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case Else:
  End Select
End Sub

Private Sub fpCmdExit_Click()
  DoEvents
  Unload frmPostBills
End Sub
Public Sub setstuff(F As Boolean)
  Fflag = F
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  If Fflag Then
    frmPostBills.Caption = "Final Bill Posting"
  End If
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

Private Sub fpCmdPost_Click()
  frmBlinkTimer.Enabled = False
  LblWarn.Visible = False
  DeActivateControls Me
  If Fflag Then
    PostFBillTrans
  Else
    PostBillTrans
  End If
  ActivateControls Me
  MsgBox "Posting Complete.", vbOKOnly, "Complete"
  fpCmdExit_Click
End Sub

Private Sub frmBlinkTimer_Timer()
  Static tog As Boolean
  tog = Not tog
  If tog Then
    LblWarn.Visible = False
  Else
    LblWarn.Visible = True
  End If

End Sub

Private Sub PostBillTrans() 'Normal Bill Post
  Dim UBSetupLen As Integer, CycleFlag As Boolean, IndianFlag As Boolean
  Dim SedgeFlag As Boolean, UBBillRecLen As Integer, UBCustRecLen As Integer
  Dim UBCust As Integer, UBBill As Integer, UBTran As Integer
  Dim NumOfTranRecs As Long, NumOfBillRecs As Long, BillCnt As Long
  Dim PostedCnt As Long, EstFlag As String, MRCnt As Integer
  Dim WhatService As Integer, TestAmt As Double, HowMuch As Double
  Dim FRFlag As Boolean, FRCnt As Integer, RevCnt As Integer
  Dim MtrCnt As Integer, CubMtr As Boolean, ReadAmt As Long
  Dim MaxMeterAmt As Long, TUse As Double, PrevLastTrans As Long
  Dim WhatCycle As Integer, NumOfCust As Long, cnt As Long
  Dim Activated As Long, Wadesboroflag As Boolean, wftemp As String
  UBLog "IN: Bill Posting."
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen      'load setup file

  TOWNNAME$ = UBSetUpRec(1).UTILNAME
  'If UBSetUpRec(1).PreByBook = "N" Then
    CycleFlag = UBSetUpRec(1).BILLCYCL = "Y"
  'End If
  'Section to check for customer modifications
  'Town of Lilesville Special Discount Situation
  wftemp = ""
  If InStr(TOWNNAME$, "INDIAN TRAIL") Then
    IndianFlag = True
  Else
    IndianFlag = False
  End If

  If InStr(TOWNNAME$, "SEDGEFIELD") Then
    SedgeFlag = True
  Else
    SedgeFlag = False
  End If
  If InStr(TOWNNAME$, "WADESBORO") Then
    Wadesboroflag = True
  Else
    Wadesboroflag = False
  End If
'  If FileSize&("UBSNDEM.DAT") > 0 Then
'    For cnt = 1 To 3
'      QPSound 1750, 2
'      QPSound 1650, 2
'    Next
'  End If
  FrmShowPctComp.Label1 = "Posting Bill Transactions"
  FrmShowPctComp.Show , Me

  UBLog "START: Posting Transactions."

  ReDim UBBillRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType

  UBBillRecLen = Len(UBBillRec(1))
  UBCustRecLen = Len(UBCustRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBBill = FreeFile
  Open UBPath$ + UBBillsFile For Random Shared As UBBill Len = UBBillRecLen
  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBBillRecLen

  NumOfTranRecs& = LOF(UBTran) \ UBBillRecLen
  NumOfBillRecs = LOF(UBBill) \ UBBillRecLen

  'ShowProcessingScrn "Posting Billing Transactions"

  If CycleFlag Then
    GoSub GetWhatCycle
  End If
'  For BillCnt = 1 To NumOfCust
'
'      If UBBillRec(1).CustAcctNo > 0 Then
'        If UBBillRec(1).CustAcctNo = 1309 Then Stop
'        Get UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
'        Select Case UBCustRec(1).Status
'        Case "B", "I", "D"
'          If UBCustRec(1).CurrBalance <> 0 Then
'            UBCustRec(1).PrevBalance = Round(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'            UBCustRec(1).CurrBalance = 0
'            Put UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
'          End If
'        End Select
'      End If
'

  For BillCnt = 1 To NumOfBillRecs
    FrmShowPctComp.ShowPctComp BillCnt, NumOfBillRecs

    Get UBBill, BillCnt, UBBillRec(1)
    If (UBBillRec(1).ActiveFlag And UBBillRec(1).Transamt > 0) Or (UBBillRec(1).NONProfit = "Y") Then
      PostedCnt& = PostedCnt& + 1
      NumOfTranRecs& = NumOfTranRecs& + 1       'point to next trans to write
      Get UBCust, BillCnt, UBCustRec(1)
      EstFlag$ = QPTrim$(UBCustRec(1).EstFlag)
      For MRCnt = 1 To 2
        WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
        If UBCustRec(1).Monthly(MRCnt).PayAmt > 0 And WhatService > 0 Then
          TestAmt# = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + UBCustRec(1).Monthly(MRCnt).PayAmt)
          If TestAmt# > UBCustRec(1).Monthly(MRCnt).AMTOWED Then
            HowMuch# = Round#(UBCustRec(1).Monthly(MRCnt).AMTOWED - UBCustRec(1).Monthly(MRCnt).TotAmtPD)
          Else
            HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
          End If
          UBCustRec(1).Monthly(MRCnt).TotAmtPD = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + HowMuch#)
        End If
      Next
      '062597 added removal of nonrecurring flat rates
      FRFlag = False
      For FRCnt = 1 To 4        'Remove non-recurring flat rates
        If UBCustRec(1).FlatRates(FRCnt).FRFREQ = "N" Then
          UBCustRec(1).FlatRates(FRCnt).FRDESC = ""
          UBCustRec(1).FlatRates(FRCnt).FRAMT = 0
          UBCustRec(1).FlatRates(FRCnt).FRFREQ = ""
          UBCustRec(1).FlatRates(FRCnt).REVSRC = 0
          UBCustRec(1).FlatRates(FRCnt).NumMin = 0
          FRFlag = True
        End If
      Next
      If FRFlag Then
        UBLog "BILL POST: Removed Flat Rate. Acct:" + Str$(BillCnt)
      End If
      '111698 Prorate
      If UBBillRec(1).ProRatePCT < 100 Then
        UBLog "BILL POST: Reset Prorate Acct:" + Str$(BillCnt) + " PCT:" + Str$(UBBillRec(1).ProRatePCT)
      End If
      UBCustRec(1).ProRatePCT = 100
      '*************
      UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBCustRec(1).CurrBalance = UBBillRec(1).Transamt
      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      For RevCnt = 1 To MaxRevsCnt
        UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      Next
      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
      UBBillRec(1).TransDesc = "Utility Billing"
      UBBillRec(1).TransDate = UBBillRec(1).BillDate
      UBBillRec(1).OperatorNumber = OPERNUM
        
      For MtrCnt = 1 To 7
        CubMtr = False
        If UBCustRec(1).LocMeters(MtrCnt).CurRead >= 0 Then
          If Len(EstFlag$) > 0 Then
            UBBillRec(1).ESTREAD(MtrCnt) = "Y"
          End If
          If UBCustRec(1).LocMeters(MtrCnt).MtrUnit = "C" Then
            CubMtr = True
          End If
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          If ReadAmt& < 0 Then  'Meter rolled over or, been misread
            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          End If
          If CubMtr Then
            ReadAmt& = ReadAmt& * 7.481
          End If
          If ReadAmt& < 1 Then
            ReadAmt& = 1
          End If
          If UBCustRec(1).LocMeters(MtrCnt).AvgUse < 1 Then
            UBCustRec(1).LocMeters(MtrCnt).AvgUse = 1
          End If
  'do it this way to keep first time cust have correct usecnt and avg
          If UBCustRec(1).LocMeters(MtrCnt).UseCnt < 1 Then
            UBCustRec(1).LocMeters(MtrCnt).UseCnt = 1
            TUse# = ReadAmt& + (UBCustRec(1).LocMeters(MtrCnt).AvgUse * UBCustRec(1).LocMeters(MtrCnt).UseCnt)
            UBCustRec(1).LocMeters(MtrCnt).AvgUse = TUse# / UBCustRec(1).LocMeters(MtrCnt).UseCnt
            UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
          Else
            TUse# = ReadAmt& + (UBCustRec(1).LocMeters(MtrCnt).AvgUse * UBCustRec(1).LocMeters(MtrCnt).UseCnt)
            '''TUse# = 0
            UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
            UBCustRec(1).LocMeters(MtrCnt).AvgUse = TUse# / UBCustRec(1).LocMeters(MtrCnt).UseCnt
            UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
          End If
          'so all reads will end in 0
          If Wadesboroflag Then
            wftemp$ = ""
            wftemp$ = Str(UBCustRec(1).LocMeters(MtrCnt).AvgUse)
            If Len(wftemp$) > 0 Then
              wftemp$ = Left$(wftemp$, (Len(wftemp$) - 1))
              wftemp$ = wftemp$ + "0"
              UBCustRec(1).LocMeters(MtrCnt).AvgUse = Val(wftemp$)
            End If
          End If
          If SedgeFlag Then
            UBCustRec(1).LocMeters(MtrCnt).CurRead = 0
            UBCustRec(1).LocMeters(MtrCnt).PrevRead = 0
            UBCustRec(1).LocMeters(MtrCnt).AvgUse = 0
          End If
        End If
      Next
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      UBCustRec(1).LastTrans = NumOfTranRecs&
      UBCustRec(1).EstFlag = ""
      If IndianFlag Then
        UBCustRec(1).USERCODE1 = ""
      End If

      Put UBCust, BillCnt, UBCustRec(1)
      Put UBTran, NumOfTranRecs&, UBBillRec(1)
      '**************
    End If
    'ShowPctComp BillCnt, NumOfBillRecs
  Next
  Close
  UBLog "  DONE: Posting Transactions."
  UBLog "POSTED:" + Str$(PostedCnt&) + " New BILL Transactions."
  'DALE
  KillFile UBPath$ + UBBillsFile
  KillFile UBPath$ + "UBBILLS.PRN"
  '**************
  UBLog "KILLED: UBBILLS.DAT & UBBILLS.PRN"

  'ShowProcessingScrn "Activating Pending Accounts."

  UBLog "ACTIVATING ACCOUNTS:"
  If CycleFlag Then
    UBLog " CYCLE:" + Str$(WhatCycle)
  End If
  FrmShowPctComp.Label1 = "Updating Customer Accounts"
  FrmShowPctComp.Show , Me

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCust& = LOF(UBCust) / UBCustRecLen

  For cnt = 1 To NumOfCust&
    Get UBCust, cnt, UBCustRec(1)

'040803 add to set avtive flags only on current cycle
    FrmShowPctComp.ShowPctComp cnt, NumOfCust&

    If CycleFlag Then
      If UBCustRec(1).BILLCYCL <> WhatCycle Then
        GoTo NotThisCycle
      End If
    End If
    If UBCustRec(1).Status = "P" Then
      UBCustRec(1).Status = "A"
      For MtrCnt = 1 To 7
        UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
      Next
      UBLog "ACTIVATED: ,ReadFlagCleared" + Str$(cnt) + "  " + UBCustRec(1).CustName
      Activated = Activated + 1
      Put UBCust, cnt, UBCustRec(1)
    End If
    If UBCustRec(1).Status = "I" Then
      For MtrCnt = 1 To 7
        UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
      Next
      Put UBCust, cnt, UBCustRec(1)
      UBLog "ReadFlagCleared: " + Str$(cnt) + "  " + UBCustRec(1).CustName
    End If
NotThisCycle:
    'ShowPctComp cnt, CInt(NumOfCust&)
  Next

  Close
  UBLog "     DONE: Activating Accounts."
  UBLog "ACTIVATED:" + Str$(Activated) + " Pending Accounts."
  'BlockClear
  'DisplayUBScrn "UPDATEOK"
 ' WaitForAction

ExitBillPost:
  UBLog "OUT: Bill Posting." + CrLf$
Exit Sub


GetWhatCycle:
  WhatCycle = 0
  For BillCnt = 1 To NumOfBillRecs
    Get UBBill, BillCnt, UBBillRec(1)
    If (UBBillRec(1).ActiveFlag And UBBillRec(1).Transamt > 0) Or (UBBillRec(1).NONProfit = "Y") Then
      Get UBCust, BillCnt, UBCustRec(1)
      WhatCycle = UBCustRec(1).BILLCYCL
      Exit For
    End If
  Next
Return

End Sub

Private Sub PostFBillTrans() 'Final
  Dim UBSetupLen As Integer, CycleFlag As Boolean, IndianFlag As Boolean
  Dim SedgeFlag As Boolean, UBBillRecLen As Integer, UBCustRecLen As Integer
  Dim UBCust As Integer, UBBill As Integer, UBTran As Integer
  Dim NumOfTranRecs As Long, NumOfBillRecs As Long, BillCnt As Long
  Dim PostedCnt As Long, EstFlag As String, MRCnt As Integer
  Dim WhatService As Integer, TestAmt As Double, HowMuch As Double
  Dim FRFlag As Boolean, FRCnt As Integer, RevCnt As Integer
  Dim MtrCnt As Integer, CubMtr As Boolean, ReadAmt As Long
  Dim MaxMeterAmt As Long, TUse As Double, PrevLastTrans As Long
  Dim WhatCycle As Integer, NumOfCust As Long, cnt As Long
  Dim Activated As Long, CleveFlag As Boolean, NextTransRec As Long
  Dim DepAppliedFlag As Boolean, UBTransRecLen As Integer
  Dim DepTranAmt As Double, CustChCnt As Long, LLCnt As Integer
  Dim DepositAmt As Double, ThisTran As Long, DZCnt As Integer
  Dim Wadesboroflag As Boolean, wftemp As String, TempDepAmt#, ThisAmt#
     Dim LugoffFlag As Boolean
 
  UBLog "IN: POST FINAL"

  ReDim deprev(1 To 15) As Double
  ReDim DepRevKept(1 To 15) As Double
  ReDim UBTempDepTran(1) As UBTransRecType

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

'STOP
'   If InStr(TOWNNAME$, "LUGOFF") Then
'    LugoffFlag = True
'  Else
'    LugoffFlag = False
'  End If

  If InStr(UBSetUpRec(1).UTILNAME, "CLEVELAND") Then
    CleveFlag = True
    UBLog "POST FINAL:  CLEVELAND Detected "
  End If
  If InStr(TOWNNAME$, "WADESBORO") Then
    Wadesboroflag = True
  Else
    Wadesboroflag = False
  End If

  ReDim UBBillRec(1) As UBTransRecType
  ReDim UBCustRec(1 To 2) As NewUBCustRecType

  UBBillRecLen = Len(UBBillRec(1))
  UBCustRecLen = Len(UBCustRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCust = LOF(UBCust) / UBCustRecLen
  UBBill = FreeFile
  Open UBPath$ + UBFinBillsFile For Random Shared As UBBill Len = UBBillRecLen

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBBillRecLen

  NumOfBillRecs = LOF(UBBill) / UBBillRecLen
  'ShowProcessingScrn "Posting Final Billing Transactions"
  '*****************************************************
  FrmShowPctComp.Label1 = "Posting Final Billing Transactions"
  FrmShowPctComp.Show , Me
      
  For BillCnt = 1 To NumOfBillRecs
    FrmShowPctComp.ShowPctComp BillCnt, NumOfBillRecs
    Get UBBill, BillCnt, UBBillRec(1)
    If UBBillRec(1).ActiveFlag Then             'AND UBBillRec(1).TransAmt > 0 then
      Get UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
      UBCustRec(1).Status = "B"
      UBCustRec(1).PrevBalance = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBCustRec(1).CurrBalance = UBBillRec(1).Transamt
      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)

      If UBBillRec(1).ApplyDepFlag = "Y" Then
        If CleveFlag Then
          GoSub ProcCleveDeposit
        Else
          GoSub ProcCustDeposit
        End If
      Else
        For RevCnt = 1 To MaxRevsCnt
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
        Next
      End If
      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
      For MtrCnt = 1 To 7
        CubMtr = False
        If UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 Then
          If UBCustRec(1).LocMeters(MtrCnt).MtrUnit = "C" Then
            CubMtr = True
          End If
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          If ReadAmt& < 0 Then  'Meter rolled over or, misread
            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          End If
          If CubMtr Then
            ReadAmt& = ReadAmt& * 7.481
          End If
          UBCustRec(1).LocMeters(MtrCnt).AvgUse = Round(UBCustRec(1).LocMeters(MtrCnt).AvgUse + ReadAmt&)
          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
          '050697 Fixed current reading not being move to previous
          UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
          'to give all reads, averages etc 0 on the end
          If Wadesboroflag Then
            wftemp$ = ""
            wftemp$ = Str(UBCustRec(1).LocMeters(MtrCnt).AvgUse)
            If Len(wftemp$) > 0 Then
              wftemp$ = Left$(wftemp$, (Len(wftemp$) - 1))
              wftemp$ = wftemp$ + "0"
              UBCustRec(1).LocMeters(MtrCnt).AvgUse = Val(wftemp$)
            End If
          End If
        End If
      Next

      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      NextTransRec& = (LOF(UBTran) \ UBBillRecLen) + 1          'point at next
      Put UBTran, NextTransRec&, UBBillRec(1)

      UBCustRec(1).LastTrans = NextTransRec&

      'detach the new vacant rec from this customer
      If UBCustRec(1).OldRec > 0 Then
        UBLog "POST FINAL: DETACHED OLD ACCT: " + Str$(UBBillRec(1).CustAcctNo) + " from NEW:" + Str$(UBCustRec(1).OldRec)
        UBCustRec(1).OldRec = 0
      End If
      UBCustRec(1).EstFlag = ""
      Put UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)

      '040997 added Transaction to show customers applied deposit
      If DepAppliedFlag Then
        GoSub MakeAppDepTrans
      End If
    Else
'      If UBBillRec(1).CustAcctNo > 0 Then
'        If UBBillRec(1).CustAcctNo = 1309 Then Stop
'        Get UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
'        Select Case UBCustRec(1).Status
'        Case "B", "I", "D"
'          If UBCustRec(1).CurrBalance <> 0 Then
'            UBCustRec(1).PrevBalance = Round(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'            UBCustRec(1).CurrBalance = 0
'            Put UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
'          End If
'        End Select
'      End If
    End If
    'ShowPctComp BillCnt, NumOfBillRecs
  Next

  Close
  KillFile UBPath$ + UBFinBillsFile
  KillFile UBPath$ + "UBFBILLS.PRN"


ExitBillPost:
  UBLog "OUT: POST FINAL"
  Exit Sub

MakeAppDepTrans:
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  UBTransRec(1).TransDate = UBBillRec(1).TransDate
  'UBTransRec(1)CustLocation = UBBillRec(1).CustAcctNo
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = UBBillRec(1).CustAcctNo
  UBTransRec(1).Transamt = DepTranAmt#
  '091198 Changed to put original deposit amounts in revenue source
  For cnt = 1 To 15
    UBTransRec(1).RevAmt(cnt) = DepRevKept(cnt)
  Next
  UBTransRec(1).OperatorNumber = OPERNUM
  UBTransRec(1).TransDesc = "Applied Deposit"
  UBTransRec(1).TransType = TranAppliedDeposit
  UBTransRec(1).RunBalance = Round#((UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance) - Abs(DepTranAmt#))
  UBCustRec(1).DepositAmt = 0
  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - Abs(DepTranAmt#))

  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&

  If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 Then
    If UBCustRec(1).Status = "B" Then
      CustChCnt = CustChCnt + 1
      UBLog "POST FINAL: SET CUST STATUS=I ACCT:" + Str$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).Status = "I"
    End If
  End If
  NextTransRec& = (LOF(UBTran) \ UBTransRecLen) + 1             'point at next
  Put UBTran, NextTransRec&, UBTransRec(1)

  UBCustRec(1).LastTrans = NextTransRec&

  Put UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)

  UBLog "POST FINAL: DEP APPLIED TRANS:" + Str$(NextTransRec&)

Return

ProcCleveDeposit:
  For LLCnt = 1 To 15
    deprev(LLCnt) = 0
  Next LLCnt

    DepAppliedFlag = False
    DepTranAmt# = -UBCustRec(1).DepositAmt
    DepositAmt# = UBCustRec(1).DepositAmt
    If DepositAmt# = 0 Then
      For RevCnt = 1 To MaxRevsCnt
        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      Next
      GoTo NoDepReturn
    End If

    ThisTran& = UBCustRec(1).LastTrans
    Do While ThisTran& > 0
      Get UBTran, ThisTran&, UBTempDepTran(1)
      If UBTempDepTran(1).TransType = TranDepositPayment Then
        For DZCnt = 1 To 15
          deprev(DZCnt) = Round#(deprev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
'added ???????
          DepRevKept(DZCnt) = deprev(DZCnt)
        Next
      End If
      ThisTran& = UBTempDepTran(1).PrevTrans
    Loop

    For RevCnt = 1 To MaxRevsCnt - 1
      UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      If deprev(RevCnt) > 0 Then
        DepAppliedFlag = True
        If UBCustRec(1).CurrRevAmts(RevCnt) < deprev(RevCnt) Then
          deprev(RevCnt) = Round#(deprev(RevCnt) - UBCustRec(1).CurrRevAmts(RevCnt))
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
        ElseIf UBCustRec(1).CurrRevAmts(RevCnt) > deprev(RevCnt) Then
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - deprev(RevCnt))
          deprev(RevCnt) = 0
        Else    'the deposit and the revenue are equal
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          deprev(RevCnt) = 0
        End If
      End If
    Next

    'If there was any deposit left after applying to the cust rev totals
    For RevCnt = 1 To MaxRevsCnt - 1
      If deprev(RevCnt) > 0 Then
        UBCustRec(1).CurrRevAmts(RevCnt) = -deprev(RevCnt)
      End If
    Next
    UBCustRec(1).DepositAmt = 0

NoDepReturn:

    Return
ProcCustDeposit:
  LoadRevs

    For LLCnt = 1 To 15
      deprev(LLCnt) = 0
    Next
    For cnt = 1 To MaxRevsCnt - 1
      DepRevKept(cnt) = 0
    Next
    TempDepAmt# = 0

    DepAppliedFlag = False
    DepTranAmt# = -UBCustRec(1).DepositAmt
    DepositAmt# = UBCustRec(1).DepositAmt
    If DepositAmt# = 0 Then
'051799 added to correct rev problem with accounts that had no deposit but
'       apply deposit to final bill was selected
      For RevCnt = 1 To MaxRevsCnt
        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      Next
      GoTo NoDepReturn
    End If
    Dim tempTot As Double
    tempTot = 0
  
    ThisTran& = UBCustRec(1).LastTrans
    Do While ThisTran& > 0
      Get UBTran, ThisTran&, UBTempDepTran(1)
      If UBTempDepTran(1).TransType = TranDepositPayment Then
        For DZCnt = 1 To 15
          deprev(DZCnt) = Round#(deprev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
          tempTot = tempTot + UBTempDepTran(1).RevAmt(DZCnt)
        Next
      ElseIf (UBTempDepTran(1).TransType = TranAppliedDeposit) Or (UBTempDepTran(1).TransType = TranRefundDeposit) Or (UBTempDepTran(1).TransType = TranDepPaymentVoid) Then
        For DZCnt = 1 To 15
          deprev(DZCnt) = Round#(deprev(DZCnt) - UBTempDepTran(1).RevAmt(DZCnt))
          tempTot = tempTot - UBTempDepTran(1).RevAmt(DZCnt)
        Next
      End If
      ThisTran& = UBTempDepTran(1).PrevTrans
    Loop
    
    TempDepAmt# = DepositAmt#
    If TempDepAmt# > 0 Then DepAppliedFlag = True
    For RevCnt = 1 To MaxRevsCnt - 1
      ThisAmt# = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      If LugoffFlag = True Then
        If tempTot > 0 Then
        If UBCustRec(1).CurrRevAmts(RevCnt) < tempTot Then
          tempTot = Round#(tempTot - UBCustRec(1).CurrRevAmts(RevCnt))
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          TempDepAmt# = Round#(TempDepAmt# - ThisAmt#)
          DepRevKept(RevCnt) = ThisAmt#
        ElseIf UBCustRec(1).CurrRevAmts(RevCnt) > tempTot Then
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - tempTot)
          TempDepAmt# = Round#(TempDepAmt# - tempTot)
          DepRevKept(RevCnt) = tempTot
          ThisAmt# = TempDepAmt#
          tempTot = 0
        Else    'the deposit and the revenue are equal
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          TempDepAmt# = Round#(TempDepAmt# - tempTot)
          DepRevKept(RevCnt) = tempTot
          tempTot = 0
        End If
        End If
      Else
      If deprev(RevCnt) > 0 Then
        If UBCustRec(1).CurrRevAmts(RevCnt) < deprev(RevCnt) Then
          deprev(RevCnt) = Round#(deprev(RevCnt) - UBCustRec(1).CurrRevAmts(RevCnt))
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          TempDepAmt# = Round#(TempDepAmt# - ThisAmt#)
          DepRevKept(RevCnt) = ThisAmt#
        ElseIf UBCustRec(1).CurrRevAmts(RevCnt) > deprev(RevCnt) Then
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - deprev(RevCnt))
          TempDepAmt# = Round#(TempDepAmt# - deprev(RevCnt))
          ThisAmt# = TempDepAmt#
          DepRevKept(RevCnt) = deprev(RevCnt)
          deprev(RevCnt) = 0
        Else    'the deposit and the revenue are equal
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          TempDepAmt# = Round#(TempDepAmt# - deprev(RevCnt))
          DepRevKept(RevCnt) = deprev(RevCnt)
          deprev(RevCnt) = 0
        End If
      End If
      End If
    Next
    
    If TempDepAmt# > 0 Then
    For cnt = 1 To MaxRevsCnt - 1
      RevCnt = DistArray(cnt).DistCnt
      ThisAmt# = UBCustRec(1).CurrRevAmts(RevCnt)
      'UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      If ThisAmt# > 0 Then
        If TempDepAmt# >= ThisAmt# Then
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          TempDepAmt# = Round#(TempDepAmt# - ThisAmt#)
          DepRevKept(RevCnt) = Round(DepRevKept(RevCnt) + ThisAmt#)
        Else
          ThisAmt# = TempDepAmt#
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - TempDepAmt#)
          DepRevKept(RevCnt) = Round(DepRevKept(RevCnt) + TempDepAmt#)
          TempDepAmt# = 0
        End If
      ElseIf TempDepAmt# = 0 Then
       'UBCustRec(1).CurrRevAmts(RevCnt) = 0
       DepRevKept(RevCnt) = Round(DepRevKept(RevCnt) + 0)
      End If
    Next
    End If
    
     If TempDepAmt# > 0 Then
        RevCnt = 1
        DepRevKept(RevCnt) = Round(DepRevKept(RevCnt) + TempDepAmt#)
        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - TempDepAmt#)
     End If
    UBCustRec(1).DepositAmt = 0

Return
'ProcCustDeposit:
'  LoadRevs
'
'    For LLCnt = 1 To 15
'      deprev(LLCnt) = 0
'    Next
'    For cnt = 1 To MaxRevsCnt - 1
'      DepRevKept(cnt) = 0
'    Next
'    TempDepAmt# = 0
'
'    DepAppliedFlag = False
'    DepTranAmt# = -UBCustRec(1).DepositAmt
'    DepositAmt# = UBCustRec(1).DepositAmt
'    If DepositAmt# = 0 Then
''051799 added to correct rev problem with accounts that had no deposit but
''       apply deposit to final bill was selected
'      For RevCnt = 1 To MaxRevsCnt
'        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
'      Next
'      GoTo NoDepReturn
'    End If
'
'
''    ThisTran& = UBCustRec(1).LastTrans
''    Do While ThisTran& > 0
''      Get UBTran, ThisTran&, UBTempDepTran(1)
''      If UBTempDepTran(1).TransType = TranDepositPayment Then
''        For DZCnt = 1 To 15
''          deprev(DZCnt) = Round#(deprev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
''          'DepRevKept(DZCnt) = deprev(DZCnt)
''        Next
''      ElseIf (UBTempDepTran(1).TransType = TranAppliedDeposit) Or (UBTempDepTran(1).TransType = TranRefundDeposit) Or (UBTempDepTran(1).TransType = TranDepPaymentVoid) Then
''        For DZCnt = 1 To 15
''          deprev(DZCnt) = Round#(deprev(DZCnt) - UBTempDepTran(1).RevAmt(DZCnt))
''          'DepRevKept(DZCnt) = deprev(DZCnt)
''        Next
''      End If
''      ThisTran& = UBTempDepTran(1).PrevTrans
''    Loop
'    TempDepAmt# = DepositAmt#
'    If TempDepAmt# > 0 Then DepAppliedFlag = True
'    For cnt = 1 To MaxRevsCnt - 1
'      RevCnt = DistArray(cnt).DistCnt
'      ThisAmt# = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
'      UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
'      If ThisAmt# > 0 Then
'        If TempDepAmt# >= ThisAmt# Then
'          UBCustRec(1).CurrRevAmts(RevCnt) = 0
'          TempDepAmt# = Round#(TempDepAmt# - ThisAmt#)
'          DepRevKept(RevCnt) = ThisAmt#
'        Else
'          ThisAmt# = TempDepAmt#
'          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - TempDepAmt#)
'          DepRevKept(RevCnt) = TempDepAmt#
'          TempDepAmt# = 0
'        End If
'      ElseIf TempDepAmt# = 0 Then
'       'UBCustRec(1).CurrRevAmts(RevCnt) = 0
'       DepRevKept(RevCnt) = 0
'      End If
'
'    Next
'     If TempDepAmt# > 0 Then
'        RevCnt = 1
'        DepRevKept(RevCnt) = Round(DepRevKept(RevCnt) + TempDepAmt#)
'        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - TempDepAmt#)
'     End If
'
'
'    UBCustRec(1).DepositAmt = 0
'
'Return

End Sub
Private Sub LoadRevs()
  Dim NumofRevs As Integer, UBSetupLen As Integer, RevCnt As Integer
  Dim InvRev As Integer, OutOfOrder As Boolean, x As Integer
  Dim tmp As DistArrayType
  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  ReDim Preserve DistArray(1 To NumofRevs) As DistArrayType

  For RevCnt = 1 To MaxRevsCnt
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
  Next
  Do
    OutOfOrder = False          'assume it's sorted
    For x = 1 To MaxRevsCnt - 1
      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
        'SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
        tmp = DistArray(x)
        DistArray(x) = DistArray(x + 1)
        DistArray(x + 1) = tmp
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder

  
End Sub

