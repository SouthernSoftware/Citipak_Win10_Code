VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBBillingMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Billing, Readings, Penalties"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   Icon            =   "frmUBBillingMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLateNotice 
      Caption         =   "&Late Notice Processing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   8
      Top             =   6784
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitUBBillingProcess 
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
      Height          =   492
      Left            =   3846
      TabIndex        =   9
      Top             =   7368
      Width           =   4524
   End
   Begin VB.CommandButton cmdRefundDeposit 
      Caption         =   "Customers &Deposit Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   6
      Top             =   5622
      Width           =   4524
   End
   Begin VB.CommandButton cmdUBAdjustments 
      Caption         =   "Utility Billing &Adjustments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   7
      Top             =   6203
      Width           =   4524
   End
   Begin VB.CommandButton cmdUtilityBillPrinting 
      Caption         =   "&Utility Bill Printing Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   2
      Top             =   3298
      Width           =   4524
   End
   Begin VB.CommandButton cmdPenaltyProcess 
      Caption         =   "Pe&nalty Processing Menu "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   5
      Top             =   5041
      Width           =   4524
   End
   Begin VB.CommandButton cmdBankDraft 
      Caption         =   "&Bank Draft Processing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   4
      Top             =   4460
      Width           =   4524
   End
   Begin VB.CommandButton cmdPostBillingTrans 
      Caption         =   "Po&st Billing Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   3
      Top             =   3879
      Width           =   4524
   End
   Begin VB.CommandButton cmdPreBilling 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Pre-Billing Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   1
      Top             =   2730
      Width           =   4524
   End
   Begin VB.CommandButton cmdMeterReadings 
      BackColor       =   &H008F8265&
      Caption         =   "&Meter Readings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2136
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   8505
      Width           =   12225
      _ExtentX        =   21564
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
            TextSave        =   "1:44 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "9/4/2008"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Utility Billing, Readings, Penalties"
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
      Left            =   3540
      TabIndex        =   10
      Top             =   1104
      Width           =   5148
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
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
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
End
Attribute VB_Name = "frmUBBillingMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class


Private Sub cmdBankDraft_Click()
If LevelPass = 1 Then
  Load frmUBDraftMenu
  DoEvents
  frmUBDraftMenu.Show
  Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdExitUBBillingProcess_Click()
  Load frmUBMainMenu
  DoEvents
  frmUBMainMenu.Show
  Unload frmUBBillingMenu
End Sub

Private Sub cmdLateNotice_Click()
  If LevelPass = 1 Then
  Load frmUBLateNoticeMenu
  DoEvents
  frmUBLateNoticeMenu.Show
  Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdMeterReadings_Click()
  If LevelPass = 1 Then
  Load frmUBMeterMenu
  DoEvents
  frmUBMeterMenu.Show
  Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdPenaltyProcess_Click()
  If LevelPass = 1 Then
  Load frmUBPenaltyMenu
  DoEvents
  frmUBPenaltyMenu.Show
  Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdPostBillingTrans_Click()
  Dim Today As String, chkthedate As Integer, entdate As Integer, UBPFile As Integer
  Dim UBBillSetuplen As Integer, BillInfoRecLen As Integer
  Dim FntSize As Integer, BLType As Integer
  ReDim BillInfoRec(1) As PrintBillInfoType
  If LevelPass = 1 Then
  BillInfoRecLen = Len(BillInfoRec(1))

  ReDim MsgText(0 To 5) As String
 'get Bill type from setup and store integer
  ReDim UBBillSetup(1) As UBBillSetupType
  UBBillSetuplen = Len(UBBillSetup(1))
  LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
  BLType = UBBillSetup(1).Bill
  If Not Exist(UBPath$ + UBBillsFile) Then
    UBLog "ERROR: UBBILLS.DAT Calculation file NOT FOUND!"
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO BILL FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
  If Not BLType = 98 And Not BLType = 97 Then
    If Not Exist(UBPath$ + "UBBILLS.PRN") Then
      UBLog "ERROR: UBBILLS.PRN Print File NOT FOUND!"
      frmMsgDialog.RetLabel = "-2"
      FntSize = frmMsgDialog.Label(3).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
      MsgText(0) = "ERROR:"
      MsgText(1) = ""
      MsgText(2) = "NO BILLS PRINTED!"
      MsgText(3) = ""
      MsgText(4) = ""
      MsgText(5) = ""
      GetOKorNot MsgText(), True
      Exit Sub
    End If
  Else
  'just extra warning
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
    MsgText(0) = "POST BILLS ?"
    MsgText(1) = ""
    MsgText(2) = "ARE YOU SURE"
    MsgText(3) = "YOU WANT TO POST"
    MsgText(4) = "BILLS???"
    MsgText(5) = "OK to continue, or Cancel."
     If GetOKorNot(MsgText()) Then
       UBLog "USER WANTS TO CONTINUE!"
     Else
       Exit Sub
    End If
  End If
  Today = Format(Now, "mm/dd/yyyy")
  chkthedate = Date2Num(Today)
  UBPFile = FreeFile
    Open UBPath$ + "UBPINFON.DAT" For Random As #UBPFile Len = BillInfoRecLen
    Get #UBPFile, 1, BillInfoRec(1)
    Close UBPFile
  entdate = BillInfoRec(1).BillDate
  If entdate > (chkthedate + 30) Or entdate < (chkthedate - 30) Then
    UBLog "Invalid Bill Date Post, give opt to cancel- OPER:" + Str$(OPERNUM)
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:"
    MsgText(1) = ""
    MsgText(2) = "Billing Date is NOT"
    MsgText(3) = "within monthly date range."
    MsgText(4) = ""
    MsgText(5) = "OK to continue, or Cancel."
    If GetOKorNot(MsgText()) Then
      UBLog "Continue Bill post with out of range date -" + Num2Date(entdate)
    Else
      UBLog "Cancel bill post so can check dates."
      Exit Sub
    End If
  End If

  frmPostBills.setstuff False
  'Load frmPostBills
  DoEvents
  frmPostBills.Show
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If

End Sub

Private Sub cmdPreBilling_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If LevelPass = 1 Then
 'Need to do warning here if have not posted prior bills!!!!
   If Exist(UBPath$ + "UBBILLS.DAT") And Exist(UBPath$ + "UBBILLS.PRN") Then
     UBLog "ERROR: UNPOSTED BILLING DETECTED!"
     UBLog "ASKING USER WANT TO CONTINUE?"
     FntSize = frmMsgDialog.Label(3).FontSize
     frmMsgDialog.Label(1).FontSize = (FntSize + 2)
     frmMsgDialog.Label(3).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = "UNPOSTED BILLING DETECTED!"
     MsgText(3) = ""
     MsgText(4) = "Are You Sure You Want To Continue?"
     MsgText(5) = ""
     If GetOKorNot(MsgText()) Then
       UBLog "USER WANTS TO CONTINUE!"
       KillFile (UBPath$ + "UBBILLS.PRN")
     Else
       UBLog "USER ABORTED PREBILLING."
       Exit Sub
    End If
  End If
 
  Load frmPreBilling
  DoEvents
  frmPreBilling.Show
  Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdRefundDeposit_Click()
  If LevelPass = 1 Then
  Load frmUBDepositMenu
  DoEvents
  frmUBDepositMenu.Show
  Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If

End Sub

Private Sub cmdUBAdjustments_Click()
  If LevelAdj = True Then
    Load frmAdjustmentEntry
    DoEvents
    frmAdjustmentEntry.Show
    Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdUtilityBillPrinting_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If LevelPass = 1 Then
  If Not Exist(UBPath$ + "UBBilSet.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO billsetup"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO BILL SETUP FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
    Load frmUBPrintBillsMenu
    DoEvents
    frmUBPrintBillsMenu.Show
    Unload frmUBBillingMenu
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpUtilityBillings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitUBBillingProcess.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via UBBillingMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
    'Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdMeterReadings.SetFocus
    Case vbKeyEnd
      cmdExitUBBillingProcess.SetFocus
    Case Else:
  End Select
End Sub

