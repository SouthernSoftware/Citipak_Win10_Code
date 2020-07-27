VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDCSetupMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decals Setup Maintenance Menu"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmUBDepositMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D0D0D0&
      Caption         =   "ReCalc Running Balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6032
      Width           =   4524
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Convert Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5272
      Width           =   4524
   End
   Begin VB.CommandButton cmdApplicationDef 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Application Letter Defaults"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4512
      Width           =   4524
   End
   Begin VB.CommandButton cmdReindexName 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Re&Index Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3752
      Width           =   4524
   End
   Begin VB.CommandButton cmdRelink 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Relink DC Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3888
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2992
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitMenu 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6792
      Width           =   4524
   End
   Begin VB.CommandButton cmdSetupInfo 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&System Default Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2232
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
            TextSave        =   "10:28 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "7/27/2005"
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
      Caption         =   "Decal Setup Maintenance Menu"
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
      Left            =   3540
      TabIndex        =   8
      Top             =   1104
      Width           =   5148
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
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
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
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
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmDCSetupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class


Private Sub cmdReindexName_Click()
  SortDCNameIndex frmDCSetupMenu
End Sub

Private Sub cmdApplicationDef_Click()
  Load frmApplicationLetter
  DoEvents
  frmApplicationLetter.Show
  Unload Me
End Sub

Private Sub cmdExitMenu_Click()
  Load frmDCMainMenu
  DoEvents
  frmDCMainMenu.Show
  Unload Me
End Sub


Private Sub cmdRelink_Click()
  RelinkDCStuff frmDCSetupMenu
End Sub

Private Sub cmdSetupInfo_Click()
  Load frmSystemSetup
  DoEvents
  frmSystemSetup.Show
  Unload Me
End Sub

'Private Sub cmdReprint_Click()
'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'  If Not Exist("UBFBILLS.PRN") Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO PRN FILE. Reprint Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL PRINT FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'
'  If Not Exist(UBFinBillsFile) Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO BILL FILE! Reprint Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'  frmBillPrinting.REPRN True, True
'  Load frmBillPrinting
'  DoEvents
'  frmBillPrinting.Show
'  Unload frmUBFinalBillPrintMenu
'
'End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via DCSEtupMenu by " + PWUser$
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
   ' Me.SetFocus
  End If
  DoEvents
End Sub

'Private Sub cmdPrnAllUBBills_Click()
'  Dim FntSize As Integer
'  ReDim MsgText(0 To 5) As String
'
'  If Not Exist(UBFinBillsFile) Then
'    frmMsgDialog.RetLabel = "-2"
'    UBLog "ERROR: NO BILL FILE! Final"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "NO BILL FILE!"
'    MsgText(3) = ""
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'    Exit Sub
'  End If
'  frmBillPrinting.REPRN False, True
'  Load frmBillPrinting
'  DoEvents
'  frmBillPrinting.Show
'  Unload frmUBFinalBillPrintMenu
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdSetupInfo.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub Command1_Click()
  If MsgBox("Continue with transaction conversion?", vbYesNo, "Continue") = vbYes Then
    ConvertTrans
  End If
End Sub
Private Sub Command2_Click()
  If MsgBox("Continue with Running Balance recalc?", vbYesNo, "Continue") = vbYes Then
    RecalcBal
  End If
End Sub

Private Sub ConvertTrans()
Dim DCTranLen  As Integer, NumOfRecs As Long
Dim DCTran  As Integer
Dim cnt As Long
If Exist(DCPath$ + "DCSetup.dat") Then
  If MsgBox("Setup file exist continue with trasaction conversion?", vbYesNo, "Continue") = vbNo Then
    GoTo donealready
  End If
End If
ReDim DCTrans(1) As DCTransRecType
If Exist(DCPath$ + "DCCust.dat") And Exist(DCPath$ + "DCTrans.dat") Then
  DCTranLen = Len(DCTrans(1))
  DCTran = FreeFile
  NumOfRecs = FileSize(DCPath$ + "DCTrans.DAT") \ DCTranLen
  Open DCPath$ + "DCTrans.DAT" For Random Shared As DCTran Len = DCTranLen
  For cnt = 1 To NumOfRecs
    Get DCTran, cnt, DCTrans(1)
    If DCTrans(1).ChkByte = Chr$(1) Then
      MsgBox "Already Converted.", vbOKOnly, "Converted"
      GoTo donealready
    End If
    DCTrans(1).ExtraDesc = ""
    DCTrans(1).VoidFlag = "N"
    DCTrans(1).ChkByte = Chr$(1)
    If (DCTrans(1).CashAmount <> 0) And (DCTrans(1).ChkAmount <> 0) Then
      DCTrans(1).TransTender = 3
    ElseIf DCTrans(1).CashAmount <> 0 Then
      DCTrans(1).TransTender = 1
    ElseIf DCTrans(1).ChkAmount <> 0 Then
      DCTrans(1).TransTender = 2
    Else
      DCTrans(1).TransTender = 0
    End If
    Put #DCTran, cnt, DCTrans(1)
  Next
  MsgBox "Transaction File Converted", vbOKOnly, "Completed"
donealready:
  Close #DCTran
Else
  MsgBox "Files Missing", vbOKOnly, "Nothing Converted"
End If 'file already exist do nothing
End Sub
Private Sub RecalcBal()
Dim DCTranLen  As Integer, NumOfRecs As Long, DCCustRecLen As Integer
Dim DCFile As Integer, PrevTranBal As Double, TrHandle As Integer
Dim cnt As Long, CntT As Long, PrevTranRec As Long
ReDim DCCustRec(1) As DCCustRecType
ReDim DCTransRec(1) As DCTransRecType
If Exist(DCPath$ + "DCCust.dat") And Exist(DCPath$ + "DCTrans.dat") Then
  DCCustRecLen = Len(DCCustRec(1))
  TrHandle = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = DCCustRecLen
  NumOfRecs = FileSize(DCPath$ + "DCCust.DAT") \ DCCustRecLen

  DCTranLen = Len(DCTransRec(1))
  DCFile = FreeFile
  Open DCPath$ + "DCTrans.DAT" For Random Shared As DCFile Len = DCTranLen
  For cnt = 1 To NumOfRecs
    Get TrHandle, cnt, DCCustRec(1)
    PrevTranRec& = DCCustRec(1).FirstTrans
    PrevTranBal = 0
    If PrevTranRec& > 0 Then
      Do While PrevTranRec& > 0
        CntT& = PrevTranRec&
        Get DCFile, CntT&, DCTransRec(1)
        If DCTransRec(1).TransType = 1 Or DCTransRec(1).TransType = 4 Then
          DCTransRec(1).BalanceAfterTrans = PrevTranBal + DCTransRec(1).TransAmount
        ElseIf DCTransRec(1).TransType = 2 Or DCTransRec(1).TransType = 3 Then
          DCTransRec(1).BalanceAfterTrans = PrevTranBal - DCTransRec(1).TransAmount
        End If
        PrevTranRec& = DCTransRec(1).NextTrans
        PrevTranBal = DCTransRec(1).BalanceAfterTrans
        Put #DCFile, CntT&, DCTransRec(1)
      Loop
    End If
  Next
  Close #DCFile
  Close #TrHandle
Else
  MsgBox "Files Missing", vbOKOnly, "Nothing Recalculated"
End If 'file already exist do nothing

End Sub


