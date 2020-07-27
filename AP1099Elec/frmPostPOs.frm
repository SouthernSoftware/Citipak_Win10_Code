VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostPOs 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Purchase Orders"
   ClientHeight    =   8616
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   12192
   Icon            =   "frmPostPOs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8616
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   2544
      Top             =   2952
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Ok"
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
      Left            =   8112
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1356
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
      Height          =   492
      Left            =   9792
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   1356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   8364
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   445
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
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   348
      Left            =   4956
      TabIndex        =   7
      Top             =   2928
      Width           =   2292
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Once Posted, PO Forms May NOT Be Printed Again. "
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
      Index           =   2
      Left            =   3288
      TabIndex        =   6
      Top             =   4008
      Width           =   5628
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Ok to Begin Posting or Exit to Escape Posting Procedure. "
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
      Index           =   1
      Left            =   2712
      TabIndex        =   5
      Top             =   4392
      Width           =   6780
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Before You Post, Make Sure You Have Printed All Purchase Order Forms Needed. "
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
      Height          =   612
      Index           =   0
      Left            =   2748
      TabIndex        =   4
      Top             =   3336
      Width           =   6708
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Post Purchase Orders"
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
      Left            =   3906
      TabIndex        =   3
      Top             =   1584
      Width           =   4476
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1344
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2172
      Left            =   2370
      Top             =   2784
      Width           =   7452
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3210
      Top             =   1224
      Width           =   5772
   End
End
Attribute VB_Name = "frmPostPOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim PO As POFORMRecType2
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim Editing As Boolean
Private Sub Timer1_Timer()
 ' Label2.Visible = Not Label2.Visible
  '&H0080FFFF&
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Label2.ForeColor = &H80FFFF
    Shape3.BackColor = &HC0&
  Else
    Label2.ForeColor = &HFFFF&
    Shape3.BackColor = &H80&
  End If
  
End Sub

Private Sub cmdExit_Click()
  Unload frmPostPOs
End Sub

Private Sub cmdOk_Click()
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  PostPOTrans
  
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload frmPostPOs
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      MainLog "Close AP"
      ClearInUse PWcnt
    End If
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
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpPostPO
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PostPOTrans()
  Dim Ready As Integer, cnt As Integer, NumFunds As Integer
  Dim VendRecNum As Integer, RecordNum As Integer, AcctDist As Integer
  Dim POEditFile As Integer, NumEdTrans As Integer, AP2Post As Integer
  Dim LedgerRecLen As Integer, DistRecLEn As Integer, VendorRecLen As Integer
  Dim Fund As Integer, FundNum As String, BadTrans As Integer
  Dim APLedgerFile As Integer, NumLedgerRecs As Long, POLogFileName As String
  Dim APDistFile As Integer, NumDistRecs As Long, d As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, OK2Kill As Boolean
  Dim PrevVendTrans As Long, FrstVendTrans As Long, ReportFile As String
  Dim PRNFile As Integer, ReportFileT As String, ReportErr As String
  ReDim APDistRec(1) As APDistRecType
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim Tr2Post(1) As GLTransRecType
 
  LedgerRecLen = Len(APLedgerRec(1))
  DistRecLEn = Len(APDistRec(1))
  VendorRecLen = Len(Vendor)
  Ready = 0

  '--Verify that there are transactions
  OpenPOEditFile POEditFile, NumEdTrans

  '--Test for valid Vendor Numbers
  For cnt = 1 To NumEdTrans
    Get POEditFile, cnt, PO
    If PO.LOCKED = False Then
      If PO.Deleted = 1 Then
      'If can't find vendor abort
        VendRecNum = FindVendorRec(PO.VNDRCODE)
        If VendRecNum > 0 Then
          Ready = Ready + 1
         'Do Stuff
        Else
          MsgBox "Invalid Vendor Code, Unable to Locate Vendor. Operation Aborted.", vbOKOnly, "Error"
          Close
          Exit Sub
        End If
      End If
    Else
      Editing = True
      Exit For
    End If
  Next
  Close

  If Not Editing Then
    '--Check for no active transactions
    If Ready = 0 Then
      '--No active transactions - tell user and get out
      MsgBox "No Purchase Orders To Post.", vbOKOnly, "Post Canceled"
      Exit Sub
    End If
    SetAttr ("APPED.DAT"), vbReadOnly
  '--make sure we're ready to post
    If MsgBox("Are You Sure You Wish To Post Now?", vbYesNo, "Continue?") = vbNo Then
      GoSub OutNow
    End If
  Else
    MsgBox "PO Editing Is In Process, Please Close PO Edit Procedures Before Trying To Post.", vbOKOnly, "Canceled"
    GoSub OutNow
  End If
  FrmShowPctComp.Label1 = "Verifying Purchase Order Transactions"
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  ReDim FundList(1) As String
  GetFundList FundList(), NumFunds
  ReDim TrFundSum#(1 To NumFunds)

  OpenPOEditFile POEditFile, NumEdTrans
  AP2Post = FreeFile
  Open "APPOIF.DAT" For Random As AP2Post Len = Len(Tr2Post(1))

  RecordNum = 0 'Reset Active counter for posting

  For cnt = 1 To NumEdTrans     'number of invoices to process
    ReDim TrFundSum#(1 To NumFunds)
    'Reinitialize transaction fund summary elements for next loop
    Get POEditFile, cnt, PO
    FrmShowPctComp.ShowPctComp cnt, NumEdTrans
    If PO.Deleted = 1 Then '1 was assigned during print po's or approval
      For AcctDist = 1 To 36
        If Len(QPTrim$(PO.ITEMS(AcctDist).ACCTNO)) Then
          RecordNum = RecordNum + 1
          Tr2Post(1).AcctNum = PO.ITEMS(AcctDist).ACCTNO
          Tr2Post(1).TRDATE = PO.PODATE
          Tr2Post(1).Desc = PO.VNDRCODE
          Tr2Post(1).Ref = PO.PONum
          Tr2Post(1).DrAmt = PO.ITEMS(AcctDist).EXT
          Tr2Post(1).CrAmt = 0
          Tr2Post(1).Src = "PO" + Format$(Now, "mmddyy")
          Put AP2Post, RecordNum, Tr2Post(1)

          '--Add this distribution to proper fund
          For Fund = 1 To NumFunds
            FundNum$ = Left$(PO.ITEMS(AcctDist).ACCTNO, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              TrFundSum#(Fund) = Round#(TrFundSum#(Fund) + PO.ITEMS(AcctDist).EXT)
              'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(PO.Items(AcctDist).ext)
              Exit For
            End If
          Next

        End If  'test for blank distribution line
      Next      'Acct'g Distribution
    End If      'not deleted

  Next          'transaction

  Close

  '--common post & link sub in comnaux
  Post2PO "APPOIF.DAT", BadTrans%, frmPostPOs, False
  If BadTrans <> 0 Then
    '--Couldn't find an account.
    '--Account was possibly deleted after entry made?
      MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
      ReportFile$ = "TempLog.PRN"
      ViewPrint ReportFile$, "Error Log"
      SetAttr ("APPED.DAT"), vbNormal
      frmCitiCancel.Show
      Unload frmPostPOs
      Unload frmPOProcessMenu
      Exit Sub

  End If
  Post2PO "Appoif.dat", BadTrans%, frmPostPOs, True
  If BadTrans <> 0 Then                  'posting problem
      MsgBox "Error, One or more transactions were not posted. Make sure the printer is ready and Press a Key to View Log.", vbOKOnly, "Posting Error"
      POLogFileName = "POlog.dat"
      ReportFile$ = "POlog.dat"
      ViewPrint ReportFile$, "Posting Log"
      SetAttr ("APPED.DAT"), vbNormal
    End If

  '--Now post transaction to apledger.dat
  SetAttr ("APPED.DAT"), vbNormal
  OpenPOEditFile POEditFile, NumEdTrans
  OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  For cnt = 1 To NumEdTrans
    Get POEditFile, cnt, PO
    If PO.Deleted = 1 Then
      '--Update PO in vendor link to apledger
      NumLedgerRecs = NumLedgerRecs + 1
      VendRecNum = PO.VNDRREC

      OpenVendorFile VendorFile, NumVRecs
      Get VendorFile, VendRecNum, Vendor

      PrevVendTrans = Vendor.LastTran
      FrstVendTrans = Vendor.FrstTran

      If FrstVendTrans = 0 Then
        Vendor.LastTran = NumLedgerRecs
        Vendor.FrstTran = NumLedgerRecs
      Else
        Vendor.LastTran = NumLedgerRecs
        Get APLedgerFile, PrevVendTrans, APLedgerRec(1)
        APLedgerRec(1).NextTrans = NumLedgerRecs
        Put APLedgerFile, PrevVendTrans, APLedgerRec(1)
      End If
      Put VendorFile, VendRecNum, Vendor
      Close VendorFile

      '--Post transaction to apledger
      APLedgerRec(1).VRecNum = PO.VNDRREC
      APLedgerRec(1).VendorCode = PO.VNDRCODE
      APLedgerRec(1).TRDATE = PO.PODATE
      APLedgerRec(1).DOCNum = PO.PONum
      APLedgerRec(1).PONum = PO.PONum
      APLedgerRec(1).TRCode = 4

      APLedgerRec(1).DeptNumb = Val(PO.REQNUM)

      'APLedgerRec(1).PayCode = VAL(PO.PayCode)
      APLedgerRec(1).GLDistDate = PO.PODATE
      APLedgerRec(1).Amt = PO.POAmt
      APLedgerRec(1).NextTrans = 0
      APLedgerRec(1).FrstDist = NumDistRecs& + 1
      APLedgerRec(1).Bankcode = 0
      For d = 1 To 36
        If Len(QPTrim$(PO.ITEMS(d).ACCTNO)) Then
          '--Post and link distributions
          ReDim APDistRec(1) As APDistRecType
          NumDistRecs& = NumDistRecs& + 1
          APDistRec(1).APLedgerRec = NumLedgerRecs
          APDistRec(1).DistAcctRec = PO.ITEMS(d).AcctRec
          APDistRec(1).DistAcctNum = PO.ITEMS(d).ACCTNO
          APDistRec(1).DistAmt = PO.ITEMS(d).EXT
          'APDistRec(1).DistCRAmt = 0
          APDistRec(1).NextDist = NumDistRecs& + 1
          Put APDistFile, NumDistRecs&, APDistRec(1)
        End If
      Next
      'Update the last distribution pointer
      Get APDistFile, NumDistRecs&, APDistRec(1)
      APDistRec(1).NextDist = 0
      Put APDistFile, NumDistRecs&, APDistRec(1)
      APLedgerRec(1).LastDist = NumDistRecs&
      Put APLedgerFile, NumLedgerRecs, APLedgerRec(1)
      PO.Deleted = -1
      Put POEditFile, cnt, PO
    End If
  Next

  OK2Kill = -1
  For cnt = 1 To NumEdTrans
    Get POEditFile, cnt, PO
    If PO.Deleted <> -1 Then
      OK2Kill = 0
      Exit For
    End If
  Next
  Close

  If OK2Kill = -1 Then
    KillFile "APPED.DAT"
  End If
  KillFile "APPOIF.DAT"
Call MainLog("PO Posting Completed")
MsgBox "Posting Complete.", vbOKOnly, "Completed"
Exit Sub
OutNow:
  SetAttr ("APPED.DAT"), vbNormal
  Exit Sub
End Sub


Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
