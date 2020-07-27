VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBDraftMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACH - Draft Processing Menu"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmUBDraftMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitUBDraft 
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
      Left            =   3840
      TabIndex        =   5
      Top             =   6120
      Width           =   4524
   End
   Begin VB.CommandButton cmdPrenote 
      Caption         =   "Prepare Draft &Prenote File"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   4160
      Width           =   4524
   End
   Begin VB.CommandButton cmdTestFile 
      Caption         =   "Prepare Draft Test File for &Bank"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   5464
      Width           =   4524
   End
   Begin VB.CommandButton cmdCustList 
      Caption         =   "Print Draft &Customer Listing"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   4812
      Width           =   4524
   End
   Begin VB.CommandButton cmdTransmission 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prepare Draft &Transmission File"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   3508
      Width           =   4524
   End
   Begin VB.CommandButton cmdAcctsDraftRpt 
      BackColor       =   &H008F8265&
      Caption         =   "&Accounts To Draft Report"
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
      Left            =   3840
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2856
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
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
            TextSave        =   "2:10 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "11/15/2004"
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
      Caption         =   "ACH - Draft Processing Menu"
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
      TabIndex        =   7
      Top             =   1104
      Width           =   5148
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
Attribute VB_Name = "frmUBDraftMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class

Private Sub cmdAcctsDraftRpt_Click()
  Load frmDraftAccountsRpt
  DoEvents
  frmDraftAccountsRpt.Show
  Unload frmUBDraftMenu
End Sub

Private Sub cmdCustList_Click()
  Load frmDraftCustList
  DoEvents
  frmDraftCustList.Show
  Unload frmUBDraftMenu
End Sub

Private Sub cmdExitUBDraft_Click()
  Load frmUBBillingMenu
  DoEvents
  frmUBBillingMenu.Show
  Unload frmUBDraftMenu
End Sub

Private Sub cmdPrenote_Click()
  frmDraftPrenote.PreSet True
  'Load frmDraftPrenote
  DoEvents
  frmDraftPrenote.Show
End Sub

Private Sub cmdTestFile_Click()
  frmDraftPrenote.Label1.Caption = "Ready to Create Draft Test File?"
  frmDraftPrenote.Caption = "Draft Test File"
  frmDraftPrenote.PreSet False
  DoEvents
  frmDraftPrenote.Show
End Sub

Private Sub cmdTransmission_Click()
  Load frmDraftPrepareFile
  DoEvents
  frmDraftPrepareFile.Show
  Unload frmUBDraftMenu
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitUBDraft.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via UBDraftMenu by " + PWUser$
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdAcctsDraftRpt.SetFocus
    Case vbKeyEnd
      cmdExitUBDraft.SetFocus
    Case Else:
  End Select
End Sub
Public Sub DoPreNote()
  DeActivateControls Me
  UBPrenote
  ActivateControls Me
End Sub
Public Sub DoDrftTest()
  DeActivateControls Me
  UBDraftTest
  ActivateControls Me
End Sub

Private Sub UBPrenote()
  Dim UBSetupLen As Integer, IndexName As String, WarrFlag As Boolean
  Dim UBCustRecLen As Integer, UBCust As Integer, NWoodFlag As Boolean
  Dim NumOfRecs As Long, cnt As Long, BalRecFlag As Boolean
  Dim CustCycle As Integer, CustOk As Boolean, CCnt As Integer
  Dim CstCnt As Long, llow As Long, hhigh As Long, BankCnt As Integer
  Dim PrevBank As String, GTotal As Double, CompanyAcct As String
  Dim DraftFile As Integer, PlyFlag As Boolean, GATot As Double
  Dim PayRecLen As Integer, PayFileName As String, Make99Flag As Boolean
  Dim BDate As Integer, Done As Boolean, DraftDate As String
  Dim NewDraftFile As String, Step1 As Boolean, Step2 As Boolean
  Dim Step3 As Boolean, Step4 As Boolean, Step5 As Boolean, CustBal As Double
  Dim FedIDNum As String, DraftFileNum As Integer, Counter As Integer
  Dim BillAmt As String, TotalAmountn As Double, AcctNumber As String
  Dim nme As String, BankAcct As String, sp As String, Trac As Integer
  Dim Trace As String, hashh As Double, Number As Integer, Pay99File As Integer
  Dim hash As String, TotalAmount As String, TotSize As Double
  Dim BlockSize As Single, BlockSizeS As String, FillSize As Single
  Dim outfile As Integer, FasFlag As Boolean, NumOfLines As Integer
  Dim ZZCnt As Integer, PreNoteFile As String, AcctRecord As Long
  'load setup file

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TOWNNAME$ = UCase$(UBSetUpRec(1).UTILNAME)

  ReDim UBDraftRec(1) As UBDraftRecType
  DraftFile = FreeFile
  Open UBPath$ + "UBSDRAFT.dat" For Random Access Read Shared As #DraftFile Len = Len(UBDraftRec(1))
  Get DraftFile, 1, UBDraftRec(1)
  Close

  PreNoteFile$ = QPTrim$(UBDraftRec(1).FileName)

  If InStr(TOWNNAME$, "WARRENTON") > 0 Then
    PreNoteFile$ = "UBDFNOTE"
    WarrFlag = True
  ElseIf Len(PreNoteFile$) = 0 Then
    PreNoteFile$ = UBPath$ + "UBDFNOTE.DAT"
  End If

  If InStr(TOWNNAME$, "NORWOOD") > 0 Then
    NWoodFlag = True
  End If
  If InStr(TOWNNAME$, "LEE") > 0 Then
    NWoodFlag = True
  End If
'  IF INSTR(TownName$, "FAISON") > 0 THEN
'    FasFlag = True
'  END IF

ProcessPrenote:
  frmDraftMsg.Label(5).Visible = False
  frmDraftMsg.Label(0).Caption = "Building Record Type 1"
  frmDraftMsg.Show 1, Me
  DoEvents

  Do
FormLoop:
    
    ' Process Record Type 1
    If Not Step1 Then
      GoSub ProcessStep1
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      DoEvents
      frmDraftMsg.Label(1).Caption = "Building Record Type 5"
      DoEvents
      GoTo FormLoop
    End If

    If Not Step2 Then
      GoSub ProcessStep2
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      DoEvents
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 "
      DoEvents
      GoTo FormLoop
    End If

    If Not Step3 Then
      GoSub ProcessStep3
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 ..Done!"
      DoEvents
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 "
      DoEvents
      GoTo FormLoop
    End If

    If Not Step4 Then
      GoSub ProcessStep4
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 ..Done!"
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 ..Done!"
      DoEvents
      frmDraftMsg.Label(4).Caption = "Building Record Type 9"
      DoEvents
      GoTo FormLoop
    End If

    If Not Step5 Then
      GoSub ProcessStep5
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 ..Done!"
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 ..Done!"
      frmDraftMsg.Label(4).Caption = "Building Record Type 9 ..Done!"
      DoEvents
      frmDraftMsg.Label(6).Caption = "File Name Is: " + PreNoteFile$
      DoEvents
      GoTo FormLoop
    End If


  Loop Until Done
  Exit Sub

Return
OpenMainDraftInfo:
  ReDim UBDraftRec(1) As UBDraftRecType
  DraftFile = FreeFile
  Open UBPath$ + "UBSDRAFT.dat" For Random Access Read Shared As #DraftFile Len = Len(UBDraftRec(1))
  Get DraftFile, 1, UBDraftRec(1)


Return

ProcessStep1:
  GoSub OpenMainDraftInfo
  FedIDNum$ = QPTrim$(UBDraftRec(1).FEDPREFX + UBDraftRec(1).FEDID) + "00"

  ReDim UBDraftRecord1(1) As UBDraftRecord1Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT1.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord1(1))
  UBDraftRecord1(1).Field1 = "1"
  UBDraftRecord1(1).Field2 = "01"
  UBDraftRecord1(1).Field3 = " " + UBDraftRec(1).BANKDEST
  UBDraftRecord1(1).Field4 = " " + UBDraftRec(1).BANKORIG
  UBDraftRecord1(1).Field5 = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  UBDraftRecord1(1).Field6 = Left$(Time$, 2) + Mid$(Time$, 4, 2)
  UBDraftRecord1(1).Field7 = "A"
  UBDraftRecord1(1).Field8 = "094"
  UBDraftRecord1(1).Field9 = "10"
  UBDraftRecord1(1).Field10 = "1"
  UBDraftRecord1(1).Field11 = QPTrim$(UCase$(UBDraftRec(1).BankName))
  UBDraftRecord1(1).Field12 = QPTrim$(UCase$(UBDraftRec(1).BANKLOC))
  UBDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  Put DraftFileNum, 1, UBDraftRecord1(1)
  Close DraftFileNum
  Step1 = True
  Return

ProcessStep2:
  ReDim UBDraftRecord5(1) As UBDraftRecord5Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT5.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord5(1))
  UBDraftRecord5(1).Field1 = "5"
  UBDraftRecord5(1).Field2 = "200"
  UBDraftRecord5(1).Field3 = Left$(TOWNNAME$, 16)
  UBDraftRecord5(1).Field4 = "                    "
  UBDraftRecord5(1).Field5 = FedIDNum$
  UBDraftRecord5(1).Field6 = "PPD"
  UBDraftRecord5(1).Field7 = "UTIL BILL"
  UBDraftRecord5(1).Field8 = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  UBDraftRecord5(1).Field9 = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  UBDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  UBDraftRecord5(1).Field11 = "1"
  UBDraftRecord5(1).Field12 = Left$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord5(1).Field13 = "0000001"
  Put DraftFileNum, 1, UBDraftRecord5(1)
  Close DraftFileNum
  Step2 = True
  Return

ProcessStep3:
  Counter = 0

  ReDim UBDraftRecord6(1) As UBDraftRecord6Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))
  Close DraftFileNum
  Kill UBPath$ + "UBDRAFT6.DAT"
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))

  'GO THRU DATA FILE HERE
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  For cnt = 1 To NumOfRecs
    Get UBCust, cnt, UBCustRec(1)
    AcctRecord = cnt

    '  Process Customer Here
    If UBCustRec(1).Status = "A" And UBCustRec(1).USEDRAFT = "Y" And UBCustRec(1).PreNoteFlag = 0 Then
      If Left$(UBCustRec(1).TRANSIT, 8) = "05318221" Then UBCustRec(1).TRANSIT = "053108221": Put UBCust, cnt, UBCustRec(1)
      Counter = Counter + 1
      AcctNumber$ = Str$(AcctRecord)
      AcctNumber$ = Right$(AcctNumber$, Len(AcctNumber$) - 1)
      If Len(AcctNumber$) < 15 Then
        AcctNumber$ = AcctNumber$ + String$(15 - Len(AcctNumber$), 32)
      End If
      nme$ = UBCustRec(1).CustName
      If Len(nme$) < 22 Then
        nme$ = nme$ + String$(22 - Len(nme$), 32)
      Else
        nme$ = Left$(nme$, 22)
      End If

      'Check for Spaces WithIn Bank Account Numbered as Entered by Customer
      BankAcct$ = QPTrim$(UBCustRec(1).BankAcct)
      sp = InStr(BankAcct$, " ")
      If sp > 0 Then
        BankAcct$ = Left$(BankAcct$, sp - 1) + Right$(BankAcct$, Len(BankAcct$) - sp)
      End If

      If Len(BankAcct$) < 17 Then BankAcct$ = BankAcct$ + String$(17 - Len(BankAcct$), 32)
      Trac = Trac + 1
      Trace$ = Str$(Trac): Trace$ = Right$(Trace$, Len(Trace$) - 1)
      If Len(Trace$) < 7 Then Trace$ = String$(7 - Len(Trace$), "0") + Trace$

      UBDraftRecord6(1).Field1 = "6"
      UBDraftRecord6(1).Field2 = "28"           ' Designates Prenote Trans
      UBDraftRecord6(1).Field3 = Left$(UBCustRec(1).TRANSIT, 8)
      UBDraftRecord6(1).Field4 = Right$(UBCustRec(1).TRANSIT, 1)
      UBDraftRecord6(1).Field5 = Left$(BankAcct$, 17)
      UBDraftRecord6(1).Field6 = "0000000000"   ' All zero's for Prenote
      UBDraftRecord6(1).Field7 = AcctNumber$
      UBDraftRecord6(1).Field8 = UCase$(nme$)
      UBDraftRecord6(1).Field9 = "  "
      UBDraftRecord6(1).Field10 = "0"
      UBDraftRecord6(1).Field11 = Left$(UBDraftRec(1).BANKORIG, 8) + Trace$
      Put DraftFileNum, Counter, UBDraftRecord6(1)
      hashh# = hashh# + Val(Left$(UBCustRec(1).TRANSIT, 8))
      UBCustRec(1).PreNoteFlag = 1
      Put UBCust, cnt, UBCustRec(1)
      Number = Number + 1
    End If
  Next cnt
  Close DraftFileNum
  Step3 = True
  Return
ProcessStep4:
  hash$ = Str$(hashh#)
  hash$ = Right$(hash$, Len(hash$) - 1)

  If Len(hash$) < 10 Then
    hash$ = String$(10 - Len(hash$), "0") + hash$
  End If
  If Len(hash$) > 10 Then
    hash$ = Right$(hash$, 10)
  End If

  If Len(Trace$) > 6 Then Trace$ = Right$(Trace$, 6)

  ReDim UBDraftRecord8(1) As UBDraftRecord8Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT8.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord8(1))
  UBDraftRecord8(1).Field1 = "8"
  UBDraftRecord8(1).Field2 = "200"
  UBDraftRecord8(1).Field3 = Trace$
  UBDraftRecord8(1).Field4 = hash$
  UBDraftRecord8(1).Field5 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field7 = FedIDNum$
  UBDraftRecord8(1).Field8 = String$(19, 32)    ' Reserved
  UBDraftRecord8(1).Field9 = String$(6, 32)     ' Reserved for Federal Reserve use
  UBDraftRecord8(1).Field10 = Left$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord8(1).Field11 = "0000001"
  Put DraftFileNum, 1, UBDraftRecord8(1)
  Close DraftFileNum
  Step4 = True
  Return

ProcessStep5:
  TotSize# = Val(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records

  If BlockSize! <> Int(BlockSize!) Then
    BlockSize! = Int(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize - 1)))
  Else
    FillSize! = 0
  End If

  BlockSizeS$ = Str$(BlockSize!)
  BlockSizeS$ = Right$(BlockSizeS$, Len(BlockSizeS$) - 1)
  If Len(BlockSizeS$) < 6 Then BlockSizeS$ = String$(6 - Len(BlockSizeS$), "0") + BlockSizeS$
  If Len(Trace$) < 8 Then Trace$ = String$(8 - Len(Trace$), "0") + Trace$

  ReDim UBDraftRecord9(1) As UBDraftRecord9Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT9.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord9(1))
  UBDraftRecord9(1).Field1 = "9"
  UBDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  UBDraftRecord9(1).Field3 = BlockSizeS$
  UBDraftRecord9(1).Field4 = Trace$
  UBDraftRecord9(1).Field5 = hash$
  UBDraftRecord9(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord9(1).Field7 = "000000000000"
  UBDraftRecord9(1).Field8 = String$(39, 32)    ' Reserved
  Put DraftFileNum, 1, UBDraftRecord9(1)
  Close DraftFileNum
  ' Now Put Them Together In File Name UBDFNOTE
  outfile = FreeFile
  Open PreNoteFile$ For Output As outfile

  'OPEN "O", OutFile, : WIDTH #OutFile, 255

  ReDim UBDraftRecord1(1) As UBDraftRecord1Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT1.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord1(1))
  Get DraftFileNum, 1, UBDraftRecord1(1)
  Print #outfile, UBDraftRecord1(1).Field1;
  Print #outfile, UBDraftRecord1(1).Field2;
  Print #outfile, UBDraftRecord1(1).Field3;
  Print #outfile, UBDraftRecord1(1).Field4;
  Print #outfile, UBDraftRecord1(1).Field5;
  Print #outfile, UBDraftRecord1(1).Field6;
  Print #outfile, UBDraftRecord1(1).Field7;
  Print #outfile, UBDraftRecord1(1).Field8;
  Print #outfile, UBDraftRecord1(1).Field9;
  Print #outfile, UBDraftRecord1(1).Field10;
  Print #outfile, UBDraftRecord1(1).Field11;
  Print #outfile, UBDraftRecord1(1).Field12;
  Print #outfile, UBDraftRecord1(1).Field13;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If

  Close DraftFileNum

  ReDim UBDraftRecord5(1) As UBDraftRecord5Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT5.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord5(1))
  Get DraftFileNum, 1, UBDraftRecord5(1)
  Print #outfile, UBDraftRecord5(1).Field1;
  Print #outfile, UBDraftRecord5(1).Field2;
  Print #outfile, UBDraftRecord5(1).Field3;
  Print #outfile, UBDraftRecord5(1).Field4;
  Print #outfile, UBDraftRecord5(1).Field5;
  Print #outfile, UBDraftRecord5(1).Field6;
  Print #outfile, UBDraftRecord5(1).Field7;
  Print #outfile, UBDraftRecord5(1).Field8;
  Print #outfile, UBDraftRecord5(1).Field9;
  Print #outfile, UBDraftRecord5(1).Field10;
  Print #outfile, UBDraftRecord5(1).Field11;
  Print #outfile, UBDraftRecord5(1).Field12;
  Print #outfile, UBDraftRecord5(1).Field13;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If

  Close DraftFileNum

  ReDim UBDraftRecord6(1) As UBDraftRecord6Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94

  For cnt = 1 To NumOfLines
    Get DraftFileNum, cnt, UBDraftRecord6(1)
    Print #outfile, UBDraftRecord6(1).Field1;
    Print #outfile, UBDraftRecord6(1).Field2;
    Print #outfile, UBDraftRecord6(1).Field3;
    Print #outfile, UBDraftRecord6(1).Field4;
    Print #outfile, UBDraftRecord6(1).Field5;
    Print #outfile, UBDraftRecord6(1).Field6;
    Print #outfile, UBDraftRecord6(1).Field7;
    Print #outfile, UBDraftRecord6(1).Field8;
    Print #outfile, UBDraftRecord6(1).Field9;
    Print #outfile, UBDraftRecord6(1).Field10;
    Print #outfile, UBDraftRecord6(1).Field11;
    If NWoodFlag = 0 And FasFlag = 0 Then
      Print #outfile,
    ElseIf FasFlag Then
      Print #outfile, Chr$(13);
    End If
  Next cnt
  Close DraftFileNum

  ReDim UBDraftRecord8(1) As UBDraftRecord8Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT8.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord8(1))
  Get DraftFileNum, 1, UBDraftRecord8(1)
  Print #outfile, UBDraftRecord8(1).Field1;
  Print #outfile, UBDraftRecord8(1).Field2;
  Print #outfile, UBDraftRecord8(1).Field3;
  Print #outfile, UBDraftRecord8(1).Field4;
  Print #outfile, UBDraftRecord8(1).Field5;
  Print #outfile, UBDraftRecord8(1).Field6;
  Print #outfile, UBDraftRecord8(1).Field7;
  Print #outfile, UBDraftRecord8(1).Field8;
  Print #outfile, UBDraftRecord8(1).Field9;
  Print #outfile, UBDraftRecord8(1).Field10;
  Print #outfile, UBDraftRecord8(1).Field11;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If

  Close DraftFileNum

  ReDim UBDraftRecord9(1) As UBDraftRecord9Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT9.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord9(1))
  Get DraftFileNum, 1, UBDraftRecord9(1)
  Print #outfile, UBDraftRecord9(1).Field1;
  Print #outfile, UBDraftRecord9(1).Field2;
  Print #outfile, UBDraftRecord9(1).Field3;
  Print #outfile, UBDraftRecord9(1).Field4;
  Print #outfile, UBDraftRecord9(1).Field5;
  Print #outfile, UBDraftRecord9(1).Field6;
  Print #outfile, UBDraftRecord9(1).Field7;
  Print #outfile, UBDraftRecord9(1).Field8;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If

  Close DraftFileNum

  If NWoodFlag = 0 Then
    If FillSize! < 0 Then
      FillSize! = 940 - Abs(FillSize!)
    End If
    For cnt = 1 To Abs(FillSize!) / 94
      If FasFlag Then
        Print #outfile, String$(94, "9"); Chr$(13);
      Else
        Print #outfile, String$(94, "9")
      End If
    Next cnt
  End If

  Close
  Step5 = True
  Done = True
  Return


End Sub

Private Sub UBDraftTest()
  Dim UBSetupLen As Integer, IndexName As String, WarrFlag As Boolean
  Dim UBCustRecLen As Integer, UBCust As Integer, NWoodFlag As Boolean
  Dim NumOfRecs As Long, cnt As Long, BalRecFlag As Boolean
  Dim CustCycle As Integer, CustOk As Boolean, CCnt As Integer
  Dim CstCnt As Long, llow As Long, hhigh As Long, BankCnt As Integer
  Dim PrevBank As String, GTotal As Double, CompanyAcct As String
  Dim DraftFile As Integer, PlyFlag As Boolean, GATot As Double
  Dim PayRecLen As Integer, PayFileName As String, Make99Flag As Boolean
  Dim BDate As Integer, Done As Boolean, DraftDate As String
  Dim NewDraftFile As String, Step1 As Boolean, Step2 As Boolean
  Dim Step3 As Boolean, Step4 As Boolean, Step5 As Boolean, CustBal As Double
  Dim FedIDNum As String, DraftFileNum As Integer, Counter As Integer
  Dim BillAmt As String, TotalAmountn As Double, AcctNumber As String
  Dim nme As String, BankAcct As String, sp As String, Trac As Integer
  Dim Trace As String, hashh As Double, Number As Integer, Pay99File As Integer
  Dim hash As String, TotalAmount As String, TotSize As Double
  Dim BlockSize As Single, BlockSizeS As String, FillSize As Single
  Dim outfile As Integer, FasFlag As Boolean, NumOfLines As Integer
  Dim ZZCnt As Integer, TestFileName As String, AcctRecord As Long

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TOWNNAME$ = UBSetUpRec(1).UTILNAME

  If InStr(TOWNNAME$, "WARRENTON") > 0 Then
    TestFileName$ = "UBDFTEST"
    WarrFlag = True
  Else
    TestFileName$ = UBPath$ + "UBDFTEST.DAT"
  End If


ProcessTest:

  frmDraftMsg.Label(5).Visible = False
  frmDraftMsg.Show 1, Me
  frmDraftMsg.Label(0).Caption = "Building Record Type 1"
  DoEvents

  Do
FormTestLoop:

    If Not Step1 Then
      GoSub TestProcessStep1
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      DoEvents
      frmDraftMsg.Label(1).Caption = "Building Record Type 5"
      DoEvents
      GoTo FormTestLoop
    End If

    If Not Step2 Then
      GoSub TestProcessStep2
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      DoEvents
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 "
      DoEvents
      GoTo FormTestLoop
    End If
    If Not Step3 Then
      GoSub TestProcessStep3
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 ..Done!"
      DoEvents
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 "
      DoEvents
      GoTo FormTestLoop
    End If

    If Not Step4 Then
      GoSub TestProcessStep4
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 ..Done!"
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 ..Done!"
      DoEvents
      frmDraftMsg.Label(4).Caption = "Building Record Type 9"
      DoEvents
      GoTo FormTestLoop
    End If
    If Not Step5 Then
      GoSub TestProcessStep5
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 ..Done!"
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 ..Done!"
      frmDraftMsg.Label(4).Caption = "Building Record Type 9 ..Done!"
      DoEvents
      frmDraftMsg.Label(6).Caption = "File Name Is: " + TestFileName$
      GoTo FormTestLoop
    End If


  Loop Until Done
  Exit Sub

  Return
OpenDraftInfo:
  ReDim UBDraftRec(1) As UBDraftRecType
  DraftFile = FreeFile
  Open UBPath$ + "UBSDRAFT.dat" For Random Access Read Shared As #DraftFile Len = Len(UBDraftRec(1))
  Get DraftFile, 1, UBDraftRec(1)
  Return

TestProcessStep1:
  GoSub OpenDraftInfo

  FedIDNum$ = QPTrim$(UBDraftRec(1).FEDPREFX + UBDraftRec(1).FEDID) + "00"

  ReDim UBDraftRecord1(1) As UBDraftRecord1Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT1.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord1(1))
  UBDraftRecord1(1).Field1 = "1"
  UBDraftRecord1(1).Field2 = "01"
  UBDraftRecord1(1).Field3 = " " + UBDraftRec(1).BANKDEST
  UBDraftRecord1(1).Field4 = " " + UBDraftRec(1).BANKORIG
  UBDraftRecord1(1).Field5 = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  UBDraftRecord1(1).Field6 = Left$(Time$, 2) + Mid$(Time$, 4, 2)
  UBDraftRecord1(1).Field7 = "A"
  UBDraftRecord1(1).Field8 = "094"
  UBDraftRecord1(1).Field9 = "10"
  UBDraftRecord1(1).Field10 = "1"
  UBDraftRecord1(1).Field11 = QPTrim$(UCase$(UBDraftRec(1).BankName))
  UBDraftRecord1(1).Field12 = QPTrim$(UCase$(UBDraftRec(1).BANKLOC))
  UBDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  Put DraftFileNum, 1, UBDraftRecord1(1)
  Close DraftFileNum
  Step1 = True
  Return

TestProcessStep2:
  ReDim UBDraftRecord5(1) As UBDraftRecord5Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT5.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord5(1))
  UBDraftRecord5(1).Field1 = "5"
  UBDraftRecord5(1).Field2 = "200"
  UBDraftRecord5(1).Field3 = "UTILITY DRAFT TEST"
  UBDraftRecord5(1).Field4 = "                    "
  UBDraftRecord5(1).Field5 = FedIDNum$
  UBDraftRecord5(1).Field6 = "PPD"
  UBDraftRecord5(1).Field7 = "UTIL BILL"
  UBDraftRecord5(1).Field8 = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  UBDraftRecord5(1).Field9 = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  UBDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  UBDraftRecord5(1).Field11 = "1"
  UBDraftRecord5(1).Field12 = Left$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord5(1).Field13 = "0000001"
  Put DraftFileNum, 1, UBDraftRecord5(1)
  Close DraftFileNum
  Step2 = True
  Return

TestProcessStep3:
  Counter = 0

  ReDim UBDraftRecord6(1) As UBDraftRecord6Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))
  Close DraftFileNum
  Kill UBPath$ + "UBDRAFT6.DAT"
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))

  'GO THRU DATA FILE HERE
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  For cnt = 1 To NumOfRecs
    Get UBCust, cnt, UBCustRec(1)
    AcctRecord = cnt
    '  Process Customer Here
    If UBCustRec(1).Status = "A" And UBCustRec(1).USEDRAFT = "Y" Then
      Counter = Counter + 1
      AcctNumber$ = Str$(AcctRecord)
      AcctNumber$ = Right$(AcctNumber$, Len(AcctNumber$) - 1)
      If Len(AcctNumber$) < 15 Then
        AcctNumber$ = AcctNumber$ + String$(15 - Len(AcctNumber$), 32)
      End If
      nme$ = UBCustRec(1).CustName
      If Len(nme$) < 22 Then
        nme$ = nme$ + String$(22 - Len(nme$), 32)
      Else
        nme$ = Left$(nme$, 22)
      End If
      'Check for Spaces WithIn Bank Account Numbered as Entered by Customer
      BankAcct$ = QPTrim$(UBCustRec(1).BankAcct)
      sp = InStr(BankAcct$, " ")
      If sp > 0 Then
        BankAcct$ = Left$(BankAcct$, sp - 1) + Right$(BankAcct$, Len(BankAcct$) - sp)
      End If
      If Len(BankAcct$) < 17 Then BankAcct$ = BankAcct$ + String$(17 - Len(BankAcct$), 32)
      Trac = Trac + 1
      Trace$ = Str$(Trac): Trace$ = Right$(Trace$, Len(Trace$) - 1)
      If Len(Trace$) < 7 Then Trace$ = String$(7 - Len(Trace$), "0") + Trace$

      UBDraftRecord6(1).Field1 = "6"
      Select Case UBCustRec(1).AcctType
      Case "S"
        UBDraftRecord6(1).Field2 = "38"       ' Designates Savings Acct
      Case Else
        UBDraftRecord6(1).Field2 = "28"       ' Designates Checking Acct
      End Select
      UBDraftRecord6(1).Field3 = Left$(UBCustRec(1).TRANSIT, 8)
      UBDraftRecord6(1).Field4 = Right$(UBCustRec(1).TRANSIT, 1)
      UBDraftRecord6(1).Field5 = Left$(BankAcct$, 17)
      UBDraftRecord6(1).Field6 = "0000000000"   ' All zero's for Prenote
      UBDraftRecord6(1).Field7 = AcctNumber$
      UBDraftRecord6(1).Field8 = UCase$(nme$)
      UBDraftRecord6(1).Field9 = "  "
      UBDraftRecord6(1).Field10 = "0"
      UBDraftRecord6(1).Field11 = Left$(UBDraftRec(1).BANKORIG, 8) + Trace$
      Put DraftFileNum, Counter, UBDraftRecord6(1)
      hashh# = hashh# + Val(Left$(UBCustRec(1).TRANSIT, 8))
      UBCustRec(1).PreNoteFlag = 1
      Number = Number + 1
    End If
  Next cnt
  Close DraftFileNum
  Step3 = True
Return
TestProcessStep4:
  hash$ = Str$(hashh#)
  hash$ = Right$(hash$, Len(hash$) - 1)

  If Len(hash$) < 10 Then
    hash$ = String$(10 - Len(hash$), "0") + hash$
  End If
  If Len(hash$) > 10 Then
    hash$ = Right$(hash$, 10)
  End If

  If Len(Trace$) > 6 Then Trace$ = Right$(Trace$, 6)

  ReDim UBDraftRecord8(1) As UBDraftRecord8Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT8.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord8(1))
  UBDraftRecord8(1).Field1 = "8"
  UBDraftRecord8(1).Field2 = "200"
  UBDraftRecord8(1).Field3 = Trace$
  UBDraftRecord8(1).Field4 = hash$
  UBDraftRecord8(1).Field5 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord8(1).Field7 = FedIDNum$
  UBDraftRecord8(1).Field8 = String$(19, 32)    ' Reserved
  UBDraftRecord8(1).Field9 = String$(6, 32)     ' Reserved for Federal Reserve
  UBDraftRecord8(1).Field10 = Left$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord8(1).Field11 = "0000001"
  Put DraftFileNum, 1, UBDraftRecord8(1)
  Close DraftFileNum
  Step4 = True
  Return

TestProcessStep5:
  TotSize# = Val(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records

  If BlockSize! <> Int(BlockSize!) Then
    BlockSize! = Int(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize! - 1)))
  Else
    FillSize! = 0
  End If

  BlockSizeS$ = Str$(BlockSize!)
  BlockSizeS$ = Right$(BlockSizeS$, Len(BlockSizeS$) - 1)
  If Len(BlockSizeS$) < 6 Then BlockSizeS$ = String$(6 - Len(BlockSizeS$), "0") + BlockSizeS$
  If Len(Trace$) < 8 Then Trace$ = String$(8 - Len(Trace$), "0") + Trace$

  ReDim UBDraftRecord9(1) As UBDraftRecord9Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT9.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord9(1))
  UBDraftRecord9(1).Field1 = "9"
  UBDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  UBDraftRecord9(1).Field3 = BlockSizeS$
  UBDraftRecord9(1).Field4 = Trace$
  UBDraftRecord9(1).Field5 = hash$
  UBDraftRecord9(1).Field6 = "000000000000"     ' zero for prenote
  UBDraftRecord9(1).Field7 = "000000000000"
  UBDraftRecord9(1).Field8 = String$(39, 32)    ' Reserved
  Put DraftFileNum, 1, UBDraftRecord9(1)
  Close DraftFileNum

  ' Now Put Them Together In File Name UBDFNOTE
  outfile = FreeFile
  Open TestFileName$ For Output As outfile

  'OPEN "O", OutFile, TestFileName$
  
  ReDim UBDraftRecord1(1) As UBDraftRecord1Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT1.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord1(1))
  Get DraftFileNum, 1, UBDraftRecord1(1)
  Print #outfile, UBDraftRecord1(1).Field1;
  Print #outfile, UBDraftRecord1(1).Field2;
  Print #outfile, UBDraftRecord1(1).Field3;
  Print #outfile, UBDraftRecord1(1).Field4;
  Print #outfile, UBDraftRecord1(1).Field5;
  Print #outfile, UBDraftRecord1(1).Field6;
  Print #outfile, UBDraftRecord1(1).Field7;
  Print #outfile, UBDraftRecord1(1).Field8;
  Print #outfile, UBDraftRecord1(1).Field9;
  Print #outfile, UBDraftRecord1(1).Field10;
  Print #outfile, UBDraftRecord1(1).Field11;
  Print #outfile, UBDraftRecord1(1).Field12;
  Print #outfile, UBDraftRecord1(1).Field13
  Close DraftFileNum

  ReDim UBDraftRecord5(1) As UBDraftRecord5Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT5.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord5(1))
  Get DraftFileNum, 1, UBDraftRecord5(1)
  Print #outfile, UBDraftRecord5(1).Field1;
  Print #outfile, UBDraftRecord5(1).Field2;
  Print #outfile, UBDraftRecord5(1).Field3;
  Print #outfile, UBDraftRecord5(1).Field4;
  Print #outfile, UBDraftRecord5(1).Field5;
  Print #outfile, UBDraftRecord5(1).Field6;
  Print #outfile, UBDraftRecord5(1).Field7;
  Print #outfile, UBDraftRecord5(1).Field8;
  Print #outfile, UBDraftRecord5(1).Field9;
  Print #outfile, UBDraftRecord5(1).Field10;
  Print #outfile, UBDraftRecord5(1).Field11;
  Print #outfile, UBDraftRecord5(1).Field12;
  Print #outfile, UBDraftRecord5(1).Field13
  Close DraftFileNum
  ReDim UBDraftRecord6(1) As UBDraftRecord6Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94

  For cnt = 1 To NumOfLines
    Get DraftFileNum, cnt, UBDraftRecord6(1)
    Print #outfile, UBDraftRecord6(1).Field1;
    Print #outfile, UBDraftRecord6(1).Field2;
    Print #outfile, UBDraftRecord6(1).Field3;
    Print #outfile, UBDraftRecord6(1).Field4;
    Print #outfile, UBDraftRecord6(1).Field5;
    Print #outfile, UBDraftRecord6(1).Field6;
    Print #outfile, UBDraftRecord6(1).Field7;
    Print #outfile, UBDraftRecord6(1).Field8;
    Print #outfile, UBDraftRecord6(1).Field9;
    Print #outfile, UBDraftRecord6(1).Field10;
    Print #outfile, UBDraftRecord6(1).Field11
  Next cnt
  Close DraftFileNum
  ReDim UBDraftRecord8(1) As UBDraftRecord8Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT8.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord8(1))
  Get DraftFileNum, 1, UBDraftRecord8(1)
  Print #outfile, UBDraftRecord8(1).Field1;
  Print #outfile, UBDraftRecord8(1).Field2;
  Print #outfile, UBDraftRecord8(1).Field3;
  Print #outfile, UBDraftRecord8(1).Field4;
  Print #outfile, UBDraftRecord8(1).Field5;
  Print #outfile, UBDraftRecord8(1).Field6;
  Print #outfile, UBDraftRecord8(1).Field7;
  Print #outfile, UBDraftRecord8(1).Field8;
  Print #outfile, UBDraftRecord8(1).Field9;
  Print #outfile, UBDraftRecord8(1).Field10;
  Print #outfile, UBDraftRecord8(1).Field11
  Close DraftFileNum

  ReDim UBDraftRecord9(1) As UBDraftRecord9Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT9.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord9(1))
  Get DraftFileNum, 1, UBDraftRecord9(1)
  Print #outfile, UBDraftRecord9(1).Field1;
  Print #outfile, UBDraftRecord9(1).Field2;
  Print #outfile, UBDraftRecord9(1).Field3;
  Print #outfile, UBDraftRecord9(1).Field4;
  Print #outfile, UBDraftRecord9(1).Field5;
  Print #outfile, UBDraftRecord9(1).Field6;
  Print #outfile, UBDraftRecord9(1).Field7;
  Print #outfile, UBDraftRecord9(1).Field8
  Close DraftFileNum
  For cnt = 1 To FillSize! / 94
    Print #outfile, String$(94, "9")
  Next cnt
  Close
  Step5 = True
  Done = True
  Return

End Sub

