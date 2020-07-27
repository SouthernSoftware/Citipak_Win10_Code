VERSION 5.00
Begin VB.Form frmCitiPakExportData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Citi-Pak Export"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8625
   Icon            =   "frmCitiPakExportData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstReqFilesPayroll 
      Height          =   1815
      ItemData        =   "frmCitiPakExportData.frx":08CA
      Left            =   960
      List            =   "frmCitiPakExportData.frx":08D1
      TabIndex        =   6
      Top             =   1200
      Width           =   2745
   End
   Begin VB.ListBox lstMissingFilesListPayroll 
      Height          =   1815
      ItemData        =   "frmCitiPakExportData.frx":08D8
      Left            =   4920
      List            =   "frmCitiPakExportData.frx":08DF
      TabIndex        =   5
      Top             =   1200
      Width           =   2745
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Export Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3345
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CheckBox chkNCTaxBilling 
      Caption         =   "NC Tax Billing"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Citipak  NC Tax Data Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Required files"
      Height          =   210
      Left            =   1192
      TabIndex        =   1
      Top             =   720
      Width           =   2160
   End
   Begin VB.Label Label2 
      Caption         =   "Missing Files"
      Height          =   210
      Left            =   5272
      TabIndex        =   0
      Top             =   720
      Width           =   2160
   End
End
Attribute VB_Name = "frmCitiPakExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Dim ErrorCode As Integer
  'Private Temp_Class As Resize_Class

Private Sub chkNCTaxBilling_Click()
  Call FillFilesToConvertList
End Sub

Private Sub cmdProcess_Click()
  Dim RptNamex As String
  Dim RptHandlex As Integer
  Dim ThisFile As String
  lblInfo.FontSize = 12
  ErrorCode = 0

  ValidateFilesExists

  If ErrorCode = 0 Then
    If chkNCTaxBilling.Value = 1 Then
      GoTo DownInHere
      Call ProcessNCTaxCust
       Call ProcessNCTaxReal
      Call ProcessNCTaxPers2
      Call ProcessNCMortCodes
      Call ProcessNCOptSearches
      Call ProcessNCTownships
      Call ProcessNCSystemSetup
      Call ProcessNCMessages
      Call ProcessNCTaxBill
      'Call ProcessNCLateLetter
      Call ProcessNCRateTables
      'Call ProcessNCGLPay
      'Call ProcessNCGLBill
      Call ProcessNCTransHist
      Call ProcessNCOptRevRateTables
DownInHere:
      'Call ProcessNCLaserStandard
      Call ProcessNCBalance
      MsgBox ("Completed successfully.")
    End If
    
    cmdProcess.Enabled = False
    
    lblInfo.Caption = "You may now exit, export completed"
  End If
'    StartPath = App.Path
'    ThisFile = "\CompletedNormally.txt"

'    If chkVATaxBilling.Value = 1 Then
'      If DirExists(StartPath + "\VATAXData") Then
'        If Exist(StartPath + "\VATAXData\" + ThisFile) Then
'          KillFile (StartPath + "\VATAXData\" + ThisFile)
'        End If
'
'        RptNamex$ = StartPath + "\VATAXData\" + ThisFile
'        RptHandlex = FreeFile
'        Open RptNamex$ For Output As #RptHandlex
'        Print #RptHandlex, "GO";
'      Else
'        StartPath = App.Path
'        ThisFile = "\CompletedNormally.txt"
'        If DirExists(StartPath + "\VATAXData") Then
'          If Exist(StartPath + "\VATAXData\" + ThisFile) Then
'            KillFile (StartPath + "\VATAXData\" + ThisFile)
'          End If
'        End If
'        RptNamex$ = StartPath + "\VATAXData\" + ThisFile
'        RptHandlex = FreeFile
'        Open RptNamex$ For Output As #RptHandlex
'        Print #RptHandlex, "Stop";
'        Close RptHandlex
'
'        Exit Sub
'      End If
'    End If
    
'    If chkNCTaxBilling.Value = 1 Then
'      If DirExists(StartPath + "\NCTAXData") Then
'        If Exist(StartPath + "\NCTAXData\" + ThisFile) Then
'          KillFile (StartPath + "\NCTAXData\" + ThisFile)
'        End If
'
'        RptNamex$ = StartPath + "\NCTAXData\" + ThisFile
'        RptHandlex = FreeFile
'        Open RptNamex$ For Output As #RptHandlex
'        Print #RptHandlex, "GO";
'      Else
'        StartPath = App.Path
'        ThisFile = "\CompletedNormally.txt"
'        If DirExists(StartPath + "\NCTAXData") Then
'          If Exist(StartPath + "\NCTAXData\" + ThisFile) Then
'            KillFile (StartPath + "\NCTAXData\" + ThisFile)
'          End If
'        End If
'        RptNamex$ = StartPath + "\NCTAXData\" + ThisFile
'        RptHandlex = FreeFile
'        Open RptNamex$ For Output As #RptHandlex
'        Print #RptHandlex, "Stop";
'        Close RptHandlex
'
'        Exit Sub
'      End If
'    End If
'
'    If DirExists(StartPath + "\PayrollData") Then
'      If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'        KillFile (StartPath + "\PayrollData\" + ThisFile)
'      End If
'    End If
'
'    RptNamex$ = StartPath + "\PayrollData\" + ThisFile
'    RptHandlex = FreeFile
'    Open RptNamex$ For Output As #RptHandlex
'    Print #RptHandlex, "GO";
'    Else
'    StartPath = App.Path
'    ThisFile = "\CompletedNormally.txt"
'    If DirExists(StartPath + "\PayrollData") Then
'      If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'        KillFile (StartPath + "\PayrollData\" + ThisFile)
'      End If
'    End If
'
'    RptNamex$ = StartPath + "\PayrollData\" + ThisFile
'    RptHandlex = FreeFile
'    Open RptNamex$ For Output As #RptHandlex
'    Print #RptHandlex, "Stop";
'  End If
'  Close RptHandlex

End Sub

' Private Sub ProcessPrSys()
'  Dim SysRec As RegDSysFileRecType
'  Dim SysHandle As Integer
'  Dim NumOfSysRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'
'  OpenSysFile SysHandle
'  NumOfSysRecs = LOF(SysHandle) / Len(SysRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\PRSys.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfSysRecs
'    Get SysHandle, x, SysRec
'    Print #RptHandle, QPTrim$(SysRec.USEIMP);
'    Print #RptHandle, B & QPTrim$(SysRec.CashAcct);
'    Print #RptHandle, B & QPTrim$(SysRec.CITIDIR);
'    Print #RptHandle, B & QPTrim$(SysRec.SplitFlag);
'    Print #RptHandle, B & QPTrim$(SysRec.EXPMETHD);
'    If Len(SysRec.FRNGRATE) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", SysRec.FRNGRATE);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(SysRec.AcctCnt) > 0 Then
'      Print #RptHandle, B & Using$("########", SysRec.AcctCnt);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(SysRec.GLActLen) > 0 Then
'      Print #RptHandle, B & Using$("########", SysRec.GLActLen);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(SysRec.CheckStyle) > 0 Then
'      Print #RptHandle, B & Using$("########", SysRec.CheckStyle);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    Print #RptHandle, B & QPTrim$(SysRec.GLCheckYN) & B
'
'    FrmShowPctComp.ShowPctComp x, NumOfSysRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Unload FrmShowPctComp
' End Sub
'Private Sub ProcessEIC1RecType()
'  Dim EICRec As EICRecType
'  Dim EICHandle As Integer
'  Dim NumOfEICRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'
'  OpenEICFile EICHandle
'  NumOfEICRecs = LOF(EICHandle) / Len(EICRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\EICRecord.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  ' I am putting both records on one line
'  ' because in the table I am moving it in it is on one line
'    Get EICHandle, 1, EICRec
'    If Len(EICRec.EIC(1).EIC1OVR0) > 0 Then
'      Print #RptHandle, Using$("###,###.##", EICRec.EIC(1).EIC1OVR0);
'    Else
'      Print #RptHandle, 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1NVR0) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1NVR0);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1AMT0) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1AMT0);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1OVR1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1OVR1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1NVR1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1NVR1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1AMT1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1AMT1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1OVR2) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1OVR2);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1NVR2) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1NVR2);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1AMT2) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1AMT2);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1LESS) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1LESS);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(1).EIC1EXES) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(1).EIC1EXES);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1OVR0) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1OVR0);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1NVR0) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1NVR0);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1AMT0) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1AMT0);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1OVR1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1OVR1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1NVR1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1NVR1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1AMT1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1AMT1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1OVR2) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1OVR2);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1NVR2) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1NVR2);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1AMT2) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1AMT2);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1LESS) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1LESS);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1EXES) > 0 Then
'        Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1EXES);
'    Else
'        Print #RptHandle, B & 0;
'    End If
'    If Len(EICRec.EIC(2).EIC1EXES) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC(2).EIC1EXES) & B;
'    Else
'      Print #RptHandle, B & 0 & B;
'    End If
'
'    FrmShowPctComp.ShowPctComp x, NumOfEICRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'
'  Unload FrmShowPctComp
'  Close EICHandle
' End Sub
'
'Private Sub ProcessRetireRecType()
'  Dim RetireRec As RetireRecType
'  Dim RetireHandle As Integer
'  Dim NumOfRetireRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'
'  OpenRetFile RetireHandle
'  NumOfRetireRecs = LOF(RetireHandle) / Len(RetireRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\PRRetireRec.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRetireRecs
'    Get RetireHandle, x, RetireRec
'    Print #RptHandle, QPTrim$(RetireRec.TYPEDES1);
'    If Len(RetireRec.TYPEWH1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", RetireRec.TYPEWH1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    If Len(RetireRec.TYPEM1) > 0 Then
'      Print #RptHandle, B & Using$("###,###.##", RetireRec.TYPEM1);
'    Else
'      Print #RptHandle, B & 0;
'    End If
'    Print #RptHandle, B & QPTrim$(RetireRec.TYPEOT1);
'    Print #RptHandle, B & QPTrim$(RetireRec.TYPETD1) & B
'
'    FrmShowPctComp.ShowPctComp x, NumOfRetireRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close RetireHandle
' End Sub
'
'Private Sub ProcessUnitFileRecType()
'  Dim UnitRec As UnitFileRecType
'  Dim UnitHandle As Integer
'  Dim NumOfUnitRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'
'  OpenUnitFile UnitHandle
'  NumOfUnitRecs = LOF(UnitHandle) / Len(UnitRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\PRUnit.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfUnitRecs
'    Get UnitHandle, x, UnitRec
'    Print #RptHandle, QPTrim$(UnitRec.UFEMPR);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFADDR1);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFADDR2);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFCITY);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFSTATE);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFZIP);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFATTN);
'    Print #RptHandle, B & QPTrim$(UnitRec.ESCRemitNum);
'    Print #RptHandle, B & QPTrim$(UnitRec.ESCEmplrNum);
'    Print #RptHandle, B & Using$("########", UnitRec.ESCRTYPE);
'    Print #RptHandle, B & QPTrim$(UnitRec.LMT401YN);
'    Print #RptHandle, B & QPTrim$(UnitRec.BankDraft);
'    Print #RptHandle, B & QPTrim$(UnitRec.SSNOnCheck);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFFEDID);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFSTAID);
'    Print #RptHandle, B & Using$("###,###.##", UnitRec.TAXWBASE);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFRETID);
'    Print #RptHandle, B & QPTrim$(UnitRec.UFRETIDL);
'    Print #RptHandle, B & Using$("###,###.##", UnitRec.GMatch401K);
'    Print #RptHandle, B & Using$("###,###.##", UnitRec.LMatch401K);
'    Print #RptHandle, B & QPTrim$(UnitRec.BBTCNTNO);
'    Print #RptHandle, B & QPTrim$(UnitRec.BBTBATCH) & B
'    FrmShowPctComp.ShowPctComp x, NumOfUnitRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close UnitHandle
'  Unload FrmShowPctComp
' End Sub
' Private Sub ProcessDraftInfo()
'  Dim DraftRec As DraftInfoFileName
'  Dim DraftHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenPRDraftFile DraftHandle
'  NumOfRecs = LOF(DraftHandle) / Len(DraftRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\DraftInfo.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get DraftHandle, x, DraftRec
'    Print #RptHandle, QPTrim$(DraftRec.BankName);
'    Print #RptHandle, B & QPTrim$(DraftRec.BANKDEST);
'    Print #RptHandle, B & QPTrim$(DraftRec.BANKORIG);
'    Print #RptHandle, B & QPTrim$(DraftRec.BANKLOC);
'    Print #RptHandle, B & QPTrim$(DraftRec.FEDPREFX);
'    Print #RptHandle, B & QPTrim$(DraftRec.FEDID) & B
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close DraftHandle
'  Unload FrmShowPctComp
' End Sub
' Private Sub ProcessErnCodeRecType()
'  Dim ErnRec As ErnCodeRecType
'  Dim ErnHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenErnCodeFile ErnHandle
'  NumOfRecs = LOF(ErnHandle) / Len(ErnRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\EarnCodes.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get ErnHandle, x, ErnRec
'    If QPTrim(ErnRec.ERNCODE1) <> "" Then
'        Print #RptHandle, QPTrim$(ErnRec.ERNCODE1);
'        Print #RptHandle, B & QPTrim$(ErnRec.ERNFWT1);
'        Print #RptHandle, B & QPTrim$(ErnRec.ERNSWT1);
'        Print #RptHandle, B & QPTrim$(ErnRec.ERNSOC1);
'        Print #RptHandle, B & QPTrim$(ErnRec.ERNMED1);
'        Print #RptHandle, B & QPTrim$(ErnRec.EarnYN);
'        Print #RptHandle, B & QPTrim$(ErnRec.ERNRET1) & B
'
'    End If
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close ErnHandle
'  Unload FrmShowPctComp
' End Sub
' Private Sub ProcessDedCodeRecType()
'  Dim DedRec As DedCodeRecType
'  Dim DedHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'  Dim z As String
'
'  OpenDedCodeFile DedHandle
'  NumOfRecs = LOF(DedHandle) / Len(DedRec)
'
'  StartPath = App.Path
'
'  B = "|"
'  z = "-"
'
'  ThisFile = "\DedCodes.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get DedHandle, x, DedRec
'    Print #RptHandle, QPTrim$(DedRec.DCDESC1);
'    Print #RptHandle, B & QPTrim$(DedRec.DCFWT1);
'    Print #RptHandle, B & QPTrim$(DedRec.DCSWT1);
'    Print #RptHandle, B & QPTrim$(DedRec.DCSOC1);
'    Print #RptHandle, B & QPTrim$(DedRec.DCMED1);
'    Print #RptHandle, B & z & QPTrim$(DedRec.DCACCT1) & B
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close DedHandle
'  Unload FrmShowPctComp
' End Sub
'  Private Sub ProcessAccrualDates()
'  Dim AccRec As AccrualDates
'  Dim AccHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenAccrualDatesFile AccHandle
'  NumOfRecs = LOF(AccHandle) / Len(AccRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\AccrualDates.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get AccHandle, x, AccRec
'    Print #RptHandle, MakeRegDate(AccRec.PreviousDate);
'    Print #RptHandle, B & MakeRegDate(AccRec.CurrentDate) & B
'
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close AccHandle
'  Unload FrmShowPctComp
' End Sub
'Private Sub ProcessPayRateType()
'  Dim PayRateRec As PayRateType
'  Dim PayRateHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenPayRateFile PayRateHandle
'  NumOfRecs = LOF(PayRateHandle) / Len(PayRateRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\PayRate.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get PayRateHandle, x, PayRateRec
'      Print #RptHandle, Using$("########", PayRateRec.EmpRecNum);
'      Print #RptHandle, B & QPTrim$(PayRateRec.EmpNo);
'      Print #RptHandle, B & QPTrim$(PayRateRec.EmpLName);
'      Print #RptHandle, B & QPTrim$(PayRateRec.EmpFName);
'      Print #RptHandle, B & MakeRegDate(PayRateRec.EMPHDATE);
'      Print #RptHandle, B & MakeRegDate(PayRateRec.EMPTDATE);
'
'    For Y = 1 To 30
'      If Y < 30 Then
'        If Val(PayRateRec.RegPayRate(Y)) <> 0 Then
'          Print #RptHandle, B & Using$("##,####.##", PayRateRec.RegPayRate(Y));
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(PayRateRec.OTPayRate(Y)) <> 0 Then
'          Print #RptHandle, B & Using$("###,###.##", PayRateRec.OTPayRate(Y));
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPFREQ(Y));
'        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPTYPE(Y));
'        If Val(PayRateRec.PayChngDate(Y)) <> 0 Then
'          Print #RptHandle, B & MakeRegDate(PayRateRec.PayChngDate(Y));
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        Print #RptHandle, B & QPTrim$(PayRateRec.EMPJOB(Y));
'      Else
'        If Val(PayRateRec.RegPayRate(Y)) <> 0 Then
'          Print #RptHandle, B & Using$("##,####.##", PayRateRec.RegPayRate(Y));
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(PayRateRec.OTPayRate(Y)) <> 0 Then
'          Print #RptHandle, B & Using$("###,###.##", PayRateRec.OTPayRate(Y));
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPFREQ(Y));
'        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPTYPE(Y));
'        If Val(PayRateRec.PayChngDate(Y)) <> 0 Then
'          Print #RptHandle, B & MakeRegDate(PayRateRec.PayChngDate(Y));
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        Print #RptHandle, B & QPTrim$(PayRateRec.EMPJOB(Y)) & B
'      End If
'    Next Y
'
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close PayRateHandle
'  Unload FrmShowPctComp
' End Sub
' Private Sub ProcessPRMessRecType()
'  Dim MessRec As PRMessRecType
'  Dim MessRecHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenEmpMessage MessRecHandle
'  NumOfRecs = LOF(MessRecHandle) / Len(MessRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\MessRec.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get MessRecHandle, x, MessRec
'    If Val(MessRec.EmpRec) <> 0 Then
'      Print #RptHandle, Using$("########", MessRec.EmpRec);
'
'      For Y = 1 To 15
'        If Y < 15 Then
'          Print #RptHandle, B & QPTrim$(MessRec.MessLine(Y).Msg) & B;
'        Else
'          Print #RptHandle, B & QPTrim$(MessRec.MessLine(Y).Msg) & B
'        End If
'      Next Y
'    End If
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close MessRecHandle
'  Unload FrmShowPctComp
' End Sub
'
' Private Sub ProcessOrbitEmpData()
'  Dim OrbitRec As OrbitEmpData
'  Dim OrbitHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenOrbEmpDataFile OrbitHandle
'  NumOfRecs = LOF(OrbitHandle) / Len(OrbitRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\OrbitEmpData.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get OrbitHandle, x, OrbitRec
'    Print #RptHandle, Using$("####", OrbitRec.EmpRecNum);
'    Print #RptHandle, B & QPTrim$(OrbitRec.EmpNum);
'    Print #RptHandle, B & QPTrim$(OrbitRec.RecType);
'    Print #RptHandle, B & QPTrim$(OrbitRec.AgencyNum);
'    Print #RptHandle, B & QPTrim$(OrbitRec.PlanCode);
'    Print #RptHandle, B & QPTrim$(OrbitRec.SSN);
'    Print #RptHandle, B & QPTrim$(OrbitRec.MemberID);
'    Print #RptHandle, B & QPTrim$(OrbitRec.FirstName);
'    Print #RptHandle, B & QPTrim$(OrbitRec.MiddleName);
'    Print #RptHandle, B & QPTrim$(OrbitRec.LastName);
'    Print #RptHandle, B & QPTrim$(OrbitRec.Suffix);
'    Print #RptHandle, B & QPTrim$(OrbitRec.Gender);
'    If Val(QPTrim$((OrbitRec.DateOfBirth))) > 0 Then
'      Print #RptHandle, B & MakeMDYDate(OrbitRec.DateOfBirth);
'    Else
'      Print #RptHandle, B & "";
'    End If
'    Print #RptHandle, B & QPTrim$(OrbitRec.AddLine1);
'    Print #RptHandle, B & QPTrim$(OrbitRec.AddLine2);
'    Print #RptHandle, B & QPTrim$(OrbitRec.City);
'    Print #RptHandle, B & QPTrim$(OrbitRec.State);
'    Print #RptHandle, B & QPTrim$(OrbitRec.Zip);
'    Print #RptHandle, B & QPTrim$(OrbitRec.OutOfCntryAdd);
'    Print #RptHandle, B & QPTrim$(OrbitRec.JobClass);
'    Print #RptHandle, B & QPTrim$(OrbitRec.DeptNum);
'    If Val(QPTrim$((OrbitRec.EmployDate))) > 0 Then
'      Print #RptHandle, B & MakeMDYDate(OrbitRec.EmployDate);
'    Else
'      Print #RptHandle, B & "";
'    End If
'    If Val(QPTrim$(OrbitRec.EligibleDate)) > 0 Then
'      Print #RptHandle, B & MakeMDYDate(OrbitRec.EligibleDate);
'    Else
'      Print #RptHandle, B & "";
'    End If
'    Print #RptHandle, B & QPTrim$(OrbitRec.PayType);
'    Print #RptHandle, B & QPTrim$(OrbitRec.VacHours);
'    Print #RptHandle, B & QPTrim$(OrbitRec.ContrPdEmpPrd);
'    If Val(OrbitRec.ContrPdEmpBegDate) > 0 Then
'      Print #RptHandle, B & MakeMDYDate(OrbitRec.ContrPdEmpBegDate);
'    Else
'      Print #RptHandle, B & "";
'    End If
'    If Val(QPTrim(OrbitRec.ContrPdEmpEndDate)) > 0 Then
'      Print #RptHandle, B & MakeMDYDate(OrbitRec.ContrPdEmpEndDate);
'    Else
'      Print #RptHandle, B & "";
'    End If
'    If Val(QPTrim(OrbitRec.TerminationDate)) > 0 Then
'      Print #RptHandle, B & MakeMDYDate(OrbitRec.TerminationDate);
'    Else
'      Print #RptHandle, B & "";
'    End If
'    Print #RptHandle, B & QPTrim$(OrbitRec.TermType);
'    Print #RptHandle, B & QPTrim$(OrbitRec.SharedPosition);
'    If OrbitRec.Deleted = True Then
'      Print #RptHandle, B & "Y";
'    Else
'      Print #RptHandle, B & "N";
'    End If
'    Print #RptHandle, B & QPTrim$(OrbitRec.Adjustment) & B
'
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close OrbitHandle
'  Unload FrmShowPctComp
' End Sub
'  Private Sub ProcessVoidedCheckType()
'  Dim VoidRec As VoidCheckType
'  Dim VoidHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenTempVoidFile VoidHandle
'  NumOfRecs = LOF(VoidHandle) / Len(VoidRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\VoidedCheck.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get VoidHandle, x, VoidRec
'    Print #RptHandle, Using$("##########", VoidRec.CheckNum);
'    Print #RptHandle, B & QPTrim$(VoidRec.EmpNum);
'    Print #RptHandle, B & MakeRegDate(VoidRec.CheckDate);
'    Print #RptHandle, B & Using$("###,###.##", VoidRec.CheckAmt) & B
'
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close VoidHandle
'  Unload FrmShowPctComp
' End Sub
'Private Sub ProcessW2ElectronicsSubRa()
'  Dim W2Rec As W2ElectronicSubRA
'  Dim W2Handle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenW2ESubRA W2Handle
'  NumOfRecs = LOF(W2Handle) / Len(W2Rec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\W2ESubRA.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get W2Handle, x, W2Rec
'    Print #RptHandle, QPTrim$(W2Rec.EINNum);
'    Print #RptHandle, B & QPTrim$(W2Rec.PersIDNum);
'    Print #RptHandle, B & QPTrim$(W2Rec.ResubID);
'    Print #RptHandle, B & QPTrim$(W2Rec.ReSubWFID);
'    Print #RptHandle, B & QPTrim$(W2Rec.SftwrCode);
'    Print #RptHandle, B & QPTrim$(W2Rec.CmpnyName);
'    Print #RptHandle, B & QPTrim$(W2Rec.LocAddr);
'    Print #RptHandle, B & QPTrim$(W2Rec.DelAddr);
'    Print #RptHandle, B & QPTrim$(W2Rec.City);
'    Print #RptHandle, B & QPTrim$(W2Rec.State);
'    Print #RptHandle, B & QPTrim$(W2Rec.Zip);
'    Print #RptHandle, B & QPTrim$(W2Rec.ZipExt);
'    Print #RptHandle, B & QPTrim$(W2Rec.SubmttrName);
'    Print #RptHandle, B & QPTrim$(W2Rec.SubLocAddr);
'    Print #RptHandle, B & QPTrim$(W2Rec.SubDelAddr);
'    Print #RptHandle, B & QPTrim$(W2Rec.SubCity);
'    Print #RptHandle, B & QPTrim$(W2Rec.SubState);
'    Print #RptHandle, B & QPTrim$(W2Rec.SubZip);
'    Print #RptHandle, B & QPTrim$(W2Rec.SubZipExt);
'    Print #RptHandle, B & QPTrim$(W2Rec.ContactName);
'    Print #RptHandle, B & QPTrim$(W2Rec.CntctPhone);
'    Print #RptHandle, B & QPTrim$(W2Rec.CntPhnExt);
'    Print #RptHandle, B & QPTrim$(W2Rec.CntEMail);
'    Print #RptHandle, B & QPTrim$(W2Rec.CntFAX);
'    Print #RptHandle, B & QPTrim$(W2Rec.CntMethod);
'    Print #RptHandle, B & QPTrim$(W2Rec.PrepCode) & B
'
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close W2Handle
'  Unload FrmShowPctComp
' End Sub
'
'Private Sub validateDedCodes()
'  Dim DedRec As DedCodeRecType
'  Dim DedHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim Y As Long
'  Dim wrkCnt As Integer
'  Dim wrkDesc As String
'
'  OpenDedCodeFile DedHandle
'  NumOfRecs = LOF(DedHandle) / Len(DedRec)
'
'  Dim dedArray(1 To 50) As String
'
'  For x = 1 To 50
'    dedArray(x) = ""
'  Next x
'
'  For x = 1 To NumOfRecs
'    Get DedHandle, x, DedRec
'    dedArray(x) = DedRec.DCDESC1
'  Next x
'
'  wrkCnt = 0
'  For x = 1 To 50
'    If wrkCnt <= 1 Then
'      If QPTrim(dedArray(x)) <> "" Then
'          wrkCnt = 0
'          wrkDesc = dedArray(x)
'
'          For Y = 1 To 50
'            If Y <> x Then
'                If dedArray(Y) = wrkDesc Then
'                  wrkCnt = wrkCnt + 1
'                End If
'            End If
'          Next
'      Else
'        wrkCnt = 0
'      End If
'    Else
'      ErrorCode = 1
'      MsgBox ("Duplicate deductions codes were found. This must be changed before the conversion can be run")
'      Exit Sub
'    End If
'  Next x
'  Close DedHandle
'End Sub
'Private Sub ValidateErnCodes()
'  Dim ErnRec As ErnCodeRecType
'  Dim ErnHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim Y As Integer
'  Dim B As String
'  Dim wrkCnt As Integer
'  Dim wrkDesc As String
'
'  OpenErnCodeFile ErnHandle
'  NumOfRecs = LOF(ErnHandle) / Len(ErnRec)
'
'  Dim ErnArray(1 To 20) As String
'
'  For x = 1 To 20
'    ErnArray(x) = ""
'  Next x
'
'  For x = 1 To NumOfRecs
'    Get ErnHandle, x, ErnRec
'    ErnArray(x) = ErnRec.ERNCODE1
'  Next x
'
'  wrkCnt = 0
'  For x = 1 To 20
'    If wrkCnt <= 1 Then
'      If QPTrim(ErnArray(x)) <> "" Then
'          wrkCnt = 0
'          wrkDesc = ErnArray(x)
'
'          For Y = 1 To 20
'            If Y <> x Then
'                If ErnArray(Y) = wrkDesc Then
'                wrkCnt = wrkCnt + 1
'                End If
'            End If
'          Next
'      Else
'        wrkCnt = 0
'      End If
'    Else
'      ErrorCode = 1
'      MsgBox ("Duplicate earn codes were found. This must be changed before the conversion can be run")
'      Exit Sub
'    End If
'  Next x
'
'  Close ErnHandle
'
' End Sub
 
Private Sub chkFixedAssets_Click()
    Call FillFilesToConvertList
End Sub

Private Sub chkPayroll_Click()
    Call FillFilesToConvertList
End Sub

Private Sub chkVehDec_Click()
    Call FillFilesToConvertList
End Sub

Private Sub FillFilesToConvertList()
  'lstReqFilesPayroll.Clear
'  If chkPayroll.Value = 1 Then
'    lstReqFilesPayroll.AddItem ("PREMP1.DAT                Payroll")
'    lstReqFilesPayroll.AddItem ("PREMP2.DAT")
'    lstReqFilesPayroll.AddItem ("PREMP3.DAT")
'    lstReqFilesPayroll.AddItem ("PRTRANSH.DAT")
'    lstReqFilesPayroll.AddItem ("PRSYS.DAT")
'    lstReqFilesPayroll.AddItem ("PREICTBL.DAT")
'    lstReqFilesPayroll.AddItem ("PRRETIRE.DAT")
'    lstReqFilesPayroll.AddItem ("PRUNIT.DAT")
'    lstReqFilesPayroll.AddItem ("PRDRAFTI.DAT")
'    lstReqFilesPayroll.AddItem ("PRERNCOD.DAT")
'    lstReqFilesPayroll.AddItem ("PRDEDCOD.DAT")
'    lstReqFilesPayroll.AddItem ("PRACCRUE.DAT")
'    lstReqFilesPayroll.AddItem ("PAYRATE.DAT")
'    lstReqFilesPayroll.AddItem ("EMPMESS.DAT")
'    lstReqFilesPayroll.AddItem ("OrbEmpData.DAT")
'    lstReqFilesPayroll.AddItem ("TEMPVOID.DAT")
'    lstReqFilesPayroll.AddItem ("W2ESUBRA.DAT")
'  End If
'
'  If chkFixedAssets = 1 Then
'    lstReqFilesPayroll.AddItem (FASetUpFileName + "      Fixed Assets")
'    lstReqFilesPayroll.AddItem (FAItemFileName)
'    lstReqFilesPayroll.AddItem (FAAssetCodeName)
'    lstReqFilesPayroll.AddItem (FADeptCodeName)
'    lstReqFilesPayroll.AddItem (FAFundCodeName)
'    lstReqFilesPayroll.AddItem (FADprHistFileName)
'  End If
  
'  If chkVehDec = 1 Then
'    lstReqFilesPayroll.AddItem (DCSetupFile + "     Vehicle Decal")
'    lstReqFilesPayroll.AddItem (DCTranFile)
'    lstReqFilesPayroll.AddItem (DCCustFile)
'    lstReqFilesPayroll.AddItem (DCVCodeFile)
'    lstReqFilesPayroll.AddItem (DCVehFile)
'  End If
'
'  If chkVATaxBilling.Value = 1 Then
''    lstReqFilesPayroll.AddItem ("TXPOPTSH.DAT")
''    lstReqFilesPayroll.AddItem ("TXROPTSH.DAT")
'    lstReqFilesPayroll.AddItem ("TAXCUST.DAT")
'    lstReqFilesPayroll.AddItem ("TAXPROP.DAT")
'    lstReqFilesPayroll.AddItem ("TAXPERS.DAT")
''    lstReqFilesPayroll.AddItem ("TAXMORT.DAT")
''    lstReqFilesPayroll.AddItem ("TXCOPTSH.DAT")
''    lstReqFilesPayroll.AddItem ("TXTWNSHP.DAT")
'    lstReqFilesPayroll.AddItem ("TAXTRANS.DAT")
'    lstReqFilesPayroll.AddItem ("TAXSETUP.DAT")
''    lstReqFilesPayroll.AddItem ("LSRPITEM.DAT")
''    lstReqFilesPayroll.AddItem ("LSRRITEM.DAT")
''    lstReqFilesPayroll.AddItem ("TAXMESS.DAT")
''    lstReqFilesPayroll.AddItem ("TXLATLTR.DAT")
'     lstReqFilesPayroll.AddItem ("TXPENRTB.DAT")
' End If
  
  If chkNCTaxBilling.Value = 1 Then
    lstReqFilesPayroll.AddItem ("TAXCUST.DAT")
    lstReqFilesPayroll.AddItem ("TAXPROP.DAT")
    lstReqFilesPayroll.AddItem ("TAXPERS.DAT")
    lstReqFilesPayroll.AddItem ("TAXTRANS.DAT")
    lstReqFilesPayroll.AddItem ("TAXSETUP.DAT")
  End If
  
End Sub

'Private Sub ProcessGlFundRecType()
'  Dim GLFundRec As GLFundRecType
'  Dim GlFundHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenFundFile GlFundHandle
'  NumOfRecs = LOF(GlFundHandle) / Len(GLFundRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\GLFund.txt"
'  If Exist(StartPath + "\GeneralLedgerData\" + ThisFile) Then
'    KillFile (StartPath + "\GeneralLedgerData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\GeneralLedgerData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "GL Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get GlFundHandle, x, GLFundRec
'    If GLFundRec.Deleted = 0 Then
'      Print #RptHandle, "False";
'    Else
'      Print #RptHandle, "True";
'    End If
'
'    Print #RptHandle, B & QPTrim$(GLFundRec.FundNum);
'    Print #RptHandle, B & QPTrim$(GLFundRec.Title) & B
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close GlFundHandle
'  Close RptHandle
'  Unload FrmShowPctComp
' End Sub
 
' Private Sub ProcessGlAcctRec()
'  Dim GLAcctRec As GLAcctRecType
'  Dim GlAcctHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenAcctFile GlAcctHandle
'  NumOfRecs = LOF(GlAcctHandle) / Len(GLAcctRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\GLAcct.txt"
'  If Exist(StartPath + "\GeneralLedgerData\" + ThisFile) Then
'    KillFile (StartPath + "\GeneralLedgerData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\GeneralLedgerData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "GL Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get GlAcctHandle, x, GLAcctRec
'    Print #RptHandle, QPTrim$(GLAcctRec.Num); '1
'    Print #RptHandle, B & QPTrim$(GLAcctRec.Title); '2
'    Print #RptHandle, B & QPTrim$(GLAcctRec.Typ); '3
'    If GLAcctRec.Deleted = 0 Then
'      Print #RptHandle, B & "False" & B '4
'    Else
'      Print #RptHandle, B & "True" & B '4
'    End If
'
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close GlAcctHandle
'  Close RptHandle
'  Unload FrmShowPctComp
' End Sub
 
' Private Sub ProcessGLTrans()
'  Dim GLTransRec As GLTransRecType
'  Dim GlTransHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenGLTransFile GlTransHandle
'  NumOfRecs = LOF(GlTransHandle) / Len(GLTransRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\GLTrans.txt"
'  If Exist(StartPath + "\GeneralLedgerData\" + ThisFile) Then
'    KillFile (StartPath + "\GeneralLedgerData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\GeneralLedgerData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "GL Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    If Len(GLTransRec.TRDATE) > 0 Then
'        Get GlTransHandle, x, GLTransRec
'        Print #RptHandle, QPTrim$(GLTransRec.AcctNum);
'        Print #RptHandle, B & MakeRegDate(GLTransRec.TRDATE);
'        Print #RptHandle, B & QPTrim$(GLTransRec.Desc);
'        Print #RptHandle, B & QPTrim$(GLTransRec.LDesc);
'        Print #RptHandle, B & Using$("###,###,###,###.##", GLTransRec.CrAmt);
'        Print #RptHandle, B & Using$("###,###,###,###.##", GLTransRec.DrAmt);
'        Print #RptHandle, B & QPTrim$(GLTransRec.Ref);
'        Print #RptHandle, B & QPTrim$(GLTransRec.Src) & B
'
'    End If
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close GlTransHandle
'  Close RptHandle
'  Unload FrmShowPctComp
'End Sub

'Private Sub ProcessGLBudgetTrans()
'  Dim GLBudgetTransRec As GLTransRecType
'  Dim GlBudgetTransHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenGLBudgetTransFile GlBudgetTransHandle
'  NumOfRecs = LOF(GlBudgetTransHandle) / Len(GLBudgetTransRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\GLBudgetTrans.txt"
'  If Exist(StartPath + "\GeneralLedgerData\" + ThisFile) Then
'    KillFile (StartPath + "\GeneralLedgerData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\GeneralLedgerData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "GL Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    If Len(GLBudgetTransRec.TRDATE) > 0 Then
'        Get GlBudgetTransHandle, x, GLBudgetTransRec
'        Print #RptHandle, QPTrim$(GLBudgetTransRec.AcctNum);
'        Print #RptHandle, B & MakeRegDate(GLBudgetTransRec.TRDATE);
'        Print #RptHandle, B & QPTrim$(GLBudgetTransRec.Desc);
'        Print #RptHandle, B & QPTrim$(GLBudgetTransRec.LDesc);
'        Print #RptHandle, B & Using$("###,###,###,###.##", GLBudgetTransRec.CrAmt);
'        Print #RptHandle, B & Using$("###,###,###,###.##", GLBudgetTransRec.DrAmt);
'        Print #RptHandle, B & "BA";
'        Print #RptHandle, B & QPTrim$(GLBudgetTransRec.Src) & B
'
'
'    End If
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close GlBudgetTransHandle
'  Close RptHandle
'  Unload FrmShowPctComp
'End Sub
 
' Private Sub ProcessGlAcctRecForBudgetPrep()
'  Dim GLAcctRec As GLAcctRecType
'  Dim GlAcctHandle As Integer
'  Dim NumOfRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  OpenAcctFile GlAcctHandle
'  NumOfRecs = LOF(GlAcctHandle) / Len(GLAcctRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\GLAcctForBP.txt"
'  If Exist(StartPath + "\GeneralLedgerData\" + ThisFile) Then
'    KillFile (StartPath + "\GeneralLedgerData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\GeneralLedgerData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "GL Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfRecs
'    Get GlAcctHandle, x, GLAcctRec
'    If Val(GLAcctRec.NYEst) <> 0 Or Val(GLAcctRec.NYReq) <> 0 Or Val(GLAcctRec.NYRec) <> 0 Or Val(GLAcctRec.NYApp) <> 0 Then
'        Print #RptHandle, QPTrim$(GLAcctRec.Num); '1
'        Print #RptHandle, B & QPTrim$(GLAcctRec.Typ); '2
'        If Val(GLAcctRec.Bal) <> 0 Then
'          Print #RptHandle, B & Using$("###,###.##", GLAcctRec.Bal); '3
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(GLAcctRec.PYAct) <> 0 Then
'          Print #RptHandle, B & Using$("###,###.##", GLAcctRec.PYAct); '4
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(GLAcctRec.Bgt) <> 0 Then
'          Print #RptHandle, B & Using("###,###.##", GLAcctRec.Bgt); '5
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(GLAcctRec.NYEst) <> 0 Then
'        Print #RptHandle, B & Using$("###,###.##", GLAcctRec.NYEst); '6
'         Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(GLAcctRec.NYReq) <> 0 Then
'          Print #RptHandle, B & Using$("###,###.##", GLAcctRec.NYReq); '7
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(GLAcctRec.NYRec) <> 0 Then
'          Print #RptHandle, B & Using$("###,###.##", GLAcctRec.NYRec); '8
'        Else
'          Print #RptHandle, B & 0;
'        End If
'        If Val(GLAcctRec.NYApp) <> 0 Then
'          Print #RptHandle, B & Using$("###,###.##", GLAcctRec.NYApp) & B '9
'        Else
'          Print #RptHandle, B & 0 & B
'        End If
'   End If
'
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close GlAcctHandle
'  Close RptHandle
'  Unload FrmShowPctComp
'End Sub
'Private Sub ProcessK401DedType()
'  Dim K401Rec As K401DedType
'  Dim K401Handle As Integer
'  Dim NumOfK401Recs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'  Open401KDedFile K401Handle
'  NumOfK401Recs = LOF(K401Handle) / Len(K401Rec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\K401Ded.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  For x = 1 To NumOfK401Recs
'    Get K401Handle, x, K401Rec
'    Print #RptHandle, QPTrim$(K401Rec.VolDed);
'    Print #RptHandle, B & QPTrim$(K401Rec.LPDed);
'    Print #RptHandle, B & QPTrim$(K401Rec.RothDed) & B
'
'
'    FrmShowPctComp.ShowPctComp x, NumOfK401Recs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close K401Handle
'  Unload FrmShowPctComp
' End Sub
' Private Sub ProcessLeaveBenefits()
'  Dim LBenRec As LeaveRecType
'  Dim LBenHandle As Integer
'  Dim NumOfLBenRecs As Long
'  Dim x As Long
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim ThisFile As String
'  Dim Y As Integer
'  Dim B As String
'
'
'  OpenLeaveFileName LBenHandle
'  NumOfLBenRecs = LOF(LBenHandle) / Len(LBenRec)
'
'  StartPath = App.Path
'
'  B = "|"
'
'  ThisFile = "\LBen.txt"
'  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
'    KillFile (StartPath + "\PayrollData\" + ThisFile)
'  End If
'
'  RptName$ = StartPath + "\PayrollData\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Payroll Export"
'  FrmShowPctComp.Show , Me
'  DoEvents
'
'  Dim Dash As String
'  Dim wrk_LeaveIdx As String
'  Dash = "-"
'
'  Dim wrk_TableNumber As Integer
'  wrk_TableNumber = 0
'  For x = 1 To NumOfLBenRecs
'    Get LBenHandle, x, LBenRec
'    wrk_TableNumber = wrk_TableNumber + 1
'
'    For Y = 1 To 20
'      If Y <= 9 Then
'        wrk_LeaveIdx = QPTrim(Str(wrk_TableNumber)) + Dash + "A" + QPTrim(Str(Y))
'      Else
'        wrk_LeaveIdx = QPTrim(Str(wrk_TableNumber)) + Dash + "B" + QPTrim(Str(Y))
'      End If
'      Print #RptHandle, wrk_LeaveIdx;
'      Print #RptHandle, B & wrk_TableNumber;
'      Print #RptHandle, B & Using$("########", LBenRec.VacMax);
'      Print #RptHandle, B & Using$("########", LBenRec.VEntry(Y).YEARS);
'      Print #RptHandle, B & Using$("######.##", LBenRec.VEntry(Y).EARN);
'      Print #RptHandle, B & Using$("########", LBenRec.SICKMAX);
'      Print #RptHandle, B & Using$("########", LBenRec.SEntry(Y).YEARS);
'      Print #RptHandle, B & Using$("######.##", LBenRec.SEntry(Y).EARN);
'      Print #RptHandle, B & Using$("########", LBenRec.HolMax);
'      Print #RptHandle, B & Using$("########", LBenRec.HEntry(Y).YEARS);
'      Print #RptHandle, B & Using$("######.##", LBenRec.HEntry(Y).EARN);
'      Print #RptHandle, B & Using$("########", LBenRec.PerMax);
'      Print #RptHandle, B & Using$("########", LBenRec.PEntry(Y).YEARS);
'      Print #RptHandle, B & Using$("######.##", LBenRec.PEntry(Y).EARN) & B
'    Next Y
'
'    FrmShowPctComp.ShowPctComp x, NumOfLBenRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'  Next x
'  Close LBenHandle
'  Unload FrmShowPctComp
' End Sub
Private Sub chkGeneralLedger_Click()

End Sub
 
 Private Sub ValidateFilesExists()
 Dim bolModuleSelected As Boolean
 bolModuleSelected = False
 lstMissingFilesListPayroll.Clear
 
 'Validate Payroll Files if the user selects the payroll module
'  If chkPayroll.Value = 1 Then
'   bolModuleSelected = True
'   If Not Exist(PRData + EmpData1Name) = True Then
'     ErrorCode = 1
'     lstMissingFilesListPayroll.AddItem ("PREMP1.DAT")
'   End If
'
'   If Not Exist(PRData + EmpData2Name) = True Then
'     ErrorCode = 1
'     lstMissingFilesListPayroll.AddItem ("PREMP2.DAT")
'   End If
'
'   If Not Exist(PRData + EmpData3Name) = True Then
'     ErrorCode = 1
'     lstMissingFilesListPayroll.AddItem ("PREMP3.DAT")
'   End If
'
'   If Not Exist(PRData + TransHistFileName) = True Then
'     ErrorCode = 1
'     lstMissingFilesListPayroll.AddItem ("PRTRANSH.DAT")
'   End If
'
'   If Not Exist(PRData + SysFileName) = True Then
'     ErrorCode = 1
'     lstMissingFilesListPayroll.AddItem ("PRSYS.DAT")
'   End If
'
'   If Not Exist(PRData + EICFileName) = True Then
'     ErrorCode = 1
'     lstMissingFilesListPayroll.AddItem ("PREICTBL.DAT")
'   End If
'
'   If Not Exist(PRData + RetireFileName) = True Then
'     ErrorCode = 1
'     lstMissingFilesListPayroll.AddItem ("PRRETIRE.DAT")
'   End If
'
'    If Not Exist(PRData + UnitFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("PRUNIT.DAT")
'    End If
'
'    If Not Exist(PRData + PRDraftFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("PRDRAFTI.DAT")
'    End If
'
'    If Not Exist(PRData + ErnCodeFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("PRERNCOD.DAT")
'    End If
'
'    If Not Exist(PRData + DedCodeFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("PRDEDCOD.DAT")
'    End If
'
'    If Not Exist(PRData + AccrueFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("PRACCRUE.DAT")
'    End If
'
'    If Not Exist(PRData + PayRateName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("PAYRATE.DAT")
'    End If
'
'    If Not Exist(PRData + MessageName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("EMPMESS.DAT")
'    End If
'
'    If Not Exist(PRData + OrbitEmpDataName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("OrbEmpData.DAT")
'    End If
'
'    If Not Exist(PRData + TempVoidFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("TEMPVOID.DAT")
'    End If
'
'    If Not Exist(PRData + W2ESubRA) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("W2ESUBRA.DAT")
'    End If
'  End If
  
'Validate GL files exists
'  If chkGeneralLedger.Value = 1 Then
'    bolModuleSelected = True
'    If Not Exist(GLData + GLFundFile) = True Then
'      ErrorCode = 1
'      lstMissingFilesListGL.AddItem ("GlFund.DAT")
'    End If
'    If Not Exist(GLData + GlAcctFile) = True Then
'      ErrorCode = 1
'      lstMissingFilesListGL.AddItem ("GlAcct.DAT")
'    End If
'    If Not Exist(GLData + TransFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListGL.AddItem ("GlTrans.DAT")
'    End If
'  End If
  
'Validate FA files exists
'  If chkFixedAssets.Value = 1 Then
'    bolModuleSelected = True
'    If Not Exist(FAData + FASetUpFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (FASetUpFileName)
'    End If
'    If Not Exist(FAData + FAItemFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (FAItemFileName)
'    End If
'    If Not Exist(FAData + FAAssetCodeName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (FAAssetCodeName)
'    End If
'    If Not Exist(FAData + FADeptCodeName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (FADeptCodeName)
'    End If
'    If Not Exist(FAData + FAFundCodeName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (FAFundCodeName)
'    End If
'    If Not Exist(FAData + FADprHistFileName) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (FADprHistFileName)
'    End If
'  End If

'Validate DC files exists
'  If chkVehDec.Value = 1 Then
'    bolModuleSelected = True
'    If Not Exist(DCData + DCCustFile) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (DCCustFile)
'    End If
'    If Not Exist(DCData + DCTranFile) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (DCTranFile)
'    End If
'    If Not Exist(DCData + DCSetupFile) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (DCSetupFile)
'    End If
'    If Not Exist(DCData + DCVCodeFile) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (DCVCodeFile)
'    End If
'    If Not Exist(DCData + DCVehFile) = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem (DCVehFile)
'    End If
'  End If
'
'  If chkVATaxBilling.Value = 1 Then
'    bolModuleSelected = True
''    If Not Exist(VATAXData + "TXPOPTSH.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TXPOPTSH.DAT")
''    End If
''    If Not Exist(VATAXData + "TXROPTSH.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TXROPTSH.DAT")
''    End If
'    If Not Exist(VATAXData + "TAXCUST.DAT") = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("TAXCUST.DAT")
'    End If
'    If Not Exist(VATAXData + "TAXPROP.DAT") = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("TAXPROP.DAT")
'    End If
'    If Not Exist(VATAXData + "TAXPERS.DAT") = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("TAXPERS.DAT")
'    End If
''    If Not Exist(VATAXData + "TAXMORT.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TAXMORT.DAT")
''    End If
''    If Not Exist(VATAXData + "TXCOPTSH.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TXCOPTSH.DAT")
''    End If
''    If Not Exist(VATAXData + "TXTWNSHP.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TXTWNSHP.DAT")
''    End If
'    If Not Exist(VATAXData + "TAXTRANS.DAT") = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("TAXTRANS.DAT")
'    End If
'    If Not Exist(VATAXData + "TAXSETUP.DAT") = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("TAXSETUP.DAT")
'    End If
''    If Not Exist(VATAXData + "LSRPITEM.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("LSRPITEM.DAT")
''    End If
''    If Not Exist(VATAXData + "LSRRITEM.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("LSRRITEM.DAT")
''    End If
''    If Not Exist(VATAXData + "TAXMESS.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TAXMESS.DAT")
''    End If
''    If Not Exist(VATAXData + "TXLATLTR.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TXLATLTR.DAT")
''    End If
''    If Not Exist(VATAXData + "TXRTTBLS.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TXRTTBLS.DAT")
''    End If
'    If Not Exist(VATAXData + "TXPENRTB.DAT") = True Then
'      ErrorCode = 1
'      lstMissingFilesListPayroll.AddItem ("TXPENRTB.DAT")
'    End If
''    If Not Exist(VATAXData + "TAXRGLBAC.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TAXRGLBAC.DAT")
''    End If
''    If Not Exist(VATAXData + "TAXPGLBAC.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TAXPGLBAC.DAT")
''    End If
''    If Not Exist(VATAXData + "TAXRGLACT.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TAXRGLACT.DAT")
''    End If
''    If Not Exist(VATAXData + "TAXPGLACT.DAT") = True Then
''      ErrorCode = 1
''      lstMissingFilesListPayroll.AddItem ("TAXPGLACT.DAT")
''    End If
'  End If
  
  If chkNCTaxBilling.Value = 1 Then
    bolModuleSelected = True
    If Not Exist(NCTAXData + "TAXCUST.DAT") = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem ("TAXCUST.DAT")
    End If
    If Not Exist(NCTAXData + "TAXPROP.DAT") = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem ("TAXPROP.DAT")
    End If
    If Not Exist(NCTAXData + "TAXPERS.DAT") = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem ("TAXPERS.DAT")
    End If
    If Not Exist(NCTAXData + "TAXTRANS.DAT") = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem ("TAXTRANS.DAT")
    End If
    If Not Exist(NCTAXData + "TAXSETUP.DAT") = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem ("TAXSETUP.DAT")
    End If
  End If

  'Validate modules were selected
  If bolModuleSelected = False Then
     ErrorCode = 1
     lblInfo.Caption = "No modules were selected to export"
     lstReqFilesPayroll.Clear
     lstMissingFilesListPayroll.Clear
     'lstReqFilesGL.Clear
     'lstMissingFilesListGL.Clear
  End If
  
  'Display messege if now files were found
  If ErrorCode = 1 And bolModuleSelected = True Then
    lblInfo.Caption = "Required files were not found, export is aborted"
  End If
End Sub


Private Sub Form_Load()
  chkNCTaxBilling.Value = Checked
  cmdProcess.Enabled = True
End Sub

