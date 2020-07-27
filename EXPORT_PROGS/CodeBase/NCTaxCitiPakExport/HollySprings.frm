VERSION 5.00
Begin VB.Form frmExportPayroll 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Export"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11640
   Icon            =   "HollySprings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Press To Begin Conversion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4253
      TabIndex        =   0
      Top             =   3930
      Width           =   3135
   End
End
Attribute VB_Name = "frmExportPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Dim ErrorCode As Integer
  'Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload Me
  Shell "CitiPakExp.exe", vbNormalFocus
End Sub

Private Sub cmdProcess_Click()
  ErrorCode = 0
  Call DetermineIfAllFilesExistsBeforeUpdate
  'Call validateDedCodes
  'Call ValidateErnCodes
  If ErrorCode = 0 Then
    Call ProcessEmployeeData
    Call ProcessPrSys
    Call ProcessEIC1RecType
    Call ProcessRetireRecType
    Call ProcessUnitFileRecType
    Call ProcessDraftInfo
    Call ProcessErnCodeRecType
    Call ProcessDedCodeRecType
    Call ProcessAccrualDates
    Call ProcessPayRateType
    Call ProcessPRMessRecType
    Call ProcessOrbitEmpData
    Call ProcessVoidedCheckType
    Call ProcessW2ElectronicsSubRa
   End If
  Exit Sub
End Sub
 Private Sub DetermineIfAllFilesExistsBeforeUpdate()
   If Not Exist(PRData + EmpData1Name) = True Then
    MsgBox ("The file 'PRData\PREMP1.DAT' cannot be found. It is required to run this program.")
    Exit Sub
  End If
    
  If Not Exist(PRData + EmpData2Name) = True Then
    ErrorCode = 1
    MsgBox ("The file 'PRData\PREMP2.DAT' cannot be found. It is required to run this program.")
    Exit Sub
  End If
  
  If Not Exist(PRData + EmpData3Name) = True Then
    ErrorCode = 1
    MsgBox ("The file 'PRData\PREMP3.DAT' cannot be found. It is required to run this program.")
    Exit Sub
  End If
  
  If Not Exist(PRData + TransHistFileName) = True Then
    ErrorCode = 1
    MsgBox ("The file 'PRData\PRTRANSH.DAT' cannot be found. It is required to run this program.")
    Exit Sub
  End If
     
  If Not Exist(PRData + SysFileName) = True Then
    ErrorCode = 1
    MsgBox ("The file 'PRData\PRSYS.DAT' cannot be found. It is required to run this program.")
    Exit Sub
  End If
  
  If Not Exist(PRData + EICFileName) = True Then
    ErrorCode = 1
    MsgBox ("The file 'PRData\PREICTBL.DAT' cannot be found. It is required to run this program.")
    Exit Sub
  End If
  If Not Exist(PRData + RetireFileName) = True Then
    ErrorCode = 1
    MsgBox ("The file 'PRData\PRRETIRE.DAT' cannot be found. It is required to run this program.")
    Exit Sub
  End If
   If Not Exist(PRData + UnitFileName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\PRUNIT.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + PRDraftFileName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\PRDRAFTI.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + ErnCodeFileName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\PRERNCOD.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + DedCodeFileName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\PRDEDCOD.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + AccrueFileName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\PRACCRUE.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + PayRateName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\PAYRATE.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + MessageName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\EMPMESS.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + OrbitEmpDataName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\OrbEmpData.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + TempVoidFileName) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\TEMPVOID.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
   If Not Exist(PRData + W2ESubRA) = True Then
     ErrorCode = 1
     MsgBox ("The file 'PRData\W2ESUBRA.DAT' cannot be found. It is required to run this program.")
     Exit Sub
   End If
  End Sub
 Private Sub ProcessEmployeeData()
  Dim Emp1Rec As EmpData1Type
  Dim EmpRec As EmpData2Type
  Dim Emp3Rec As EmpData3Type
  Dim DraftRec As PRPPDraftInfoType
  Dim E1Handle As Integer
  Dim E2Handle As Integer
  Dim E3Handle As Integer
  Dim THandle As Integer
  Dim NumOfE1Recs As Long
  Dim NumOfERecs As Long
  Dim NumOfE3Recs As Long
  Dim NumOfDraftInfoRecs
  Dim TransRec As TransRecType
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim Rpt1Handle As Integer
  Dim RptHandle As Integer
  Dim Rpt1Name As String
  Dim RptName As String
  Dim Rpt3Name As String
  Dim RptHandleT As Integer
  Dim Rpt3Handle As Integer
  Dim RptNameT As String
  Dim RptHandleM As Integer
  Dim RptNameM As String
  Dim ThisFile1 As String
  Dim ThisFile As String
  Dim ThisFileT As String
  Dim ThisFile3 As String
  Dim ThisFileM As String
  Dim NextRec As Long
  Dim B As String
  Dim DRec As DedCodeRecType
  Dim DHandle As Integer
  Dim NumOfDRecs As Integer
  Dim AERec As ErnCodeRecType
  Dim NumOfAERecs As Integer
  Dim AE2Handle As Integer
  Dim Y As Integer
  Dim Message As String
  Dim FedStats As String
  Dim StateStats As String
  Dim IncludeOT As String
  

  OpenEmpData1File E1Handle
  NumOfE1Recs = LOF(E1Handle) / Len(Emp1Rec)
    
  OpenEmpData2File E2Handle
  NumOfERecs = LOF(E2Handle) / Len(EmpRec)
  
  OpenEmpData3File E3Handle
  NumOfE3Recs = LOF(E3Handle) / Len(Emp3Rec)
         
  StartPath = App.Path
  
  B = "|"
    
  ThisFile1 = "\EmployeeData1.txt"
  ThisFile = "\EmployeeData.txt"
  ThisFile3 = "\EmployeeData3.txt"
  ThisFileT = "\TransData.txt"
  If DirExists(StartPath + "\PayrollData") Then
    If Exist(StartPath + "\PayrollData\" + ThisFile1) Then
      KillFile (StartPath + "\PayrollData\" + ThisFile1)
    End If
    If Exist(StartPath + "\PayrollData\" + ThisFile) Then
      KillFile (StartPath + "\PayrollData\" + ThisFile)
    End If
    If Exist(StartPath + "\PayrollData\" + ThisFileT) Then
      KillFile (StartPath + "\PayrollData\" + ThisFileT)
    End If
    If Exist(StartPath + "\PayrollData\" + ThisFile3) Then
      KillFile (StartPath + "\PayrollData\" + ThisFile3)
    End If
  Else
    MkDir StartPath + "\PayrollData"
  End If
  
  Rpt1Name$ = StartPath + "\PayrollData\" + ThisFile1
  Rpt1Handle = FreeFile
  Open Rpt1Name$ For Output As #Rpt1Handle
  
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  Rpt3Name$ = StartPath + "\PayrollData\" + ThisFile3
  Rpt3Handle = FreeFile
  Open Rpt3Name$ For Output As #Rpt3Handle
   
  RptNameT$ = StartPath + "\PayrollData\" + ThisFileT
  RptHandleT = FreeFile
  Open RptNameT$ For Output As #RptHandleT
  
  OpenTransHistFile THandle
  NumOfTRecs = LOF(THandle) / Len(TransRec)
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfERecs
    Get E2Handle, x, EmpRec
    NextRec = EmpRec.LastTransRec
    Do While NextRec > 0
        Get THandle, NextRec, TransRec
        Print #RptHandleT, Using$("##########", TransRec.EmpPin); '1
        Print #RptHandleT, B & Using$("##########", TransRec.CheckNum); '2
        Print #RptHandleT, B & MakeRegDate(TransRec.PayPdStart); '3
        Print #RptHandleT, B & MakeRegDate(TransRec.PayPdEnd); '4
        Print #RptHandleT, B & MakeRegDate(TransRec.CheckDate); '5
        Print #RptHandleT, B & MakeRegDate(TransRec.PostDate); '6
        Print #RptHandleT, B & QPTrim$(TransRec.PayType); '7
        Print #RptHandleT, B & Using$("#.####", TransRec.BaseRate); '8
        Print #RptHandleT, B & Using$("#.####", TransRec.OTRate); '9
        Print #RptHandleT, B & Using$("####.##", TransRec.RegHrsWork); '10
        Print #RptHandleT, B & Using$("####.##", TransRec.VacUsed); '11
        Print #RptHandleT, B & Using$("####.##", TransRec.SickUsed); '12
        Print #RptHandleT, B & Using$("####.##", TransRec.CompUsed); '13
        Print #RptHandleT, B & Using$("####.##", TransRec.HOLHOURS); ' 14
        Print #RptHandleT, B & Using$("####.##", TransRec.PerHours); '15
        Print #RptHandleT, B & Using$("####.##", TransRec.RegHrsPaid); '16
        Print #RptHandleT, B & Using$("####.##", TransRec.OTHours); ' 17
        Print #RptHandleT, B & Using$("####.##", TransRec.OTHrsPaid); '18
        Print #RptHandleT, B & Using$("####.##", TransRec.OT2Comp); '19
        Print #RptHandleT, B & Using$("###,###.##", TransRec.TotRegWage); '20
        Print #RptHandleT, B & Using$("####.##", TransRec.TotOTWage); '21
        Print #RptHandleT, B & Using$("###,###.##", TransRec.GrossWage); '22
        Print #RptHandleT, B & Using$("####.##", TransRec.TotAdditEarn); '23
        Print #RptHandleT, B & Using$("###,###,###.##", TransRec.GrossPay); '24
        Print #RptHandleT, B & Using$("####.##", TransRec.SocGrossPay); '25
        Print #RptHandleT, B & Using$("####.##", TransRec.MedGrossPay); ' 26
        Print #RptHandleT, B & Using$("####.##", TransRec.FedGrossPay); '27
        Print #RptHandleT, B & Using$("####.##", TransRec.StaGrossPay); '28
        Print #RptHandleT, B & Using$("####.##", TransRec.SocTaxAmt); '29
        Print #RptHandleT, B & Using$("####.##", TransRec.MedTaxAmt); '30
        Print #RptHandleT, B & Using$("####.##", TransRec.FedTaxAmt); '31
        Print #RptHandleT, B & Using$("####.##", TransRec.StaTaxAmt); '32
        Print #RptHandleT, B & Using$("####.##", TransRec.TotTaxAmt); '33
        Print #RptHandleT, B & Using$("####.##", TransRec.RetireAmt); '34
        Print #RptHandleT, B & Using$("####.##", TransRec.TotDedAmt); '35
        Print #RptHandleT, B & Using$("###,###,###.##", TransRec.EICAmt); '36
        Print #RptHandleT, B & Using$("###,###,###.##", TransRec.NetPay); '37
        Print #RptHandleT, B & Using$("####.##", TransRec.MatchRetAmt); '38
        Print #RptHandleT, B & Using$("####.##", TransRec.MatchSocAmt); '39
        Print #RptHandleT, B & Using$("####.##", TransRec.MatchMedAmt); '40
        Print #RptHandleT, B & Using$("####.##", TransRec.RetGrossPay); '41
        Print #RptHandleT, B & Using$("###,###,###.##", TransRec.TaxFring); '42
        Print #RptHandleT, B & Using$("##########", TransRec.VoidRec); '43
        Print #RptHandleT, B & QPTrim(EmpRec.EmpFName); '44
        Print #RptHandleT, QPTrim(EmpRec.EmpLName); '44
        Print #RptHandleT, B & QPTrim(EmpRec.EmpNo); '45
        
        
        For Y = 1 To 33 '46 Through 78
          If TransRec.DAmt(Y) = 0 Then
            Print #RptHandleT, B & 0;
          Else
            Print #RptHandleT, B & Using$("####.##", TransRec.DAmt(Y));
          End If
        Next Y
        
        For Y = 1 To 8 '79 Thourgh 118
          If TransRec.TDist(Y).DRHrs = 0 Then
            Print #RptHandleT, B & 0;
          Else
            Print #RptHandleT, B & Using$("####.##", TransRec.TDist(Y).DRHrs);
          End If
          If TransRec.TDist(Y).DOHrs = 0 Then
            Print #RptHandleT, B & 0;
          Else
            Print #RptHandleT, B & Using$("####.##", TransRec.TDist(Y).DOHrs);
          End If
          If TransRec.TDist(Y).DPct = 0 Then
           Print #RptHandleT, B & 0;
          Else
            Print #RptHandleT, B & Using$("####.##", TransRec.TDist(Y).DPct);
          End If
          If TransRec.TDist(Y).DRWage = 0 Then
            Print #RptHandleT, B & 0;
          Else
            Print #RptHandleT, B & Using$("####.##", TransRec.TDist(Y).DRWage);
          End If
          If TransRec.TDist(Y).DOWage = 0 Then
            Print #RptHandleT, B & 0;
          Else
            Print #RptHandleT, B & Using$("####.##", TransRec.TDist(Y).DOWage);
          End If
        Next Y
        For Y = 1 To 6 ' 119 Through 130
          If Y < 6 Then
            Print #RptHandleT, B & QPTrim$(TransRec.EDist(Y).EAcct);
            Print #RptHandleT, B & Using$("####.##", TransRec.EDist(Y).EAmt);
          Else
            Print #RptHandleT, B & QPTrim$(TransRec.EDist(Y).EAcct);
            Print #RptHandleT, B & Using$("####.##", TransRec.EDist(Y).EAmt) & B
          End If
        Next Y
      NextRec = TransRec.PrevTransRec
    Loop
    
    'Employee 1 data
    Get E1Handle, x, Emp1Rec
    If Emp1Rec.EmpNo <> "" Then
      Print #Rpt1Handle, QPTrim$(Emp1Rec.EmpNo);
    Else
      Print #Rpt1Handle, B & 0;
    End If
    Print #Rpt1Handle, B & QPTrim$(Emp1Rec.EmpLName);
    Print #Rpt1Handle, B & QPTrim$(Emp1Rec.EmpFName);
    If Emp1Rec.Data1RecNum <> 0 Then
      Print #Rpt1Handle, B & Using$("########", Emp1Rec.Data1RecNum);
    Else
      Print #Rpt1Handle, B & 0;
    End If
    If Emp1Rec.TransRecNum <> 0 Then
      Print #Rpt1Handle, B & Using$("#######", Emp1Rec.TransRecNum);
    Else
      Print #Rpt1Handle, B & 0;
    End If
    If Emp1Rec.Deleted = 1 Then
      Print #Rpt1Handle, B & "N" & B
    Else
      Print #Rpt1Handle, B & "Y" & B
    End If
    
    
    'Employee 3 data
    Get E3Handle, x, Emp3Rec
    Print #Rpt3Handle, Using$("########", EmpRec.EmpNo);
    
    If Emp3Rec.YTDGrossPay <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDGrossPay);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDSocGrossPay <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDSocGrossPay);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDMedGrossPay <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDMedGrossPay);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDFedGrossPay <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDFedGrossPay);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDStaGrossPay <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDStaGrossPay);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDOTPay <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDOTPay);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDRegPay <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDRegPay);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDNet <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDNet);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDSocial <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDSocial);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDMedicare <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDMedicare);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDFederal <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDFederal);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDState <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDState);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDRetire <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDRetire);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDDAmtT <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDDAmtT);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDEarnT <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDEarnT);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDEIC <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDEIC);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDOther2 <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDOther2);
    Else
      Print #Rpt3Handle, B & 0;
    End If

    
    For Y = 1 To 50
      If Emp3Rec.YTDDAmt(Y) <> 0 Then
        Print #Rpt3Handle, B & Y;
        Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDDAmt(Y));
      Else
        Print #Rpt3Handle, B & Y;
        Print #Rpt3Handle, B & 0;
      End If
    Next Y

    If Emp3Rec.YTDEarn1 <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDEarn1);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDEarn2 <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDEarn2);
    Else
      Print #Rpt3Handle, B & 0;
    End If
    If Emp3Rec.YTDEarn3 <> 0 Then
      Print #Rpt3Handle, B & Using$("###,###.##", Emp3Rec.YTDEarn3) & B
    Else
      Print #Rpt3Handle, B & 0 & B
    End If
       
    
     
    'Employee Table
    Print #RptHandle, QPTrim$(EmpRec.EmpNo);
    Print #RptHandle, B & QPTrim$(EmpRec.EmpSSN);
    Print #RptHandle, B & QPTrim$(EmpRec.EmpLName);
    Print #RptHandle, B & QPTrim$(EmpRec.EmpFName);
    Print #RptHandle, B & QPTrim$(EmpRec.EmpAddr1);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPADDR2);
    Print #RptHandle, B & QPTrim$(EmpRec.EmpCity);
    Print #RptHandle, B & QPTrim$(EmpRec.EmpState);
    Print #RptHandle, B & QPTrim$(EmpRec.EmpZip);
    Print #RptHandle, B & MakeRegDate(EmpRec.EMPBDAY);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPGENDR);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPRACE);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPRETNO);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPRETTP);
    
    If EmpRec.DRAFTCOD <> "" Then
      Print #RptHandle, B & "Y";
    Else
      Print #RptHandle, B & "N";
    End If
        
    Print #RptHandle, B & QPTrim$(EmpRec.EMPJOB);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPWCCLS);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPSTATS);
    Print #RptHandle, B & Using$("###.##", EmpRec.EMPBCODE);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPPTYPE);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPPFREQ);
    Print #RptHandle, B & Using$("##,###.##", EmpRec.EMPPRATE);
    Print #RptHandle, B & Using$("##,###.##", EmpRec.EMPORATE);
        
    Print #RptHandle, B & MakeRegDate(EmpRec.EMPHDATE);
    Print #RptHandle, B & MakeRegDate(EmpRec.EMPRDATE);
    Print #RptHandle, B & MakeRegDate(EmpRec.EMPTDATE);
        
    Print #RptHandle, B & QPTrim$(EmpRec.EMPFEDX);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPFEDO2);
    Print #RptHandle, B & Using$("###.##", EmpRec.EMPFEDO1);
      
    Print #RptHandle, B & QPTrim$(EmpRec.EMPFEDS);
    Print #RptHandle, B & Using$("###", EmpRec.EMPFEDA);
    Print #RptHandle, B & Using$("###.##", EmpRec.EMPFEDAA);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPSTAX);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPSTAO2);
    Print #RptHandle, B & Using$("###.##", EmpRec.EMPSTAO1);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPSTAS);
    Print #RptHandle, B & Using$("###", EmpRec.EMPSTAA);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPSTAAA);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPSOCX);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPMEDX);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPEIC);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPVACE);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPVUSED);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPVBAL);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPSLE);
      
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPSLUSE);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPSLBAL);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPCTE);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPCTUSE);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPCTBAL);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.PERERN);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.PerUsed);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.PERBAL);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.HOLERN);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.HolUsed);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.HOLBAL);
    Print #RptHandle, B & Using$("##", EmpRec.LeaveTbl);
    Print #RptHandle, B & QPTrim$(EmpRec.ExcludeESC);

    If EmpRec.EMPTDATE = 0 Or EmpRec.Deleted = 0 Then
      Print #RptHandle, B & "Y";
    Else
      Print #RptHandle, B & "N";
    End If
    
    Print #RptHandle, B & Using$("########", EmpRec.EmpPin);
    Print #RptHandle, B & QPTrim$(EmpRec.YN401K);
    Print #RptHandle, B & QPTrim$(EmpRec.PrimeDept);
    Print #RptHandle, B & QPTrim$(EmpRec.Comment);
    Print #RptHandle, B & QPTrim$(EmpRec.DRAFTCOD);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPDDACC);
    Print #RptHandle, B & QPTrim$(EmpRec.PRENOTED);
    Print #RptHandle, B & QPTrim$(EmpRec.BankName);
    Print #RptHandle, B & QPTrim$(EmpRec.BANKLOC);
    Print #RptHandle, B & QPTrim$(EmpRec.TRANSIT);
    
    For Y = 1 To 33
      If QPTrim$(EmpRec.EmpDed(Y).DPct) = "AMOUNT" Then
        If EmpRec.EmpDed(Y).DAmt > 0 Then
          Print #RptHandle, B & Using$("###.##", EmpRec.EmpDed(Y).DAmt);
          Print #RptHandle, B & 0;
        Else
          Print #RptHandle, B & 0;
          Print #RptHandle, B & 0;
        End If
      Else
        If EmpRec.EmpDed(Y).DAmt > 0 Then
          Print #RptHandle, B & 0;
          Print #RptHandle, B & Using$("###.##", EmpRec.EmpDed(Y).DAmt);
        Else
          Print #RptHandle, B & 0;
          Print #RptHandle, B & 0;
        End If
      End If
      Print #RptHandle, B & QPTrim$(EmpRec.EmpDed(Y).DOTI);
    Next Y

    Print #RptHandle, B & QPTrim$(EmpRec.EMPEACT1);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPEAMT1);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPEACT2);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPEAMT2);
    Print #RptHandle, B & QPTrim$(EmpRec.EMPEACT3);
    Print #RptHandle, B & Using$("#,###.##", EmpRec.EMPEAMT3);
    
    For Y = 1 To 8
      Print #RptHandle, B & QPTrim$(EmpRec.EDist(Y).DAcct);
      If EmpRec.EDist(Y).DAmt <> 0 Then
        Print #RptHandle, B & Using$("#,###.##", EmpRec.EDist(Y).DAmt);
      Else
       Print #RptHandle, B & 0;
      End If
    Next Y
    
    Print #RptHandle, B & QPTrim$(EmpRec.HomePhone);
    Print #RptHandle, B & QPTrim$(EmpRec.EmrgncyCntctName);
    Print #RptHandle, B & QPTrim$(EmpRec.EmrgncyCntctPhnNum);
    Print #RptHandle, B & QPTrim$(EmpRec.EmrgncyCntctRelation) & B

  

    FrmShowPctComp.ShowPctComp x, NumOfERecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload FrmShowPctComp
  
End Sub
 Private Sub ProcessPrSys()
  Dim SysRec As RegDSysFileRecType
  Dim SysHandle As Integer
  Dim NumOfSysRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    

  OpenSysFile SysHandle
  NumOfSysRecs = LOF(SysHandle) / Len(SysRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\PRSys.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
  
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfSysRecs
    Get SysHandle, x, SysRec
    Print #RptHandle, QPTrim$(SysRec.USEIMP);
    Print #RptHandle, B & QPTrim$(SysRec.CashAcct);
    Print #RptHandle, B & QPTrim$(SysRec.CITIDIR);
    Print #RptHandle, B & QPTrim$(SysRec.SplitFlag);
    Print #RptHandle, B & QPTrim$(SysRec.EXPMETHD);
    If Len(SysRec.FRNGRATE) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", SysRec.FRNGRATE);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(SysRec.AcctCnt) > 0 Then
      Print #RptHandle, B & Using$("########", SysRec.AcctCnt);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(SysRec.GLActLen) > 0 Then
      Print #RptHandle, B & Using$("########", SysRec.GLActLen);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(SysRec.CheckStyle) > 0 Then
      Print #RptHandle, B & Using$("########", SysRec.CheckStyle);
    Else
      Print #RptHandle, B & 0;
    End If
    Print #RptHandle, B & QPTrim$(SysRec.GLCheckYN) & B
    
    FrmShowPctComp.ShowPctComp x, NumOfSysRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
 End Sub
Private Sub ProcessEIC1RecType()
  Dim EICRec As EIC1RecType
  Dim EICHandle As Integer
  Dim NumOfEICRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    

  OpenEICFile EICHandle
  NumOfEICRecs = LOF(EICHandle) / Len(EICRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\EICRecord.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
  
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  ' I am putting both records on one line
  ' because in the table I am moving it in it is on one line
  For x = 1 To 2
    Get EICHandle, x, EICRec
    If x = 1 Then
      If Len(EICRec.EIC1OVR0) > 0 Then
        Print #RptHandle, Using$("###,###.##", EICRec.EIC1OVR0);
      Else
        Print #RptHandle, 0;
      End If
    Else
      If Len(EICRec.EIC1OVR0) > 0 Then
        Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1OVR0);
      Else
        Print #RptHandle, B & 0;
      End If
    End If
    If Len(EICRec.EIC1NVR0) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1NVR0);
    Else
      Print #RptHandle, B & 0;
    End If
        If Len(EICRec.EIC1AMT0) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1AMT0);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(EICRec.EIC1OVR1) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1OVR1);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(EICRec.EIC1NVR1) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1NVR1);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(EICRec.EIC1AMT1) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1AMT1);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(EICRec.EIC1OVR2) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1OVR2);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(EICRec.EIC1NVR2) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1NVR2);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(EICRec.EIC1AMT2) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1AMT2);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(EICRec.EIC1LESS) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1LESS);
    Else
      Print #RptHandle, B & 0;
    End If
    If x = 1 Then
      If Len(EICRec.EIC1EXES) > 0 Then
        Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1EXES);
      Else
        Print #RptHandle, B & 0;
      End If
    Else
      If Len(EICRec.EIC1EXES) > 0 Then
        Print #RptHandle, B & Using$("###,###.##", EICRec.EIC1EXES) & B;
      Else
        Print #RptHandle, B & 0 & B;
      End If
    End If
    
    FrmShowPctComp.ShowPctComp x, NumOfEICRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
 End Sub
 
Private Sub ProcessRetireRecType()
  Dim RetireRec As RetireRecType
  Dim RetireHandle As Integer
  Dim NumOfRetireRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    

  OpenSysFile RetireHandle
  NumOfRetireRecs = LOF(RetireHandle) / Len(RetireRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\PRRetireRec.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
  
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRetireRecs
    Get RetireHandle, x, RetireRec
    Print #RptHandle, QPTrim$(RetireRec.TYPEDES1);
    If Len(RetireRec.TYPEWH1) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", RetireRec.TYPEWH1);
    Else
      Print #RptHandle, B & 0;
    End If
    If Len(RetireRec.TYPEM1) > 0 Then
      Print #RptHandle, B & Using$("###,###.##", RetireRec.TYPEM1);
    Else
      Print #RptHandle, B & 0;
    End If
    Print #RptHandle, B & QPTrim$(RetireRec.TYPEOT1);
    Print #RptHandle, B & QPTrim$(RetireRec.TYPETD1) & B
    
    FrmShowPctComp.ShowPctComp x, NumOfRetireRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close RetireHandle
 End Sub
 
Private Sub ProcessUnitFileRecType()
  Dim UnitRec As UnitFileRecType
  Dim UnitHandle As Integer
  Dim NumOfUnitRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    

  OpenUnitFile UnitHandle
  NumOfUnitRecs = LOF(UnitHandle) / Len(UnitRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\PRUnit.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
  
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfUnitRecs
    Get UnitHandle, x, UnitRec
    Print #RptHandle, QPTrim$(UnitRec.UFEMPR);
    Print #RptHandle, B & QPTrim$(UnitRec.UFADDR1);
    Print #RptHandle, B & QPTrim$(UnitRec.UFADDR2);
    Print #RptHandle, B & QPTrim$(UnitRec.UFCITY);
    Print #RptHandle, B & QPTrim$(UnitRec.UFSTATE);
    Print #RptHandle, B & QPTrim$(UnitRec.UFZIP);
    Print #RptHandle, B & QPTrim$(UnitRec.UFATTN);
    Print #RptHandle, B & QPTrim$(UnitRec.ESCRemitNum);
    Print #RptHandle, B & QPTrim$(UnitRec.ESCEmplrNum);
    Print #RptHandle, B & Using$("########", UnitRec.ESCRTYPE);
    Print #RptHandle, B & QPTrim$(UnitRec.LMT401YN);
    Print #RptHandle, B & QPTrim$(UnitRec.BankDraft);
    Print #RptHandle, B & QPTrim$(UnitRec.SSNOnCheck);
    Print #RptHandle, B & QPTrim$(UnitRec.UFFEDID);
    Print #RptHandle, B & QPTrim$(UnitRec.UFSTAID);
    Print #RptHandle, B & Using$("###,###.##", UnitRec.TAXWBASE);
    Print #RptHandle, B & QPTrim$(UnitRec.UFRETID);
    Print #RptHandle, B & QPTrim$(UnitRec.UFRETIDL);
    Print #RptHandle, B & Using$("###,###.##", UnitRec.GMatch401K);
    Print #RptHandle, B & Using$("###,###.##", UnitRec.LMatch401K);
    Print #RptHandle, B & QPTrim$(UnitRec.BBTCNTNO);
    Print #RptHandle, B & QPTrim$(UnitRec.BBTBATCH) & B
    
    
    
    FrmShowPctComp.ShowPctComp x, NumOfUnitRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close UnitHandle
  Unload FrmShowPctComp
 End Sub
 Private Sub ProcessDraftInfo()
  Dim DraftRec As DraftInfoFileName
  Dim DraftHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenPRDraftFile DraftHandle
  NumOfRecs = LOF(DraftHandle) / Len(DraftRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\DraftInfo.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get DraftHandle, x, DraftRec
    Print #RptHandle, QPTrim$(DraftRec.BankName);
    Print #RptHandle, B & QPTrim$(DraftRec.BANKDEST);
    Print #RptHandle, B & QPTrim$(DraftRec.BANKORIG);
    Print #RptHandle, B & QPTrim$(DraftRec.BANKLOC);
    Print #RptHandle, B & QPTrim$(DraftRec.FEDPREFX);
    Print #RptHandle, B & QPTrim$(DraftRec.FEDID) & B
       
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close DraftHandle
  Unload FrmShowPctComp
 End Sub
 Private Sub ProcessErnCodeRecType()
  Dim ErnRec As ErnCodeRecType
  Dim ErnHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenErnCodeFile ErnHandle
  NumOfRecs = LOF(ErnHandle) / Len(ErnRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\EarnCodes.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get ErnHandle, x, ErnRec
    Print #RptHandle, QPTrim$(ErnRec.ERNCODE1);
    Print #RptHandle, B & QPTrim$(ErnRec.ERNFWT1);
    Print #RptHandle, B & QPTrim$(ErnRec.ERNSWT1);
    Print #RptHandle, B & QPTrim$(ErnRec.ERNSOC1);
    Print #RptHandle, B & QPTrim$(ErnRec.ERNMED1);
    Print #RptHandle, B & QPTrim$(ErnRec.ERNRET1) & B
       
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close ErnHandle
  Unload FrmShowPctComp
 End Sub
 Private Sub ProcessDedCodeRecType()
  Dim DedRec As DedCodeRecType
  Dim DedHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenDedCodeFile DedHandle
  NumOfRecs = LOF(DedHandle) / Len(DedRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\DedCodes.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get DedHandle, x, DedRec
    Print #RptHandle, QPTrim$(DedRec.DCDESC1);
    Print #RptHandle, B & QPTrim$(DedRec.DCFWT1);
    Print #RptHandle, B & QPTrim$(DedRec.DCSWT1);
    Print #RptHandle, B & QPTrim$(DedRec.DCSOC1);
    Print #RptHandle, B & QPTrim$(DedRec.DCMED1) & B
    
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close DedHandle
  Unload FrmShowPctComp
 End Sub
  Private Sub ProcessAccrualDates()
  Dim AccRec As AccrualDates
  Dim AccHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenAccrualDatesFile AccHandle
  NumOfRecs = LOF(AccHandle) / Len(AccRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\AccrualDates.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get AccHandle, x, AccRec
    Print #RptHandle, MakeRegDate(AccRec.PreviousDate);
    Print #RptHandle, B & MakeRegDate(AccRec.CurrentDate) & B

       
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close AccHandle
  Unload FrmShowPctComp
 End Sub
Private Sub ProcessPayRateType()
  Dim PayRateRec As PayRateType
  Dim PayRateHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenPayRateFile PayRateHandle
  NumOfRecs = LOF(PayRateHandle) / Len(PayRateRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\PayRate.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get PayRateHandle, x, PayRateRec
      Print #RptHandle, Using$("########", PayRateRec.EmpRecNum);
      Print #RptHandle, B & QPTrim$(PayRateRec.EmpNo);
      Print #RptHandle, B & QPTrim$(PayRateRec.EmpLName);
      Print #RptHandle, B & QPTrim$(PayRateRec.EmpFName);
      Print #RptHandle, B & MakeRegDate(PayRateRec.EMPHDATE);
      Print #RptHandle, B & MakeRegDate(PayRateRec.EMPTDATE);
      
    For Y = 1 To 30
      If Y < 30 Then
        If Val(PayRateRec.RegPayRate(Y)) > 0 Then
          Print #RptHandle, B & Using$("##,####.##", PayRateRec.RegPayRate(Y));
        Else
          Print #RptHandle, B & 0;
        End If
        If Val(PayRateRec.OTPayRate(Y)) > 0 Then
          Print #RptHandle, B & Using$("###,###.##", PayRateRec.OTPayRate(Y));
        Else
          Print #RptHandle, B & 0;
        End If
        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPFREQ(Y));
        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPTYPE(Y));
        If Len(PayRateRec.PayChngDate(Y)) > 0 Then
          Print #RptHandle, B & MakeRegDate(PayRateRec.PayChngDate(Y));
        Else
          Print #RptHandle, B & 0;
        End If
        Print #RptHandle, B & QPTrim$(PayRateRec.EMPJOB(Y));
      Else
        If Val(PayRateRec.RegPayRate(Y)) > 0 Then
          Print #RptHandle, B & Using$("##,####.##", PayRateRec.RegPayRate(Y));
        Else
          Print #RptHandle, B & 0;
        End If
        If Val(PayRateRec.OTPayRate(Y)) > 0 Then
          Print #RptHandle, B & Using$("###,###.##", PayRateRec.OTPayRate(Y));
        Else
          Print #RptHandle, B & 0;
        End If
        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPFREQ(Y));
        Print #RptHandle, B & QPTrim$(PayRateRec.EMPPTYPE(Y));
        If Len(PayRateRec.PayChngDate(Y)) > 0 Then
          Print #RptHandle, B & MakeRegDate(PayRateRec.PayChngDate(Y));
        Else
          Print #RptHandle, B & 0;
        End If
        Print #RptHandle, B & QPTrim$(PayRateRec.EMPJOB(Y)) & B
      End If
    Next Y
    
         
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close PayRateHandle
  Unload FrmShowPctComp
 End Sub
 Private Sub ProcessPRMessRecType()
  Dim MessRec As PRMessRecType
  Dim MessRecHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenEmpMessage MessRecHandle
  NumOfRecs = LOF(MessRecHandle) / Len(MessRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\MessRec.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get MessRecHandle, x, MessRec
    Print #RptHandle, Using$("########", MessRec.EmpRec);
    For Y = 1 To 15
      Print #RptHandle, B & QPTrim$(MessRec.MessLine(Y).Msg) & B
    Next Y
           
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close MessRecHandle
  Unload FrmShowPctComp
 End Sub
 
 Private Sub ProcessOrbitEmpData()
  Dim OrbitRec As OrbitEmpData
  Dim OrbitHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenOrbEmpDataFile OrbitHandle
  NumOfRecs = LOF(OrbitHandle) / Len(OrbitRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\OrbitEmpData.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get OrbitHandle, x, OrbitRec
    Print #RptHandle, Using$("####", OrbitRec.EmpRecNum);
    Print #RptHandle, B & QPTrim$(OrbitRec.EmpNum);
    Print #RptHandle, B & QPTrim$(OrbitRec.RecType);
    Print #RptHandle, B & QPTrim$(OrbitRec.AgencyNum);
    Print #RptHandle, B & QPTrim$(OrbitRec.PlanCode);
    Print #RptHandle, B & QPTrim$(OrbitRec.SSN);
    Print #RptHandle, B & QPTrim$(OrbitRec.MemberID);
    Print #RptHandle, B & QPTrim$(OrbitRec.FirstName);
    Print #RptHandle, B & QPTrim$(OrbitRec.MiddleName);
    Print #RptHandle, B & QPTrim$(OrbitRec.LastName);
    Print #RptHandle, B & QPTrim$(OrbitRec.Suffix);
    Print #RptHandle, B & QPTrim$(OrbitRec.Gender);
    If Val(QPTrim$((OrbitRec.DateOfBirth))) > 0 Then
      Print #RptHandle, B & MakeMDYDate(OrbitRec.DateOfBirth);
    Else
      Print #RptHandle, B & "";
    End If
    Print #RptHandle, B & QPTrim$(OrbitRec.AddLine1);
    Print #RptHandle, B & QPTrim$(OrbitRec.AddLine2);
    Print #RptHandle, B & QPTrim$(OrbitRec.City);
    Print #RptHandle, B & QPTrim$(OrbitRec.State);
    Print #RptHandle, B & QPTrim$(OrbitRec.Zip);
    Print #RptHandle, B & QPTrim$(OrbitRec.OutOfCntryAdd);
    Print #RptHandle, B & QPTrim$(OrbitRec.JobClass);
    Print #RptHandle, B & QPTrim$(OrbitRec.DeptNum);
    If Val(QPTrim$((OrbitRec.EmployDate))) > 0 Then
      Print #RptHandle, B & MakeMDYDate(OrbitRec.EmployDate);
    Else
      Print #RptHandle, B & "";
    End If
    If Val(QPTrim$(OrbitRec.EligibleDate)) > 0 Then
      Print #RptHandle, B & MakeMDYDate(OrbitRec.EligibleDate);
    Else
      Print #RptHandle, B & "";
    End If
    Print #RptHandle, B & QPTrim$(OrbitRec.PayType);
    Print #RptHandle, B & QPTrim$(OrbitRec.VacHours);
    Print #RptHandle, B & QPTrim$(OrbitRec.ContrPdEmpPrd);
    If Val(OrbitRec.ContrPdEmpBegDate) > 0 Then
      Print #RptHandle, B & MakeMDYDate(OrbitRec.ContrPdEmpBegDate);
    Else
      Print #RptHandle, B & "";
    End If
    If Val(QPTrim(OrbitRec.ContrPdEmpEndDate)) > 0 Then
      Print #RptHandle, B & MakeMDYDate(OrbitRec.ContrPdEmpEndDate);
    Else
      Print #RptHandle, B & "";
    End If
    If Val(QPTrim(OrbitRec.TerminationDate)) > 0 Then
      Print #RptHandle, B & MakeMDYDate(OrbitRec.TerminationDate);
    Else
      Print #RptHandle, B & "";
    End If
    Print #RptHandle, B & QPTrim$(OrbitRec.TermType);
    Print #RptHandle, B & QPTrim$(OrbitRec.SharedPosition);
    If OrbitRec.Deleted = True Then
      Print #RptHandle, B & "Y";
    Else
      Print #RptHandle, B & "N";
    End If
    Print #RptHandle, B & QPTrim$(OrbitRec.Adjustment) & B

           
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close OrbitHandle
  Unload FrmShowPctComp
 End Sub
  Private Sub ProcessVoidedCheckType()
  Dim VoidRec As VoidCheckType
  Dim VoidHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenTempVoidFile VoidHandle
  NumOfRecs = LOF(VoidHandle) / Len(VoidRec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\VoidedCheck.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get VoidHandle, x, VoidRec
    Print #RptHandle, Using$("##########", VoidRec.CheckNum);
    Print #RptHandle, B & QPTrim$(VoidRec.EmpNum);
    Print #RptHandle, B & MakeRegDate(VoidRec.CheckDate);
    Print #RptHandle, B & Using$("###,###.##", VoidRec.CheckAmt) & B

           
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close VoidHandle
  Unload FrmShowPctComp
 End Sub
   Private Sub ProcessW2ElectronicsSubRa()
  Dim W2Rec As W2ElectronicSubRA
  Dim W2Handle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
    
  OpenW2ESubRA W2Handle
  NumOfRecs = LOF(W2Handle) / Len(W2Rec)
    
  StartPath = App.Path
  
  B = "|"
    
  ThisFile = "\W2ESubRA.txt"
  If Exist(StartPath + "\PayrollData\" + ThisFile) Then
    KillFile (StartPath + "\PayrollData\" + ThisFile)
  End If
    
  RptName$ = StartPath + "\PayrollData\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  FrmShowPctComp.Label1 = "Payroll Export"
  FrmShowPctComp.Show , Me
  DoEvents
  
  For x = 1 To NumOfRecs
    Get W2Handle, x, W2Rec
    Print #RptHandle, QPTrim$(W2Rec.EINNum);
    Print #RptHandle, B & QPTrim$(W2Rec.PersIDNum);
    Print #RptHandle, B & QPTrim$(W2Rec.ResubID);
    Print #RptHandle, B & QPTrim$(W2Rec.ReSubWFID);
    Print #RptHandle, B & QPTrim$(W2Rec.SftwrCode);
    Print #RptHandle, B & QPTrim$(W2Rec.CmpnyName);
    Print #RptHandle, B & QPTrim$(W2Rec.LocAddr);
    Print #RptHandle, B & QPTrim$(W2Rec.DelAddr);
    Print #RptHandle, B & QPTrim$(W2Rec.City);
    Print #RptHandle, B & QPTrim$(W2Rec.State);
    Print #RptHandle, B & QPTrim$(W2Rec.Zip);
    Print #RptHandle, B & QPTrim$(W2Rec.ZipExt);
    Print #RptHandle, B & QPTrim$(W2Rec.SubmttrName);
    Print #RptHandle, B & QPTrim$(W2Rec.SubLocAddr);
    Print #RptHandle, B & QPTrim$(W2Rec.SubDelAddr);
    Print #RptHandle, B & QPTrim$(W2Rec.SubCity);
    Print #RptHandle, B & QPTrim$(W2Rec.SubState);
    Print #RptHandle, B & QPTrim$(W2Rec.SubZip);
    Print #RptHandle, B & QPTrim$(W2Rec.SubZipExt);
    Print #RptHandle, B & QPTrim$(W2Rec.ContactName);
    Print #RptHandle, B & QPTrim$(W2Rec.CntctPhone);
    Print #RptHandle, B & QPTrim$(W2Rec.CntPhnExt);
    Print #RptHandle, B & QPTrim$(W2Rec.CntEMail);
    Print #RptHandle, B & QPTrim$(W2Rec.CntFAX);
    Print #RptHandle, B & QPTrim$(W2Rec.CntMethod);
    Print #RptHandle, B & QPTrim$(W2Rec.PrepCode) & B

           
    FrmShowPctComp.ShowPctComp x, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close W2Handle
  Unload FrmShowPctComp
 End Sub
 
Private Sub validateDedCodes()
  Dim DedRec As DedCodeRecType
  Dim DedHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim Y As Long
  Dim wrkCnt As Integer
  Dim wrkDesc As String
        
  OpenDedCodeFile DedHandle
  NumOfRecs = LOF(DedHandle) / Len(DedRec)
  
  Dim dedArray(1 To 50) As String

  For x = 1 To 50
    dedArray(x) = ""
  Next x
     
  For x = 1 To NumOfRecs
    Get DedHandle, x, DedRec
    dedArray(x) = DedRec.DCDESC1
  Next x
  
    
  wrkCnt = 0
  For x = 1 To 50
    If wrkCnt <= 1 Then
      If dedArray(x) <> "" Then
          wrkCnt = 0
          wrkDesc = dedArray(x)
        
          For Y = 1 To 50
            If dedArray(Y) = wrkDesc Then
              wrkCnt = wrkCnt + 1
            End If
          Next
      Else
        wrkCnt = 0
      End If
    Else
      ErrorCode = 1
      MsgBox ("Duplicate deductions codes were found. This must be changed before the conversion can be run")
      Exit Sub
    End If
  Next x
  Close DedHandle
End Sub
Private Sub ValidateErnCodes()
  Dim ErnRec As ErnCodeRecType
  Dim ErnHandle As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim Y As Integer
  Dim B As String
  Dim wrkCnt As Integer
  Dim wrkDesc As String
    
  OpenErnCodeFile ErnHandle
  NumOfRecs = LOF(ErnHandle) / Len(ErnRec)
 
  Dim ErnArray(1 To 20) As String
  
  For x = 1 To 20
    ErnArray(x) = ""
  Next x
         
  For x = 1 To NumOfRecs
    Get ErnHandle, x, ErnRec
    ErnArray(x) = ErnRec.ERNCODE1
  Next x
    
  wrkCnt = 0
  For x = 1 To 20
    If wrkCnt <= 1 Then
      If ErnArray(x) <> "" Then
          wrkCnt = 0
          wrkDesc = ErnArray(x)
        
          For Y = 1 To 20
            If ErnArray(Y) = wrkDesc Then
              wrkCnt = wrkCnt + 1
            End If
          Next
      Else
        wrkCnt = 0
      End If
    Else
      ErrorCode = 1
      MsgBox ("Duplicate earn codes were found. This must be changed before the conversion can be run")
      Exit Sub
    End If
  Next x

  Close ErnHandle
  
 End Sub


