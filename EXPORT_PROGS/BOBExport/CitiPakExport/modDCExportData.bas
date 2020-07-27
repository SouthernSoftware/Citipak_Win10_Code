Attribute VB_Name = "modDCExportData"
Option Explicit

Public Sub ProcessDCData()
  
  Dim DCTrans As DCTransRecType
  Dim DCCust As DCCustRecType
  Dim DCSysRec As DCSetupType
  Dim DCVCode As DCCatCodeRecType
  Dim DCVehic As DCVehType
  
  Dim DCCustOutFile As String
  Dim DCVCodeOutFile As String
  Dim DCTranOutFile As String
  Dim DCVehicOutFile  As String
  Dim DCSysOutFile As String
  Dim NumOfRecs As Long
  Dim CNumOfRecs As Long
  Dim TNumOfRecs As Long
  Dim VNumOfRecs As Long
  Dim RecLen As Integer
  Dim CRecLen As Integer
  Dim TRecLen As Integer
  Dim VRecLen As Integer
  Dim Cnt As Integer
  Dim CustCnt As Long
  'Dim x As Long
  Dim PrevTranRec As Long
  Dim PrevVehRec As Long
  Dim RptHandle As Integer
  Dim CRptHandle As Integer
  Dim TRptHandle As Integer
  Dim VRptHandle As Integer
  Dim RptName As String
  Dim CRptName As String
  Dim TRptName As String
  Dim VRptName As String
  
  Dim FileHandle As Integer
  Dim CFileHandle As Integer
  Dim TFileHandle As Integer
  Dim VFileHandle As Integer
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
  Dim DCOutPath As String
  Dim tStr As String
  Dim FmtC As String, FmtI As String
  Dim BadCnt As Integer
  Dim VehCnt As Integer
  
  CRecLen = Len(DCCust)
  TRecLen = Len(DCTrans)
  VRecLen = Len(DCVehic)
  FmtC = "############.##"
  FmtI = "#######"
  DCOutPath = "\VehicleDecalData\"
  DCCustOutFile = "DCCUST.TXT"
  DCVCodeOutFile = "DCVCODE.TXT"
  DCTranOutFile = "DCTRANS.TXT"
  DCVehicOutFile = "DCVEHIC.TXT"
  DCSysOutFile = "DCSETUP.TXT"
  StartPath = App.Path
  B = "|"
  
'----------------------------------------------------
'Public Const DCCustFile = "DCCust.dat"
'Public Const DCTranFile = "DCTrans.dat"
'Public Const DCSetupFile = "DCSetup.dat"
'Public Const DCVCodeFile = "DCCODE.dat"
'Public Const DCVehFile = "DCVEH.dat"

'DC System Setup*****************
  RecLen = Len(DCSysRec)
  RptName$ = StartPath + DCOutPath + DCSysOutFile
  If Exist(RptName$) Then
    KillFile (RptName$)
  End If
   
  FileHandle = FreeFile
  Open DCData + DCSetupFile For Random Shared As FileHandle Len = RecLen
  NumOfRecs = LOF(FileHandle) / RecLen
  
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  For Cnt = 1 To NumOfRecs    'will only be one record
    Get FileHandle, Cnt, DCSysRec
    Print #RptHandle, QPTrim$(DCSysRec.DCTNNAME); B; '1
    Print #RptHandle, QPTrim$(DCSysRec.GLInterface); B; '1
    Print #RptHandle, QPTrim(Using$(FmtI, DCSysRec.AppType)); B;
    Print #RptHandle, QPTrim$(DCSysRec.DCVers); B; '1
    Print #RptHandle, QPTrim$(DCSysRec.Taxbalchk); B
'    DefLook      As String * 1
  Next
  Close
'SS-------------------------------

'DC Vehicle Codes ********************
'----------------------------------
  RecLen = Len(DCVCode)
  RptName$ = StartPath + DCOutPath + DCVCodeOutFile
  If Exist(RptName$) Then
    KillFile (RptName$)
  End If
    
  FrmShowPctComp.Label1 = "DC Vehicle Codes"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  
  FileHandle = FreeFile
  Open DCData + DCVCodeFile For Random Shared As FileHandle Len = RecLen
  NumOfRecs = LOF(FileHandle) / RecLen
  
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  For Cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp Cnt, NumOfRecs
    DoEvents
    Get FileHandle, Cnt, DCVCode
    Print #RptHandle, QPTrim$(DCVCode.CATCODE); B; '1
    Print #RptHandle, QPTrim$(DCVCode.CODEDESC); B; '1
    Print #RptHandle, QPTrim(Using$(FmtI, DCVCode.APPNUMB)); B;
    Print #RptHandle, QPTrim(Using$(FmtI, DCVCode.BILLCODE)); B;
    Print #RptHandle, QPTrim$(DCVCode.REVGLNUM); B; '1
    Print #RptHandle, QPTrim$(DCVCode.CASHACCT); B; '1
    Print #RptHandle, QPTrim(Using$(FmtC, DCVCode.Fee)); B
  Next
  Close

'----------------------------------
'VC  ********************

'DC Customers ----------------------------------
  
  CRptName$ = StartPath + DCOutPath + DCCustOutFile
  If Exist(CRptName$) Then
    KillFile (CRptName$)
  End If
  TRptName$ = StartPath + DCOutPath + DCTranOutFile
  If Exist(TRptName$) Then
    KillFile (TRptName$)
  End If
  VRptName$ = StartPath + DCOutPath + DCVehicOutFile
  If Exist(VRptName$) Then
    KillFile (VRptName$)
  End If
  
  'Dim TNumOfRecs As Long
    
  FrmShowPctComp.Label1 = "DC Cust, Trans & Vehs"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  
  CFileHandle = FreeFile
  Open DCData + DCCustFile For Random Shared As CFileHandle Len = CRecLen
  CNumOfRecs = LOF(CFileHandle) / CRecLen
  TFileHandle = FreeFile
  Open DCData + DCTranFile For Random Shared As TFileHandle Len = TRecLen
  TNumOfRecs = LOF(TFileHandle) / TRecLen
  VFileHandle = FreeFile
  Open DCData + DCVehFile For Random Shared As VFileHandle Len = VRecLen
  VNumOfRecs = LOF(VFileHandle) / VRecLen
  
  CRptHandle = FreeFile
  Open CRptName$ For Output As #CRptHandle
  TRptHandle = FreeFile
  Open TRptName$ For Output As #TRptHandle
  VRptHandle = FreeFile
  Open VRptName$ For Output As #VRptHandle
    
  For Cnt = 1 To CNumOfRecs
    FrmShowPctComp.ShowPctComp Cnt, CNumOfRecs
    DoEvents
    Get CFileHandle, Cnt, DCCust
    If DCCust.Deleted = "Y" Then
      GoTo NotThisDCCust
    End If
    CustCnt = CustCnt + 1
    Print #CRptHandle, QPTrim(Using$(FmtI, CustCnt)); B; '1
    Print #CRptHandle, QPTrim$(DCCust.CUSTNUMB); B; '2
    Print #CRptHandle, QPTrim$(DCCust.BILLNAME); B; '3
    Print #CRptHandle, QPTrim$(DCCust.ADDRESS1); B; '4
    Print #CRptHandle, QPTrim$(DCCust.ADDRESS2); B; '5
    Print #CRptHandle, QPTrim$(DCCust.City); B; '5
    Print #CRptHandle, QPTrim$(DCCust.State); B; '7
    Print #CRptHandle, QPTrim$(DCCust.ZIPCODE); B; '8
    Print #CRptHandle, QPTrim$(DCCust.SOSEC); B; '9
    Print #CRptHandle, QPTrim$(DCCust.DRVLIC); B; '10
    Print #CRptHandle, QPTrim$(MakeRegDate(DCCust.DATEOPED)); B; '11
    Print #CRptHandle, QPTrim$(DCCust.CASHONLY); B; '12
    Print #CRptHandle, QPTrim$(DCCust.resident); B; '13
    Print #CRptHandle, QPTrim$(DCCust.Owner); B; '14
    Print #CRptHandle, QPTrim$(DCCust.HPHONE); B; '15
    Print #CRptHandle, QPTrim$(DCCust.WPHONE); B; '16
    Print #CRptHandle, QPTrim$(DCCust.LICENSE); B; '17
    Print #CRptHandle, QPTrim(Using$(FmtI, DCCust.Valid)); B; '18
    Print #CRptHandle, QPTrim$(DCCust.Deleted); B; '19
    'Print #cRptHandle, QPTrim(Using$(FmtI, DCCust.FirstTrans)); B;
    'Print #cRptHandle, QPTrim(Using$(FmtI, DCCust.LastTrans)); B;
    'Print #CRptHandle, QPTrim(Using$(FmtI, DCCust.FirstCar)); B;
    'Print #CRptHandle, QPTrim(Using$(FmtI, DCCust.LastCar)); B;
    Print #CRptHandle, QPTrim$(DCCust.SocSec1); B; '20
    Print #CRptHandle, QPTrim$(DCCust.OtherName); B  '21
    If DCCust.FirstTrans > 0 Then
      PrevTranRec& = DCCust.FirstTrans
      Do While PrevTranRec& > 0
        Get TFileHandle, PrevTranRec&, DCTrans
'        If DCTrans.CustomerNumber = 35 Then Stop
        If DCTrans.TransType = 1 Then
          If DCTrans.VoidFlag <> "Y" Then
'            If DCTrans.VehRecord = 0 Then Stop
            Print #TRptHandle, QPTrim(Using$(FmtI, DCTrans.TransType)); B;
            Print #TRptHandle, QPTrim(Using$(FmtI, CustCnt)); B;
            Print #TRptHandle, QPTrim$(MakeRegDate(DCTrans.TransDate)); B;
            Print #TRptHandle, QPTrim(Using$(FmtC, DCTrans.TransAmount)); B;
            Print #TRptHandle, QPTrim$(DCTrans.TRVinDesc); B;   '1
            Print #TRptHandle, QPTrim$(DCTrans.ExtraDesc); B;   '1
            Print #TRptHandle, QPTrim(Using$(FmtC, DCTrans.CashAmount)); B;
            Print #TRptHandle, QPTrim(Using$(FmtC, DCTrans.ChkAmount)); B;
            Print #TRptHandle, QPTrim(Using$(FmtC, DCTrans.BalanceAfterTrans)); B;
            Print #TRptHandle, QPTrim$(DCTrans.makemodel); B;   '1
            Print #TRptHandle, QPTrim$(DCTrans.StateTag); B;
            Print #TRptHandle, QPTrim$(MakeRegDate(DCTrans.ExpireDate)); B;
            Print #TRptHandle, QPTrim$(DCTrans.Sticker); B;
            Print #TRptHandle, QPTrim(Using$(FmtI, DCTrans.OperNum)); B;
            Print #TRptHandle, QPTrim$(DCTrans.GLInterfaced); B;
            Print #TRptHandle, QPTrim$(DCTrans.DecalCat); B;
            Print #TRptHandle, QPTrim(Using$(FmtI, DCTrans.TransTender)); B;
            Print #TRptHandle, QPTrim$(DCTrans.VoidFlag); B;
            Print #TRptHandle, QPTrim(Using$(FmtI, DCTrans.CustomerNumber)); B;
            Print #TRptHandle, QPTrim(Using$(FmtI, DCTrans.VehRecord)); B
'            Else
'              BadCnt = BadCnt + 1
'            End If
        Else
'          Stop
        End If
          End If
        PrevTranRec& = DCTrans.NextTrans
        
      Loop
    End If
    
    If DCCust.FirstCar > 0 Then
      PrevVehRec = DCCust.FirstCar
      Do While PrevVehRec > 0
        Get VFileHandle, PrevVehRec, DCVehic
'        If DCVehic.MasterRecord = 35 Then Stop
        If DCVehic.Active <> "Y" Then
          GoTo CarSkip
        End If
        VehCnt = VehCnt + 1
        Print #VRptHandle, QPTrim(Using$(FmtI, CustCnt)); B;
        Print #VRptHandle, QPTrim$(DCVehic.DecalCat); B; '1
        Print #VRptHandle, QPTrim$(DCVehic.makemodel); B; '1
        Print #VRptHandle, QPTrim$(DCVehic.StateTag); B; '1
        Print #VRptHandle, QPTrim$(MakeRegDate(DCVehic.ExpireDate)); B;
        Print #VRptHandle, QPTrim$(DCVehic.Sticker); B; '1
        Print #VRptHandle, QPTrim$(DCVehic.Valid); B; '1
        Print #VRptHandle, QPTrim$(DCVehic.Active); B; '1
        Print #VRptHandle, QPTrim$(DCVehic.Notes); B; '1
        Print #VRptHandle, QPTrim$(DCVehic.Desc); B; '1
        Print #VRptHandle, QPTrim(Using$(FmtC, DCVehic.Fee)); B;
        Print #VRptHandle, QPTrim(Using$(FmtI, DCVehic.MasterRecord)); B; 'Actual Cust Record number
        Print #VRptHandle, QPTrim(Using$(FmtI, PrevVehRec)); B            'Actual Record Number
        'Print #VRptHandle, QPTrim(Using$(FmtI, VehCnt)); B
        'Print #VRptHandle, QPTrim(Using$(FmtI, DCVehic.NextRec)); B
CarSkip:
        PrevVehRec = DCVehic.NextRec
      Loop
    End If
NotThisDCCust:
  Next
  Close

'CU ********************
'----------------------------------
End Sub
