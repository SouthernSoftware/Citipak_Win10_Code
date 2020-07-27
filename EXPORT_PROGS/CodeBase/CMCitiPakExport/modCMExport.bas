Attribute VB_Name = "modCMExport"
Option Explicit

Public Sub ProcessCMData()
  
  Dim CMTransRec As CMTransRecTypeII
  Dim CMMiscCode As MiscCodeRecType
  Dim CMUBSysRec As UBSetupRecType
  
  Dim CMTransRecOutFile As String
  Dim CMMiscCodeOutFile As String
  
  Dim UNumOfRecs As Long
  Dim TNumOfRecs As Long
  Dim CNumOfRecs As Long
  
  Dim URecLen As Integer
  Dim TRecLen As Integer
  Dim CRecLen As Integer
  
  Dim CMOutPath As String
  
  Dim UBServ As String
  
  Dim FmtC As String
  Dim FmtI As String
  Dim FmtL As String
  
  Dim CMUBSysHandle As Integer
  Dim CMTransHandle As Integer
  Dim CMCodesHandle As Integer
  
  Dim OutFileHandle As Integer
  
  Dim TCnt As Long
  Dim tLop As Integer
  Dim tUBCnt As Integer
  Dim NumCMTransRecs As Long
  Dim NumCMCodesRecs As Long
  Dim B As String
  
  B = "|"
  FmtC = "############.##"
  FmtI = "#######"
  FmtL = "###########"
  CMTransRecOutFile = "CMTRANS.TXT"
  CMMiscCodeOutFile = "CMCODES.TXT"
  
  URecLen = Len(CMUBSysRec)
  TRecLen = Len(CMTransRec)
  CRecLen = Len(CMMiscCode)
  
  CMUBSysHandle = FreeFile         'load ub system data
  Open CMUBSysFile For Random As CMUBSysHandle Len = URecLen
  Get CMUBSysHandle, 1, CMUBSysRec
  Close CMUBSysHandle
  
  CMCodesHandle = FreeFile
  Open CMCodeFile For Random As CMCodesHandle Len = CRecLen
  NumCMCodesRecs = LOF(CMCodesHandle) / CRecLen
    
  OutFileHandle = FreeFile
  Open CMMiscCodeOutFile For Output As OutFileHandle
  
  For TCnt = 1 To NumCMCodesRecs
    Get CMCodesHandle, TCnt, CMMiscCode
    Print #OutFileHandle, QPTrim(Str$(TCnt)); B;
    Print #OutFileHandle, QPTrim(CMMiscCode.MiscCode); B;
'    MiscCode As String * 7
    Print #OutFileHandle, QPTrim(CMMiscCode.Description); B;
'    Description As String * 25
    Print #OutFileHandle, QPTrim(CMMiscCode.GlAcctNumb); B;
'    GlAcctNumb As String * 14
    Print #OutFileHandle, QPTrim(CMMiscCode.InactiveFlag); B
'    InactiveFlag As String * 1
'    NotUsed As String * 17
  Next
  tUBCnt = NumCMCodesRecs
  
  For TCnt = 1 To 15
    UBServ = QPTrim(CMUBSysRec.Revenues(TCnt).RevName)
    If Len(UBServ) > 0 Then
      tUBCnt = tUBCnt + 1
      Print #OutFileHandle, QPTrim(Str$(tUBCnt)); B;
      Print #OutFileHandle, QPTrim(Str$(TCnt)); B;
      Print #OutFileHandle, QPTrim(UBServ); B;
      Print #OutFileHandle, QPTrim("UB"); B;
      Print #OutFileHandle, QPTrim("N"); B
    End If
  Next
'  Call CMAddUBCodes2Data
  
  Close
  FrmShowPctComp.Label1 = "Cash Management History"
  FrmShowPctComp.Show , frmCitiPakExportData
  
  CMTransHandle = FreeFile
  Open CMTranFile For Random As CMTransHandle Len = TRecLen
  NumCMTransRecs = LOF(CMTransHandle) / TRecLen
  
  OutFileHandle = FreeFile
  Open CMTransRecOutFile For Output As OutFileHandle
  
  For TCnt = 1 To NumCMTransRecs
    If (TCnt Mod 20) = 0 Then
      FrmShowPctComp.ShowPctComp TCnt, NumCMTransRecs
      DoEvents
    End If
    Get CMTransHandle, TCnt, CMTransRec
    'If CMTransRec.TransSource = 1 Or CMTransRec.TransSource = 201 Then
    
    Print #OutFileHandle, QPTrim(Str$(TCnt)); B;
    Print #OutFileHandle, QPTrim$(MakeRegDate(CMTransRec.TransDate)); B;
    Print #OutFileHandle, QPTrim(Using$(FmtC, CMTransRec.TransAmount)); B;
    Print #OutFileHandle, QPTrim(Using$(FmtC, CMTransRec.TransCash)); B;
    Print #OutFileHandle, QPTrim(Using$(FmtC, CMTransRec.TransCheck)); B;
    Print #OutFileHandle, QPTrim(Using$(FmtC, CMTransRec.TransAmtOwed)); B;
    Print #OutFileHandle, QPTrim(CMTransRec.TransDesc); B;
    Print #OutFileHandle, QPTrim(Using$(FmtI, CMTransRec.TransSource)); B;
                          '1-Misc 24-Util 27-UtilDep 31-Tax 131-Newtax 41-License 141-NewBL 51-decal
                          '201-void Misc 224-void util 227-void dep 241-void lic 231-void tax 251-void Decal
    Print #OutFileHandle, QPTrim(CMTransRec.TransName); B;
    Print #OutFileHandle, QPTrim(Using$(FmtL, CMTransRec.TransAcctNum)); B;
    Print #OutFileHandle, QPTrim(Using$(FmtL, CMTransRec.TransDetNum)); B;
    
    For tLop = 1 To 15
      Print #OutFileHandle, QPTrim(Using$(FmtC, CMTransRec.TransRevAmt(tLop))); B;
    Next
    Print #OutFileHandle, QPTrim(Using$(FmtL, CMTransRec.TransOperNum)); B;
    Print #OutFileHandle, QPTrim(CMTransRec.Trans2GL); B;
    Print #OutFileHandle, QPTrim(Using$(FmtI, CMTransRec.TransTender)); B;
'    If CMTransRec.TransTender = 4 Then
'
'    Stop
'    End If
                         'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
    Print #OutFileHandle, QPTrim(Using$(FmtL, CMTransRec.TransVoidNum)); B
                           'Voided trans link to record voided or void trans
   'End If
  Next
  Close
  FrmShowPctComp.ShowPctComp 1, 1

'    TransDate    As Integer
'    TransAmount  As Double
'    TransCash    As Double
'    TransCheck   As Double
'    TransAmtOwed As Double
'    TransDesc    As String * 25
'    TransName    As String * 25
'    TransAcctNum As Long               'Holds Master Acct Record Number in Mod
'    TransDetNum  As Long               'Holds Record Number of Transaction Det
'    TransRevAmt(1 To 15) As Double
'    TransOperNum As Long
'    Trans2GL     As String * 1
'    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
'    TransVoidNum As Long        'Voided trans link to record voided or void trans
'    ChkByte      As String * 1
'    TransPad     As String * 18

'    ChkByte      As String * 1
'    TransPad     As String * 18

End Sub

