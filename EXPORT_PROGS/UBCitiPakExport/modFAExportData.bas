Attribute VB_Name = "modFAExportData"
Option Explicit

Public Sub ProcessFAData()
  Dim FATrans As DprHistType
  Dim FAItem As FAItemRecType
  Dim FASysRec As FASetupRecType
  Dim FAItemOutFile As String
  Dim FACodeOutFile As String
  Dim FADeptOutFile As String
  Dim FASysOutFile As String
  Dim FAFundOutFile As String
  Dim FATranOutFile As String
  
  Dim FACode As FAAssetCodeRecType
  Dim FADept As FADeptCodeType
  Dim FAFund As FAFundCodeType
  Dim NumOfRecs As Long
  Dim ftNumOfRecs As Long
  Dim RecLen As Integer
  Dim ftRecLen As Integer
  Dim TRRec As Long
  Dim Cnt As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim ftRptHandle As Integer
  Dim ftRptName As String
  Dim FileHandle As Integer
  Dim ftFileHandle As Integer
  Dim ThisFile As String
  Dim Y As Integer
  Dim B As String
  Dim FAOutPath As String
  Dim tStr As String
  Dim Fmt1 As String, Fmt2 As String
  Dim ITCnt As Integer
  Fmt1 = "############.##"
  Fmt2 = "#######"
  FAOutPath = "\FixedAssetData\"
  FAItemOutFile = "FAITEMS.TXT"
  FACodeOutFile = "FAGCODES.TXT"
  FADeptOutFile = "FADEPTS.TXT"
  FASysOutFile = "FASETUP.TXT"
  FAFundOutFile = "FAFUNDS.TXT"
  FATranOutFile = "FATRANS.TXT"
  StartPath = App.Path
  B = "|"
'----------------------------------------------------
  RecLen = Len(FASysRec)
  RptName$ = StartPath + FAOutPath + FASysOutFile
  If Exist(RptName$) Then
    KillFile (RptName$)
  End If
  
  FileHandle = FreeFile
  Open FAData + FASetUpFileName For Random Shared As FileHandle Len = RecLen
    
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  Get FileHandle, 1, FASysRec
  Print #RptHandle, QPTrim$(FASysRec.TOWNNAME); B; '1
  Print #RptHandle, QPTrim(Using$(Fmt1, FASysRec.Pct1St)); B;
  Print #RptHandle, QPTrim(FASysRec.PRate1St); B;
  Print #RptHandle, QPTrim$(FASysRec.DeprType); B
  Close
'----------------------------------------------------
  RecLen = Len(FADept)
  RptName$ = StartPath + FAOutPath + FADeptOutFile
  If Exist(RptName$) Then
    KillFile (RptName$)
  End If
    
  FileHandle = FreeFile
  Open FAData + FADeptCodeName For Random Shared As FileHandle Len = RecLen
  NumOfRecs = LOF(FileHandle) / RecLen
  
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  For Cnt = 1 To NumOfRecs
    Get FileHandle, Cnt, FADept
    Print #RptHandle, QPTrim$(FADept.DeptDesc); B; '1
    Print #RptHandle, QPTrim(Using$(Fmt2, FADept.DeptNum)); B
  Next
  Close


'----------------------------------------------------
  RecLen = Len(FAFund)
  RptName$ = StartPath + FAOutPath + FAFundOutFile
  If Exist(RptName$) Then
    KillFile (RptName$)
  End If
    
  FileHandle = FreeFile
  Open FAData + FAFundCodeName For Random Shared As FileHandle Len = RecLen
  NumOfRecs = LOF(FileHandle) / RecLen
  
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  For Cnt = 1 To NumOfRecs
    Get FileHandle, Cnt, FAFund
    Print #RptHandle, QPTrim$(FAFund.FundDesc); B; '1
    Print #RptHandle, QPTrim(Using$(Fmt2, FAFund.FundNum)); B
  Next
  Close

'  Public Const FAYearEndName = "FAYEAR.DAT"
'  Public Const FADeprEditName = "FADPREDT.DAT"
'  Public Const TempDprFileName = "FATEMPDPR.DAT"
'  Public Const TempDispDateName = "FATEMPDISPDATE.DAT"
'------------------------------------------------------
'----------------------------------------------------
  ITCnt = 0
  ftRecLen = Len(FATrans)
  ftRptName$ = StartPath + FAOutPath + FATranOutFile
  If Exist(ftRptName$) Then
    KillFile (ftRptName$)
  End If
  
  ftFileHandle = FreeFile
  Open FAData + FADprHistFileName For Random Shared As ftFileHandle Len = ftRecLen
  ftNumOfRecs = LOF(FileHandle) / ftRecLen
  
  FrmShowPctComp.Label1 = "FA Asset History"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  
  ftRptHandle = FreeFile
  Open ftRptName$ For Output As #ftRptHandle
  
  RecLen = Len(FAItem)
  RptName$ = StartPath + FAOutPath + FAItemOutFile
  If Exist(RptName$) Then
    KillFile (RptName$)
  End If
    
  FileHandle = FreeFile
  Open FAData + FAItemFile For Random Shared As FileHandle Len = RecLen
  NumOfRecs = LOF(FileHandle) / RecLen
  
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  FrmShowPctComp.Label1 = "FA Asset Items"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For Cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp Cnt, NumOfRecs
    DoEvents
    Get FileHandle, Cnt, FAItem
    ITCnt = ITCnt + 1
0    Print #RptHandle, QPTrim(Using$(Fmt2, ITCnt)); B;
1    Print #RptHandle, QPTrim$(FAItem.ItemTag); B; '1
2    Print #RptHandle, QPTrim$(FAItem.ISTATUS); B; '1
3    Print #RptHandle, QPTrim$(FAItem.DEPYN); B; '1  'depreciate Y/N
4    Print #RptHandle, QPTrim$(MakeRegDate(FAItem.AQURDATE)); B;
5    Print #RptHandle, QPTrim$(FAItem.IDESC1); B;
6    Print #RptHandle, QPTrim$(FAItem.IDESC2); B;
7    Print #RptHandle, QPTrim$(FAItem.GLAcct); B;
8    Print #RptHandle, QPTrim(Using$(Fmt2, FAItem.IDEPT)); B;
9    Print #RptHandle, QPTrim$(FAItem.ASSETCODE); B;
10   Print #RptHandle, QPTrim(Using$(Fmt2, FAItem.ILIFE)); B;
11   Print #RptHandle, QPTrim(Using$(Fmt1, FAItem.ORGCOST)); B;
12   Print #RptHandle, QPTrim(Using$(Fmt1, FAItem.DEP2DATE)); B;
13   Print #RptHandle, QPTrim(Using$(Fmt1, FAItem.CURRVAL)); B;
14   Print #RptHandle, QPTrim$(MakeRegDate(FAItem.CDEPDATE)); B;
15   Print #RptHandle, QPTrim$(MakeRegDate(FAItem.DispDate)); B;
16   Print #RptHandle, QPTrim$(FAItem.VENDOR); B;
17   Print #RptHandle, QPTrim$(FAItem.SERIALNO); B;
18   Print #RptHandle, QPTrim$(FAItem.ITEMMFG); B;
19   Print #RptHandle, QPTrim$(FAItem.CONTACT); B;
20   Print #RptHandle, QPTrim$(FAItem.ITEMLOC); B;
21   Print #RptHandle, QPTrim$(MakeRegDate(FAItem.EOLDATE)); B;
22   Print #RptHandle, QPTrim$(FAItem.VHCLMAKE); B;
23   Print #RptHandle, QPTrim$(FAItem.VHCLMODL); B;
24   Print #RptHandle, QPTrim$(FAItem.VHCLVIN); B;
25   Print #RptHandle, QPTrim$(FAItem.VHCLTAG); B;
26   Print #RptHandle, QPTrim$(FAItem.VHCLCOLR); B;
27   Print #RptHandle, QPTrim$(MakeRegDate(FAItem.WARRXDAT)); B;
28   Print #RptHandle, QPTrim$(FAItem.PHONE); B;
29   Print #RptHandle, QPTrim(Using$(Fmt2, FAItem.FundNum)); B;
30   Print #RptHandle, QPTrim(Using$(Fmt1, FAItem.DisposAmt)); B;
31   Print #RptHandle, QPTrim(Using$(Fmt2, FAItem.LastDprRec)); B;
32   Print #RptHandle, QPTrim(Using$(Fmt2, FAItem.LifeLeft)); B;
33   Print #RptHandle, QPTrim$(FAItem.PONum); B;
34   Print #RptHandle, QPTrim$(FAItem.CheckNum); B;
35   Print #RptHandle, QPTrim(Using$(Fmt2, FAItem.DsplFlag)); B;
36   Print #RptHandle, QPTrim$(FAItem.DsplMethod); B
      If FAItem.LastDprRec > 0 Then
        TRRec = FAItem.LastDprRec
        Do While TRRec > 0
          Get ftFileHandle, TRRec, FATrans
          Print #ftRptHandle, QPTrim(Using$(Fmt2, ITCnt)); B;
          If FATrans.PrevDprRec <> 0 Then
            Print #ftRptHandle, QPTrim(Using$(Fmt2, FATrans.PrevDprRec)); B;
          Else
            Print #ftRptHandle, QPTrim("0"); B;
          End If
          Print #ftRptHandle, QPTrim(Using$(Fmt2, FATrans.ThisDept)); B;
          Print #ftRptHandle, QPTrim(Using$(Fmt1, FATrans.DprAmt)); B;
          Print #ftRptHandle, QPTrim(FATrans.DprYear); B;
          Print #ftRptHandle, QPTrim(FATrans.ItemTag); B;
          Print #ftRptHandle, QPTrim(Using$(Fmt1, FATrans.DprToDate)); B;
          Print #ftRptHandle, QPTrim(FATrans.ThisDesc1); B;
          Print #ftRptHandle, QPTrim(Using$(Fmt1, FATrans.BookTotal)); B;
          Print #ftRptHandle, QPTrim(Using$(Fmt1, FATrans.OrigCost)); B;
          Print #ftRptHandle, QPTrim(Using$(Fmt1, FATrans.Life)); B;
          Print #ftRptHandle, QPTrim(FATrans.PurchYear); B;
          Print #ftRptHandle, QPTrim(Using$(Fmt2, FATrans.LifeLeft)); B;
          Print #ftRptHandle, FATrans.SoSoftFlag; B
          If FATrans.PrevDprRec = TRRec Then
            Exit Do
          End If
          TRRec = FATrans.PrevDprRec
        Loop
      End If
  Next
  
  Close
  
  FrmShowPctComp.Label1 = "FA Group Codes"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
    
  RptName$ = StartPath + FAOutPath + FACodeOutFile
  If Exist(RptName$) Then
    KillFile (RptName$)
  End If
  RecLen = Len(FACode)
  FileHandle = FreeFile
  Open FAData + FACodeFile For Random Shared As FileHandle Len = RecLen
  
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  NumOfRecs = LOF(FileHandle) / RecLen
  For Cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp Cnt, NumOfRecs
    DoEvents
    Get FileHandle, Cnt, FACode
    Print #RptHandle, QPTrim$(FACode.ASSETCODE); B; '1
    Print #RptHandle, QPTrim$(FACode.AssetStatus); B; '1
    Print #RptHandle, QPTrim$(FACode.AssetDesc); B
  Next
  Close
  
End Sub
