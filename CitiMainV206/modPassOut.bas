Attribute VB_Name = "modPassOut"
Option Explicit
Dim Citipass As CitiPassType
Global LevelPass As Integer
Global PWcnt As Integer
Global PWUser As String
Global CPAdminhand As Integer
Global UBPath As String
Global PWfromMdl As Integer
Global TypeSysOP As Integer
Global PassP As String
'did this for johnston co ? maybe all later on
'Public Const PassP$ = "C:\CPWork\PassTemp.dat"
'Public Const PassP$ = "C:\PassTemp.dat"

'LevelPass 1 is fullaccess, 2 is reports only
'Constants for Modules  By Menu Option Order
Public Const BL = 1 'Business License
Public Const AP = 2 'Accounts Payable
Public Const GL = 3 'General Ledger
Public Const PR = 4 'Payroll
Public Const FA = 5 'Fixed Assets
Public Const TX = 6 'Taxes
Public Const IC = 7 'Inventory Control
Public Const CM = 8 'Cash Management
Public Const UB = 9 'Utility Billing
Public Const DC = 10 'Vehicle Decals
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Boolean
Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
  szCSDVersion As String * 128
End Type

Type UserPriviType
  FullAccess    As Boolean
  ReportsOnly   As Boolean
  PaymentAccess As Boolean
  Adjustments   As Boolean
  ZDummy2       As Boolean
End Type

Type CitiPassType
  PassNum   As Integer      'Unique Numeric Value
  username  As String * 15
  PassWord  As String * 10
  Administ  As Boolean
  DelFlag   As Boolean
  SaveSpace As String * 19  'Save For Future Use
  Module(1 To 15) As UserPriviType
  InUseFlag As Boolean
  CompName  As String * 50
  FlagMod   As Integer       'Set to Number of Module Signed on to
  Flag2     As Integer       'Use for payroll or other modules limiting full access users
  Pad       As String * 46
End Type
Type ReceiptPRNType
  RcpPort   As String * 40
  PrnDefYN  As Integer
  CtlDefYN  As Integer
  PaymDate  As Integer
  RValidate As Integer
  ZExtra    As String * 16
End Type
'this type is temporary file kept on local drive for passing values between exe's
Type CitiPassTempType
  usernum   As Integer
  username  As String * 15
  frommdl   As Integer   'this is to indicate to citipak ok to have file
End Type

Public Sub OpenCitiPassFile(CitiPassFile, NumPassRecs)
  Dim PassRecLen As Integer
  On Error GoTo PassError
  PassRecLen = Len(Citipass)
  CitiPassFile = FreeFile
  Open "CitiPass.dat" For Random Shared As CitiPassFile Len = PassRecLen
  NumPassRecs = LOF(CitiPassFile) \ PassRecLen

'  Lock #BgtEditFileNum
  Exit Sub

PassError:
  CitiPassFile = -1
  
  MsgBox "Password Maintenance Is In Process, See Password Administrator.", vbOKOnly, "Access Denied"
End Sub
''' Dim pz As String, z As String, cnt4 As Integer
''' pz$ = "NONORGANIC"
''' z$ = ""
''' For cnt4 = 1 To Len(pz$)
'''  z$ = z$ + Chr$(Asc(Mid$(pz$, cnt4, 1)) Xor 127)
'''Next
'''
'''Stop


Public Sub ClearInUse(PWcnt)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    If PWcnt > 0 Then
    SetAttr ("Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, Citipass
    Citipass.InUseFlag = False
    Citipass.FlagMod = 0
    Citipass.Flag2 = 0
    Citipass.CompName = ""
    Put CitiPassFile, PWcnt, Citipass
    Close CitiPassFile
    End If
    KillFile (PassP$)
  End If
End Sub

Public Sub ResetInUse()
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    SetAttr ("CitiPass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    For cnt = 1 To NumPassRecs
      Get CitiPassFile, cnt, Citipass
      Citipass.InUseFlag = False
      Citipass.Flag2 = 0
      Citipass.FlagMod = 0
      Citipass.CompName = ""
      Put CitiPassFile, cnt, Citipass
    Next
   Close CitiPassFile
  End If
End Sub
Public Sub SetTemp()
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
  KillFile$ (PassP$)
  'lentemp = Len(Tempfile)
  Tempfile = FreeFile
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp
  PassTemp.usernum = PWcnt
  PassTemp.username = PWUser
  PassTemp.frommdl = 0
  Put Tempfile, 1, PassTemp
  Close
End Sub
Public Sub GetTemp()
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
  
  'lentemp = Len(Tempfile)
  Tempfile = FreeFile
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp
  Get Tempfile, 1, PassTemp
  PWUser = QPTrim(PassTemp.username)
  PWcnt = PassTemp.usernum
  PWfromMdl = PassTemp.frommdl
  Close

End Sub
Public Sub SetToGo()
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
  
  Tempfile = FreeFile
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp
  PassTemp.usernum = PWcnt
  PassTemp.username = PWUser
  PassTemp.frommdl = 0
  Put Tempfile, 1, PassTemp
  Close
End Sub

Public Sub doexitstuff()
  'frmMainMenu.Enabled = False
  
  DoEvents
  Unload frmMainMenu
  Set frmMainMenu = Nothing
  'End ' frmMainMenu
End Sub

Public Sub SetToComeBack()
On Error GoTo Cancel
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
10:
  Tempfile = FreeFile
11:  Open PassP$ For Random Shared As Tempfile ' Len = lentemp
12:  PassTemp.usernum = PWcnt
13:  PassTemp.username = PWUser
14:  PassTemp.frommdl = 99
15:  Put Tempfile, 1, PassTemp
16:  Close
Cancel:
   If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (CitipakMain Settocomeback - Line:" & Erl & ")"
  End If
  Close
  Exit Sub
End Sub


Public Sub SetTempPWPath()
  Dim rOsVersionInfo As OSVERSIONINFO
  Dim sOperatingSystem As String
  sOperatingSystem = "NONE"
  rOsVersionInfo.dwOSVersionInfoSize = Len(rOsVersionInfo)
  If GetVersionEx(rOsVersionInfo) Then
    TypeSysOP = rOsVersionInfo.dwMajorVersion
    If TypeSysOP >= 6 Then
      PassP$ = "C:\CPWork\PassTemp.dat"
    Else
      PassP$ = "C:\PassTemp.dat"
    End If
  End If
End Sub

