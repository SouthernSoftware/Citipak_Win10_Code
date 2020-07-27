Attribute VB_Name = "modPassOut"
Option Explicit
Dim Citipass As CitiPassType
Global LevelPass As Integer
Global PWcnt As Integer
Global PWUser As String
Global CPAdminhand As Integer
Global LevelAdj As Boolean
'for Johnston temp Maybe for everyone later?
'Public Const PassP$ = "C:\CPWork\passtemp.dat"
''Public Const PassP$ = "C:\passtemp.dat"

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

Type UserPriviType
  FullAccess    As Boolean
  ReportsOnly   As Boolean
  PaymentAccess As Boolean
  Adjustments   As Boolean
  ZDummy2       As Boolean
End Type

Type CitiPassType
  PassNum   As Integer      'Unique Numeric Value
  UserName  As String * 15
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
'this type is temporary file kept on local drive for passing values between exe's
Type CitiPassTempType
  usernum   As Integer
  UserName  As String * 15
  frommdl   As Integer   'this is to indicate to citipak ok to have file
End Type
'&*&(*&(*&(*&(*&*(
'01/10/2008  The following code used for Vista problem with writing to c:\
Global RcptFileName As String
Global PassP As String
Global TypeSysOP As Integer
Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
  szCSDVersion As String * 128
End Type
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Boolean
Public Sub SetTempPWPath()
  Dim rOsVersionInfo As OSVERSIONINFO
  Dim sOperatingSystem As String
  sOperatingSystem = "NONE"
  rOsVersionInfo.dwOSVersionInfoSize = Len(rOsVersionInfo)
  If GetVersionEx(rOsVersionInfo) Then
    TypeSysOP = rOsVersionInfo.dwMajorVersion
    If TypeSysOP >= 6 Then
      PassP$ = "C:\CPWork\PassTemp.dat"
      RcptFileName$ = "C:\CPWork\RcptPrn.dat"
    Else
      PassP$ = "C:\PassTemp.dat"
      RcptFileName$ = "C:\RcptPrn.dat"
    End If
  End If
End Sub
'*&*&(*&*&(*&(*&(&(*&

Public Sub OpenCitiPassFile(CitiPassFile, NumPassRecs)
  Dim PassRecLen As Integer
  On Error GoTo PassError
  PassRecLen = Len(Citipass)
  CitiPassFile = FreeFile
  Open UBPath$ + "CitiPass.dat" For Random Shared As CitiPassFile Len = PassRecLen
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
  If Exist(UBPath$ + "CitiPass.dat") Then
    If PWcnt > 0 Then
    SetAttr (UBPath$ + "Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, Citipass
    Citipass.InUseFlag = False
    Citipass.FlagMod = 0
    Citipass.Flag2 = 0
    Citipass.CompName = ""
    Put CitiPassFile, PWcnt, Citipass
    Close CitiPassFile
    End If
  End If
  KillFile (PassP$)
End Sub

Public Sub ResetInUse()
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist(UBPath$ + "CitiPass.dat") Then
    SetAttr (UBPath$ + "CitiPass.dat"), vbNormal
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
Public Sub Ready4others(PWcnt)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist(UBPath$ + "CitiPass.dat") Then
    If PWcnt > 0 Then
    SetAttr (UBPath$ + "Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, Citipass
    'Citipass.InUseFlag = False
    Citipass.FlagMod = 0
    Citipass.Flag2 = 0
    'Citipass.CompName = ""
    Put CitiPassFile, PWcnt, Citipass
    Close CitiPassFile
    End If
  End If
  SetToGo
End Sub
Public Sub SetToGo()
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
  
  Tempfile = FreeFile
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp
  PassTemp.usernum = PWcnt
  PassTemp.UserName = PWUser
  PassTemp.frommdl = UB
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
  PWUser = QPTrim(PassTemp.UserName)
  PWcnt = PassTemp.usernum
  Close

End Sub


