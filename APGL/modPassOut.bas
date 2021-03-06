Attribute VB_Name = "modPassOut"
Option Explicit
Dim CitiPass As CitiPassType
Global LevelPass As Integer
Global PWcnt As Integer
Global PWUser As String
Global CPAdminhand As Integer
Global CloseAccess As Boolean
'Dim tmpmdl As Integer
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
  ZDummy1       As Boolean
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

Public Sub OpenCitiPassFile(CitiPassFile, NumPassRecs)
  Dim PassRecLen As Integer
  On Error GoTo PassError
  PassRecLen = Len(CitiPass)
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
'use this one to close gl and open citipak main
Public Sub Ready4others(PWcnt)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    If PWcnt > 0 Then
    SetAttr ("Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, CitiPass
    'Citipass.InUseFlag = False
    tmpmdl = CitiPass.FlagMod
    CitiPass.FlagMod = 0
    CitiPass.Flag2 = 0
    'Citipass.CompName = ""
    Put CitiPassFile, PWcnt, CitiPass
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
  PassTemp.frommdl = tmpmdl
  Put Tempfile, 1, PassTemp
  Close
End Sub

Public Sub ClearInUse(PWcnt)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    If PWcnt > 0 Then
      SetAttr ("Citipass.dat"), vbNormal
      OpenCitiPassFile CitiPassFile, NumPassRecs
      Get CitiPassFile, PWcnt, CitiPass
      CitiPass.InUseFlag = False
      CitiPass.FlagMod = 0
      CitiPass.Flag2 = 0
      CitiPass.CompName = ""
      Put CitiPassFile, PWcnt, CitiPass
      Close CitiPassFile
    End If
  End If
  'Kill (PassP$)
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




