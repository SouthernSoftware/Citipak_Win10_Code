Attribute VB_Name = "modPassOut"
Option Explicit
Dim CitiPass As CitiPassType
Global LevelPass As Integer
Global PWcnt As Integer
Global PWUser As String
Global CPAdminhand As Integer
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
    On Error GoTo 0
End Sub
Public Sub ClearInUsePRX()
  Dim x As Integer
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    SetAttr ("Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    For x = 1 To NumPassRecs
      Get CitiPassFile, x, CitiPass
      If CitiPass.FlagMod = 4 Then
        CitiPass.InUseFlag = False
        CitiPass.FlagMod = 0
        CitiPass.Flag2 = 0
        CitiPass.CompName = ""
        Put CitiPassFile, x, CitiPass
      End If
    Next x
    Close CitiPassFile
    KillFile ("C:\passtemp.dat")
  End If
End Sub
Public Sub ClearInUsePRReg(PWcnt As Integer)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If PWcnt = 0 Then Exit Sub
  If PWcnt = -3 Then Exit Sub 'sosoft password
  If Exist("CitiPass.dat") Then
    SetAttr ("Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
      Get CitiPassFile, PWcnt, CitiPass
      CitiPass.InUseFlag = False
      CitiPass.FlagMod = 0
      CitiPass.Flag2 = 0
      CitiPass.CompName = ""
      Put CitiPassFile, PWcnt, CitiPass
    Close CitiPassFile
    Call SetToGo
    KillFile ("C:\passtemp.dat")
  End If
End Sub

Public Sub ClearInUse(PWcnt As Integer)
  Dim x As Integer
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    SetAttr ("Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, CitiPass 'clears this users password
    'data if he terminates abnormally
    CitiPass.InUseFlag = False
    CitiPass.FlagMod = 0
    CitiPass.Flag2 = 0
    CitiPass.CompName = ""
    Put CitiPassFile, PWcnt, CitiPass
    Close CitiPassFile
    KillFile ("C:\passtemp.dat")
  End If
End Sub
Public Sub ResetInUse()
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    SetAttr ("CitiPass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    For cnt = 1 To NumPassRecs
      Get CitiPassFile, cnt, CitiPass
      CitiPass.InUseFlag = False
      CitiPass.Flag2 = 0
      CitiPass.FlagMod = 0
      CitiPass.CompName = ""
      Put CitiPassFile, cnt, CitiPass
    Next
   Close CitiPassFile
  End If
End Sub

Public Sub SetToGo()
End Sub

Public Sub Ready4others(PWcnt)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    If PWcnt > 0 Then 'sosoft = -3
      SetAttr ("Citipass.dat"), vbNormal
      OpenCitiPassFile CitiPassFile, NumPassRecs
      Get CitiPassFile, PWcnt, CitiPass
      'Citipass.InUseFlag = False
      CitiPass.FlagMod = 0
      CitiPass.Flag2 = 0
      'Citipass.CompName = ""
      Put CitiPassFile, PWcnt, CitiPass
      Close CitiPassFile
      SetToGo
    ElseIf PWcnt <= 0 Then
      SetToGo
    End If
  End If
End Sub

