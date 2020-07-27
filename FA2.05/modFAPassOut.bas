Attribute VB_Name = "modFAPassOut"
Option Explicit
Dim CitiPass As CitiPassType
Global LevelPass As Integer
Global PWcnt As Integer
Global PWUser As String
Global CPAdminhand As Integer
'LevelPass 1 is fullaccess, 2 is reports only
'Constants for Modules  By Menu Option Order
Public Const BL = 0 'Business License
Public Const AP = 1 'Accounts Payable
Public Const GL = 2 'General Ledger
Public Const PR = 3 'Payroll
Public Const FA = 4 'Fixed Assets
Public Const TX = 5 'Taxes
Public Const IC = 6 'Inventory Control
Public Const CM = 7 'Cash Management
Public Const UB = 8 'Utility Billing
Public Const DC = 9 'Vehicle Decals

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
End Sub


Public Sub ClearInUse(PWcnt)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  If Exist("CitiPass.dat") Then
    If PWcnt > 0 Then
    SetAttr ("Citipass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, CitiPass
    CitiPass.InUseFlag = False
    CitiPass.CompName = ""
    CitiPass.FlagMod = 0
    Put CitiPassFile, PWcnt, CitiPass
    Close CitiPassFile
    End If
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
      CitiPass.CompName = ""
      Put CitiPassFile, cnt, CitiPass
    Next
   Close CitiPassFile
  End If
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

