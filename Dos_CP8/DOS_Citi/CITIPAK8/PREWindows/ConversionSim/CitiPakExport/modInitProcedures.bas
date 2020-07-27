Attribute VB_Name = "modInitProcedures"
Option Explicit

'*******************
' ShowWindow() Commands
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_MAXIMIZE = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10
Const SW_MAX = 10

'added 11/08/04---------------
  Global Const hlpFrontPage = 9999
  Global Const hlpPRMain = 1
  Global Const hlpControlFile = 2
  Global Const hlpEmployerFile = 3
  Global Const hlpSystemFile = 4
  Global Const hlpStateTaxTable = 5
  Global Const hlpFederalTax = 6
  Global Const hlpEarnedIncome = 7
  Global Const hlpLeaveBenefitT = 8
  Global Const hlpDeductionCode = 9
  Global Const hlpAdditionalEarning = 10
  Global Const hlpRetirementFile = 11
  Global Const hlpACHBankDraftM = 12
  Global Const hlpACHBankDraft = 13
  Global Const hlpPrepareDraft = 14
  Global Const hlpPrintDraftEmployee = 15
  Global Const hlpPrinterSetupAnd = 16
  Global Const hlpInitializeNewYear = 17
  Global Const hlpEmployeeFile = 18
  Global Const hlpAddANewEmployee = 19
  Global Const hlpDirectDeposit = 20
  Global Const hlpJobDescription = 21
  Global Const hlpTaxWithholding = 22
  Global Const hlpMiscDeduction = 23
  Global Const hlpAlternate = 24
  Global Const hlpWageDistribution = 25
  Global Const hlpBenefit = 26
  Global Const hlpEditViewEmployee = 27
  Global Const hlpPrintEmployeeData = 28
  Global Const hlpPrintEmployeeList = 29
  Global Const hlpPrintTerminated = 30
  Global Const hlpPrintEmpEmergency = 31
  Global Const hlpQuickMaintenance = 32
  Global Const hlpPayroll = 33
  Global Const hlpAccrueLeave = 34
  Global Const hlpSetPayPeriod = 35
  Global Const hlpTransactionEntry = 36
  Global Const hlpEnterEditTime = 37
'  Global Const hlpCheckPrinting = 38
'  Global Const hlpPrintPayroll = 39
'  Global Const hlpReprintSelected = 40
  Global Const hlpVoidAPostedPayroll = 41
  Global Const hlpManualTransaction = 42
  Global Const hlpEnterManual = 43
  Global Const hlpManualRegister = 44
  Global Const hlpPostManual = 45
  Global Const hlpClearCurrent = 46
  Global Const hlpPayrollReports = 47
  Global Const hlpEarnings = 48
  Global Const hlpGrossWage = 49
  Global Const hlpPayrollDeductions = 50
  Global Const hlpESCReport = 51
  Global Const hlpReprintReports = 52
  Global Const hlpTaxFringe = 53
  Global Const hlpLeaveBenefit = 54
  Global Const hlpYearToDateWage = 55
  Global Const hlpChecksIssued = 56
  Global Const hlpChecksByNumber = 57
  Global Const hlpRetirement = 58
  Global Const hlpSupplemental = 59
  Global Const hlpAnnualWorkers = 60
  Global Const hlpSEPP = 61
  Global Const hlpEmployeePayRate = 62
 'added 11/08/04^^^^^^^^^^^^^^^^^^^^^^^^


Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hwnd As Long)
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const GW_HWNDPREV = 3
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Public Sub ActivatePrevInstance()
     Dim OldTitle As String
     Dim PrevHndl As Long
     Dim result As Long
     Dim Temp As Integer
       'Save the title of the application.
     OldTitle = App.Title
     'Rename the title of this application so FindWindow
     'will not find this application instance.
     App.Title = "unwanted instance"
'     'Check if found
     If PrevHndl = 0 Then
     'Attempt to get window handle using VB6 class name
       PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
     End If
     'Check if found
     If PrevHndl = 0 Then
        'No previous instance found.
        Exit Sub
     End If
     'Get handle to previous window.
      PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
     'Activate the application.
     'Restore the program.
      Call ShowWindow(PrevHndl, SW_MAXIMIZE)
      Call SetForegroundWindow(PrevHndl)
     End
End Sub

Sub Main()
'  Dim CitiPassFile As Integer, NumPassRecs As Integer
'  Dim CitiPass As CitiPassType
  Dim cnt&, dl&
  Dim FromPayCheck As Boolean
  
  App.HelpFile = "helpfiles\PAYROLL.hlp"
  FromPayCheck = False
  FromPR = False
  CurrCitiPath = App.Path
  
'  If Exist("sosoftpw.dat") Then
'    KillFile "sosoftpw.dat"
'  End If
  
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
'  If Exist("C:\passtemp.dat") Then
'    'GetTemp
'    If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
'      PWcnt = -3
'      LevelPass = 1
'      If Exist("paycheckmain.dat") Then
'        FromPayCheck = True
'        KillFile "paycheckmain.dat"
'      End If
'      GoTo SoSoft
'    End If
'    If Len(Dir$("Citipass.dat")) Then
'      SetAttr ("CitiPass.dat"), vbNormal
'      OpenCitiPassFile CitiPassFile, NumPassRecs
'      If Not CitiPassFile = -1 Then
'        Get CitiPassFile, PWcnt, CitiPass
'        If CitiPass.Flag2 = -1 Then
'          CitiPass.Flag2 = 0
'          FromPayCheck = True
'        End If
'        If Not CitiPass.DelFlag Then
'            If CitiPass.Module(4).FullAccess = True Then
'               LevelPass = 1
'            ElseIf CitiPass.Module(4).ReportsOnly = True Then
'               LevelPass = 2
'            End If
'            If CitiPass.Module(4).PaymentAccess = True Then
'               'this means can access close option
''               CloseAccess = True
'            End If
'        End If
'      End If
'      Close CitiPassFile
'    End If
'  End If
SoSoft:
'    If LevelPass > 0 Then
'      Call MainLog("In PR, with Level " & LevelPass)
'      DoEvents
'      If FromPayCheck = True Then
'        FromPR = True
'        frmPayrollProcessingMenu.Show
'      Else
        Load frmCitiPakExportData
        frmCitiPakExportData.Show
'      End If
      DoEvents
'    Else
'      Shell "Citipak.exe", vbMaximizedFocus
'      DoEvents
'    End If
End Sub
