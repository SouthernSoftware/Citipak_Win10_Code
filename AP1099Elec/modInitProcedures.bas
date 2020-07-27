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
'@@@@@@@@@@@@@@@@@@@@@@@@
'Help file constants
  Global Const hlpFrontPage = 9999
  Global Const hlpCitipakMain = 1
  Global Const hlpGeneralLedger = 2
  Global Const hlpGJ = 3
  Global Const hlpGJEdit = 4
  Global Const hlpGJRegister = 5
  Global Const hlpGJPost = 6
  Global Const hlpCashReceipts = 7
  Global Const hlpEnterEditCash = 8
  Global Const hlpPostCashR = 9
  Global Const hlpCashDisbursements = 10
  Global Const hlpEntryEditCash = 11
  Global Const hlpPrintCash = 12
  Global Const hlpPostCash = 13
  Global Const hlpGLReports = 14
  Global Const hlpTrialBalance = 15
  Global Const hlpCashBalance = 16
  Global Const hlpAccountBalance = 17
  Global Const hlpAccountHistory = 18
  Global Const hlpBudgetHistory = 19
  Global Const hlpBalanceSheet = 20
  Global Const hlpBudgetVsActual = 21
  Global Const hlpDepartmentBudgetVs = 22
  Global Const hlpQueryGL = 23
  Global Const hlpExportFiles = 24
  Global Const hlpFunctionReport = 25
  Global Const hlpFunctionBudtVAct = 26
  Global Const hlpBudgetMaintenance = 27
  Global Const hlpEditBudgetMenu = 28
  Global Const hlpPrintBudget = 29
  Global Const hlpPostBudgetEntries = 30
  Global Const hlpBudgetPreparation = 31
  Global Const hlpBudPrepOptions = 32
  Global Const hlpBudPrepWSOptions = 33
  Global Const hlpBankReconciliation = 34
  Global Const hlpOutstandingCheck = 35
  Global Const hlpSelectChecksTo = 36
  Global Const hlpPrintCanceled = 37
  Global Const hlpRemoveCanceled = 38
  Global Const hlpAddOutstanding = 39
  Global Const hlpSortOutstanding = 40
  Global Const hlpGetDistributions = 41
  Global Const hlpGrabTransactions = 42
  Global Const hlpTransactionJournal = 43
  Global Const hlpPostInterface = 44
  Global Const hlpTransferToGeneral = 45
  Global Const hlpInitialize = 46
  Global Const hlpGLSetupAnd = 47
  Global Const hlpFundMaintenance = 48
  Global Const hlpAddChangeDelete = 49
  Global Const hlpPrintFund = 50
  Global Const hlpFundIndex = 51
  Global Const hlpChartOfAccountsMain = 52
  Global Const hlpAccountEntryEdit = 53
  Global Const hlpChartOfAccounts = 54
  Global Const hlpAccountIndex = 55
  Global Const hlpDepartment = 56
  Global Const hlpDepartmentEntry = 57
  Global Const hlpDeptList = 58
  Global Const hlpDeptIndex = 59
  Global Const hlpBankMaintenance = 60
  Global Const hlpAddABankMenu = 61
  Global Const hlpEditABankMenu = 62
  Global Const hlpPrintBank = 63
  Global Const hlpFunction = 64
  Global Const hlpFunctionEntryEdit = 65
  Global Const hlpPrintFunction = 66
  Global Const hlpFunctionIndex = 67
  Global Const hlpSetAllowable = 68
  Global Const hlpGLClosing = 69
  Global Const hlpSelectFundsToClose = 70
  Global Const hlpPreClosingMenu = 71
  Global Const hlpSystem = 72
  Global Const hlpAccountSetup = 73
  Global Const hlpSetFiscalPeriod = 74
  Global Const hlpInvoiceTaxSetup = 75
  Global Const hlpGLUtil = 76
  Global Const hlpGLPriorYear = 77
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  'Global Const hlpFrontPage = 9999
  Global Const hlpAccountsPayable = 1
  Global Const hlpAPVendMaint = 2
  Global Const hlpAddVend = 3
  Global Const hlpEditVend = 4
  Global Const hlpDefDist = 5
  Global Const hlpVendList = 6
  Global Const hlpVendFile = 7
  Global Const hlpVendLab = 8
  Global Const hlpPrnDefDist = 9
  Global Const hlpVendUtil = 10
  Global Const hlpPOProcess = 11
  Global Const hlpEnterPO = 12
  Global Const hlpAppPO = 13
  Global Const hlpNApprPO = 14
  Global Const hlpApprPO = 15
  Global Const hlpPOForms = 16
  Global Const hlpPostPO = 17
  Global Const hlpCancPO = 18
  Global Const hlpPOControl = 19
  Global Const hlpInvProcess = 20
  Global Const hlpEnterInv = 21
  Global Const hlpInvReg = 22
  Global Const hlpInvRegVO = 23
  Global Const hlpPostInv = 24
  Global Const hlpVoidInv = 25
  Global Const hlpChkProcess = 26
  Global Const hlpOpenPayRep = 27
  Global Const hlpInvforPay = 28
  Global Const hlpPreAudit = 29
  Global Const hlpPrintAPChks = 30
  Global Const hlpReprintAPChks = 31
  Global Const hlpCancelChks = 32
  Global Const hlpChkReg = 33
  Global Const hlpChkRegDet = 34
  Global Const hlpPostAPChk = 35
  Global Const hlpVoidPostedAPChk = 36
  Global Const hlpAPReports = 37
  Global Const hlpVendHist = 38
  Global Const hlpOpenPay = 39
  Global Const hlpOpenPayDate = 40
  Global Const hlpOpenPO = 41
  Global Const hlpPOHistGLAcct = 42
  Global Const hlpAPChklisting = 43
  Global Const hlpPSL = 44
  Global Const hlp1099Proc = 45
  Global Const hlpPayerInfo = 46
  Global Const hlpExtract1099 = 47
  Global Const hlpAdd1099 = 48
  Global Const hlpEdit1099 = 49
  Global Const hlpPrint1099Rep = 50
  Global Const hlpPrint1099Forms = 51
  Global Const hlpSalesTx = 52
  Global Const hlpCoSalesTx = 53



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
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
Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Const GW_HWNDPREV = 3
Global ComputerName As String
Global Twiddle As String
Global DelayExit As Boolean
Global tmpmdl As Integer
Global GLUBKill As Integer
'&*&(*&(*&(*&(*&*(
'01/10/2008  The following code used for Vista problem with writing to c:\
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
    Else
      PassP$ = "C:\PassTemp.dat"
    End If
  End If
End Sub
'*&*&(*&*&(*&(*&(&(*&
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
