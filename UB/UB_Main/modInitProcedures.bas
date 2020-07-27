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

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hWnd As Long)
Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Const GW_HWNDPREV = 3
Global ComputerName As String
Global DelayExit As Boolean
'************Help Constants
  Global Const hlpUtilityBillingMain = 1
  Global Const hlpCustMaint = 2
  Global Const hlpAddACustomer = 3
  Global Const hlpEditACustomer = 4
  Global Const hlpEnterMeterCoord = 5
  Global Const hlpSetACustToFinal = 6
  Global Const hlpDeleteCustomer = 7
  Global Const hlpCustomerInfo = 8
  Global Const hlpCustomerConsumpt = 9
  Global Const hlpCustomerTrans = 10
  Global Const hlpQuickCustList = 11
  Global Const hlpQuickCustListRate = 12
  Global Const hlpPaymentDate = 13
  Global Const hlpPaymentsDeposits = 14
  Global Const hlpPaymentTransaction = 15
  Global Const hlpDepositTransaction = 16
  Global Const hlpDeleteAPayment = 17
  Global Const hlpDeleteADeposit = 18
  Global Const hlpUtilityBillings = 19
  Global Const hlpMeterReadings = 20
  Global Const hlpManualMeterReadEntry = 21
  Global Const hlpManualMeterReading = 22
  Global Const hlpHandHeldMeter = 23
  Global Const hlpImportExport = 24
  Global Const hlpPrintMeterReading = 25
  Global Const hlpEstimatedMeter = 26
  Global Const hlpPrintMeterReadingList = 27
  Global Const hlpPrintMeterReadingReport = 28
  Global Const hlpPrintHighLow = 29
  Global Const hlpInactive = 30
  Global Const hlpPrintMeterReadingNotes = 31
  Global Const hlpPrintMeterReadingSheets = 32
  Global Const hlpMetersNoSerial = 33
  Global Const hlpStoredAverage = 34
  Global Const hlpPrebilling = 35
  Global Const hlpUtilityBill = 36
  Global Const hlpPrintAllUtility = 37
  Global Const hlpReprintSelected = 38
  Global Const hlpPrintBStatus = 39
  Global Const hlpPostedBill = 40
  Global Const hlpBankDraft = 41
  Global Const hlpAccountsToDraft = 42
  Global Const hlpPrepareDraft = 43
  Global Const hlpPrintDraftCustomer = 44
  Global Const hlpPenalty = 45
  Global Const hlpCalculatePenalty = 46
  Global Const hlpEditPenalty = 47
  Global Const hlpCustomerDeposit = 48
  Global Const hlpApplyDepositTo = 49
  Global Const hlpRefundCustomer = 50
  Global Const hlpDepositCredit = 51
  Global Const hlpUtilityBilling = 52
  Global Const hlpLateNotice = 53
  Global Const hlpPrintLate = 54
  Global Const hlpPrintLateNotice = 55
  Global Const hlpFinalBill = 56
  Global Const hlpFinalMeterReading = 57
  Global Const hlpPreBillingReport = 58
  Global Const hlpFinalBillPrintingMenu = 59
  Global Const hlpPrintAllFinal = 60
  Global Const hlpReprintFinalBills = 61
  Global Const hlpCustomer = 62
  Global Const hlpTransactionJournal = 63
  Global Const hlpTransactionSum = 64
  Global Const hlpPaymentSummary = 65
  Global Const hlpCustomerCutOff = 66
  Global Const hlpCustomerStreet = 67
  Global Const hlpBillPaymentTax = 68
  Global Const hlpCycleCount = 69
  Global Const hlpPumpCodeReport = 70
  Global Const hlpMasterCustomer = 71
  Global Const hlpMasterBalance = 72
  Global Const hlpMasterDeposit = 73
  Global Const hlpMembershipFees = 74
  Global Const hlpFlatRate = 75
  Global Const hlpMailingLabels = 76
  Global Const hlpMeterInstalledReport = 77
  Global Const hlpStatistical = 78
  Global Const hlpConsumptionByRate = 79
  Global Const hlpConsumptionByRateCyc = 80
  Global Const hlpTopTen = 81
  Global Const hlpConsumptionByRange = 82
  Global Const hlpWorkOrderP = 83
  Global Const hlpEnterEditWork = 84
  Global Const hlpPrintWorkOrdersBy = 85
  Global Const hlpPrinWorkOrder = 86
  Global Const hlpPrintSelectedWork = 87
  Global Const hlpUtilitySystemSetup = 88
  Global Const hlpBilling = 89
  Global Const hlpRateTableMenu = 90
  Global Const hlpAddANewRate = 91
  Global Const hlpEditAnExistingRate = 92
  Global Const hlpDeleteAnExisting = 93
  Global Const hlpBankDraftSetup = 94
  Global Const hlpBillInformation = 95
  Global Const hlpGroupCode = 96
  Global Const hlpExport = 97
  Global Const hlpExportCustomer = 98
  Global Const hlpExportConsumption = 99
  Global Const hlpRecalc = 100
  Global Const hlpWorkOrder = 101
  Global Const hlpEditLateNotice = 102
  Global Const hlpFAQ = 103
'******************************************
  
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

