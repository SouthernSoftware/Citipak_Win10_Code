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
  Global Const hlpTaxBillingMain = 1
  Global Const hlpTaxSystemSetupAnd = 2
  Global Const hlpTaxSystemDefault = 3
  Global Const hlpMortgageCodeM = 4
  Global Const hlpTaxGLAccountsSetupP = 5
  Global Const hlpTaxGLAccountsSetup = 6
  Global Const hlpOptionalRevenue = 7
  Global Const hlpAddNewRateCode = 8
  Global Const hlpEditExistingRate = 9
  Global Const hlpDeleteAnExisting = 10
  Global Const hlpCustomerM = 11
  Global Const hlpAddANew = 12
  Global Const hlpPersonal = 13
  Global Const hlpRealEstate = 14
  Global Const hlpCustomerLookup = 15
  Global Const hlpEditAnExisting = 16
  Global Const hlpExportCustomer = 17
  Global Const hlpAbstract = 18
  Global Const hlpPaymentOperator = 19
  Global Const hlpTaxPaymentMenu = 20
  Global Const hlpEnterEdit = 21
  Global Const hlpPrintTransaction = 22
  Global Const hlpDeleteTax = 23
  Global Const hlpRefundFor = 24
  Global Const hlpTaxBillingMenu = 25
  Global Const hlpTaxPrebilling = 26
  Global Const hlpTaxBillPrint = 27
  Global Const hlpPrintTaxBills = 28
  Global Const hlpReprintTax = 29
  Global Const hlpCreateMortgage = 30
  Global Const hlpTaxPenaltyBilling = 31
  Global Const hlpTaxPenalty = 32
  Global Const hlpEditPenalty = 33
  Global Const hlpTaxInterestBilling = 34
  Global Const hlpCalculateI = 35
  Global Const hlpEditInterest = 36
  Global Const hlpTaxLateNotice = 37
  Global Const hlpPrintLate = 38
  Global Const hlpReprintLate = 39
  Global Const hlpPrintLateNotice = 40
  Global Const hlpReprintPostedTax = 41
  Global Const hlpTaxReportsMenu = 42
  Global Const hlpCustomerI = 43
  Global Const hlpMasterCustomer = 44
  Global Const hlpMasterValuation = 45
  Global Const hlpTransaction = 46
  Global Const hlpLateListing = 47
  Global Const hlpExemption = 48
  Global Const hlpPrintAdvertising = 49
  Global Const hlpCustomerInfo = 50
  Global Const hlpPayment = 51
  Global Const hlpCustomer = 52
  Global Const hlpMasterBalance = 53
  Global Const hlpMailingLabels = 54
  Global Const hlpMortgageCode = 55
  Global Const hlpRealPropertyH = 56
  Global Const hlpCollectionRate = 57
  Global Const hlpRealProperty = 58
  Global Const hlpRealPropClass = 59
  Global Const hlpManualTaxBilling = 60
  Global Const hlpEnterTaxBill = 61
  Global Const hlpEditTransaction = 62
  Global Const hlpTaxBilling = 63
  Global Const hlpTaxAdvertising = 64
  Global Const hlpCalculate = 65
  Global Const hlpEditAdvertising = 66
  Global Const hlpAdvertisingMailing = 67
  Global Const hlpPPTRARemovalMenu = 68
  Global Const hlpPPTRARemoval = 69
  Global Const hlpDMVProcessing = 70
  Global Const hlpPrepareDMV = 71
  Global Const hlpReprocessDMV = 72
  Global PassP As String '2/14/08
  Global TypeSysOP As Integer '2/14/08
  Global RcptFileName As String '2/14/08
  
  Type OSVERSIONINFO '2/14/08
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
  End Type
  
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
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Boolean '2/14/08

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
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim CitiPass As CitiPassType
  Dim cnt&, dl&
  
  App.HelpFile = "helpfiles\VATAX.hlp"
  CurrCitiPath = App.Path
  
  If Exist("sosoftpw.dat") Then
    KillFile "sosoftpw.dat"
  End If
  
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
'  OperNum = 12 'for testing purposes only
'  If Exist("C:\passtemp.dat") Then
  Call SetTempPWPath '2/14/08
  If Exist(PassP$) Then '2/14/08
    GetTemp
    If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
      PWcnt = -3
      LevelPass = 1
      GoTo SoSoft
    End If
    If Len(Dir$("Citipass.dat")) Then
      SetAttr ("CitiPass.dat"), vbNormal
      OpenCitiPassFile CitiPassFile, NumPassRecs
      If Not CitiPassFile = -1 Then
        Get CitiPassFile, PWcnt, CitiPass
        OperNum = CitiPass.PassNum
        If CitiPass.Flag2 = -1 Then
          CitiPass.Flag2 = 0
          FromTX = True
        End If
        If Not CitiPass.DelFlag Then
          If CitiPass.Module(6).FullAccess = True Then
             LevelPass = 1
          ElseIf CitiPass.Module(6).PaymentAccess = True Then
             LevelPass = 3
          ElseIf CitiPass.Module(6).ReportsOnly = True Then
             LevelPass = 2
          End If
'            If CitiPass.Module(6).PaymentAccess = True Then
'               'this means can access close option
'               CloseAccess = True
'            End If
        End If
      End If
      Close CitiPassFile
    End If
  End If
SoSoft:
    Call MainLog("In Taxes, with Level " & LevelPass & " ")
    DoEvents
    Load frmVATaxMainMenu
    frmVATaxMainMenu.Show
    Call CheckDirs
    Call CheckInt
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
      RcptFileName$ = "C:\CPWork\RcptPrn.dat" 'don't touch 2/14/08
    Else
      PassP$ = "C:\PassTemp.dat"
      RcptFileName$ = "C:\RcptPrn.dat" 'don't touch 2/14/08
    End If
  End If
End Sub
