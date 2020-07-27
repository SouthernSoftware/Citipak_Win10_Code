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

  Global Const hlpVehicleDecals = 1
  Global Const hlpDecalCategory = 2
  Global Const hlpAddADecalCode = 3
  Global Const hlpEditExistingDecal = 4
  Global Const hlpCustomerM = 5
  Global Const hlpAddNewCustomer = 6
  Global Const hlpEditExisting = 7
  Global Const hlpPrint = 8
  Global Const hlpVoidDecal = 9
  Global Const hlpReportsMenu = 10
  Global Const hlpCustomer = 11
  Global Const hlpCustomerBalance = 12
  Global Const hlpTransaction = 13
  Global Const hlpDetailedCustomer = 14
  Global Const hlpDecalListing = 15
  Global Const hlpExpiredDecal = 16
  Global Const hlpResidentReport = 17
  Global Const hlpOwnerReport = 18
  Global Const hlpNonOwnerReport = 19
  Global Const hlpPresentDecal = 20
  Global Const hlpPaymentDate = 21
  Global Const hlpPurchaseDecals = 22
  Global Const hlpEnterEditDecal = 23
  Global Const hlpDeleteDecal = 24
  Global Const hlpSetDefault = 25
  Global Const hlpDecalSetup = 26
  Global Const hlpSystemDefault = 27
  Global Const hlpEditLateNotice = 28


Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hwnd As Long)
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Const GW_HWNDPREV = 3
Global ComputerName As String
Global DelayExit As Boolean

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

