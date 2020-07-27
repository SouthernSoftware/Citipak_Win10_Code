Attribute VB_Name = "modFAInitProcedures"
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
Global RcptFileName As String '2/14/08
Global PassP As String '2/14/08
Global TypeSysOP As Integer '2/14/08
Type OSVERSIONINFO '2/14/08
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
  szCSDVersion As String * 128
End Type
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Boolean '2/14/08
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
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim CitiPass As CitiPassType
  Dim cnt&, dl&
  Dim FromPayCheck As Boolean
  
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
  Call SetTempPWPath
'  If Exist("C:\passtemp.dat") Then
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
        If CitiPass.Flag2 = -1 Then
          CitiPass.Flag2 = 0
          FromPayCheck = True
        End If
        If Not CitiPass.DelFlag Then
            If CitiPass.Module(5).FullAccess = True Then
               LevelPass = 1
            ElseIf CitiPass.Module(5).ReportsOnly = True Then
               LevelPass = 2
            End If
'            If Citipass.Module(4).PaymentAccess = True Then
'               'this means can access close option
'               CloseAccess = True
'            End If
        End If
      End If
      Close CitiPassFile
    End If
  End If
SoSoft:
'    If LevelPass > 0 Then
      Call MainLog("In Fixed Assets, with Level " & LevelPass)
      DoEvents
      FromFA = False
      Load frmFAMainMenu
      frmFAMainMenu.Show
      DoEvents
'    Else
'      Shell "Citipak.exe", vbMaximizedFocus
'      DoEvents
'    End If
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
      RcptFileName$ = "C:\CPWork\RcptPrn.dat"
    Else
      PassP$ = "C:\PassTemp.dat"
      RcptFileName$ = "C:\RcptPrn.dat"
    End If
  End If
End Sub
