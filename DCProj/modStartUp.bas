Attribute VB_Name = "modStartUp"
Option Explicit
'Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Public Const SND_SYNC = &H0
'Public Const SND_ASYNC = &H1
'Public Const SND_NODEFAULT = &H2
'Public Const SND_LOOP = &H8
'Public Const SND_NOSTOP = &H10
'

Sub Main()
  Dim RetValue As Integer
  Dim dcSetUpRec(1) As DCSetupType
  Dim RecLen As Integer
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim Citipass As CitiPassType
  Dim cnt&, dl&

  DebugMode = False
  
  Twiddle = "||//--\\"
  
  App.TaskVisible = False        'don't show in task list
  DCPath$ = QPTrim$(App.Path)    'start up path
  
  If Right$(DCPath$, 1) <> "\" Then
    DCPath$ = DCPath$ + "\"
  End If
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  SetTempPWPath
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  cnt& = 199
  OK4Secure = False
  '(*&(^%^&$%^&*&^%$#@#$%^&*^%%$#@##$%^&*^*&%^$%#$@#
  'BackColor = &HD0D0D0
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  App.HelpFile = "helpfiles\DECALS.hlp"

  CrLf = Chr$(13) + Chr$(10)
  FF = Chr$(12)
  Chr9 = Chr$(9)
  LoadDCSetUpFile dcSetUpRec(), RecLen
  TOWNNAME$ = QPTrim$(dcSetUpRec(1).DCTNNAME)
  If DebugMode = False Then
  If Exist(PassP$) Then
    GetTemp
    If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
      LevelPass = 1
      OK4Secure = True
      PWUser = "Sosoft Support"
      PWcnt = 0
      OperNum = 0
      DCLog "Support Sign in"
      Load frmDCMainMenu
      DoEvents
      frmDCMainMenu.Show
    ElseIf PWcnt > 0 Then
    If Len(Dir$("Citipass.dat")) Then
      SetAttr ("CitiPass.dat"), vbNormal
      OpenCitiPassFile CitiPassFile, NumPassRecs
      If Not CitiPassFile = -1 Then
        Get CitiPassFile, PWcnt, Citipass
        If Not Citipass.DelFlag Then
         If Citipass.Administ = True Then OK4Secure = True
          If Citipass.Module(10).FullAccess = True Then
            LevelPass = 1
          ElseIf Citipass.Module(10).PaymentAccess = True Then
            LevelPass = 2
          ElseIf Citipass.Module(10).ReportsOnly = True Then
            LevelPass = 3
          End If
          OperNum = Citipass.PassNum
          PWUser = QPTrim(Citipass.UserName)
        End If
      End If
      Close CitiPassFile
    End If
  End If
  End If
    If LevelPass > 0 Then
      Call DCLog("In dc, with Level " & LevelPass)
      DelayExit = True
      Load frmDCMainMenu
      DoEvents
      frmDCMainMenu.Show
    End If
  Else
    LevelPass = 1
    PWUser = "Sosoft Support"
    PWcnt = 0
    OperNum = 0
    DelayExit = True
    LevelAdj = True
    Load frmDCMainMenu
    frmDCMainMenu.Show
    DoEvents
  End If
'Only for testing
'    LevelPass = 1
'    PWUser = "Sosoft Support"
'    PWcnt = 0
'    OPERNUM = 0
'    DelayExit = True
'    Load frmUBMainMenu
'    frmUBMainMenu.Show
'    DoEvents
End Sub

Public Sub DoTheTime()
  Dim sec As Long
  sec = Timer
  Do
  Loop Until (sec + 1) < Timer
End Sub
