Attribute VB_Name = "ubStartUp"
Option Explicit
'Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Public Const SND_SYNC = &H0
'Public Const SND_ASYNC = &H1
'Public Const SND_NODEFAULT = &H2
'Public Const SND_LOOP = &H8
'Public Const SND_NOSTOP = &H10
Sub Main()
  On Error GoTo Cancel
' Dim RetValue As Integer
  Dim CMSetUpRec(1) As CMSetupType
  Dim RecLen As Integer
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim CitiPass As CitiPassType
  Dim cnt&, dl&
10:
  DebugMode = False
  Twiddle = "||//--\\"
  
  App.TaskVisible = False        'don't show in task list
  UBPath$ = QPTrim$(App.Path)    'start up path
15:
  If Right$(UBPath$, 1) <> "\" Then
    UBPath$ = UBPath$ + "\"
  End If
18:
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
21:   SetTempPWPath
22:
  TempIndexName = UBPath$ + "UBTEMP.IDX"
  BookIndexFile = UBPath$ + "UBCUSTBK.IDX"
  NameIndexFile = UBPath$ + "UBCUSTNM.IDX"
  UBCustFile = UBPath$ + "UBCUST.DAT"
  UBOwnerFile = UBPath$ + "UBOWNER.DAT"
27:
  CrLf = Chr$(13) + Chr$(10)
  FF = Chr$(12)
  Chr9 = Chr$(9)
  
  Call CheckHasTaxes(intHasTaxes)

' Call ConvertData
' Stop
30:
     LoadCMSetUpFile CMSetUpRec(), RecLen
31:  TownName$ = QPTrim$(CMSetUpRec(1).CMTOWNNAME)
32:  If DebugMode = False Then
33:    If Exist(PassP$) Then
34:      GetTemp
35:      If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
36:        LevelPass = 1
37:        PWUser = "Sosoft Support"
38:        PWcnt = 0
39:        OperNum = 0
40:        CMLog "Support Sign in"
41:        Load frmCMMainMenu
42:        DoEvents
43:        frmCMMainMenu.Show
44:      Else
45:        If Len(Dir$("Citipass.dat")) Then
46:          SetAttr ("CitiPass.dat"), vbNormal
47:          OpenCitiPassFile CitiPassFile, NumPassRecs
48:          If Not CitiPassFile = -1 Then
'MsgBox ("PWcnt = " + Str(PWcnt))

49:            Get CitiPassFile, PWcnt, CitiPass
50:            If Not CitiPass.DelFlag Then
51:              If CitiPass.Module(8).FullAccess = True Then
52:                LevelPass = 1
53:              ElseIf CitiPass.Module(8).ReportsOnly = True Then
54:                  LevelPass = 3
55:              ElseIf CitiPass.Module(8).PaymentAccess = True Then
56:                  LevelPass = 2
57:              End If
58:              OperNum = CitiPass.PassNum
59:              PWUser = QPTrim(CitiPass.UserName)
60:            End If
61:          End If
62:          Close CitiPassFile
63:        End If
64:      End If
65:    End If
66:      If LevelPass > 0 Then
67:        Call CMLog("In CM, with Level " & LevelPass)
68:        DelayExit = True
69:        Load frmCMMainMenu
70:        DoEvents
71:        frmCMMainMenu.Show
72:      End If
73:
      Else
74:    LevelPass = 1
75:    PWUser = "Sosoft Support"
76:   PWcnt = 0
77:    OperNum = 0
78:    CMLog "Support Sign in"
79:    Load frmCMMainMenu
80:    frmCMMainMenu.Show
81:    DoEvents
  End If
  
Cancel:
   If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (CMSubMain - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub

Public Sub DoTheTime()
  Dim sec As Long
  sec = Timer
  Do
  Loop Until (sec + 1) < Timer
End Sub


