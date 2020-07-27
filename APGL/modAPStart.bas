Attribute VB_Name = "modAPStart"
Sub Main()
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim CitiPass As CitiPassType
  Dim cnt&, dl&
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  OKtoPO = False
  cnt& = 199
  SetTempPWPath
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  App.HelpFile = "helpfiles\AP.hlp"
  If Exist(PassP$) Then
    GetTemp
    If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
      tmpmdl = 2
      LevelPass = 1
      OKtoPO = True
    ElseIf PWcnt > 0 Then
    If Len(Dir$("Citipass.dat")) Then
      SetAttr ("CitiPass.dat"), vbNormal
      OpenCitiPassFile CitiPassFile, NumPassRecs
      If Not CitiPassFile = -1 Then
        Get CitiPassFile, PWcnt, CitiPass
        If Not CitiPass.DelFlag Then
            If CitiPass.Module(2).FullAccess = True Then
               LevelPass = 1
            ElseIf CitiPass.Module(2).ReportsOnly = True Then
               LevelPass = 2
            End If
            If CitiPass.Module(2).PaymentAccess = True Then
               OKtoPO = True
            End If
        End If
      End If
      Close CitiPassFile
    End If
    End If
  End If

'LevelPass = 1
    If LevelPass > 0 Or OKtoPO = True Then
      Call MainLog("In AP, with Level " & LevelPass)
      DoEvents
      DelayExit = True
      Load frmAPMainMenu
      frmAPMainMenu.Show
      DoEvents
'    Else
'      Shell "Citipak.exe", vbMaximizedFocus
'      DoEvents
    End If
End Sub


