Attribute VB_Name = "modGLStart"
Sub Main()
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim Citipass As CitiPassType
  Dim cnt&, dl&
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  screenW = (Screen.Width / Screen.TwipsPerPixelX)
  cnt& = 199
  
  '(*&(^%^&$%^&*&^%$#@#$%^&*^%%$#@##$%^&*^*&%^$%#$@#
  'BackColor = &HD0D0D0
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  App.HelpFile = "helpfiles\GL.hlp"
  If Exist("C:\passtemp.dat") Then
    GetTemp
    If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
      LevelPass = 1
      CloseAccess = True
    Else
    If Len(Dir$("Citipass.dat")) Then
      SetAttr ("CitiPass.dat"), vbNormal
      OpenCitiPassFile CitiPassFile, NumPassRecs
      If Not CitiPassFile = -1 Then
        Get CitiPassFile, PWcnt, Citipass
        If Not Citipass.DelFlag Then
            If Citipass.Module(3).FullAccess = True Then
               LevelPass = 1
            ElseIf Citipass.Module(3).ReportsOnly = True Then
               LevelPass = 2
            End If
            If Citipass.Module(3).PaymentAccess = True Then
               'this means can access close option
               CloseAccess = True
            End If
        End If
      End If
      Close CitiPassFile
    End If
  End If
  End If
'    If LevelPass > 0 Then
      Call MainLog("In GL, with Level " & LevelPass)
      DoEvents
      DelayExit = True
      Load frmGLMainMenu
      frmGLMainMenu.Show
      DoEvents
'    Else
'      Shell "Citipak.exe", vbMaximizedFocus
'      DoEvents
'    End If
End Sub

