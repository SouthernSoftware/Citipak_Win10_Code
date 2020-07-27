Attribute VB_Name = "modWin2WinCommon"
Option Explicit
Public XFiles(1 To 12)

Public Function Exist(FileName$) As Boolean
  Dim FileHandle As Integer
  Dim TempSize As Long

  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  TempSize = LOF(FileHandle)
  Close FileHandle
  If TempSize <= 0 Then
    Kill FileName$
    Exist = False
  Else
    Exist = True
  End If

End Function

Public Sub KillFile(FileName As String)
  If Exist(FileName$) Then
    Kill FileName$
  End If
End Sub

Public Function OldRound#(n As Double)
  If n < -2000000000# Then n = 0
  OldRound# = Round(n, 2)
End Function

Public Function CheckForAllFiles() As Boolean
  Dim x As Integer

  x = 1
  CheckForAllFiles = True
  
  If Not Exist("PRDATA\PREMP2.DAT") Then
    XFiles(x) = "PRDATA\PREMP2.DAT"
    x = x + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRUNIT.DAT") Then
    XFiles(x) = "PRDATA\PRUNIT.DAT"
    x = x + 1
    CheckForAllFiles = False
  End If

End Function

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim ThisChar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    ThisChar = Asc(Mid$(Text, cnt, 1))
    If ThisChar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

Public Function FileSize(FileName$) As Long
  Dim FileHandle As Integer
  If Exist(FileName$) Then
    FileHandle = FreeFile
    Open FileName$ For Binary As FileHandle
    FileSize = LOF(FileHandle)
    Close FileHandle
  Else
    FileSize = 0
  End If
End Function

Public Sub UnloadAllFormsAndOpn()
  Unload frmConvertWin2Win
  Unload frmWin2WinInProg
  Unload frmWin2WinMissingFiles
End Sub

