VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "modCommonWrtsvlBchCnvrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Const PRData = "prdata\"
Public Const LeaveFileName = "PRLEAVE.DAT"
Public Sub OpenOldLeaveFileName(LeaveHandle As Integer)
  Dim LeaveRec As OldLeaveRecType
  Dim LeaveRecLen As Integer
  LeaveRecLen = Len(LeaveRec)
  LeaveHandle = FreeFile
  Open PRData + LeaveFileName For Random Shared As LeaveHandle Len = LeaveRecLen
End Sub
Public Sub OpenLeaveFileName(LeaveHandle As Integer)
  Dim LeaveRec As LeaveRecType
  Dim LeaveRecLen As Integer
  LeaveRecLen = Len(LeaveRec)
  LeaveHandle = FreeFile
  Open PRData + LeaveFileName For Random Shared As LeaveHandle Len = LeaveRecLen
End Sub

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



