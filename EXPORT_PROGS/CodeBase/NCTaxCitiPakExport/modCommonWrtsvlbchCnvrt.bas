Attribute VB_Name = "modCommonWrtsvlbchCnvrt"
Option Explicit
  Public Const PRData = "prdata\"
  Public Const EarnNoMatchName = "PRDATA\PRERNOMC.DAT"
  Public Const ErnCodeFileName = "PRERNCOD.DAT"
  Public Const UnitFileName = "PRUNIT.DAT"
  Public Const TransWorkFileName = "PRTRANST.DAT"
  Public Const EmpData2Name = "PREMP2.DAT"
  Public Const TransHistFileName = "PRTRANSH.DAT"
  Public Const LeaveFileName = "PRLEAVE.DAT"
Public Sub OpenEmpData2File(EmpData2FileHandle As Integer)
  Dim EmpData2FileRec As EmpData2Type
  Dim EmpData2RecLen As Integer
  EmpData2RecLen = Len(EmpData2FileRec)
  EmpData2FileHandle = FreeFile
  Open PRData + EmpData2Name For Random Shared As EmpData2FileHandle Len = EmpData2RecLen
End Sub
Public Sub OpenOldTransWorkFile(OldTransWorkFileHandle As Integer)
  Dim OldTransWorkFileRec As OldTransRecType
  Dim OldTransWorkRecLen As Integer
  OldTransWorkRecLen = Len(OldTransWorkFileRec)
  OldTransWorkFileHandle = FreeFile
  Open PRData + TransWorkFileName For Random Shared As OldTransWorkFileHandle Len = OldTransWorkRecLen
End Sub
Public Sub OpenOldTransHistFile(OldTransHistFileHandle As Integer)
  Dim OldTransHistFileRec As OldTransRecType
  Dim OldTransHistRecLen As Integer
  OldTransHistRecLen = Len(OldTransHistFileRec)
  OldTransHistFileHandle = FreeFile
  Open PRData + TransHistFileName For Random Shared As OldTransHistFileHandle Len = OldTransHistRecLen
End Sub
Public Sub OpenTransWorkFile(TransWorkFileHandle As Integer)
  Dim TransWorkFileRec As TransRecType
  Dim TransWorkRecLen As Integer
  TransWorkRecLen = Len(TransWorkFileRec)
  TransWorkFileHandle = FreeFile
  Open PRData + TransWorkFileName For Random Shared As TransWorkFileHandle Len = TransWorkRecLen
End Sub
Public Sub OpenTransHistFile(TransHistFileHandle As Integer)
  Dim TransHistFileRec As TransRecType
  Dim TransHistRecLen As Integer
  TransHistRecLen = Len(TransHistFileRec)
  TransHistFileHandle = FreeFile
  Open PRData + TransHistFileName For Random Shared As TransHistFileHandle Len = TransHistRecLen
End Sub
Public Sub OpenDosErnCodeFile(DosErnCodeFileHandle As Integer)
  Dim DosErnCodeFileRec As DosErnCodeRecType
  Dim DosErnCodeRecLen As Integer
  DosErnCodeRecLen = Len(DosErnCodeFileRec)
  DosErnCodeFileHandle = FreeFile
  Open PRData + ErnCodeFileName For Random Shared As DosErnCodeFileHandle Len = DosErnCodeRecLen
End Sub
Public Sub OpenErnCodeFile(ErnCodeFileHandle As Integer)
  Dim ErnCodeFileRec As ErnCodeRecType
  Dim ErnCodeRecLen As Integer
  ErnCodeRecLen = Len(ErnCodeFileRec)
  ErnCodeFileHandle = FreeFile
  Open PRData + ErnCodeFileName For Random Shared As ErnCodeFileHandle Len = ErnCodeRecLen
End Sub
Public Sub OpenEarnNoMatchFile(EarnNoMatchHandle As Integer) '12/12/02
  Dim EarnNoMatchLen As Integer
  Dim EarnNoMatch As EarnNoMatchType
  EarnNoMatchLen = Len(EarnNoMatch)
  EarnNoMatchHandle = FreeFile
  Open EarnNoMatchName For Random Shared As EarnNoMatchHandle Len = EarnNoMatchLen
End Sub
Public Sub OpenUnitFile(FileHandle As Integer)
  Dim UnitFileRec As UnitFileRecType
  Dim UnitRecLen As Integer
  UnitRecLen = Len(UnitFileRec)
  FileHandle = FreeFile
  Open PRData + UnitFileName For Random Shared As FileHandle Len = UnitRecLen
End Sub
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



