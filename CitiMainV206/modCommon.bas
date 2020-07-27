Attribute VB_Name = "modCommon"
Option Explicit
Global rptopt As Integer

Public Static Sub DoTheTime()
  Dim sec As Long
  sec = Timer
  Do
  Loop Until (sec + 2) < Timer
End Sub
Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
  ' Load frmLoadingRpt
   frmViewPrint.ReportName = ReportFile$
   frmViewPrint.Caption = Title
   frmViewPrint.PgNum = PgNum
   If ForceSBar Then
     frmViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmViewPrint.cmdAlignment.Enabled = True
     frmViewPrint.AlignRpt = AlgnRptfile$
    Else
      frmViewPrint.cmdAlignment.Enabled = False
    End If
   frmViewPrint.NoPbox = False
 '  Unload frmLoadingRpt
   frmViewPrint.Show 1
End Sub

Public Function QPTrim$(Text As String)
  'Dim CPos As Long
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
Public Function Exist(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
  Dim testFile As String
  testFile = UCase$(FileName$)
    
  On Local Error GoTo FileError
  
  FileHandle = FreeFile
  
  Open testFile For Input Shared As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
  End If
  GoTo ExistExit
FileError:
  Close FileHandle
  Exist = False
  If UCase(Error) <> "FILE NOT FOUND" Then
    MsgBox "Error " & Error$ & " " & testFile, vbOKOnly, "Error"
  End If
ExistExit:
  On Error GoTo 0
End Function

Public Sub KillFile(FileName$)
  If Exist(FileName) Then
    Kill FileName$
  End If
End Sub

