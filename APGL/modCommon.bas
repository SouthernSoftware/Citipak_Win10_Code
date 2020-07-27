Attribute VB_Name = "modCommon"
Option Explicit

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
  'On Local Error GoTo FileError
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
    Kill FileName$
  End If
'  Exit Function
'FileError:
'  MsgBox "Error " & Error & " Has Occured", vbOKOnly, "Error"
End Function


