Attribute VB_Name = "modShellFile"

Option Explicit

Const cstrModuleName As String = "modShellAPIs"

' // Shell File Operations

Const FO_MOVE = &H1
Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_RENAME = &H4
Const FOF_MULTIDESTFILES = &H1
Const FOF_CONFIRMMOUSE = &H2
Const FOF_SILENT = &H4                      '  don't create progress/report
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
                                      '  Must be freed using SHFreeNameMappings
Const FOF_ALLOWUNDO = &H40
Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs

'       the global confirmation settings

Private Type SHFILEOPSTRUCT
   hWnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Function SH_MoveFile(ByVal strSource As String, ByVal strTarget As String, Optional ByVal blnProgress As Boolean = True, Optional ByVal blnConfirm As Boolean = False) As Boolean
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim lngRetVal As Long
   
   With SHFileOp
      .wFunc = FO_MOVE
      .pFrom = strSource
      .pTo = strTarget
      If (Not blnProgress And blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_SILENT
      ElseIf (Not blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
      Else
         .fFlags = FOF_ALLOWUNDO
      End If
   End With
   lngRetVal = SHFileOperation(SHFileOp)
   
   If (lngRetVal > 0) Then
      '---  File does not exist...
      SH_MoveFile = False
   ElseIf (SHFileOp.fAnyOperationsAborted) Then
      '---  Operation aborted, file not moved...
      SH_MoveFile = False
   Else
      SH_MoveFile = True
   End If
End Function

Public Function SH_CopyFile(ByVal strSource As String, ByVal strTarget As String, Optional ByVal blnProgress As Boolean = True, Optional ByVal blnConfirm As Boolean = False) As Boolean
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim lngRetVal As Long
   
   With SHFileOp
      .wFunc = FO_COPY
      .pFrom = strSource
      .pTo = strTarget
      If (Not blnProgress And blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_SILENT
      ElseIf (Not blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
      Else
         .fFlags = FOF_ALLOWUNDO
      End If
   End With
   lngRetVal = SHFileOperation(SHFileOp)
   
   If (lngRetVal > 0) Then
      '---  File does not exist...
      SH_CopyFile = False
   ElseIf (SHFileOp.fAnyOperationsAborted) Then
      '---  Operation aborted, file not copied...
      SH_CopyFile = False
   Else
      SH_CopyFile = True
   End If
End Function

Public Function SH_Recycle(ByVal strFileName As String, Optional ByVal blnProgress As Boolean = True, Optional ByVal blnConfirm As Boolean = False) As Boolean
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim lngRetVal As Long
   
   With SHFileOp
      .wFunc = FO_DELETE
      .pFrom = strFileName
      If (Not blnProgress And blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_SILENT
      ElseIf (Not blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
      Else
         .fFlags = FOF_ALLOWUNDO
      End If
   End With
   lngRetVal = SHFileOperation(SHFileOp)
   
   If (lngRetVal > 0) Then
      '---  File does not exist...
      SH_Recycle = False
   ElseIf (SHFileOp.fAnyOperationsAborted) Then
      '---  Operation aborted, file not sent to recycle bin...
      SH_Recycle = False
   Else
      SH_Recycle = True
   End If
End Function

Public Function SH_Rename(ByVal strFileName As String, ByVal strNewname As String, Optional ByVal blnProgress As Boolean = True, Optional ByVal blnConfirm As Boolean = False) As Boolean
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim lngRetVal As Long
   
   With SHFileOp
      .wFunc = FO_RENAME
      .pFrom = strFileName
      .pTo = strNewname
      If (Not blnProgress And blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_SILENT
      ElseIf (Not blnConfirm) Then
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
      Else
         .fFlags = FOF_ALLOWUNDO
      End If
   End With
   lngRetVal = SHFileOperation(SHFileOp)
   
   If (lngRetVal > 0) Then
      '---  File does not exist...
      SH_Rename = False
   ElseIf (SHFileOp.fAnyOperationsAborted) Then
      '---  Operation aborted, file not copied...
      SH_Rename = False
   Else
      SH_Rename = True
   End If
End Function

Public Sub Sendkeys(Text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), wait
   Set WshShell = Nothing
End Sub



