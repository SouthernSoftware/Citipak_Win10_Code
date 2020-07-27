Attribute VB_Name = "modNetFunctions"
Option Explicit

Private Const mstrcModuleName As String = "modNetFunctions"

Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function Net_UserName(Optional ByVal strWhichConnection As String = vbNullString) As String
  Const strcProcName As String = "Net_UserName"
  On Error GoTo Proc_Error
  
  '--- Your code goes here...
  
  Dim strUser As String
  Dim lngSize As Long
  Dim lngRet As Long
  
  lngSize = 255
  strUser = String(lngSize, 255)
  lngRet = WNetGetUser(strWhichConnection, strUser, lngSize)
  If (lngRet = 0) Then
    lngSize = InStr(1, strUser, vbNullChar)
    If (lngSize > 0) Then
      strUser = Left(strUser, lngSize - 1)
    Else
      strUser = ""
    End If
  Else
    strUser = ""
  End If

  '--- Your "Normal" exit stuff goes here...
  
Proc_Exit:
  '--- Cleanup code goes here...
  Net_UserName = strUser
  GoTo Proc_End
  
Proc_Error:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, mstrcModuleName, strcProcName, Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select
  
Proc_End:
End Function

Public Function Net_ComputerName() As String
  Const strcProcName As String = "Net_ComputerName"
  On Error GoTo Proc_Error
  
  '--- Your code goes here...
  
  Dim strComputer As String
  Dim lngSize As Long
  Dim lngRet As Long
  
  lngSize = 255
  strComputer = String(lngSize, 255)
  lngRet = GetComputerName(strComputer, lngSize)
  If (lngRet > 0) Then
    lngSize = InStr(1, strComputer, vbNullChar)
    If (lngSize > 0) Then
      strComputer = Left(strComputer, lngSize - 1)
    Else
      strComputer = ""
    End If
  Else
    strComputer = ""
  End If

  '--- Your "Normal" exit stuff goes here...
  
Proc_Exit:
  '--- Cleanup code goes here...
  Net_ComputerName = strComputer
  GoTo Proc_End
  
Proc_Error:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, mstrcModuleName, strcProcName, Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select
  
Proc_End:
End Function

