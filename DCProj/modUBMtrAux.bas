Attribute VB_Name = "modUBMtrAux"
Option Explicit

Public Function Chk4HuskyError%()
  ReDim MsgText(0 To 5) As String
  Dim HuskErrHere As Boolean
  Dim EndFileName As String
  Dim EnvString As String
  Dim Indx As Integer
  EndFileName = ""
  HuskErrHere = True 'assume a problem
  
  MsgText(0) = "ERROR CONNECTION FAILURE!"
  MsgText(1) = "ERROR CODE:  "
  MsgText(2) = ""
  MsgText(3) = ""
  MsgText(4) = "Please call Southern Software support"
  MsgText(5) = "if you need assistance."
'  Stop
  DoEvents

  Indx = 1   ' Initialize index to 1.
  Do
    EnvString = UCase(Environ(Indx))   ' Get environment variable.
    If InStr(EnvString, "WINDIR") > 0 Then
      EndFileName = Mid$(EnvString, 8) + "\hcomw.end"
      Exit Do
    Else
      Indx = Indx + 1   ' Not PATH entry,
    End If   ' so increment.
  Loop Until EnvString = ""
  
  If Len(EndFileName) > 0 Then
    Dim HuskyErr As String
    Dim hFile As Integer
    If Exist(EndFileName) Then
      hFile = FreeFile
      Open EndFileName For Input As #hFile
      Line Input #hFile, HuskyErr
      Close hFile
      HuskyErr = Left$(HuskyErr, 3)
      If HuskyErr <> "000" Then
        MsgText(1) = MsgText(1) + HuskyErr
      End If
  '    Stop
      Select Case HuskyErr
      Case "000"
        HuskErrHere = False
      Case "001"
        MsgText(2) = "Invalid command-line parameter."
      Case "003"
        MsgText(2) = "Invalid port number."
      Case "006"
        MsgText(2) = "Invalid filename specified."
      Case "007"
        MsgText(2) = "Invalid computer specified in /DOWNLOAD"
      Case "008"
        MsgText(2) = "Binary file not found for /DOWNLOAD"
      Case "010"
        MsgText(2) = "Unable to open file."
      Case "011"
        MsgText(2) = "Unable to read file."
      Case "020"
        MsgText(2) = "Unable to connect to remote computer."
      Case "022"
        MsgText(0) = "WARNING. . ."
        MsgText(2) = "User Aborted Transfer Operation."
        MsgText(3) = ""
        
      Case "024"
        MsgText(2) = "Unable to write file!"
      Case "026"
        MsgText(2) = "Cannot log to new drive using /CD"
      Case "028"
        MsgText(2) = "Cannot find file!"
      Case "300"
        MsgText(2) = "Only one command-line process allowed."
      ' Command restricted on unregistered version."
      Case Else
        MsgText(2) = "Command restricted."
      End Select
    Else
      HuskErrHere = True
    End If
  Else
    HuskErrHere = True
  End If
  
  If HuskErrHere Then
    Chk4HuskyError = True
    MsgText(2) = UCase$(MsgText(2))
    GetOKorNot MsgText(), True, True
  Else
    Chk4HuskyError = False
  End If
End Function

Public Function MakeExpCoordinate(MtrPos As Double) As String
  Dim Coord As String * 11
  Dim TempCoord As String
  Dim PPos As Integer
  
  TempCoord = QPTrim(Str(MtrPos))
  If InStr(TempCoord, ".") <= 0 Then  'it mtr coord is a whole numb. unlikely but
    TempCoord = TempCoord + ".00000"
    GoTo GotCoord
  End If
  'not a whole number allready has a decimal just pad enough to cover 5 places
  TempCoord = TempCoord + "00000"
  'now find decimal point and trim up
  PPos = InStr(TempCoord, ".")
  
  TempCoord = Left$(TempCoord, PPos + 5)
  
GotCoord:
  RSet Coord = TempCoord
  MakeExpCoordinate = Coord
  
End Function
