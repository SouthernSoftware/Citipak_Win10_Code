Attribute VB_Name = "Module2"
Sub ConvertData()
'  Stop
  Dim tmpCustRec As NewUBCustRecType
  Dim UBHandle As Integer, CustRecLen As Integer
  Dim NumCust As Long, CCnt As Long, Cnt As Integer
  Dim TmpPNum As String, SPPos As Integer
  CustRecLen = Len(tmpCustRec)
  
  UBHandle = FreeFile
  Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
  NumCust = LOF(UBHandle) / CustRecLen
  For CCnt = 1 To NumCust
    Get #UBHandle, CCnt, tmpCustRec
    'If CCnt = 727 Then Stop
    TmpPNum = tmpCustRec.HPHONE      'do the home phone
    If Len(QPTrim$(TmpPNum)) > 6 Then
      If Mid$(TmpPNum, 1, 3) = "   " Then
        GoSub FmtPhone1
        LSet tmpCustRec.HPHONE = TmpPNum
      Else
        GoSub FmtPhone2
        LSet tmpCustRec.HPHONE = TmpPNum
      End If
    End If
'    Stop
    TmpPNum = tmpCustRec.WPHONE      'do the work phone
    If Len(QPTrim$(TmpPNum)) > 6 Then
      If Mid$(TmpPNum, 1, 3) = "   " Then
        GoSub FmtPhone1
        LSet tmpCustRec.WPHONE = TmpPNum
      Else
        GoSub FmtPhone2
        LSet tmpCustRec.WPHONE = TmpPNum
      End If
    End If
    TmpPNum = tmpCustRec.SOSEC      'do the work phone
    If Len(QPTrim$(TmpPNum)) >= 9 Then
      TmpPNum = Left$(tmpCustRec.SOSEC, 3) + "-" + Mid$(tmpCustRec.SOSEC, 4, 2) + "-" + Mid$(tmpCustRec.SOSEC, 6, 4)
      tmpCustRec.SOSEC = TmpPNum
    End If
    Put #UBHandle, CCnt, tmpCustRec
  Next
  
  
  Close UBHandle
    
  'Stop
  GoTo ExitConvert

FmtPhone1:
  TmpPNum = "(000) " + Mid$(TmpPNum, 4, 3) + "-" + Mid$(TmpPNum, 7, 4)
Return

FmtPhone2:
  TmpPNum = "(" + Left$(TmpPNum, 3) + ") " + Mid$(TmpPNum, 4, 3) + "-" + Mid$(TmpPNum, 7, 4)
Return

ExitConvert:
End Sub
