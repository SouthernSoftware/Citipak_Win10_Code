Attribute VB_Name = "Module1"
Option Explicit

'!!! Procedures below Needed for reports!!! Mark with!!!
'Make sure to check w/Dale  PS
'!!! Added Round on 4-17-03
Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function
'loads Work Order Defaults into fpcombos


Public Function GetBillYear(intDate%)
  On Error GoTo BadNum2Date
  Dim WrkDate As String
  If intDate% = -32767 Then
    GetBillYear = -1
  Else
    WrkDate = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If
  GetBillYear = Val(Right(WrkDate, 4))
  Exit Function
BadNum2Date:
  On Error GoTo 0
  GetBillYear = -1
End Function


Public Function Num2Date$(intDate%)
  On Error GoTo BadNum2Date
  If intDate% = -32767 Then
    Num2Date$ = ""
  Else
    Num2Date$ = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If
  Exit Function
BadNum2Date:
  On Error GoTo 0
  Num2Date = ""
End Function

Public Function Date2Num%(txtDate$)
  On Error GoTo BadDate2Num
  If Len(QPTrim$(txtDate$)) = 10 Then
    Date2Num% = DateDiff("d", "12/31/1979", txtDate$)
  Else
    Date2Num% = -32767
  End If
  Exit Function

BadDate2Num:
  On Error GoTo 0
  Date2Num% = -32767
End Function


Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

Public Function ShowPctComp(ByVal cnt As Long, ByVal TotalCnt As Long) As String
  Dim PctComp As Long
  Dim Pct$
  If TotalCnt = 0 Then
    TotalCnt = 1
    cnt = 1
  End If
  PctComp = Int((cnt / TotalCnt) * 100)
  If PctComp > 100 Then PctComp = 100
  Pct$ = Str$(PctComp) + "%"
  ShowPctComp = Pct$
End Function


