Attribute VB_Name = "ModuleISEVEN"
Public Function ISEven(amt$) As Boolean 'DblNum#) As Boolean
Dim Temp As String
Temp = amt$ 'Str$(DblNum#)
Select Case Right$(Temp$, 1)
Case "0", "2", "4", "6", "8"
  ISEven = True
Case Else
  ISEven = False
End Select
End Function
