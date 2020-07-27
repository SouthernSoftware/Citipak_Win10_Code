Attribute VB_Name = "modUBMod"
Option Explicit
Dim V1Long As Long
Dim V2Long As Long
Dim TmpNum As Double

Public Static Function UBMod&(ByVal Var1 As Double, ByVal Var2 As Double)
  
  If Var2 = 0 Then    'Check for zero here
    UBMod& = 0        'Prevents division by zero error below
  Else
    V1Long = Var1     'Basic will round this while in assignment
    V2Long = Var2     'same
    
    TmpNum = V1Long / V2Long
    
    TmpNum = (Fix(TmpNum)) * V2Long   'fix removes the fractional part
    TmpNum = V1Long - TmpNum
    UBMod& = TmpNum
  End If
End Function
