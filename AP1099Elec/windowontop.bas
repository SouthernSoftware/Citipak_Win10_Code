
Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long      

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2  
Const SWP_NOSIZE = &H1


Public Function PutWindowOnTop (pFrm as Form)
  Dim lngWindowPosition As Long
  
  LngWindowPosition = SetWindowPos(pFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub