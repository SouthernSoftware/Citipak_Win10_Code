Attribute VB_Name = "modProgStartUp"
Option Explicit

Sub Main()

Dim UserRec As Integer     'Holds the password record number
Dim FileHandle As Integer  'if program was started from 'CM'

If Exist("c:\from_cm.$$$") Then      'See if program was started via 'CM'
  FileHandle = FreeFile
  Open "c:\from_cm.$$$" For Random As FileHandle Len = 2
  Get FileHandle, 1, UserRec         'UserRec now has the password rec number
  Close FileHandle
  KillFile "c:\from_cm.$$$"
  'do other startup stuff here  (i.e. get password record)
  'now do payment entry
  'TaxPayment UBPayment etc.
Else 'program was not started from 'CM'
  'do main menu startup here
  'ProgMainMenu  etc.
End If

End Sub
