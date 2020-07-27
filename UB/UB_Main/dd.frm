VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim aa$
Dim alen As Integer
Dim CPos As Integer, RCNt As Integer

Open "Fixme.csv" For Input As #1
Open "Fixedit.csv" For Output As #2

Do
  Line Input #1, A$
  CPos = InStr(A$, Chr$(9))
  Do While CPos > 0
    A$ = Mid$(A$, 1, CPos - 1) + " " + Mid$(A$, CPos + 1)
    CPos = InStr(A$, Chr$(9))
    RCNt = RCNt + 1
  Loop
   
  Print #2, A$

Loop Until EOF(1)

Close
Print RCNt
'End
End Sub
