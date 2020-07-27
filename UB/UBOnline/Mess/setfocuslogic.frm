VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   2700
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open "c:\genesis\source\prod\windows9x\law enforcement suite\sinciden.frm" For Input As #1
Open "c:\genesis\source\prod\windows9x\law enforcement suite\binciden.frm" For Output As #2
While Not EOF(1)
    Line Input #1, a$
    b$ = LCase(a$)
    c = InStr(b$, ".setfocus")
    If c = 0 Or Left(a$, 1) = "'" Then
        Print #2, a$
    Else
        d$ = ""
        For t = c - 1 To 1 Step -1
            If Mid(a$, t, 1) <> " " Then
                d$ = Mid(a$, t, 1) + d$
            End If
        Next t
        Print #2, "'---- setfocus logic ----"
        Print #2, "'         " + a$
        Print #2, "          if " + d$ + ".visible then"
        Print #2, "              " + d$ + ".setfocus"
        Print #2, "           end if"
    End If
Wend
Close
End

        
End Sub
