VERSION 5.00
Begin VB.Form frmBackGround 
   BackColor       =   &H008F8265&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2688
      Top             =   720
   End
End
Attribute VB_Name = "frmBackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    MainLog ("Payroll.exe terminated via menu bar on frmBackGround.")
    Call Terminate
    End
  End If
End Sub

