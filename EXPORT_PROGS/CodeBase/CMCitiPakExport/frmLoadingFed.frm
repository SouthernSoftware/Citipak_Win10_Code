VERSION 5.00
Begin VB.Form frmLoadingFed 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1884
   ClientLeft      =   36
   ClientTop       =   36
   ClientWidth     =   3936
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1884
   ScaleMode       =   0  'User
   ScaleWidth      =   4491.077
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Federal Tax Screen......"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   114
      TabIndex        =   0
      Top             =   744
      Width           =   3708
   End
End
Attribute VB_Name = "frmLoadingFed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Call Terminate
    MainLog ("Payroll.exe terminated via menu bar on frmLoadingFed.")
    End
  End If
End Sub

