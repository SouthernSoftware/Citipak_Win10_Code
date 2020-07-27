VERSION 5.00
Begin VB.Form frmLoadingEmpEdit 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1956
   ClientLeft      =   -12
   ClientTop       =   -12
   ClientWidth     =   4008
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1956
   ScaleMode       =   0  'User
   ScaleWidth      =   4030.803
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Employee Edit Form......"
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
      Left            =   6
      TabIndex        =   0
      Top             =   780
      Width           =   3996
   End
End
Attribute VB_Name = "frmLoadingEmpEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Call Terminate
    MainLog ("Payroll.exe terminated via menu bar on frmLoadingEmpEdit.")
    End
  End If
End Sub

