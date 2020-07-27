VERSION 5.00
Begin VB.Form frmLoadingPRTransForm 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1956
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4008
   LinkTopic       =   "Form1"
   ScaleHeight     =   1956
   ScaleWidth      =   4008
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading  Payroll Transaction Screen......"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   24
      TabIndex        =   0
      Top             =   684
      Width           =   3960
   End
End
Attribute VB_Name = "frmLoadingPRTransForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    MainLog ("Payroll.exe terminated via menu bar on frmLoadingPRTransForm.")
    Call Terminate
    End
  End If
End Sub

