VERSION 5.00
Begin VB.Form frmFALoadingRpt 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1884
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3936
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1884
   ScaleWidth      =   3936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ......"
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
      Left            =   834
      TabIndex        =   0
      Top             =   780
      Width           =   2604
   End
End
Attribute VB_Name = "frmFALoadingRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Unload frmFALoadingRpt
  End If
End Sub

