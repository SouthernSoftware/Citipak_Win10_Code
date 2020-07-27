VERSION 5.00
Begin VB.Form frmBLSaving 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5268
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   5268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saving..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   756
      TabIndex        =   0
      Top             =   1050
      Width           =   3756
   End
End
Attribute VB_Name = "frmBLSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsBLTextBoxOverrider

Private Sub Form_Load()
'Dim RetVal As Long, winhand As Long
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  MakeWindowTopMost Me.hwnd, True
End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = &HC0FFFF
  Else
    Me.BackColor = &H80FFFF
  End If
End Sub


