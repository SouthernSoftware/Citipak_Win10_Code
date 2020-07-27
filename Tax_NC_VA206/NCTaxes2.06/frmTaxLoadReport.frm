VERSION 5.00
Begin VB.Form frmTaxLoadReport 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   240
      Top             =   982
   End
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
      Height          =   390
      Left            =   990
      TabIndex        =   1
      Top             =   840
      Width           =   2610
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "**"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   906
      TabIndex        =   0
      Top             =   1438
      Width           =   2604
   End
End
Attribute VB_Name = "frmTaxLoadReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Twiddle = "||//--\\"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Unload Me
  End If
End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  Static whatchar%, thischar$
  tog = Not tog
  If tog Then
    If whatchar% > 8 Then
      whatchar% = 1
    End If
    whatchar% = whatchar% + 1
    thischar$ = Mid$(Twiddle$, whatchar%, 1)
  End If
  DoEvents
  Label2.Caption = thischar$

End Sub

