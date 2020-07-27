VERSION 5.00
Begin VB.Form frmLoadingRpt 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Test"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   240
      Top             =   120
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
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   2610
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
      Left            =   360
      TabIndex        =   0
      Top             =   780
      Width           =   3210
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLoadingRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub Form_Load()
  Twiddle = "||//--\\"
  DoEvents
  Label2.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Unload Me
  End If
End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  Static whatchar%, thischar$
  
  Label2.Visible = True
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

