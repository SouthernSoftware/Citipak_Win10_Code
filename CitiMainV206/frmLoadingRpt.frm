VERSION 5.00
Begin VB.Form frmLoadingRpt 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1884
   ClientLeft      =   36
   ClientTop       =   36
   ClientWidth     =   3936
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1884
   ScaleWidth      =   3936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "**"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1152
      TabIndex        =   1
      Top             =   1104
      Width           =   1644
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
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
      Left            =   672
      TabIndex        =   0
      Top             =   660
      Width           =   2604
   End
End
Attribute VB_Name = "frmLoadingRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fromx As Form
Dim setfrom As Boolean
Private Sub Form_Load()
  Me.Show
  DoEvents
  MakeWindowTopMost Me.hwnd, True
  DoEvents
End Sub
Public Sub setwherefrom(x As Form)
'This is to deactivate controls for reports that take longer to process
'in the AR Report process - will reactivate in queryunload below.
  Set fromx = x
  setfrom = True
  DeActivateControls fromx, True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  MakeWindowTopMost Me.hwnd, False
  If setfrom = True Then
    ActivateControls fromx
  End If
  DoEvents
  setfrom = False
End Sub
Public Sub ShowHowMuch()
Static whatchar%, ThisChar$
whatchar% = whatchar% + 1
If whatchar% > 8 Then
  whatchar% = 1
End If
ThisChar$ = Mid$(Twiddle$, whatchar%, 1)
Label2.Caption = ThisChar$
'If Label2.Alignment = 0 Then
'  Label2.Alignment = 2
'ElseIf Label2.Alignment = 2 Then
'  Label2.Alignment = 1
'ElseIf Label2.Alignment = 1 Then
'  Label2.Alignment = 0
'End If
DoEvents
End Sub

'Private Sub Timer1_Timer()
'  Static tog As Boolean
' ' Stop
'  tog = Not tog
'  DoEvents
'  If tog Then
''If Line1.Visible = True Then
'    Line1.Visible = False
'    Line2.Visible = True
'  Else
'    Line1.Visible = True
'    Line2.Visible = False
'  End If
'DoEvents
'End Sub

