VERSION 5.00
Begin VB.Form frmModuleInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Southern Software - CITIPAK"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   12195
   Icon            =   "frmModuleInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Ok"
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
      Left            =   8184
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5976
      Width           =   948
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact 1-800-842-8190 For Ordering Information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2976
      TabIndex        =   5
      Top             =   6084
      Width           =   4956
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   2982
      TabIndex        =   3
      Top             =   3600
      Width           =   6228
   End
   Begin VB.Image Image1 
      Height          =   1332
      Left            =   3216
      Picture         =   "frmModuleInfo.frx":08CA
      Stretch         =   -1  'True
      Top             =   4536
      Width           =   1812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Southern"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   28.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   5250
      TabIndex        =   2
      Top             =   4530
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Software, Inc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   28.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   5256
      TabIndex        =   1
      Top             =   5136
      Width           =   3732
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CITIPAK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Index           =   0
      Left            =   4656
      TabIndex        =   0
      Top             =   2688
      Width           =   2892
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   15
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   4356
      Left            =   2718
      Top             =   2304
      Width           =   6756
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5244
      Left            =   2154
      Top             =   1824
      Width           =   7884
   End
End
Attribute VB_Name = "frmModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub
Private Sub cmdOk_Click()
  Unload frmModuleInfo
End Sub


Private Sub Form_Load()
  
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select
End Sub

'Private Sub Timer1_Timer()
' ' Label2.Visible = Not Label2.Visible
'  ' &H00C0C0C0&
'  ' &H00E0E0E0&
'  Static tog As Boolean
'  tog = Not tog
'  If tog Then
'    Shape2.FillColor = &HE0E0E0
'    Shape1.BorderColor = &HC0C0C0
'  Else
'  End If
'    Shape2.FillColor = &HC0C0C0
'    Shape1.BorderColor = &HE0E0E0
'
'End Sub
