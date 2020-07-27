VERSION 5.00
Begin VB.Form frmSplashCitipak 
   BackColor       =   &H008F8265&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8265
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplashCitipak.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   5352
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   8496
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CitiPak v2.06"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CitiPak v2.06"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   735
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   840
         Width           =   4695
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
         Left            =   3720
         TabIndex        =   4
         Top             =   3720
         Width           =   3732
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
         Height          =   732
         Index           =   0
         Left            =   3360
         TabIndex        =   3
         Top             =   3120
         Width           =   3252
      End
      Begin VB.Image Image1 
         Height          =   1332
         Left            =   1680
         Picture         =   "frmSplashCitipak.frx":08CA
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   1812
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6372
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   9612
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00A79F85&
      BorderStyle     =   0  'None
      Height          =   7092
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   10332
   End
End
Attribute VB_Name = "frmSplashCitipak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class

Private Sub Form_Load()
  Dim CompName As String
  Dim MaxLen As Long, RetVal As Long
'''Put this in so will look in winpak folder, not sure if will need later
  'ChDir "C:\Program Files\Microsoft Visual Studio\VB98\WinPak"
  'ChDir "C:\WinPak"
  MaxLen = 255
  CompName = Space$(256)
  RetVal = GetComputerName(CompName, MaxLen)
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
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
Private Sub Form_KeyPress(KeyAscii As Integer)
    Shell "Citipak.exe", vbMaximizedFocus
    Unload Me
End Sub

Private Sub Frame1_Click()
    Shell "Citipak.exe", vbMaximizedFocus
    Unload Me
End Sub
