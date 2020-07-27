VERSION 5.00
Begin VB.Form frmAPVendLablOPt 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Vendor Labels"
   ClientHeight    =   2556
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   6480
   Icon            =   "frmAPVendLablOPt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2556
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSingle 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Single Column"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1092
   End
   Begin VB.CommandButton cmd3Column 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Three Column"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   1776
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1092
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5304
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2064
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Do You Wish To Print 3 Column Sheet or Single Column Labels ? "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   828
      Left            =   912
      TabIndex        =   3
      Top             =   288
      Width           =   4788
   End
End
Attribute VB_Name = "frmAPVendLablOPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Temp_Class As Resize_Class
'Dim Over As clsTextBoxOverRider
'
'Private Sub cmdExit_Click()
'  Unload frmAPVendLablOPt
'End Sub
'
'Private Sub cmdsingle_Click()
'  Unload frmAPVendLablOPt
'  frmAPVendMaintMenu.PrintVendorLabels
'End Sub
'Private Sub cmd3Column_Click()
'  Unload frmAPVendLablOPt
'  frmAPVendMaintMenu.PrnVendLabelsLaser
'
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
'    Case vbKeyEscape:
'      cmdExit_Click
'      KeyCode = 0
'
'    Case Else:
'  End Select
'
'End Sub
'
'
'Private Sub Form_Load()
'  Set Over = New clsTextBoxOverRider
'  Over.OverRide Me
''  Set Temp_Class = New Resize_Class
''  Temp_Class.InitResizeClass Me
'End Sub
''Private Sub Form_Resize()
''  If Me.WindowState <> vbMinimized Then
''    Me.Visible = False
''    Temp_Class.ResizeControls Me
''    Me.Visible = True
''    Me.SetFocus
''  End If
''End Sub
''
''
'
'
