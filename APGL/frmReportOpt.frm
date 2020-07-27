VERSION 5.00
Begin VB.Form frmReportOpt 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Option"
   ClientHeight    =   2556
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   6528
   Icon            =   "frmReportOpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2556
   ScaleWidth      =   6528
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5184
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2088
      Width           =   1092
   End
   Begin VB.CommandButton cmdGraphic 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Graphics"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   1656
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1344
      Width           =   1260
   End
   Begin VB.CommandButton cmdText 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3456
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1344
      Width           =   1284
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Format  - Graphics or Text  "
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
      Height          =   468
      Left            =   360
      TabIndex        =   3
      Top             =   552
      Width           =   5748
   End
End
Attribute VB_Name = "frmReportOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub cmdExit_Click()
  Unload frmReportOpt
End Sub

Private Sub cmdText_Click()
  Unload frmReportOpt
  rptopt = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    
    Case Else:
  End Select

End Sub

Private Sub cmdGraphic_Click()
  Unload frmReportOpt
  rptopt = 1
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  rptopt = 0
'  Set Temp_Class = New Resize_Class
'  Temp_Class.InitResizeClass Me
End Sub




