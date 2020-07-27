VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDraftPrenote 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draft Prenote "
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   ControlBox      =   0   'False
   Icon            =   "frmDraftPrenote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Esc &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6426
      TabIndex        =   5
      Top             =   4848
      Width           =   1764
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "F10 &Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4026
      TabIndex        =   4
      Top             =   4848
      Width           =   1764
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "3:33 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/12/2014"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press ""ESC"" to Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   3552
      Width           =   5652
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press ""F10"" to Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      Index           =   3
      Left            =   3240
      TabIndex        =   2
      Top             =   4032
      Width           =   5652
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   3024
      X2              =   9216
      Y1              =   2112
      Y2              =   2112
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   3024
      X2              =   3024
      Y1              =   2136
      Y2              =   5832
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready to Create PRENOTE File?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   396
      Left            =   3372
      TabIndex        =   1
      Top             =   2832
      Width           =   5484
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000015&
      BorderWidth     =   3
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   3732
      Left            =   3012
      Top             =   2112
      Width           =   6204
   End
End
Attribute VB_Name = "frmDraftPrenote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Pre As Boolean



Private Sub cmdCancel_Click()
  Unload Me
End Sub
Public Sub PreSet(x As Boolean)
  Pre = x
End Sub
Private Sub cmdProcess_Click()
  If Pre = True Then
    frmUBDraftMenu.DoPreNote
  Else
    frmUBDraftMenu.DoDrftTest
  End If
  Unload Me
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdCancel_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

