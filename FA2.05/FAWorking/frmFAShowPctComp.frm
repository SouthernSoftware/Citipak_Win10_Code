VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFAShowPctComp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2616
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5496
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2616
   ScaleWidth      =   5496
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   1920
      Width           =   1164
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   396
      Left            =   816
      TabIndex        =   1
      Top             =   1260
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   699
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " 00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2208
      TabIndex        =   5
      Top             =   924
      Width           =   732
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "% Complete."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3072
      TabIndex        =   4
      Top             =   924
      Width           =   1572
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   864
      TabIndex        =   3
      Top             =   924
      Width           =   1596
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Left            =   558
      TabIndex        =   2
      Top             =   204
      Width           =   4380
   End
End
Attribute VB_Name = "frmFAShowPctComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsFATextBoxOverRider
Dim vWidth%, vHeight%, vTop%, vLeft%
Public Out As Boolean
Private Sub Form_Initialize()
  vLeft = (Screen.Width * 0.5)  ' Set width of form.
  vTop = (Screen.Height * 0.5) ' Set height of form.
  vWidth = 525 '(Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vHeight = 280 '((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
Out = False
End Sub

Private Sub cmdCancel_Click()
  If MsgBox("Are You Sure You Want To Cancel?", vbYesNo + vbSystemModal, "Cancel Processing") = vbYes Then
    frmFAShowPctComp.Out = True
  Else
    MakeWindowTopMost Me.hwnd, True
  End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long, winhand As Long
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  MakeWindowTopMost Me.hwnd, True
  ProgressBar1.Value = 0
End Sub

Public Sub ShowPctComp(ByVal Cnt As Long, ByVal TotalCnt As Long)
  Dim PctComp As Long
  PctComp = Int((Cnt / TotalCnt) * 100)
  frmFAShowPctComp.Label3 = PctComp
  ProgressBar1.Value = PctComp
  If PctComp = 100 Then
    MakeWindowTopMost Me.hwnd, False
    Unload frmFAShowPctComp
    DoEvents
  Else
    DoEvents
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Unload frmFAShowPctComp
    DoEvents
  End If
End Sub


