VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowProgBar 
   Caption         =   "Form1"
   ClientHeight    =   3468
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   6324
   LinkTopic       =   "Form1"
   ScaleHeight     =   3468
   ScaleWidth      =   6324
   StartUpPosition =   3  'Windows Default
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
      Left            =   2580
      TabIndex        =   0
      Top             =   2298
      Width           =   1164
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   396
      Left            =   1236
      TabIndex        =   1
      Top             =   1638
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
      Left            =   2646
      TabIndex        =   5
      Top             =   1302
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
      Left            =   3492
      TabIndex        =   4
      Top             =   1302
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
      Left            =   1284
      TabIndex        =   3
      Top             =   1302
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
      Height          =   540
      Left            =   972
      TabIndex        =   2
      Top             =   678
      Width           =   4380
   End
End
Attribute VB_Name = "frmShowProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
'Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
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
    FrmShowPctComp.Out = True
  Else
    MakeWindowTopMost Me.hWnd, True
  End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long, winhand As Long
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
'  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  MakeWindowTopMost Me.hWnd, True
  ProgressBar1.Value = 0

End Sub

Public Sub ShowPctComp(ByVal cnt As Long, ByVal TotalCnt As Long)
  Dim PctComp As Long
  PctComp = Int((cnt / TotalCnt) * 100)
  FrmShowPctComp.Label3 = PctComp
  ProgressBar1.Value = PctComp
  If PctComp = 100 Then
    MakeWindowTopMost Me.hWnd, False
    Unload FrmShowPctComp
    DoEvents
  Else
    DoEvents
  End If
End Sub


