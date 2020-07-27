VERSION 5.00
Begin VB.Form frmChartAcctMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart of Accounts"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmChartAcctMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitAcctMaintMenu 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Account Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctSort 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Account Index &Utility"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctPrintList 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Print Chart of Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdAcctAddEdit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Add/Change/Delete Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Index           =   1
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Index           =   0
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CHART OF ACCOUNTS MAINTENANCE MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   6852
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
End
Attribute VB_Name = "frmChartAcctMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdAcctPrintList_Click()
  If Exist("GLacct.dat") Then
    frmPrintChartAcctSelectFund.Show
    Unload frmChartAcctMaintMenu
  Else
    MsgBox "You Must Setup Your Accounts Before Printing Them", vbOKOnly, "No Accounts"
  End If
End Sub
Private Sub cmdAcctSort_Click()
  If Exist("GLacct.dat") Then
    frmSortAcct.Show 1
  Else
    MsgBox "You Must Setup Your Accounts Before Sorting Them", vbOKOnly, "No Accounts"
  End If
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpChartOfAccountsMain
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdAcctAddEdit_Click()
  frmChartAcctEntryEdit.Show
  Unload frmChartAcctMaintMenu
End Sub
Private Sub cmdExitAcctMaintMenu_Click()
  frmGLSetupMenu.Show
  Unload frmChartAcctMaintMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitAcctMaintMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitAcctMaintMenu.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub
Public Sub OKFromSort()
  DeActivateControls frmChartAcctMaintMenu
  QSortAcctIndex frmChartAcctMaintMenu
  Call MainLog("Sorted Accounts Complete - Menu Option.")
  ActivateControls frmChartAcctMaintMenu
End Sub
